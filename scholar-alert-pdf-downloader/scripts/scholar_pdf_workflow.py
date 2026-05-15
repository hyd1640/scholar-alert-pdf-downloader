#!/usr/bin/env python3
"""OOXML and browser helpers for Scholar Alert PDF workflows.

Standard-library only: no pandas/openpyxl/playwright dependency.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from collections import Counter, defaultdict
from html import escape as html_escape
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlparse
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape
from zipfile import ZIP_DEFLATED, ZipFile

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_ID = f"{{{NS_REL}}}id"

TITLE = "\u6807\u9898"
AUTHORS = "\u4f5c\u8005"
JOURNAL = "\u671f\u520a"
LINK = "\u94fe\u63a5"
STATUS = "PDF\u4e0b\u8f7d\u72b6\u6001"
HEADERS = [TITLE, AUTHORS, JOURNAL, LINK, STATUS]
SUCCESS = "\u6210\u529f"
FAIL = "\u5931\u8d25"
UNMARKED = "\u672a\u6807\u8bb0"
ALL_DONE = "\u5168\u90e8\u5b8c\u6210"
NO_LOGIN_BROWSER_PUBLISHERS = {
    "MDPI",
    "arXiv",
    "Copernicus / EGUsphere",
    "Preprints",
    "Science",
    "Taylor & Francis",
}
SCIENCEDIRECT_PUBLISHER = "ScienceDirect / Elsevier"
SCIENCEDIRECT_MANUAL_REASON = "requires manual PDF click or human verification"
RESEARCHGATE_STATUS = f"{FAIL}: 查询正常来源 (ResearchGate is not a journal source)"
DEFAULT_PROFILE_DIR = Path.home() / ".codex" / "scholar_pdf_chrome_profile"
HTTP_DOWNLOAD_RETRIES = 5
HTTP_CHUNK_SIZE = 1024 * 256

MDPI_ISSN_SLUGS = {
    "2072-4292": "remotesensing",
    "2624-795X": "geohazards",
}

MDPI_JOURNAL_SLUGS = {
    "acoustics": "acoustics",
    "actuators": "actuators",
    "administrative sciences": "admsci",
    "agriculture": "agriculture",
    "algorithms": "algorithms",
    "animals": "animals",
    "antibiotics": "antibiotics",
    "antioxidants": "antioxidants",
    "applied sciences": "applsci",
    "atmosphere": "atmosphere",
    "batteries": "batteries",
    "biomedicines": "biomedicines",
    "biosensors": "biosensors",
    "buildings": "buildings",
    "cancers": "cancers",
    "catalysts": "catalysts",
    "cells": "cells",
    "chemosensors": "chemosensors",
    "climate": "climate",
    "crystals": "crystals",
    "diagnostics": "diagnostics",
    "diversity": "diversity",
    "drugs and drug candidates": "ddc",
    "electronics": "electronics",
    "energies": "energies",
    "entropy": "entropy",
    "fermentation": "fermentation",
    "foods": "foods",
    "forests": "forests",
    "fractals and fractional": "fractals",
    "future internet": "fi",
    "games": "games",
    "gels": "gels",
    "genes": "genes",
    "geohazards": "geohazards",
    "geosciences": "geosciences",
    "healthcare": "healthcare",
    "horticulturae": "horticulturae",
    "humanities": "humanities",
    "insects": "insects",
    "international journal of molecular sciences": "ijms",
    "journal of clinical medicine": "jcm",
    "journal of imaging": "jimaging",
    "journal of marine science and engineering": "jmse",
    "land": "land",
    "life": "life",
    "machines": "machines",
    "marine drugs": "marinedrugs",
    "materials": "materials",
    "mathematics": "mathematics",
    "metabolites": "metabolites",
    "metals": "metals",
    "microorganisms": "microorganisms",
    "minerals": "minerals",
    "molecules": "molecules",
    "nanomaterials": "nanomaterials",
    "nutrients": "nutrients",
    "pharmaceutics": "pharmaceutics",
    "plants": "plants",
    "polymers": "polymers",
    "processes": "processes",
    "remote sensing": "remotesensing",
    "resources": "resources",
    "sensors": "sensors",
    "societies": "societies",
    "sustainability": "sustainability",
    "symmetry": "symmetry",
    "toxins": "toxins",
    "vaccines": "vaccines",
    "water": "water",
}


def col_name(n: int) -> str:
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def cell_text(cell: ET.Element) -> str:
    text = cell.find(f".//{{{NS_MAIN}}}t")
    return text.text if text is not None and text.text is not None else ""


def publisher(url: str) -> str:
    host = urlparse(url).netloc.lower().replace("www.", "")
    if "arxiv.org" in host:
        return "arXiv"
    if "sciencedirect.com" in host:
        return SCIENCEDIRECT_PUBLISHER
    if "ieeexplore.ieee.org" in host:
        return "IEEE Xplore"
    if "onlinelibrary.wiley.com" in host:
        return "Wiley Online Library"
    if "tandfonline.com" in host:
        return "Taylor & Francis"
    if "science.org" in host:
        return "Science"
    if "mdpi.com" in host:
        return "MDPI"
    if "lyellcollection.org" in host:
        return "Lyell Collection"
    if "researchgate.net" in host:
        return "ResearchGate"
    if "proquest.com" in host:
        return "ProQuest"
    if "springer.com" in host:
        return "Springer"
    if "copernicus.org" in host or "egusphere" in host:
        return "Copernicus / EGUsphere"
    if "preprints.org" in host:
        return "Preprints"
    return host or "unknown"


def load_records(xlsx: Path) -> tuple[dict[str, bytes], list[dict[str, str]]]:
    with ZipFile(xlsx) as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    sheet = ET.fromstring(files["xl/worksheets/sheet1.xml"])
    rels = ET.fromstring(files["xl/worksheets/_rels/sheet1.xml.rels"])
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in list(rels)}

    link_map: dict[str, str] = {}
    for hyperlink in sheet.findall(f".//{{{NS_MAIN}}}hyperlink"):
        rid = hyperlink.attrib.get(REL_ID)
        link_map[hyperlink.attrib.get("ref", "")] = rel_map.get(rid, "")

    records: list[dict[str, str]] = []
    rows = sheet.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row")
    for row_el in rows[1:]:
        values: dict[str, str] = {}
        for cell in list(row_el):
            ref = cell.attrib.get("r", "")
            match = re.match(r"[A-Z]+", ref)
            if match:
                values[match.group(0)] = cell_text(cell)
        row_num = int(row_el.attrib["r"])
        target_url = link_map.get(f"D{row_num}") or values.get("D", "")
        records.append(
            {
                "excel_row": str(row_num),
                TITLE: values.get("A", ""),
                AUTHORS: values.get("B", ""),
                JOURNAL: values.get("C", ""),
                LINK: values.get("D", "") or target_url,
                STATUS: values.get("E", ""),
                "_target_url": target_url,
                "_publisher": publisher(target_url),
            }
        )
    return files, records


def make_cell(ref: str, text: str, style: int | None = None) -> str:
    attrs = f' r="{ref}" t="inlineStr"'
    if style is not None:
        attrs += f' s="{style}"'
    return f"<c{attrs}><is><t>{xml_escape(str(text))}</t></is></c>"


def write_records(xlsx: Path, files: dict[str, bytes], records: list[dict[str, str]]) -> None:
    sheet_rows = [
        '<row r="1">'
        + "".join(make_cell(f"{col_name(i)}1", header, 1) for i, header in enumerate(HEADERS, 1))
        + "</row>"
    ]
    for row_index, record in enumerate(records, start=2):
        values = [record.get(header, "") for header in HEADERS]
        sheet_rows.append(
            f'<row r="{row_index}">'
            + "".join(make_cell(f"{col_name(col)}{row_index}", value) for col, value in enumerate(values, 1))
            + "</row>"
        )

    hyperlinks = []
    rels = []
    for row_index, record in enumerate(records, start=2):
        rid = f"rId{row_index - 1}"
        hyperlinks.append(f'<hyperlink ref="D{row_index}" r:id="{rid}"/>')
        rels.append(
            f'<Relationship Id="{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            f'Target="{xml_escape(record.get("_target_url", ""))}" TargetMode="External"/>'
        )

    last_row = len(records) + 1
    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols><col min="1" max="1" width="72" customWidth="1"/><col min="2" max="2" width="42" customWidth="1"/><col min="3" max="3" width="36" customWidth="1"/><col min="4" max="4" width="80" customWidth="1"/><col min="5" max="5" width="60" customWidth="1"/></cols>
  <sheetData>{''.join(sheet_rows)}</sheetData>
  <autoFilter ref="A1:E{last_row}"/>
  <hyperlinks>{''.join(hyperlinks)}</hyperlinks>
</worksheet>'''.encode("utf-8")
    rels_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{NS_PKG_REL}">{''.join(rels)}</Relationships>'''.encode(
        "utf-8"
    )

    backup = xlsx.with_suffix(".before_pdf_status.xlsx")
    if not backup.exists():
        shutil.copy2(xlsx, backup)

    tmp = xlsx.with_suffix(".tmp.xlsx")
    with ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        for name, data in files.items():
            if name == "xl/worksheets/sheet1.xml":
                data = sheet_xml
            elif name == "xl/worksheets/_rels/sheet1.xml.rels":
                data = rels_xml
            zout.writestr(name, data)
    tmp.replace(xlsx)


def summarize(xlsx: Path) -> None:
    _, records = load_records(xlsx)
    statuses = Counter(
        SUCCESS if r[STATUS].startswith(SUCCESS) else FAIL if r[STATUS].startswith(FAIL) else UNMARKED
        for r in records
    )
    remaining = Counter(r["_publisher"] for r in records if r[STATUS].startswith(FAIL))
    print(
        json.dumps(
            {"records": len(records), "statuses": statuses, "remaining_by_publisher": remaining},
            ensure_ascii=False,
            indent=2,
        )
    )


def build_queue(xlsx: Path, out: Path, json_out: Path | None = None) -> None:
    _, records = load_records(xlsx)
    failures = [r for r in records if r[STATUS].startswith(FAIL)]
    order = [
        "MDPI",
        "arXiv",
        "Copernicus / EGUsphere",
        "Preprints",
        "Taylor & Francis",
        "Science",
        SCIENCEDIRECT_PUBLISHER,
        "Wiley Online Library",
        "IEEE Xplore",
        "Lyell Collection",
        "ProQuest",
        "Springer",
        "ResearchGate",
    ]
    failures.sort(key=lambda r: (order.index(r["_publisher"]) if r["_publisher"] in order else 99, int(r["excel_row"])))
    if json_out:
        json_out.write_text(json.dumps(failures, ensure_ascii=False, indent=2), encoding="utf-8")

    groups: dict[str, list[dict[str, str]]] = defaultdict(list)
    for record in failures:
        groups[record["_publisher"]].append(record)

    summary = "".join(f"<span class='pill'>{html_escape(pub)}: {len(items)}</span>" for pub, items in groups.items())
    if not summary:
        summary = f"<span class='pill'>{ALL_DONE}</span>"

    sections = []
    for pub in [p for p in order if p in groups] + [p for p in groups if p not in order]:
        rows = []
        if pub == "ResearchGate":
            hint = "ResearchGate is not treated as the primary journal source. Use these rows to search for a normal publisher or journal source."
        elif pub in NO_LOGIN_BROWSER_PUBLISHERS:
            hint = "No login normally needed. Run the no-login browser round before manual publisher login."
        elif pub == SCIENCEDIRECT_PUBLISHER:
            hint = "Use the saved session first. If a row reports manual PDF click or human verification, complete the visible ScienceDirect check or click Download PDF in the Chrome tabs left open, then rerun only this group."
        else:
            hint = "Log in for this publisher/platform, then return to Codex and run only this group."
        for record in groups[pub]:
            rows.append(
                f"<tr><td>{record['excel_row']}</td>"
                f"<td><a target='_blank' href='{html_escape(record['_target_url'], quote=True)}'>{html_escape(record[TITLE])}</a>"
                f"<div class='sub'>{html_escape(record[JOURNAL])}</div></td>"
                f"<td>{html_escape(record[STATUS])}</td></tr>"
            )
        sections.append(
            f"<section><h2>{html_escape(pub)} <small>{len(groups[pub])} papers</small></h2>"
            f"<p class='hint'>{html_escape(hint)}</p>"
            f"<table><thead><tr><th>Excel row</th><th>Paper</th><th>Status</th></tr></thead><tbody>{''.join(rows)}</tbody></table></section>"
        )

    html = f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>PDF download queue</title><style>
body{{font-family:Segoe UI,Microsoft YaHei,Arial,sans-serif;margin:24px;color:#17202a}}h1{{font-size:22px}}h2{{font-size:18px;margin-top:24px;border-top:1px solid #dfe6ee;padding-top:18px}}small{{color:#667085;font-weight:400}}.notice{{background:#f1f7ff;border:1px solid #b8d6ff;padding:12px 14px;border-radius:6px;margin:14px 0 18px;line-height:1.6}}.pill{{display:inline-block;background:#eef2f7;border:1px solid #d7dee8;border-radius:999px;padding:5px 10px;margin:4px 6px 4px 0}}.hint{{color:#344054;background:#fff8e6;border-left:4px solid #f1b434;padding:8px 10px}}table{{border-collapse:collapse;width:100%;font-size:14px}}th,td{{border-bottom:1px solid #e4e8ee;padding:9px 8px;text-align:left;vertical-align:top}}th{{background:#f7f9fb}}a{{color:#075eb5;font-weight:600}}.sub{{color:#637083;margin-top:4px;font-size:12px}}</style></head><body>
<h1>Remaining PDF download queue</h1><div class="notice">Grouped by publisher. Log in to one publisher at a time; successful items disappear from the next queue.</div><div>{summary}</div>{''.join(sections)}</body></html>"""
    out.write_text(html, encoding="utf-8")
    print(out)


def mark_status(xlsx: Path, row: int, status: str) -> None:
    files, records = load_records(xlsx)
    for record in records:
        if int(record["excel_row"]) == row:
            record[STATUS] = status
            write_records(xlsx, files, records)
            return
    raise SystemExit(f"Row not found: {row}")


def find_chrome() -> str:
    candidates = [
        os.environ.get("CHROME_PATH", ""),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        str(Path.home() / r"AppData\Local\Google\Chrome\Application\chrome.exe"),
        "google-chrome",
        "chrome",
        "chromium",
    ]
    for candidate in candidates:
        if not candidate:
            continue
        if Path(candidate).exists() or shutil.which(candidate):
            return candidate
    raise SystemExit("Chrome executable not found. Set CHROME_PATH or install Chrome.")


def default_profile_dir() -> Path:
    override = os.environ.get("SCHOLAR_PDF_CHROME_PROFILE")
    if override:
        return Path(override).expanduser()
    return DEFAULT_PROFILE_DIR


def deep_merge_dict(base: dict, update: dict) -> dict:
    for key, value in update.items():
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            deep_merge_dict(base[key], value)
        else:
            base[key] = value
    return base


def merge_chrome_preferences(prefs_path: Path, prefs_update: dict) -> None:
    prefs = {}
    if prefs_path.exists():
        try:
            prefs = json.loads(prefs_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            backup = prefs_path.with_name(f"{prefs_path.name}.bak.{int(time.time())}")
            shutil.copy2(prefs_path, backup)
            print(f"Backed up unreadable Chrome Preferences to {backup}")
    deep_merge_dict(prefs, prefs_update)
    tmp_path = prefs_path.with_suffix(".tmp")
    tmp_path.write_text(json.dumps(prefs), encoding="utf-8")
    os.replace(tmp_path, prefs_path)


def launch_browser(queue: Path, profile: Path, download_dir: Path, port: int) -> None:
    profile.mkdir(parents=True, exist_ok=True)
    download_dir.mkdir(parents=True, exist_ok=True)
    default_dir = profile / "Default"
    default_dir.mkdir(parents=True, exist_ok=True)
    prefs_path = default_dir / "Preferences"
    prefs = {
        "download": {
            "default_directory": str(download_dir.resolve()),
            "prompt_for_download": False,
            "directory_upgrade": True,
        },
        "plugins": {"always_open_pdf_externally": True},
        "profile": {"default_content_setting_values": {"automatic_downloads": 1, "popups": 1}},
    }
    merge_chrome_preferences(prefs_path, prefs)
    url = queue.resolve().as_uri()
    args = [
        find_chrome(),
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-popup-blocking",
        "--new-window",
        url,
    ]
    subprocess.Popen(args)
    time.sleep(2)
    try:
        http_json(f"http://127.0.0.1:{port}/json/version")
        print(f"Chrome launched on port {port}: {url}")
        print(f"Using persistent Chrome profile: {profile.resolve()}")
    except Exception as exc:
        print(f"Chrome launched, but debug endpoint is not ready: {exc}")
        print(f"Using persistent Chrome profile: {profile.resolve()}")


def http_json(url: str, method: str = "GET") -> dict:
    request = Request(url, method=method)
    with urlopen(request, timeout=10) as response:
        return json.loads(response.read().decode("utf-8", errors="ignore"))


def open_tab(url: str, port: int) -> dict:
    return http_json(f"http://127.0.0.1:{port}/json/new?{quote(url, safe='')}", method="PUT")


def close_tab(tab_id: str, port: int) -> None:
    try:
        urlopen(Request(f"http://127.0.0.1:{port}/json/close/{tab_id}"), timeout=5).read()
    except Exception:
        pass


def leave_sciencedirect_manual_tabs(record: dict[str, str], port: int) -> list[str]:
    urls: list[str] = []
    for candidate in candidate_urls(record["_target_url"], record[JOURNAL]):
        parsed = urlparse(candidate)
        if "sciencedirect.com" not in parsed.netloc.lower():
            continue
        if candidate not in urls:
            urls.append(candidate)
        if len(urls) >= 2:
            break

    opened: list[str] = []
    for url in urls:
        try:
            open_tab(url, port)
            opened.append(url)
        except Exception as exc:
            print(f"Could not leave ScienceDirect tab open for row {record['excel_row']}: {exc}")
    return opened


def sanitize_filename(name: str, idx: int) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().strip(".")
    if not cleaned:
        cleaned = f"paper_{idx:02d}"
    if len(cleaned) > 150:
        cleaned = cleaned[:150].rstrip()
    return f"{idx:02d}_{cleaned}.pdf"


def is_pdf(path: Path) -> bool:
    try:
        return path.read_bytes()[:5].startswith(b"%PDF")
    except Exception:
        return False


def mdpi_direct_http_candidate(url: str) -> bool:
    parsed = urlparse(url)
    host = parsed.netloc.lower()
    if host.endswith("mdpi-res.com"):
        return parsed.path.lower().endswith(".pdf")
    if host.endswith("mdpi.com"):
        return parsed.path.rstrip("/").lower().endswith("/pdf")
    return False


def http_pdf_headers(url: str, resume_from: int = 0) -> dict[str, str]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "application/pdf,application/octet-stream;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
        "Referer": "https://www.mdpi.com/",
        "Connection": "keep-alive",
    }
    if resume_from:
        headers["Range"] = f"bytes={resume_from}-"
    return headers


def download_pdf_http(url: str, target: Path, timeout: int) -> tuple[bool, str]:
    part = target.with_name(target.name + ".part")
    last_reason = "HTTP download did not produce a PDF"
    target.parent.mkdir(parents=True, exist_ok=True)

    for attempt in range(1, HTTP_DOWNLOAD_RETRIES + 1):
        resume_from = part.stat().st_size if part.exists() else 0
        request = Request(url, headers=http_pdf_headers(url, resume_from))
        try:
            with urlopen(request, timeout=timeout) as response:
                status = response.getcode()
                mode = "ab" if resume_from and status == 206 else "wb"
                with part.open(mode) as handle:
                    while True:
                        chunk = response.read(HTTP_CHUNK_SIZE)
                        if not chunk:
                            break
                        handle.write(chunk)

            if stable_file(part) and is_pdf(part):
                os.replace(part, target)
                return True, ""

            last_reason = "HTTP response was not a PDF"
            if part.exists() and not is_pdf(part):
                part.unlink(missing_ok=True)
        except HTTPError as exc:
            if exc.code == 416 and part.exists() and is_pdf(part):
                os.replace(part, target)
                return True, ""
            last_reason = f"HTTP {exc.code}"
        except (TimeoutError, URLError, OSError) as exc:
            last_reason = f"{type(exc).__name__}: {str(exc)[:80]}"

        if attempt < HTTP_DOWNLOAD_RETRIES:
            time.sleep(min(2 * attempt, 8))

    return False, last_reason


def stable_file(path: Path) -> bool:
    try:
        first = path.stat().st_size
        time.sleep(0.8)
        second = path.stat().st_size
        return first == second and second > 1000
    except Exception:
        return False


def current_pdfs(download_dir: Path) -> set[Path]:
    return {p.resolve() for p in download_dir.glob("*.pdf")}


def find_new_pdf(download_dir: Path, before: set[Path], start: float, timeout: int) -> Path | None:
    deadline = time.time() + timeout
    candidate = None
    while time.time() < deadline:
        partials = list(download_dir.glob("*.crdownload"))
        fresh = []
        for pdf in download_dir.glob("*.pdf"):
            if pdf.resolve() not in before or pdf.stat().st_mtime >= start - 1:
                fresh.append(pdf)
        fresh.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        for pdf in fresh:
            if stable_file(pdf) and is_pdf(pdf):
                return pdf
            candidate = pdf
        if not partials and time.time() - start > 10 and not fresh:
            return None
        time.sleep(1)
    if candidate and is_pdf(candidate):
        return candidate
    return None


def normalize_mdpi_journal_name(journal: str) -> str:
    journal = re.sub(r"\s*[-,;:|].*$", "", journal)
    journal = journal.replace("&", "and")
    journal = re.sub(r"\([^)]*\)", "", journal)
    journal = re.sub(r"[^a-zA-Z0-9]+", " ", journal).strip().lower()
    return re.sub(r"\s+", " ", journal)


def mdpi_slug_candidates(url: str, journal: str) -> list[str]:
    parsed = urlparse(url)
    first_path_part = parsed.path.strip("/").split("/", 1)[0].upper()
    slugs: list[str] = []

    def add(slug: str | None) -> None:
        if slug and slug not in slugs:
            slugs.append(slug)

    add(MDPI_ISSN_SLUGS.get(first_path_part))

    normalized = normalize_mdpi_journal_name(journal)
    add(MDPI_JOURNAL_SLUGS.get(normalized))
    if normalized:
        add(normalized.replace(" ", ""))

    return slugs


def mdpi_res_candidate_urls(url: str, journal: str) -> list[str]:
    parsed = urlparse(url)
    path_parts = [part for part in parsed.path.strip("/").split("/") if part]
    if len(path_parts) < 4:
        return []

    volume, article = path_parts[1], path_parts[3]
    if not (volume.isdigit() and article.isdigit()):
        return []

    candidates: list[str] = []
    for slug in mdpi_slug_candidates(url, journal):
        stem = f"{slug}-{int(volume):02d}-{int(article):05d}"
        candidates.append(
            f"https://mdpi-res.com/d_attachment/{slug}/{stem}/article_deploy/{stem}.pdf"
        )
    return candidates


def candidate_urls(url: str, journal: str = "") -> list[str]:
    parsed = urlparse(url)
    host = parsed.netloc.lower()
    path = parsed.path
    out: list[str] = []

    def add(value: str) -> None:
        if value and value not in out:
            out.append(value)

    if "mdpi.com" not in host:
        add(url)
    if "arxiv.org" in host:
        if "/abs/" in path:
            add(url.replace("/abs/", "/pdf/") + ("" if url.endswith(".pdf") else ".pdf"))
        if "/pdf/" in path and not path.endswith(".pdf"):
            add(url.rstrip("/") + ".pdf")
    if "link.springer.com" in host and "/article/" in path:
        doi = path.split("/article/", 1)[1].strip("/")
        add(f"https://link.springer.com/content/pdf/{doi}.pdf")
    if "sciencedirect.com" in host and "/science/article/pii/" in path:
        add(url.rstrip("/") + "/pdfft?isDTMRedir=true&download=true")
        add(url.rstrip("/") + "/pdf")
    if "science.org" in host:
        if "/doi/full/" in path:
            add(url.replace("/doi/full/", "/doi/pdf/") + "?download=true")
        if "/doi/abs/" in path:
            add(url.replace("/doi/abs/", "/doi/pdf/") + "?download=true")
        if "/doi/pdf/" in path and "download=true" not in parsed.query:
            add(url + ("&" if parsed.query else "?") + "download=true")
    if "onlinelibrary.wiley.com" in host:
        if "/doi/abs/" in path:
            add(url.replace("/doi/abs/", "/doi/pdf/"))
            add(url.replace("/doi/abs/", "/doi/epdf/"))
        if "/doi/pdf/" in path:
            add(url + "?download=true")
    if "mdpi.com" in host:
        for mdpi_url in mdpi_res_candidate_urls(url, journal):
            add(mdpi_url)
        add(url)
        add(url.rstrip("/") + "/pdf")
        add(url.rstrip("/") + "/pdf?download=1")
    if "preprints.org" in host:
        add(url.rstrip("/") + "/download")
        add(url.rstrip("/") + "/download_pub")
        if "/manuscript/" in path:
            add(url.replace("/manuscript/", "/manuscript/download/"))
    if "copernicus.org" in host:
        article = re.search(r"/articles/([^/]+)/([^/]+)/([^/]+)/?$", path)
        if article:
            journal = host.split(".")[0]
            volume, first_page, year = article.groups()
            add(f"https://{host}/articles/{volume}/{first_page}/{year}/{journal}-{volume}-{first_page}-{year}.pdf")
    if "egusphere" in host and not path.endswith(".pdf"):
        preprint = re.search(r"/preprints/([^/]+)/([^/]+)/?$", path)
        if preprint:
            year, paper_id = preprint.groups()
            add(f"https://{host}/preprints/{year}/{paper_id}/{paper_id}.pdf")
    if "lyellcollection.org" in host and "/doi/abs/" in path:
        add(url.replace("/doi/abs/", "/doi/pdf/"))
    if "tandfonline.com" in host:
        if "/doi/full/" in path:
            add(url.replace("/doi/full/", "/doi/pdf/"))
        if "/doi/abs/" in path:
            add(url.replace("/doi/abs/", "/doi/pdf/"))
        if "/doi/pdf/" in path and "download=true" not in parsed.query:
            add(url + ("&" if parsed.query else "?") + "download=true")
    return out


def mark_researchgate_sources(xlsx: Path, queue_out: Path | None = None, json_out: Path | None = None) -> None:
    files, records = load_records(xlsx)
    changed = 0
    for record in records:
        if record["_publisher"] == "ResearchGate" and not record[STATUS].startswith(SUCCESS):
            record[STATUS] = RESEARCHGATE_STATUS
            changed += 1
    if changed:
        write_records(xlsx, files, records)
    if queue_out:
        build_queue(xlsx, queue_out, json_out)
    print(json.dumps({"marked_researchgate": changed}, ensure_ascii=False, indent=2))


def download_publisher(
    xlsx: Path,
    publisher_name: str,
    download_dir: Path,
    port: int,
    timeout: int,
    attempt_label: str,
    max_attempts: int | None,
    queue_out: Path | None,
    json_out: Path | None,
) -> None:
    if publisher_name == "ResearchGate":
        mark_researchgate_sources(xlsx, queue_out, json_out)
        return
    download_matching(
        xlsx=xlsx,
        label=publisher_name,
        match=lambda record: record["_publisher"] == publisher_name,
        download_dir=download_dir,
        port=port,
        timeout=timeout,
        queue_out=queue_out,
        json_out=json_out,
        failure_suffix=f" | {publisher_name} {attempt_label}",
        max_attempts=max_attempts,
    )


def download_open_access(
    xlsx: Path,
    download_dir: Path,
    port: int,
    timeout: int,
    queue_out: Path | None,
    json_out: Path | None,
) -> None:
    download_matching(
        xlsx=xlsx,
        label="no-login browser publishers",
        match=lambda record: record["_publisher"] in NO_LOGIN_BROWSER_PUBLISHERS,
        download_dir=download_dir,
        port=port,
        timeout=timeout,
        queue_out=queue_out,
        json_out=json_out,
        failure_suffix=" | no-login browser retry no download",
        max_attempts=None,
    )


def download_matching(
    xlsx: Path,
    label: str,
    match,
    download_dir: Path,
    port: int,
    timeout: int,
    queue_out: Path | None,
    json_out: Path | None,
    failure_suffix: str,
    max_attempts: int | None,
) -> None:
    try:
        http_json(f"http://127.0.0.1:{port}/json/version")
    except URLError as exc:
        raise SystemExit(f"Chrome debug endpoint unavailable on port {port}: {exc}") from exc

    download_dir.mkdir(parents=True, exist_ok=True)
    files, records = load_records(xlsx)
    attempted = 0
    new_success = 0
    manual_intervention = 0

    for idx, record in enumerate(records, start=1):
        if not record[STATUS].startswith(FAIL) or not match(record):
            continue
        if max_attempts is not None and attempted >= max_attempts:
            continue
        expected = download_dir / sanitize_filename(record[TITLE], idx)
        if expected.exists() and expected.stat().st_size > 1000 and is_pdf(expected):
            record[STATUS] = f"{SUCCESS}: {expected.name}"
            new_success += 1
            continue

        attempted += 1
        ok = False
        last_reason = "no PDF download from browser"
        for candidate in candidate_urls(record["_target_url"], record[JOURNAL]):
            if mdpi_direct_http_candidate(candidate):
                try:
                    downloaded, reason = download_pdf_http(candidate, expected, timeout)
                    if downloaded:
                        record[STATUS] = f"{SUCCESS}: {expected.name}"
                        print(f"OK row {record['excel_row']}: {expected.name}")
                        new_success += 1
                        ok = True
                        break
                    last_reason = reason
                except Exception as exc:
                    last_reason = f"{type(exc).__name__}: {str(exc)[:80]}"

            before = current_pdfs(download_dir)
            start = time.time()
            tab_id = None
            try:
                tab = open_tab(candidate, port)
                tab_id = tab.get("id")
                downloaded = find_new_pdf(download_dir, before, start, timeout)
                if downloaded:
                    if downloaded.resolve() != expected.resolve():
                        if expected.exists():
                            expected.unlink()
                        try:
                            downloaded.rename(expected)
                        except OSError:
                            shutil.copy2(downloaded, expected)
                            downloaded.unlink(missing_ok=True)
                    record[STATUS] = f"{SUCCESS}: {expected.name}"
                    print(f"OK row {record['excel_row']}: {expected.name}")
                    new_success += 1
                    ok = True
                    break
            except Exception as exc:  # keep batch moving
                last_reason = f"{type(exc).__name__}: {str(exc)[:80]}"
            finally:
                if tab_id:
                    close_tab(tab_id, port)
            time.sleep(0.5)

        if not ok:
            stop_for_manual = False
            if record["_publisher"] == SCIENCEDIRECT_PUBLISHER:
                manual_intervention += 1
                last_reason = SCIENCEDIRECT_MANUAL_REASON
                opened = leave_sciencedirect_manual_tabs(record, port)
                record[STATUS] = f"{FAIL}: {last_reason}{failure_suffix}"
                stop_for_manual = True
                print(
                    "ACTION REQUIRED row "
                    f"{record['excel_row']}: ScienceDirect may need a visible download click "
                    "or human verification. Complete it in the Chrome tab(s), then rerun "
                    f'--publisher "{SCIENCEDIRECT_PUBLISHER}".'
                )
                if opened:
                    print("Left open: " + " | ".join(opened))
            else:
                if failure_suffix not in record[STATUS]:
                    record[STATUS] = record[STATUS] + failure_suffix
            print(f"FAIL row {record['excel_row']}: {last_reason} :: {record[TITLE][:70]}")
            if stop_for_manual:
                break

    write_records(xlsx, files, records)
    if queue_out:
        build_queue(xlsx, queue_out, json_out)

    success_total = sum(r[STATUS].startswith(SUCCESS) for r in records)
    fail_total = sum(r[STATUS].startswith(FAIL) for r in records)
    remaining_target = sum(r[STATUS].startswith(FAIL) and match(r) for r in records)
    print(
        json.dumps(
            {
                "publisher": label,
                "group": label,
                "attempted": attempted,
                "new_success": new_success,
                "manual_intervention": manual_intervention,
                "total_success": success_total,
                "remaining_fail": fail_total,
                "remaining_for_group": remaining_target,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


def main() -> None:
    parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("summarize")
    p.add_argument("xlsx", type=Path)

    p = sub.add_parser("queue")
    p.add_argument("xlsx", type=Path)
    p.add_argument("--out", type=Path, default=Path("pdf_download_queue_remaining.html"))
    p.add_argument("--json-out", type=Path)

    p = sub.add_parser("mark")
    p.add_argument("xlsx", type=Path)
    p.add_argument("--row", type=int, required=True)
    p.add_argument("--status", required=True)

    p = sub.add_parser("launch-browser")
    p.add_argument("--queue", type=Path, required=True)
    p.add_argument(
        "--profile",
        type=Path,
        default=default_profile_dir(),
        help="Persistent Chrome profile used to reuse publisher login sessions.",
    )
    p.add_argument("--download-dir", type=Path, required=True)
    p.add_argument("--port", type=int, default=9223)

    p = sub.add_parser("download-publisher")
    p.add_argument("xlsx", type=Path)
    p.add_argument("--publisher", required=True)
    p.add_argument("--download-dir", type=Path, required=True)
    p.add_argument("--port", type=int, default=9223)
    p.add_argument("--timeout", type=int, default=90)
    p.add_argument(
        "--attempt-label",
        default="login retry no download",
        help="Status suffix used when this publisher round still does not download.",
    )
    p.add_argument("--queue-out", type=Path)
    p.add_argument("--json-out", type=Path)
    p.add_argument(
        "--max-attempts",
        type=int,
        help="Limit browser attempts for this publisher, useful for saved-session probes.",
    )

    p = sub.add_parser("download-open-access")
    p.add_argument("xlsx", type=Path)
    p.add_argument("--download-dir", type=Path, required=True)
    p.add_argument("--port", type=int, default=9223)
    p.add_argument("--timeout", type=int, default=90)
    p.add_argument("--queue-out", type=Path)
    p.add_argument("--json-out", type=Path)

    p = sub.add_parser("mark-researchgate")
    p.add_argument("xlsx", type=Path)
    p.add_argument("--queue-out", type=Path)
    p.add_argument("--json-out", type=Path)

    args = parser.parse_args()
    if args.cmd == "summarize":
        summarize(args.xlsx)
    elif args.cmd == "queue":
        build_queue(args.xlsx, args.out, args.json_out)
    elif args.cmd == "mark":
        mark_status(args.xlsx, args.row, args.status)
    elif args.cmd == "launch-browser":
        launch_browser(args.queue, args.profile, args.download_dir, args.port)
    elif args.cmd == "download-publisher":
        download_publisher(
            args.xlsx,
            args.publisher,
            args.download_dir,
            args.port,
            args.timeout,
            args.attempt_label,
            args.max_attempts,
            args.queue_out,
            args.json_out,
        )
    elif args.cmd == "download-open-access":
        download_open_access(
            args.xlsx,
            args.download_dir,
            args.port,
            args.timeout,
            args.queue_out,
            args.json_out,
        )
    elif args.cmd == "mark-researchgate":
        mark_researchgate_sources(args.xlsx, args.queue_out, args.json_out)


if __name__ == "__main__":
    main()
