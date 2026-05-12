---
name: "scholar-alert-pdf-downloader"
description: "End-to-end workflow for Google Scholar alert emails: search Gmail Scholar alerts for a user-specified date or date range, extract paper metadata into an Excel workbook, download accessible PDFs, launch a dedicated Chrome window for no-login and institutional download rounds, remove successes from the next queue, and update the workbook with PDF status. Use when the user asks to process Google Scholar alerts, Scholar Alert emails, paper lists, paper Excel files, or download PDFs from Scholar alert links."
---

# Scholar Alert PDF Downloader

## Overview

Use this skill to turn Google Scholar alert emails into a de-duplicated Excel paper table and a local PDF library. The workflow is not finished when the queue HTML is created: the agent must open a dedicated Chrome window, run no-login browser download rounds for public publishers, let the user log in to one gated publisher at a time, run the matching download round, then regenerate the queue until no accessible PDFs remain.

## Workflow

1. **Confirm date scope.** Require the user request or conversation to provide a target date or date range. Do not hard-code a date. If missing, ask for it before Gmail search. Resolve relative dates using the current date and the user's timezone.

2. **Search Gmail precisely.** Use Gmail when available:
   - Sender: `scholaralerts-noreply@google.com`
   - Subject phrase: user-provided, such as `新的相关研究工作`
   - Date bounds: convert the user-specified local date/range into Gmail query bounds, then verify returned timestamps against the user's timezone.

3. **Parse Scholar alert bodies.** Each paper usually appears as a Markdown title link followed by `Authors - Journal, Year`. Extract:
   - `Title`
   - `Authors`
   - `Journal`
   - `Link`

   Prefer the target URL embedded in `scholar_url?url=...` over the Google Scholar redirect URL. De-duplicate by normalized title plus target URL.

4. **Create the Excel workbook.** Use the four required columns above. Once PDF work begins, add `PDF下载状态`. Preserve clickable hyperlinks in `Link`. Keep a backup before rewrites.

5. **Download public PDFs first.** Try direct public PDF routes without login. Use [references/publisher-patterns.md](references/publisher-patterns.md) for publisher routes. Do not bypass paywalls, credential screens, CAPTCHA, MFA, or access controls. Mark blocked items as `失败: <reason>`.

   Treat MDPI, arXiv, Copernicus / EGUsphere, and Preprints as no-login browser publishers. For MDPI, prefer the `mdpi-res.com/d_attachment/.../article_deploy/...pdf` route described in publisher patterns before the older `/pdf` fallbacks. If a plain HTTP/direct route fails for these publishers, do not present them as login-needed yet; retry them in the dedicated Chrome window before asking the user to log in to paid or institution-gated publishers.

6. **Generate the remaining queue.** Use:

   ```bash
   python scripts/scholar_pdf_workflow.py queue papers.xlsx --out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
   ```

7. **Launch the dedicated Chrome download window.** Do this immediately after generating the queue; do not stop at the HTML file.

   ```bash
   python scripts/scholar_pdf_workflow.py launch-browser \
     --queue pdf_download_queue_remaining.html \
     --profile chrome_pdf_download_profile \
     --download-dir papers_pdf \
     --port 9223
   ```

8. **Run the no-login browser round.** Before asking the user to log in, retry open-access/browser-sensitive publishers in the dedicated Chrome window:

   ```bash
   python scripts/scholar_pdf_workflow.py download-open-access papers.xlsx \
     --download-dir papers_pdf \
     --port 9223 \
     --queue-out pdf_download_queue_remaining.html \
     --json-out failed_pdf_queue_remaining.json
   ```

   This should catch MDPI, arXiv, Copernicus / EGUsphere, and Preprints items that are public but need real browser behavior, cookies, redirects, or JavaScript-triggered downloads. Successful items disappear from the regenerated queue. Only remaining failures from these groups should be described as browser/manual-click failures, not as login failures.

9. **Ask the user to log in for a gated publisher.** Tell the user to click one remaining publisher group in that Chrome window and complete institutional login manually. Never ask for institutional passwords.

10. **Run one publisher download round.** After the user says something like `已登录 ScienceDirect` or `已登录 Wiley`, run only that publisher:

   ```bash
   python scripts/scholar_pdf_workflow.py download-publisher papers.xlsx \
     --publisher "ScienceDirect / Elsevier" \
     --download-dir papers_pdf \
     --port 9223 \
     --queue-out pdf_download_queue_remaining.html \
     --json-out failed_pdf_queue_remaining.json
   ```

   The command updates `PDF下载状态`, renames successful PDFs, and regenerates the queue so successes disappear from the next list.

11. **Repeat by publisher.** Reopen or refresh the regenerated queue page, ask the user to log in to the next publisher, and run `download-publisher` for only that group. Stop a group if it still does not download after login; explain likely causes such as no entitlement, extra manual PDF button, anti-automation, or a non-PDF landing page.

12. **Finish with counts and paths.** Report total papers, successful PDF count, remaining failures by publisher, workbook path, PDF folder path, and backup path.

## Browser Login Rules

- Use a dedicated Chrome profile for this workflow.
- Configure the download folder to the paper PDF folder.
- Set PDFs to download externally instead of opening in Chrome's PDF viewer.
- Run no-login browser publishers before asking the user to log in.
- Let the user handle SSO, MFA, CAPTCHA, and consent screens.
- Process only the gated publisher the user explicitly says is logged in.
- Do not repeatedly open pages for publishers the user has not logged in to.

## Script Helpers

Use `scripts/scholar_pdf_workflow.py` for the mechanical workbook, queue, browser, no-login, and publisher-round steps:

```bash
python scripts/scholar_pdf_workflow.py summarize papers.xlsx
python scripts/scholar_pdf_workflow.py queue papers.xlsx --out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
python scripts/scholar_pdf_workflow.py launch-browser --queue pdf_download_queue_remaining.html --profile chrome_pdf_download_profile --download-dir papers_pdf --port 9223
python scripts/scholar_pdf_workflow.py download-open-access papers.xlsx --download-dir papers_pdf --port 9223 --queue-out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
python scripts/scholar_pdf_workflow.py download-publisher papers.xlsx --publisher "ScienceDirect / Elsevier" --download-dir papers_pdf --port 9223 --queue-out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
python scripts/scholar_pdf_workflow.py mark papers.xlsx --row 7 --status "成功: paper.pdf"
```

The script uses only the Python standard library and edits simple `.xlsx` files directly through OOXML.

## Status Conventions

Use these prefixes so future rounds can filter reliably:

- `成功: <filename>.pdf`
- `失败: <reason>`
- `失败: <reason> | no-login browser retry no download`
- `失败: <reason> | <publisher> login retry no download`

## Quality Checks

Before finalizing:

- Validate the `.xlsx` is a valid zip package and has the expected number of rows.
- Count statuses by prefix.
- Verify every downloaded file starts with `%PDF`.
- Confirm the queue page contains only remaining failures.
- Confirm the user-facing summary distinguishes public-download successes, no-login browser successes, and login-round successes.
