---
name: "scholar-alert-pdf-downloader"
description: "End-to-end workflow for Google Scholar alert emails: search Gmail Scholar alerts for a user-specified date or date range, extract paper metadata into an Excel workbook, download accessible PDFs, launch a dedicated Chrome window for no-login and institutional download rounds, remove successes from the next queue, and update the workbook with PDF status. Use when the user asks to process Google Scholar alerts, Scholar Alert emails, paper lists, paper Excel files, or download PDFs from Scholar alert links."
---

# Scholar Alert PDF Downloader

## Overview

Use this skill to turn Google Scholar alert emails into a de-duplicated Excel paper table and a local PDF library. The workflow is not finished when the queue HTML is created: the agent must open a dedicated Chrome window, run no-login browser download rounds for public publishers, mark non-journal sources, probe saved login state cheaply, let the user log in to one gated publisher at a time, run the matching download round, then regenerate the queue until no accessible PDFs remain.

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

   Treat MDPI, arXiv, Copernicus / EGUsphere, Preprints, Science, and Taylor & Francis as no-login browser publishers. For MDPI, prefer the `mdpi-res.com/d_attachment/.../article_deploy/...pdf` route described in publisher patterns before the older `/pdf` fallbacks. For Science, derive `/doi/pdf/<doi>?download=true` from `/doi/full/<doi>`. For Taylor & Francis, derive `/doi/pdf/<doi>` from `/doi/full/<doi>` or `/doi/abs/<doi>`. If a plain HTTP/direct route fails for these publishers, do not present them as login-needed yet; retry them in the dedicated Chrome window before asking the user to log in to paid or institution-gated publishers.

6. **Generate the remaining queue.** Use:

   ```bash
   python scripts/scholar_pdf_workflow.py queue papers.xlsx --out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
   ```

7. **Launch the dedicated Chrome download window.** Do this immediately after generating the queue; do not stop at the HTML file.

   ```bash
   python scripts/scholar_pdf_workflow.py launch-browser \
     --queue pdf_download_queue_remaining.html \
     --download-dir papers_pdf \
     --port 9223
   ```

   By default this uses the persistent Chrome profile at `~/.codex/scholar_pdf_chrome_profile`. Reuse this profile across tasks so publisher cookies, SSO sessions, and local browser storage from manual logins can be reused when they are still valid. Use `--profile <path>` only when the user explicitly wants a separate browser identity.

8. **Run the no-login browser round.** Before asking the user to log in, retry open-access/browser-sensitive publishers in the dedicated Chrome window:

   ```bash
   python scripts/scholar_pdf_workflow.py download-open-access papers.xlsx \
     --download-dir papers_pdf \
     --port 9223 \
     --queue-out pdf_download_queue_remaining.html \
     --json-out failed_pdf_queue_remaining.json
   ```

   This should catch MDPI, arXiv, Copernicus / EGUsphere, Preprints, Science, and Taylor & Francis items that are public but need real browser behavior, cookies, redirects, or JavaScript-triggered downloads. Successful items disappear from the regenerated queue. Only remaining failures from these groups should be described as browser/manual-click failures, not as login failures. Do not add a separate broad direct-PDF sweep over every remaining failure after this step; candidate PDF routes are already tried inside the no-login and publisher rounds, and broad sweeps are slow and usually just record repeated `HTTP 403` or non-PDF responses.

9. **Mark ResearchGate rows as source-search tasks.** ResearchGate is not treated as the primary journal source. Do not spend browser automation time on ResearchGate downloads. Keep these rows in the workbook and mark them as needing a normal source lookup:

   ```bash
   python scripts/scholar_pdf_workflow.py mark-researchgate papers.xlsx \
     --queue-out pdf_download_queue_remaining.html \
     --json-out failed_pdf_queue_remaining.json
   ```

10. **Probe saved login state cheaply for gated publishers.** Before asking the user to log in again, run only one representative paper per remaining gated publisher in the persistent Chrome profile. This can reuse still-valid cookies or SSO state from a prior manual login, without extracting passwords or bypassing access controls. Do not run every paper in a publisher group just to test whether the login is still valid:

   ```bash
   python scripts/scholar_pdf_workflow.py download-publisher papers.xlsx \
     --publisher "ScienceDirect / Elsevier" \
     --download-dir papers_pdf \
     --port 9223 \
     --max-attempts 1 \
     --attempt-label "saved session retry no download" \
     --queue-out pdf_download_queue_remaining.html \
     --json-out failed_pdf_queue_remaining.json
   ```

   If the sample succeeds, continue with the regenerated queue and decide whether the group should be run fully. If it fails because the saved session expired, then ask the user to click that publisher group in the Chrome window and complete institutional login manually. Never ask for institutional passwords.

11. **Run one publisher download round after manual login.** After the user says something like `已登录 ScienceDirect` or `已登录 Wiley`, run only that publisher:

   ```bash
   python scripts/scholar_pdf_workflow.py download-publisher papers.xlsx \
     --publisher "ScienceDirect / Elsevier" \
     --download-dir papers_pdf \
     --port 9223 \
     --queue-out pdf_download_queue_remaining.html \
     --json-out failed_pdf_queue_remaining.json
   ```

   The command updates `PDF下载状态`, renames successful PDFs, and regenerates the queue so successes disappear from the next list.

12. **Repeat by publisher.** Reopen or refresh the regenerated queue page, probe the saved session with one paper for the next publisher first, then ask the user to log in only if the saved session is missing or expired. For ScienceDirect / Elsevier, if a browser attempt reaches a page that still needs a visible PDF click or human verification, the script leaves ScienceDirect tabs open, marks the row as needing manual action, and stops that publisher group; ask the user to complete the visible check or click the PDF/download button, then rerun only ScienceDirect. Stop any group if it still does not download after login; explain likely causes such as no entitlement, extra manual PDF button, anti-automation, or a non-PDF landing page.

13. **Finish with counts and paths.** Report total papers, successful PDF count, remaining failures by publisher, workbook path, PDF folder path, and backup path.

## Browser Login Rules

- Use the persistent dedicated Chrome profile for this workflow: `~/.codex/scholar_pdf_chrome_profile`.
- Reuse saved browser session state before asking the user to log in again.
- Do not extract, print, copy, or store passwords, MFA codes, or raw cookie values; only let Chrome reuse its own profile data.
- Configure the download folder to the paper PDF folder.
- Set PDFs to download externally instead of opening in Chrome's PDF viewer.
- Run no-login browser publishers before asking the user to log in.
- Let the user handle SSO, MFA, CAPTCHA, and consent screens.
- For ScienceDirect / Elsevier human-verification or manual PDF-button pages, leave the visible tab open and ask the user to complete the page; do not try to bypass the verification.
- Process only the gated publisher being handled in the saved-session or user-confirmed login round.
- When only checking saved login state, attempt one representative paper per publisher with `--max-attempts 1`; run the full group only after the user confirms login or the probe clearly succeeds.
- Do not repeatedly open pages for publishers the user has not logged in to.

## Script Helpers

Use `scripts/scholar_pdf_workflow.py` for the mechanical workbook, queue, browser, no-login, and publisher-round steps:

```bash
python scripts/scholar_pdf_workflow.py summarize papers.xlsx
python scripts/scholar_pdf_workflow.py queue papers.xlsx --out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
python scripts/scholar_pdf_workflow.py launch-browser --queue pdf_download_queue_remaining.html --download-dir papers_pdf --port 9223
python scripts/scholar_pdf_workflow.py download-open-access papers.xlsx --download-dir papers_pdf --port 9223 --queue-out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
python scripts/scholar_pdf_workflow.py mark-researchgate papers.xlsx --queue-out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
python scripts/scholar_pdf_workflow.py download-publisher papers.xlsx --publisher "ScienceDirect / Elsevier" --download-dir papers_pdf --port 9223 --max-attempts 1 --attempt-label "saved session retry no download" --queue-out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
python scripts/scholar_pdf_workflow.py download-publisher papers.xlsx --publisher "ScienceDirect / Elsevier" --download-dir papers_pdf --port 9223 --queue-out pdf_download_queue_remaining.html --json-out failed_pdf_queue_remaining.json
python scripts/scholar_pdf_workflow.py mark papers.xlsx --row 7 --status "成功: paper.pdf"
```

The script uses only the Python standard library and edits simple `.xlsx` files directly through OOXML.

## Status Conventions

Use these prefixes so future rounds can filter reliably:

- `成功: <filename>.pdf`
- `失败: <reason>`
- `失败: <reason> | no-login browser retry no download`
- `失败: <reason> | <publisher> saved session retry no download`
- `失败: <reason> | <publisher> login retry no download`
- `失败: 查询正常来源 (ResearchGate is not a journal source)`

- `失败: requires manual PDF click or human verification | ScienceDirect / Elsevier login retry no download`

## Quality Checks

Before finalizing:

- Validate the `.xlsx` is a valid zip package and has the expected number of rows.
- Count statuses by prefix.
- Verify every downloaded file starts with `%PDF`.
- Confirm the queue page contains only remaining failures.
- Confirm the user-facing summary distinguishes public-download successes, no-login browser successes, and login-round successes.
