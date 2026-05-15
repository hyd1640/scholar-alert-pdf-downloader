# Publisher PDF Patterns

Use these patterns only for legitimate access. Do not bypass paywalls, login flows, CAPTCHA, MFA, or institutional access controls.

## Common routes

- Springer article: `https://link.springer.com/article/<doi>`
  - PDF candidate: `https://link.springer.com/content/pdf/<doi>.pdf`
  - Springer may require a cookie-setting identity redirect even for page access. Let the user log in or initialize the site in the browser.

- ScienceDirect article: `https://www.sciencedirect.com/science/article/pii/<pii>`
  - PDF candidate: append `/pdfft?isDTMRedir=true&download=true`
  - Alternate: append `/pdf`
  - Some ScienceDirect sessions land on a visible verification page or an article page that still needs a `Download PDF` click after authentication. Leave the page open for the user and retry after they complete the visible step; do not bypass human verification.

- Wiley abstract: `https://onlinelibrary.wiley.com/doi/abs/<doi>`
  - PDF candidate: replace `/doi/abs/` with `/doi/pdf/`
  - Alternate reader: replace `/doi/abs/` with `/doi/epdf/`
  - Some Wiley pages require a visible `PDF` or `Download PDF` click after authentication.

- MDPI article: `https://www.mdpi.com/<journal>/<volume>/<issue>/<article>`
  - Current MDPI article pages may block plain non-browser requests to `www.mdpi.com/.../pdf`.
  - Preferred PDF candidate: `https://mdpi-res.com/d_attachment/<journal-slug>/<journal-slug>-<volume-2digits>-<article-5digits>/article_deploy/<journal-slug>-<volume-2digits>-<article-5digits>.pdf`
  - Example: `https://www.mdpi.com/2072-4292/18/10/1515` -> `https://mdpi-res.com/d_attachment/remotesensing/remotesensing-18-01515/article_deploy/remotesensing-18-01515.pdf`
  - The `?version=...` query string seen on MDPI pages is not required for the PDF download.
  - Fallback candidates: append `/pdf`, then `/pdf?download=1`.
  - When the page URL uses an ISSN instead of the journal slug, derive the slug from known ISSN mappings first, then from the workbook journal name. Common irregular slugs include `Remote Sensing -> remotesensing`, `Applied Sciences -> applsci`, `International Journal of Molecular Sciences -> ijms`, `Journal of Clinical Medicine -> jcm`, and `Journal of Marine Science and Engineering -> jmse`.
  - Treat as no-login/open-access. If direct HTTP fails, retry through the dedicated Chrome window before adding it to a manual login queue.

- arXiv abstract: `https://arxiv.org/abs/<id>`
  - PDF candidate: replace `/abs/` with `/pdf/` and append `.pdf` when needed.
  - Treat as no-login/open-access.

- Preprints pages on `preprints.org`
  - PDF candidates: append `/download` or `/download_pub`; manuscript URLs may also support replacing `/manuscript/` with `/manuscript/download/`.
  - Treat as no-login/open-access.

- Taylor & Francis PDF: `https://www.tandfonline.com/doi/pdf/<doi>`
  - Full-text page: `https://www.tandfonline.com/doi/full/<doi>`
  - PDF candidate: replace `/doi/full/` or `/doi/abs/` with `/doi/pdf/`
  - If already on `/doi/pdf/`, optionally append `?download=true`
  - Many Taylor & Francis PDF links download directly in the dedicated browser and do not need a login round.

- Science full text: `https://www.science.org/doi/full/<doi>`
  - PDF candidate: replace `/doi/full/` with `/doi/pdf/` and append `?download=true`
  - Example: `https://www.science.org/doi/full/10.1126/science.aec3187` -> `https://www.science.org/doi/pdf/10.1126/science.aec3187?download=true`
  - Treat Science as a no-login/browser-sensitive publisher first; only escalate if the browser route clearly requires entitlement.

- Lyell Collection abstract: `https://www.lyellcollection.org/doi/abs/<doi>`
  - PDF candidate: replace `/doi/abs/` with `/doi/pdf/`

- Copernicus/EGUsphere/NHESS pages often expose direct PDF links and usually do not need login.
  - Copernicus article pages commonly use `<journal>-<volume>-<first-page>-<year>.pdf`.
  - EGUsphere preprint pages commonly use `<paper-id>.pdf` inside the preprint directory.
  - Treat as no-login/open-access.

- IEEE Xplore PDFs may require institution entitlement and may block non-browser downloads. Use the browser after login.

- ProQuest often requires interactive pages. Prefer visible download buttons and stop if the page requires account-specific access.

- ResearchGate is not treated as the primary journal source. Do not spend browser automation time on ResearchGate downloads. Keep the row in the workbook and mark it as `失败: 查询正常来源 (ResearchGate is not a journal source)` so the user can search for the normal publisher or journal source.

## Failure language

Use short workbook-safe reasons:

- `HTTP 403`
- `not a public PDF`
- `no PDF download from browser`
- `requires manual PDF click`
- `requires manual PDF click or human verification`
- `no institutional entitlement`
- `login required`
