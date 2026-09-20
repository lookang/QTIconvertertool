# Contributing

Contributions that improve DOCX parsing, QTI validity, accessibility, privacy, documentation or browser compatibility are welcome.

## Set up the project

Clone the repository and create a Python virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
```

Run the Python test suite:

```bash
python -m unittest discover -s tests -v
```

Serve the browser application:

```bash
python -m http.server 8080
```

Open `http://localhost:8080/docx_to_qti.html`. For the optional activity endpoint, serve the same directory with PHP 8 or later instead:

```bash
php -S 127.0.0.1:8080
```

With the HTTP server running, exercise every tracked `.docx` fixture through the browser parser and QTI ZIP builder:

```powershell
npx --yes --package @playwright/cli playwright-cli -s=qti-regression open http://127.0.0.1:8080/docx_to_qti.html
npx --yes --package @playwright/cli playwright-cli -s=qti-regression run-code --filename tests/browser_docx_batch_regression.js
npx --yes --package @playwright/cli playwright-cli -s=qti-regression eval "() => window.__qtiBatchRegressionResults"
```

Every row should report `ok: true`. Question-paper rows also verify that the generated ZIP contains one QTI item per parsed question, all extracted assets, valid image references, `imsmanifest.xml`, and `assessment_test.xml`.

## Source boundaries

- `converter.py`, `parser.py` and `qti_generator.py` contain the Python conversion path.
- `docx_to_qti.html` contains the browser-side parser, editor and QTI builder.
- `conversion-hero-3d.js` and `activity-dashboard.js` contain the Three.js interfaces.
- `activity.php` is optional and must never receive document contents, filenames or user names.
- `vendor/` contains pinned browser runtime assets. Update its notices whenever a dependency changes.

## Before submitting a change

1. Add a focused regression test for every parsing bug.
2. Run the complete Python test suite.
3. Check JavaScript syntax with `node --check activity-dashboard.js` and `node --check conversion-hero-3d.js`.
4. Exercise document selection, conversion, review and ZIP export in a browser.
5. Check desktop and narrow mobile layouts and review the browser console.
6. Do not commit confidential exam papers, generated ZIP packages, analytics data, credentials or private keys.

Small synthetic fixtures are preferred. If a real document is required to reproduce a bug, obtain permission before publishing it and remove personal or confidential content.

## Pull requests

Keep each pull request focused. Describe the document structure or UI behaviour being addressed, the expected result and the validation performed. Preserve the browser-only privacy guarantee unless the change is explicitly proposing a separately reviewed server workflow.
