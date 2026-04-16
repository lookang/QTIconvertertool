# QTI Converter Tool

Convert Singapore exam papers (Word `.docx`) into IMS QTI 2.1 packages ready to import into **Student Learning Space (SLS)** and any standards-compliant LMS.

---

## 🎬 Tutorial

<div align="center">
  <a href="https://www.youtube.com/watch?v=0KFOVRdJfX0" target="_blank">
    <img src="https://img.youtube.com/vi/0KFOVRdJfX0/maxresdefault.jpg"
         alt="Watch: DOCX to QTI Converter — How to use"
         width="720">
  </a>
  <br><br>
  <a href="https://www.youtube.com/watch?v=0KFOVRdJfX0" target="_blank">
    <img src="https://img.shields.io/badge/▶%20Watch%20on%20YouTube-FF0000?style=for-the-badge&logo=youtube&logoColor=white"
         alt="Watch on YouTube">
  </a>
  &nbsp;
  <a href="https://iwant2study.org/lookangejss/QTIlowJunHua/docx_to_qti.html" target="_blank">
    <img src="https://img.shields.io/badge/🚀%20Try%20Live%20Demo-4F46E5?style=for-the-badge"
         alt="Try Live Demo">
  </a>
</div>

---

## Features

- **Automatic question extraction** — detects MCQ and structured questions from table-based exam layouts
- **Mark scheme integration** — upload the matching mark scheme alongside the question paper to auto-populate `<correctResponse>` in every QTI item
- **Image extraction** — images embedded in table cells are extracted and bundled in the output ZIP
- **Rich text preservation** — superscript, subscript, bold, and italic formatting (e.g. `v²`, `10⁻⁶`) is retained via HTML tags in QTI item stems
- **Filename passthrough** — output ZIP uses the same name as the input file (e.g. `2022 JC2 H1 Physics P1.docx` → `2022 JC2 H1 Physics P1.zip`)
- **Two usage modes** — Flask web server for local/hosted use, or a single self-contained HTML file for fully offline use (no server required)

---

## Usage

### Option A — Static HTML (no install required)

1. Open `docx_to_qti.html` in any modern browser
2. Drop your question paper `.docx` onto the main drop zone
3. Optionally drop the matching mark scheme `.docx` onto the second drop zone
4. Click **Convert Document**
5. The QTI ZIP downloads automatically

No Python, no server, no internet connection needed.

---

### Option B — Flask Web Server

#### Requirements

- Python 3.8+
- Install dependencies:

```bash
pip install -r requirements.txt
```

#### Run

```bash
python app.py
```

Then open [http://localhost:5050](http://localhost:5050) in your browser.

#### API endpoint

`POST /convert`

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `file` | `.docx` | Yes | Question paper |
| `ms_file` | `.docx` | No | Mark scheme |

Returns a `.zip` file with the same base name as the uploaded question paper.

---

## Output Structure

```
2022 JC2 H1 Physics P1.zip
├── imsmanifest.xml          # QTI package manifest
├── items/
│   ├── Q001_<id>.xml        # One assessmentItem per question
│   ├── Q002_<id>.xml
│   └── ...
└── assets/
    ├── img_0.png            # Images extracted from the document
    └── img_1.png
```

Each `assessmentItem` follows IMS QTI 2.1:

- **MCQ** → `choiceInteraction` with `simpleChoice` identifiers `a`/`b`/`c`/`d`
- **Structured** → `extendedTextInteraction`
- **Correct answer** → `<correctResponse>` populated from mark scheme (if provided) or inline ANSWERS section

---

## Document Format Requirements

For best results, format the source Word document as follows:

| Element | Convention |
|---------|-----------|
| Question number | Each question begins with its number, e.g. `1 What is ...` |
| MCQ options | Labelled `A. B. C. D.` or `A) B) C) D)` in text, or one-per-row in a table with the letter in the first cell |
| Images | Embedded directly in the document; overlaid text should be replaced with screenshots |
| Inline answers | Append an `ANSWERS` heading at the end; list answers as `1 B`, `2 A`, etc. |
| Mark scheme | Upload as a separate `.docx`; supports P1 MCQ grid format and P2/P3 structured format |

---

## Mark Scheme Formats

### Paper 1 (MCQ) — Grid format

A 6×10+ table where alternating rows are question numbers and answers:

```
| 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 |
| B | C | A | D | B | A | C | D | B | A  |
| 11| 12| ...                              |
| D | A | ...                              |
```

### Paper 2 / Paper 3 (Structured) — Row format

Rows of `[part label, answer text, marks]`:

```
| 1a   | The net force is zero.         | 2 |
| 1bi  | By Newton's third law ...      | 3 |
| 2a   | ...                            | 1 |
```

---

## Known Handling for Edge Cases

| Issue | Solution |
|-------|----------|
| Year in document title (e.g. `2022`) detected as question number | `MAX_QNUM = 200` guard ignores numbers above 200 |
| Word cross-reference field codes prepend letters (e.g. `XX23`) | Regex `^[A-Z]{0,5}(\d{1,3})$` strips the prefix |
| Sub-questions embedded inside a larger table (e.g. Q25 inside Q24's table) | Sub-question detection in inner row loop flushes and starts a new question |
| Q1-style MCQ where each option row starts with the letter (`A`, `B`, `C`, `D`) | Dedicated branch detects single-letter first cell and treats remaining cells as option text |

---

## Project Structure

```
QTIconvertertool/
├── app.py               # Flask server entry point
├── converter.py         # Core DOCX parser and QTI builder (Python)
├── docx_to_qti.html     # Self-contained browser-side converter (JavaScript + JSZip)
├── templates/
│   └── index.html       # Flask frontend with dual drop zones
├── static/
│   └── sample_doc.png   # Sample document screenshot shown in UI
├── requirements.txt     # Python dependencies
└── Procfile             # For deployment on Heroku / Render
```

---

## Dependencies

### Python (Flask server)
- `flask` — web framework
- `python-docx` — DOCX parsing
- `lxml` — XML/XPath for run-level formatting

### Browser (static HTML)
- [JSZip](https://stuk.github.io/jszip/) — loaded from CDN, used to read `.docx` files client-side

---

## Deployment

The app includes a `Procfile` for one-click deployment to platforms like **Render** or **Heroku**:

```
web: python app.py
```

Set the `PORT` environment variable if needed (defaults to `5050`).

---

## License

MIT License. Free to use, modify, and distribute.

---

## Live Demo

The static HTML version is deployed at:

**https://iwant2study.org/lookangejss/QTIlowJunHua/docx_to_qti.html**

No installation required — open in any modern browser and convert directly.

---

## Acknowledgements

This tool was built upon the first prototype by **Low Jun Hua**:
[https://github.com/junhualow/docx-qti-converter](https://github.com/junhualow/docx-qti-converter)

Built to support Singapore teachers converting examination papers for use in digital learning platforms.
