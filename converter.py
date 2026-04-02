from docx import Document
from docx.oxml.ns import qn
import os
import re


# ─────────────────────────────────────────────────────────────────────────────
# MARK-SCHEME PARSER
# ─────────────────────────────────────────────────────────────────────────────

def parse_mark_scheme(ms_docx_path):
    """
    Parse a mark-scheme .docx and return a dict  {qnum_str: answer}.

    MCQ papers (P1):
        The first table is a 6-row × 10-col grid of alternating
        [question-numbers, answer-letters] rows. Each answer letter is stored
        as lowercase ('a'–'d') to match QTI simpleChoice identifiers.

    Structured papers (P2 / P3):
        The table rows have the form  [part_label, answer_text, marks…].
        All parts belonging to the same question number are concatenated.
        Result: { '1': '1a: …\n1bi: …\n…', '2': '2a: …', … }
    """
    doc    = Document(ms_docx_path)
    tables = doc.tables
    if not tables:
        return {}

    first  = tables[0]
    nrows  = len(first.rows)
    ncols  = len(first.columns) if first.rows else 0

    # --- MCQ grid detection ---
    if nrows >= 6 and ncols >= 10:
        first_cell = first.rows[0].cells[0].text.strip()
        if first_cell.isdigit():
            return _parse_mcq_grid(first)

    # --- Structured MS ---
    return _parse_structured_ms(tables)


def _parse_mcq_grid(grid_table):
    """
    Alternating rows: [q-numbers …], [answer-letters …], [q-numbers …], …
    Returns { '1': 'b', '2': 'c', … } (lowercase letters).
    """
    answers = {}
    rows = [[c.text.strip() for c in r.cells] for r in grid_table.rows]

    for i in range(0, len(rows) - 1, 2):
        nums_row = rows[i]
        ans_row  = rows[i + 1]
        for qn, ans in zip(nums_row, ans_row):
            qn  = qn.strip()
            ans = ans.strip().upper()
            if qn.isdigit() and ans in ('A', 'B', 'C', 'D'):
                answers[qn] = ans.lower()

    return answers


def _parse_structured_ms(all_tables):
    """
    Rows: [part_label, answer_text, marks…]
    e.g. ['1a', 'rate of change of velocity', 'B1']
         ['1bi', 'Taking upwards …', 'C1\nA1']
    Returns { '1': '1a: rate of change…\n1bi: Taking upwards…', … }
    """
    answers = {}

    for table in all_tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 2:
                continue
            label = cells[0]
            text  = cells[1]
            if not label or not text:
                continue
            m = re.match(r'^(\d+)', label)
            if not m:
                continue
            qnum = m.group(1)
            entry = f"{label}: {text}"
            if qnum in answers:
                answers[qnum] += "\n" + entry
            else:
                answers[qnum] = entry

    return answers


# ─────────────────────────────────────────────────────────────────────────────
# MAIN CONVERTER
# ─────────────────────────────────────────────────────────────────────────────

def convert_docx_to_qti(input_docx, job_dir, ms_answers=None):
    """
    Convert a question-paper .docx to a QTI 2.1 ZIP package.

    Parameters
    ----------
    input_docx  : path to the question-paper .docx
    job_dir     : working directory for output files
    ms_answers  : optional dict from parse_mark_scheme(); keys are question
                  number strings ('1', '2', …); values are answer strings.
                  For MCQ items the value should be a single lowercase letter
                  ('a'–'d') matching the simpleChoice identifier.
    """
    doc = Document(input_docx)

    assets_dir = os.path.join(job_dir, "assets")
    items_dir  = os.path.join(job_dir, "items")
    os.makedirs(assets_dir, exist_ok=True)

    if ms_answers is None:
        ms_answers = {}

    # -----------------------------
    # IMAGE HANDLING
    # -----------------------------
    image_counter = 0

    def save_image(doc, rId):
        nonlocal image_counter
        image_part = doc.part.related_parts[rId]
        filename   = f"img_{image_counter}.jpg"
        with open(os.path.join(assets_dir, filename), "wb") as f:
            f.write(image_part.blob)
        image_counter += 1
        return filename

    # -----------------------------
    # ITERATE WORD CONTENT
    # -----------------------------

    def extract_run_text(run):
        """Extract text from a single run, applying sup/sub/bold/italic HTML tags."""
        text = "".join(node.text or "" for node in run.xpath(".//w:t"))
        if not text:
            return ""
        is_sup    = run.xpath(".//w:vertAlign[@w:val='superscript']")
        is_sub    = run.xpath(".//w:vertAlign[@w:val='subscript']")
        is_bold   = run.xpath(".//w:b")
        is_italic = run.xpath(".//w:i")
        if is_sup:    text = f"<sup>{text}</sup>"
        if is_sub:    text = f"<sub>{text}</sub>"
        if is_bold:   text = f"<strong>{text}</strong>"
        if is_italic: text = f"<em>{text}</em>"
        return text

    def merge_adjacent_tags(text):
        """Merge adjacent identical inline tags: <sup>a</sup><sup>b</sup> → <sup>ab</sup>."""
        for tag in ('sup', 'sub', 'strong', 'em'):
            text = re.sub(f'</{tag}><{tag}>', '', text)
        return text

    def extract_cell_text(cell):
        """Extract formatted text (sup/sub/bold/italic) from a table cell element."""
        raw = "".join(extract_run_text(r) for r in cell.xpath(".//w:r")).strip()
        return merge_adjacent_tags(raw)

    def iter_block_items(doc):
        body = doc.element.body
        for child in body.iterchildren():

            # ── Paragraph ──────────────────────────────────────────────
            if child.tag.endswith('p'):

                paragraph_text = "".join(
                    extract_run_text(r) for r in child.xpath("./w:r")
                ).strip()
                paragraph_text = merge_adjacent_tags(paragraph_text)

                if paragraph_text:
                    yield ("text", paragraph_text)

                for node in child.iter():
                    if node.tag.endswith('blip'):
                        rId = node.get(qn('r:embed'))
                        if rId:
                            yield ("image", rId)

            # ── Table ──────────────────────────────────────────────────
            if child.tag.endswith('tbl'):

                table_data      = []
                seen_image_rids = []   # deduplicate images within one table

                for row in child.xpath(".//w:tr"):

                    row_data = []

                    for cell in row.xpath(".//w:tc"):
                        # FIX: use formatted text (preserves <sup>/<sub>)
                        row_data.append(extract_cell_text(cell))

                        # FIX: collect images from table cells
                        for node in cell.iter():
                            if node.tag.endswith('blip'):
                                rId = node.get(qn('r:embed'))
                                if rId and rId not in seen_image_rids:
                                    seen_image_rids.append(rId)

                    table_data.append(row_data)

                yield ("table", table_data)

                # Yield images discovered inside table cells so they are
                # attached to the current question token list
                for rId in seen_image_rids:
                    yield ("image", rId)

    # -----------------------------
    # CELL DEDUPLICATION
    # Merged cells in Word tables repeat identical text — keep only first
    # occurrence in each row so question text isn't duplicated.
    # -----------------------------
    def dedup_cells(cells):
        seen  = set()
        parts = []
        for c in cells:
            c = c.strip()
            if c and c not in seen:
                seen.add(c)
                parts.append(c)
        return ' '.join(parts)

    # Options row: ['', 'A   text', 'B   text', 'C   text', 'D   text']
    OPT_CELL_RE  = re.compile(r'^([A-D])\s+(.*)', re.DOTALL | re.I)
    # Single-letter MCQ label that IS an option identifier (not part of a word)
    OPT_LABEL_RE = re.compile(r'^[A-D]$', re.I)

    def extract_opts_from_row(row):
        """Return list of option texts from a table options row (format: ['', 'A text', ...]).
        Also handles bold option letters, e.g. ['', '<strong>A</strong>  text', ...]."""
        opts = []
        for cell in row[1:]:          # skip first (empty) cell
            cell = cell.strip()
            if not cell:
                continue
            # Strip HTML tags to check for option-letter prefix (handles bold letters)
            cell_plain = re.sub(r'<[^>]+>', '', cell)
            m = OPT_CELL_RE.match(cell_plain)
            if m:
                # Remove the leading letter (possibly HTML-wrapped) from original cell
                opt_raw = re.sub(r'^(?:<[^>]+>)*[A-D](?:</[^>]+>)*\s*', '', cell,
                                 flags=re.IGNORECASE)
                # Strip dangling closing tags at start (e.g. '</strong>' when letter
                # shares an opening tag with the content: '<strong>A </strong>text')
                opt_raw = re.sub(r'^(\s*</[^>]+>)+', '', opt_raw).strip()
                opts.append(opt_raw or m.group(2).strip())
            # lone 'A'/'B' etc (image-only options) → skip (image captured separately)
        return opts

    # -----------------------------
    # REGEX PATTERNS (STRICT)
    # -----------------------------
    question_regex  = re.compile(r'^(\d+)\s+(.*)')
    option_regex    = re.compile(r'^[A-D][\.\)]\s*(.*)')
    # Table first-cell is a bare question number, optionally with a short
    # non-digit prefix from Word cross-reference field codes (e.g. "XX23").
    qnum_only_regex = re.compile(r'^[A-Z]{0,5}(\d{1,3})$')

    MAX_QNUM = 200   # ignore "question numbers" > this (e.g. the year 2022)

    def parse_qnum(first_cell):
        """Return int qnum from first_cell if it looks like a question-start,
        else return None.  Handles bare '25', prefixed 'XX23', and bold
        '<strong>1</strong>' cells (strip HTML tags before matching)."""
        clean = re.sub(r'<[^>]+>', '', first_cell).strip()
        m = qnum_only_regex.match(clean)
        if m:
            n = int(m.group(1))
            if n <= MAX_QNUM:
                return str(n)   # normalise to plain digit string
        return None

    # -----------------------------
    # STORAGE
    # -----------------------------
    mcq_questions        = []
    structured_questions = []

    current_qnum   = None
    current_tokens = []
    options        = []

    answers              = {}
    reading_answers      = False
    current_answer_q     = None
    current_answer_tokens = []

    # -----------------------------
    # PARSE DOCUMENT
    # -----------------------------
    for item_type, value in iter_block_items(doc):

        if reading_answers:

            if item_type == "image" and current_answer_q is not None:
                filename = save_image(doc, value)
                current_answer_tokens.append(("image", filename))
                continue

            if item_type == "table":
                current_answer_tokens.append(("table", value))
                continue

            if item_type == "text":

                text   = value.replace('\xa0', ' ').strip()
                amatch = re.match(r'^(\d+)', text)

                if amatch:

                    if current_answer_q is not None:
                        answers[current_answer_q] = current_answer_tokens

                    current_answer_q      = amatch.group(1)
                    cleaned               = re.sub(r'^\d+\s*', '', text)
                    current_answer_tokens = []

                    if cleaned:
                        current_answer_tokens.append(("text", cleaned))

                else:

                    if current_answer_q is not None:
                        current_answer_tokens.append(("text", text))

            continue

        if item_type == "image":
            filename = save_image(doc, value)
            current_tokens.append(("image", filename))
            continue

        if item_type == "table":
            rows = value
            if not rows:
                continue

            first_cell = rows[0][0].strip() if rows[0] else ''
            detected_qnum = parse_qnum(first_cell)

            # helper: flush current question into appropriate bucket
            def flush_current():
                if current_qnum is not None:
                    if options:
                        mcq_questions.append({
                            "qnum":    current_qnum,
                            "tokens":  current_tokens,
                            "options": options
                        })
                    else:
                        structured_questions.append({
                            "qnum":   current_qnum,
                            "tokens": current_tokens
                        })

            # ── Question-start table: first cell is a question number ──────
            if detected_qnum is not None:

                flush_current()

                current_qnum   = detected_qnum
                current_tokens = []
                options        = []

                # Question text from remaining cells of row 0 (dedup merged cells)
                q_text = dedup_cells(rows[0][1:])
                if q_text:
                    current_tokens.append(("text", q_text))

                # Process additional rows in the same table
                for row in rows[1:]:
                    if not row:
                        continue
                    row0     = row[0].strip()
                    row0_plain = re.sub(r'<[^>]+>', '', row0).strip()  # strip HTML for detection
                    sub_qnum = parse_qnum(row0)

                    if sub_qnum is not None:
                        # ── A new question starts inside this same table ──
                        # (e.g. Q25 embedded in Q24's big table)
                        flush_current()
                        current_qnum   = sub_qnum
                        current_tokens = []
                        options        = []
                        q_text = dedup_cells(row[1:])
                        if q_text:
                            current_tokens.append(("text", q_text))

                    elif OPT_LABEL_RE.match(row0_plain):
                        # FIX: MCQ option row where the letter IS the first cell
                        # e.g. ['A', '10–6', '10–9', '10–12']  (Q1 format)
                        # or   ['A', 'I1',   'I4R3']            (Q24 format)
                        opt_text = dedup_cells(row[1:])
                        if opt_text:
                            options.append(opt_text)
                        # (images in option rows are already queued by iter_block_items)

                    elif not row0:
                        # Candidate MCQ options row: ['', 'A text', 'B text', ...]
                        opts = extract_opts_from_row(row)
                        if len(opts) >= 2:
                            options.extend(opts)
                        else:
                            # Not recognisable options — treat as extra question text
                            extra = dedup_cells(row)
                            if extra:
                                current_tokens.append(("text", extra))
                    else:
                        # Sub-part continuation inside the same table
                        extra = dedup_cells(row)
                        if extra:
                            current_tokens.append(("text", extra))

            # ── Continuation table: first cell is not a question number ────
            elif current_qnum is not None:
                for row in rows:
                    row_text = dedup_cells(row)
                    if row_text:
                        current_tokens.append(("text", row_text))

            continue

        if item_type == "text":

            text = value.replace('\xa0', ' ').replace('\t', ' ')
            text = re.sub(r'\s+', ' ', text).strip()

            if text.upper() == "ANSWERS":
                reading_answers = True
                continue

            if not text:
                continue

            qmatch = question_regex.match(text)

            if qmatch and int(qmatch.group(1)) <= MAX_QNUM:

                if current_qnum is not None:

                    if options:
                        mcq_questions.append({
                            "qnum":    current_qnum,
                            "tokens":  current_tokens,
                            "options": options
                        })
                    else:
                        structured_questions.append({
                            "qnum":   current_qnum,
                            "tokens": current_tokens
                        })

                current_qnum   = qmatch.group(1)
                current_tokens = [("text", qmatch.group(2))]
                options        = []
                continue

            opt = option_regex.match(text)

            if opt:
                options.append(opt.group(1))
                continue

            current_tokens.append(("text", text))

    if current_answer_q is not None:
        answers[current_answer_q] = current_answer_tokens

    if current_qnum is not None:
        if options:
            mcq_questions.append({
                "qnum":    current_qnum,
                "tokens":  current_tokens,
                "options": options
            })
        else:
            structured_questions.append({
                "qnum":   current_qnum,
                "tokens": current_tokens
            })

    # ─────────────────────────────────────────────────────────────────
    # HELPER: resolve answer for a question number
    # Priority: ms_answers (external MS file) > inline answers section
    # ─────────────────────────────────────────────────────────────────
    def resolve_answer(qnum, is_mcq=False):
        """Return answer string for the given question number."""
        # External mark-scheme takes priority
        if qnum in ms_answers:
            return ms_answers[qnum]
        # Inline ANSWERS section
        if qnum in answers:
            return answer_tokens_to_string(answers[qnum])
        return "None"

    # -----------------------------
    # QTI 2.1 OUTPUT GENERATION
    # -----------------------------
    import zipfile
    import random
    import string
    import xml.sax.saxutils as saxutils


    os.makedirs(items_dir, exist_ok=True)

    def random_id(length=3):
        return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

    item_refs   = []
    item_images = {}

    # -----------------------------
    # HELPER: convert tokens to html
    # -----------------------------
    def tokens_to_html(tokens):

        html = ""

        for ttype, val in tokens:

            if ttype == "text":

                safe  = saxutils.escape(val)
                safe  = safe.replace("&lt;", "<").replace("&gt;", ">")
                html += f"<div>{safe}</div>\n"

            elif ttype == "image":

                html += f'''
<div>
<img alt="diagram" src="../assets/{val}"/>
</div>
'''

            elif ttype == "table":

                html += "<table border='1'>\n"
                for row in val:
                    html += "<tr>"
                    for cell in row:
                        safe  = saxutils.escape(cell)
                        safe  = safe.replace("&lt;", "<").replace("&gt;", ">")
                        html += f"<td>{safe}</td>"
                    html += "</tr>\n"
                html += "</table>\n"

        return html

    # -----------------------------
    # CONVERT ANSWER TOKENS TO STRING
    # -----------------------------
    def answer_tokens_to_string(tokens):

        result = ""

        for ttype, val in tokens:
            if ttype == "text":
                result += val + " "
            elif ttype == "image":
                result += f"[image:{val}] "
            elif ttype == "table":
                result += "[table] "

        return result.strip()

    letters = ["a", "b", "c", "d", "e", "f"]

    # -----------------------------
    # MCQ QUESTIONS → QTI
    # -----------------------------
    for q in mcq_questions:

        answer_text = resolve_answer(q["qnum"], is_mcq=True)

        item_id  = f"Q{int(q['qnum']):03d}_{random_id()}"
        filename = os.path.join(items_dir, f"{item_id}.xml")

        stem_html = tokens_to_html(q["tokens"])

        title_text = ""
        for ttype, val in q["tokens"]:
            if ttype == "text":
                title_text = val
                break

        # Strip HTML tags for the XML title attribute (plain text only)
        title_plain = re.sub(r'<[^>]+>', '', title_text)[:70]
        safe_title = saxutils.escape(title_plain)

        xml = f'''<assessmentItem xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
adaptive="false"
identifier="{item_id}"
timeDependent="false"
title="Q{int(q['qnum']):03d} {safe_title}"
xsi:schemaLocation="http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_v2p1.xsd">
<responseDeclaration identifier="RESPONSE" cardinality="single" baseType="identifier">
<correctResponse>
<value>{saxutils.escape(answer_text)}</value>
</correctResponse>
</responseDeclaration>
<outcomeDeclaration baseType="float" cardinality="single" identifier="SCORE"/>
<itemBody>
<choiceInteraction responseIdentifier="RESPONSE" shuffle="true" maxChoices="1">
<prompt>
<div>
{stem_html}
</div>
</prompt>
'''

        for i, opt in enumerate(q["options"]):
            if i >= len(letters):
                break
            safe = saxutils.escape(opt)
            safe = safe.replace("&lt;", "<").replace("&gt;", ">")
            xml += f'<simpleChoice identifier="{letters[i]}">{safe}</simpleChoice>\n'

        xml += '''
</choiceInteraction>
</itemBody>
<responseProcessing template="http://www.imsglobal.org/question/qti_v2p1/rptemplates/match_correct"/>
</assessmentItem>
'''

        with open(filename, "w", encoding="utf-8") as f:
            f.write(xml)

        item_refs.append(item_id)
        imgs = [v for t, v in q["tokens"] if t == "image"]
        item_images[item_id] = imgs

    # -----------------------------
    # STRUCTURED QUESTIONS → QTI
    # -----------------------------
    for q in structured_questions:

        answer_text = resolve_answer(q["qnum"])

        item_id  = f"Q{int(q['qnum']):03d}_{random_id()}"
        filename = os.path.join(items_dir, f"{item_id}.xml")

        stem_html = tokens_to_html(q["tokens"])

        xml = f'''<assessmentItem xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
adaptive="false"
identifier="{item_id}"
timeDependent="false"
title="Q{int(q['qnum']):03d}"
xsi:schemaLocation="http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_v2p1.xsd">
<responseDeclaration identifier="RESPONSE" cardinality="single" baseType="string"/>
<correctResponse>
<value>{saxutils.escape(answer_text)}</value>
</correctResponse>
<outcomeDeclaration baseType="float" cardinality="single" identifier="SCORE"/>
<itemBody>
<prompt>
{stem_html}
</prompt>
<extendedTextInteraction responseIdentifier="RESPONSE" expectedLength="400"/>
</itemBody>
<responseProcessing template="http://www.imsglobal.org/question/qti_v2p1/rptemplates/match_correct"/>
</assessmentItem>
'''

        with open(filename, "w", encoding="utf-8") as f:
            f.write(xml)

        item_refs.append(item_id)
        imgs = [v for t, v in q["tokens"] if t == "image"]
        item_images[item_id] = imgs

    # -----------------------------
    # assessment_test.xml
    # -----------------------------
    assessment_xml = '''
<assessmentTest xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
identifier="TEST1" title="Converted Test" xsi:schemaLocation="http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_v2p1.xsd">
<testPart identifier="part1" navigationMode="linear" submissionMode="individual">
<assessmentSection identifier="section1" title="Converted Test" visible="true">
'''

    for item in item_refs:
        assessment_xml += f'<assessmentItemRef href="items/{item}.xml" identifier="{item}"/>\n'

    assessment_xml += '''
</assessmentSection>
</testPart>
</assessmentTest>
'''

    with open(os.path.join(job_dir, "assessment_test.xml"), "w") as f:
        f.write(assessment_xml)

    # -----------------------------
    # imsmanifest.xml
    # -----------------------------
    manifest = '''
<manifest
xmlns="http://www.imsglobal.org/xsd/imscp_v1p1"
xmlns:imsmd="http://www.imsglobal.org/xsd/imsmd_v1p2"
xmlns:imsqti="http://www.imsglobal.org/xsd/imsqti_metadata_v2p1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
identifier="MANIFEST1"
xsi:schemaLocation="
http://www.imsglobal.org/xsd/imscp_v1p1 http://www.imsglobal.org/xsd/imscp_v1p1.xsd
http://www.imsglobal.org/xsd/imsmd_v1p2 http://www.imsglobal.org/xsd/imsmd_v1p2p4.xsd
http://www.imsglobal.org/xsd/imsqti_metadata_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_metadata_v2p1.xsd">
<metadata>
<schema>QTIv2.1 Package</schema>
<schemaversion>1.0.0</schemaversion>
</metadata>
<organizations/>
<resources>
'''

    manifest += '''
<resource identifier="RES_TEST" type="imsqti_test_xmlv2p1" href="assessment_test.xml">
<metadata/>
<file href="assessment_test.xml"/>
</resource>
'''

    for item in item_refs:

        manifest += f'''
<resource identifier="RES_{item}" type="imsqti_item_xmlv2p1" href="items/{item}.xml">
<metadata/>
<file href="items/{item}.xml"/>
'''

        for img in item_images.get(item, []):
            manifest += f'    <file href="assets/{img}"/>\n'

        manifest += "</resource>\n"

    manifest += '''
</resources>
</manifest>
'''

    with open(os.path.join(job_dir, "imsmanifest.xml"), "w", encoding="utf-8") as f:
        f.write(manifest)

    # -----------------------------
    # ZIP QTI PACKAGE
    # -----------------------------
    zip_path = os.path.join(job_dir, "qti_package.zip")
    zipf     = zipfile.ZipFile(zip_path, "w")

    for root, dirs, files in os.walk(job_dir):
        for file in files:
            if file.endswith((".zip", ".docx")):
                continue
            path    = os.path.join(root, file)
            arcname = os.path.relpath(path, job_dir)
            zipf.write(path, arcname)

    zipf.close()

    print(f"QTI 2.1 package created: {zip_path}")
    return zip_path
