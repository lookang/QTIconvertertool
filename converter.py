from docx import Document
from docx.oxml.ns import qn
import os
import re


# ─────────────────────────────────────────────────────────────────────────────
# MARK-SCHEME PARSER
# ─────────────────────────────────────────────────────────────────────────────

def parse_mark_scheme(ms_docx_path):
    """
    Parse a mark scheme document. Returns a dictionary mapping
    question numbers (strings) to answer text or options.
    It scans tables for both row-style and col-style MCQ grids,
    and scans paragraphs for both MCQ and essay answers.
    """
    import re
    doc = Document(ms_docx_path)
    answers = {}

    # 1. Parse tables for MCQ grid and structured
    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                ctext = cell.text.strip().upper()
                mQ = re.match(r'^Q?0*(\d+)$', ctext)
                if mQ:
                    qnum_str = mQ.group(1)
                    found_ans = False
                    
                    # Check right cell (column-style)
                    if c_idx + 1 < len(row.cells):
                        right = row.cells[c_idx+1].text.strip()
                        right_upper = right.upper()
                        if re.match(r'^[ABCD]$', right_upper):
                            answers[qnum_str] = right.lower()
                            found_ans = True
                        elif re.match(r'^[1234]$', right_upper):
                            answers[qnum_str] = chr(96 + int(right))
                            found_ans = True
                        elif right:
                            if qnum_str in answers and len(answers[qnum_str]) > 1:
                                answers[qnum_str] += '\n' + right
                            else:
                                answers[qnum_str] = right
                            found_ans = True
                    
                    # Check bottom cell (row-style)
                    if not found_ans and r_idx + 1 < len(table.rows):
                        bottom = table.rows[r_idx+1].cells[c_idx].text.strip().upper()
                        if re.match(r'^[ABCD]$', bottom):
                            answers[qnum_str] = bottom.lower()
                        elif re.match(r'^[1234]$', bottom):
                            answers[qnum_str] = chr(96 + int(bottom))
                else:
                    struct_match = re.match(r'^[Qq]?0*(\d+)[.:：]\s*(.*)', cell.text)
                    if struct_match:
                        qnum_str = struct_match.group(1)
                        text = struct_match.group(2).strip()
                        if text:
                            answers[qnum_str] = text

    # 2. Parse paragraphs for answers (Essay style & list style)
    current_qnum = None
    for para in doc.paragraphs:
        pt = para.text.strip()
        if not pt:
            continue
            
        qnum = None
        ans_text = ""
        m1 = re.match(r'^[Qq]0*(\d+)[.:：\s]*(.*)', pt)
        m2 = re.match(r'^0*(\d+)[.:：]+\s*(.*)', pt)
        if m1:
            qnum = m1.group(1)
            ans_text = m1.group(2).strip()
        elif m2:
            qnum = m2.group(1)
            ans_text = m2.group(2).strip()

        if qnum:
            is_mcq = False
            clean_ans = ans_text.upper().replace('.', '').strip()
            if re.match(r'^[ABCD]$', clean_ans):
                ans_text = clean_ans.lower()
                is_mcq = True
            elif re.match(r'^[1234]$', clean_ans):
                ans_text = chr(96 + int(clean_ans))
                is_mcq = True

            if qnum in answers and re.match(r'^[a-d]$', answers[qnum]):
                if is_mcq:
                    answers[qnum] = ans_text
                current_qnum = None
                continue
                
            if qnum in answers and len(answers[qnum]) > 1 and not is_mcq and answers[qnum] != ans_text:
                if ans_text not in answers[qnum]:
                    answers[qnum] += '\n' + pt
            else:
                answers[qnum] = ans_text or pt
            
            current_qnum = None if is_mcq else qnum
        else:
            if current_qnum and current_qnum in answers:
                answers[current_qnum] += '\n' + pt

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

    # -----------------------------
    # PRE-PROCESS HANYU PINYIN (W:RUBY)
    # -----------------------------
    import html
    from docx.oxml import OxmlElement
    for ruby in doc.element.xpath('.//*[local-name()="ruby"]'):
        rt = ruby.xpath('.//*[local-name()="rt"]')
        base = ruby.xpath('.//*[local-name()="rubyBase"]')
        # Extract text from descendants
        rt_text = "".join(t.text or "" for r in rt for t in r.xpath('.//*[local-name()="t"]')) if rt else ""
        base_text = "".join(t.text or "" for r in base for t in r.xpath('.//*[local-name()="t"]')) if base else ""
        
        rt_safe = html.escape(rt_text)
        base_safe = html.escape(base_text)
        
        fake_r = OxmlElement('w:r')
        fake_t = OxmlElement('w:t')
        fake_t.text = f"<ruby>{base_safe}<rp>(</rp><rt>{rt_safe}</rt><rp>)</rp></ruby>"
        fake_r.append(fake_t)
        
        ruby.addprevious(fake_r)
        ruby.getparent().remove(ruby)

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
            
        rPr = run.find(qn('w:rPr'))
        if rPr is not None:
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is not None:
                ascii_font = rFonts.get(qn('w:ascii'))
                if ascii_font and ascii_font.lower() == 'hypy':
                    mapping = {'C': 'ǎ', 'H': 'ǐ', 'K': 'ē', 'N': 'è', 'P': 'ē', 'Q': 'ó', 'S': 'è', 'A': 'ā', 'B': 'á', 'D': 'à', 'E': 'ō', 'F': 'ó', 'G': 'ǒ', 'I': 'ī', 'J': 'í', 'L': 'ì', 'M': 'ū', 'O': 'ǔ', 'R': 'ú', 'T': 'ù', 'U': 'ǖ', 'V': 'ǘ', 'W': 'ǚ', 'X': 'ǜ', 'Y': 'ě', 'Z': 'ě'}
                    text = "".join(mapping.get(c, c) for c in text)

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

    def halve_if_doubled(s):
        """Recursively halve a string that is an exact repeat of itself.
        Fixes Word XML duplicated diagram labels: 'cablecable'→'cable',
        'tension in chain'×4 → 'tension in chain'×2 → 'tension in chain'."""
        n = len(s)
        if n >= 10 and n % 2 == 0:
            mid = n >> 1
            if s[:mid] == s[mid:]:
                return halve_if_doubled(s[:mid])   # recurse for ×4, ×8 …
        return s

    def extract_cell_text(cell):
        """Extract formatted text from DIRECT paragraphs of a table cell only.
        • Does NOT descend into nested <w:tbl> elements.
        • Paragraphs are joined with a space so multi-line option cells like
          ['A', 'option text'] become 'A option text' for OPT_CELL_RE matching.
        • Detects text doubling via the HTML-stripped plain text and replaces
          with the deduplicated plain version.
        • Deduplicates across paragraphs within the cell (removes cross-paragraph
          repetition from Word merged-cell artefacts)."""
        seen_paras = set()
        paras = []
        for para in cell.iterchildren(qn('w:p')):
            runs  = "".join(extract_run_text(r) for r in para.xpath(".//w:r"))
            runs  = merge_adjacent_tags(runs).strip()
            # Check via HTML-stripped text so '<strong>X</strong><strong>X</strong>'
            # → merged '<strong>XX</strong>' is still detected as doubled.
            plain    = re.sub(r'<[^>]+>', '', runs)
            deduped  = halve_if_doubled(plain)
            if deduped != plain:
                runs = deduped  # use plain deduplicated text (HTML stripped)
            runs = runs.strip()
            # Cross-paragraph dedup: skip paragraphs whose text we've already seen.
            if runs and runs not in seen_paras:
                seen_paras.add(runs)
                paras.append(runs)
        return " ".join(paras)

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

                    for cell in row.iterchildren(qn('w:tc')):
                        # Extract text from direct paragraphs only (no nested table content)
                        row_data.append(extract_cell_text(cell))

                        # Collect images from ALL descendants of this cell
                        # (includes images in nested tables — images from nested option rows
                        #  are also picked up when we process those inner rows directly,
                        #  but seen_image_rids prevents duplicates)
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
    question_regex  = re.compile(r'^(?:[Qq](\d+)(?:[\.\)\s:]*\s+)?|\((\d+)\)|（(\d+)）|(\d+)(?:[\.\)]\s+|\s+))(.*)')
    option_regex    = re.compile(r'^(?:[A-H][\.\)]\s*|[\u2474-\u247B]\s*|[\u2460-\u2467]\s*|[（(]\s*[1-8]\s*[）)]\s*|[1-8][\.\)．]\s*)(.*)')
    OPTIONS_HEADER_RE = re.compile(r'^(?:options|choices)(?:\s*[:：])?$', re.IGNORECASE)
    STANDALONE_OPTION_LABEL_RE = re.compile(r'^[A-D]$', re.IGNORECASE)
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

    # Pattern to detect labelled question tables (notes-style docs)
    # e.g. "Planning Question (Example 1)", "Practice Question", "Additional question 2"
    labelled_q_regex = re.compile(
        r'^(planning question|practice question|additional question)',
        re.IGNORECASE
    )
    labelled_q_counter = [0]   # mutable so nested flush_current can read it

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

            # ── Labelled question table (notes-style docs) ────────────────
            # e.g. "Planning Question (Example 1)", "Practice Question"
            # Strip only first line (the label) and treat rest as question text.
            first_cell_plain = re.sub(r'<[^>]+>', '', first_cell).strip().split('\n')[0].strip()
            is_labelled_q = labelled_q_regex.match(first_cell_plain)

            if is_labelled_q and detected_qnum is None:
                flush_current()
                labelled_q_counter[0] += 1
                current_qnum = str(labelled_q_counter[0])
                current_tokens = []
                options = []
                # The full first-cell text (after stripping the label line) is the question body
                # We keep everything — the converter will include the label too which helps context
                full_text = first_cell.replace('\r', '').strip()
                if full_text:
                    current_tokens.append(("text", full_text))
                # Additional rows in same table that are NOT answer rows go in too
                ANSWER_LABELS = re.compile(
                    r'^(suggested answer|marking scheme|possible answer)',
                    re.IGNORECASE
                )
                for row in rows[1:]:
                    row0 = row[0].strip() if row else ''
                    row0_plain = re.sub(r'<[^>]+>', '', row0).strip().split('\n')[0]
                    if ANSWER_LABELS.match(row0_plain):
                        break   # Stop before model answers
                    extra = dedup_cells(row)
                    if extra:
                        current_tokens.append(("text", extra))

            # ── Question-start table: first cell is a question number ──────
            elif detected_qnum is not None:

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
            qnum_str = qmatch.group(1) or qmatch.group(2) or qmatch.group(3) or qmatch.group(4) if qmatch else None
            opt = option_regex.match(text)

            if current_qnum is not None and opt:
                options.append(opt.group(1))
                continue

            if qmatch and qnum_str and int(qnum_str) <= MAX_QNUM:

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

                current_qnum   = qnum_str
                current_tokens = [("text", qmatch.group(5) or "")]
                options        = []
                continue

            if OPTIONS_HEADER_RE.match(text):
                continue
            if current_qnum is not None and options and STANDALONE_OPTION_LABEL_RE.match(text):
                continue

            range_match = re.search(r'[Qq]\d+\s*[-－到至\u2013\u2014]\s*[Qq]\d+', text)
            inline_qmatch = re.search(r'(?:^|[^A-Za-z0-9])[Qq](\d{1,3})(?:[^A-Za-z0-9]|$)', text)
            
            if inline_qmatch and not range_match and int(inline_qmatch.group(1)) <= MAX_QNUM:
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
                current_qnum = inline_qmatch.group(1)
                current_tokens = [("text", text)]
                options = []
                
                cloze_match = re.search(r'[Qq]\d{1,3}\s*[（\(](.*?)[）\)]', text)
                if cloze_match:
                    inner_txt = cloze_match.group(1).strip()
                    # Split horizontally using python regex
                    parts = [p.strip() for p in re.split(r'(?:^|\s+)[a-hA-H1-8][\.\)\s]+', inner_txt) if p.strip()]
                    if 2 <= len(parts) <= 8:
                        options = parts
                continue

            if opt and current_qnum is not None:
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

                safe_val = val if val is not None else ""
                safe  = saxutils.escape(safe_val)
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

    # ---------------------------------------------------------------
    # MERGE + DEDUPLICATE + SORT all questions before generating XML.
    #
    # Problem: iterating mcq_questions first then structured_questions
    # produces item_refs in two unordered batches, so SLS imports them
    # in wrong order (Q5 shows as SLS question 22, etc.).
    # Additionally, some papers produce a false Q1 from cover-page
    # text ("1 hour …") giving two Q001 entries.
    #
    # Fix: build a dict keyed by qnum — MCQ entries win over structured
    # (real MCQ questions have options; cover-page false positives do not).
    # Then sort numerically so assessment_test.xml lists Q001, Q002 … in order.
    # ---------------------------------------------------------------
    q_by_num = {}   # qnum → ('mcq'|'structured', q_dict)
    for q in mcq_questions:
        q_by_num[q['qnum']] = ('mcq', q)
    for q in structured_questions:
        if q['qnum'] not in q_by_num:          # don't overwrite real MCQ with cover-page structured
            q_by_num[q['qnum']] = ('structured', q)

    all_questions = sorted(q_by_num.values(), key=lambda p: int(p[1]['qnum']))

    # -----------------------------
    # ALL QUESTIONS → QTI (in numerical order)
    # -----------------------------
    for qtype, q in all_questions:

        if qtype == 'mcq':

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
<choiceInteraction responseIdentifier="RESPONSE" shuffle="false" maxChoices="1">
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

        else:  # structured

            answer_text = resolve_answer(q["qnum"])

            item_id  = f"Q{int(q['qnum']):03d}_{random_id()}"
            filename = os.path.join(items_dir, f"{item_id}.xml")

            stem_html = tokens_to_html(q["tokens"])

            # SLS-compatible open-ended/essay item:
            # - responseDeclaration is self-closing (no correctResponse)
            # - Question HTML placed directly in itemBody as <div> blocks
            # - <prompt> goes INSIDE <extendedTextInteraction>
            # - No <responseProcessing> — teacher marks manually in SLS
            qnum_str = f"Q{int(q['qnum']):03d}"
            xml = (
                f'<assessmentItem xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"\n'
                f'  adaptive="false" identifier="{item_id}" timeDependent="false" title="{qnum_str}"\n'
                f'  xsi:schemaLocation="http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_v2p1.xsd">\n'
                '<responseDeclaration identifier="RESPONSE" cardinality="single" baseType="string">\n'
                f'  <correctResponse><value>{saxutils.escape(answer_text)}</value></correctResponse>\n'
                '</responseDeclaration>\n'
                '<outcomeDeclaration baseType="float" cardinality="single" identifier="SCORE"/>\n'
                '<outcomeDeclaration identifier="FEEDBACK" cardinality="single" baseType="identifier"/>\n'
                '<itemBody>\n'
                f'  <rubricBlock view="scorer"><div>{saxutils.escape(answer_text)}</div></rubricBlock>\n'
                '  <prompt>\n'
                f'{stem_html}'
                '  </prompt>\n'
                '  <extendedTextInteraction responseIdentifier="RESPONSE" expectedLength="400"/>\n'
                '</itemBody>\n'
                '<modalFeedback outcomeIdentifier="FEEDBACK" showHide="show" identifier="modelAnswer" title="Suggested Answer">\n'
                f'  <div>{saxutils.escape(answer_text)}</div>\n'
                '</modalFeedback>\n'
                '<responseProcessing>\n'
                '  <setOutcomeValue identifier="FEEDBACK">\n'
                '    <baseValue baseType="identifier">modelAnswer</baseValue>\n'
                '  </setOutcomeValue>\n'
                '</responseProcessing>\n'
                '</assessmentItem>\n'
            )

        with open(filename, "w", encoding="utf-8") as f:
            f.write(xml)

        item_refs.append(item_id)
        imgs = [v for t, v in q["tokens"] if t == "image"]
        item_images[item_id] = imgs

    # -----------------------------
    # assessment_test.xml
    # -----------------------------
    assessment_xml = '''<assessmentTest xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  identifier="TEST1" title="Converted Test"
  xsi:schemaLocation="http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_v2p1.xsd">
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
    manifest = '''<manifest xmlns="http://www.imsglobal.org/xsd/imscp_v1p1"
  xmlns:imsmd="http://www.imsglobal.org/xsd/imsmd_v1p2"
  xmlns:imsqti="http://www.imsglobal.org/xsd/imsqti_metadata_v2p1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  identifier="MANIFEST1"
  xsi:schemaLocation="http://www.imsglobal.org/xsd/imscp_v1p1 http://www.imsglobal.org/xsd/imscp_v1p1.xsd">
<metadata><schema>QTIv2.1 Package</schema><schemaversion>1.0.0</schemaversion></metadata>
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
