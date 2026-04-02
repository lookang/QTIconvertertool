"""
QTI 1.2 Package Generator
Produces an IMS QTI 1.2 ZIP package importable into Canvas, Moodle, etc.

ZIP structure:
  imsmanifest.xml
  assessment.xml
"""

import io
import re
import uuid
import zipfile
from xml.dom import minidom
import xml.etree.ElementTree as ET


# ── Namespace helpers ────────────────────────────────────────────────────────

QTI_NS  = 'http://www.imsglobal.org/xsd/ims_qtiasiv1p2'
CP_NS   = 'http://www.imsglobal.org/xsd/imscp_v1p1'

ET.register_namespace('',    QTI_NS)
ET.register_namespace('cp',  CP_NS)


def _prettify(root):
    """Return indented XML string from an ElementTree Element."""
    rough = ET.tostring(root, encoding='unicode', xml_declaration=False)
    reparsed = minidom.parseString(rough)
    return reparsed.toprettyxml(indent='  ', encoding=None)


# ── Item builders ────────────────────────────────────────────────────────────

def _mcq_item(question):
    """Build an <item> element for a multiple-choice question."""
    item = ET.Element('item')
    item.set('ident',  f"item_{question['number']}")
    item.set('title',  f"Question {question['number']}")

    # Metadata
    meta = ET.SubElement(item, 'itemmetadata')
    qmeta = ET.SubElement(meta, 'qtimetadata')
    field = ET.SubElement(qmeta, 'qtimetadatafield')
    ET.SubElement(field, 'fieldlabel').text = 'question_type'
    ET.SubElement(field, 'fieldentry').text = 'multiple_choice_question'

    # Presentation
    presentation = ET.SubElement(item, 'presentation')
    mat  = ET.SubElement(presentation, 'material')
    mtext = ET.SubElement(mat, 'mattext')
    mtext.set('texttype', 'text/plain')
    mtext.text = question['text']

    resp = ET.SubElement(presentation, 'response_lid')
    resp.set('ident',        f"response_{question['number']}")
    resp.set('rcardinality', 'Single')

    render = ET.SubElement(resp, 'render_choice')
    for key in ['A', 'B', 'C', 'D']:
        if key not in question['options']:
            continue
        label = ET.SubElement(render, 'response_label')
        label.set('ident', key)
        lmat  = ET.SubElement(label, 'material')
        ltext = ET.SubElement(lmat, 'mattext')
        ltext.set('texttype', 'text/plain')
        ltext.text = question['options'][key]

    # Response processing (scoring)
    resprocessing = ET.SubElement(item, 'resprocessing')
    outcomes = ET.SubElement(resprocessing, 'outcomes')
    decvar = ET.SubElement(outcomes, 'decvar')
    decvar.set('maxvalue', '100')
    decvar.set('minvalue', '0')
    decvar.set('varname',  'SCORE')
    decvar.set('vartype',  'Decimal')

    answer = question.get('answer', '')
    if answer and re.match(r'^[A-D]$', answer.strip().upper()):
        cond = ET.SubElement(resprocessing, 'respcondition')
        cond.set('continue', 'No')
        condvar = ET.SubElement(cond, 'conditionvar')
        varequal = ET.SubElement(condvar, 'varequal')
        varequal.set('respident', f"response_{question['number']}")
        varequal.text = answer.strip().upper()
        setvar = ET.SubElement(cond, 'setvar')
        setvar.set('action',  'Set')
        setvar.set('varname', 'SCORE')
        setvar.text = '100'

    return item


def _essay_item(question):
    """Build an <item> element for an open-ended / structured question."""
    item = ET.Element('item')
    item.set('ident',  f"item_{question['number']}")
    item.set('title',  f"Question {question['number']}")

    # Metadata
    meta = ET.SubElement(item, 'itemmetadata')
    qmeta = ET.SubElement(meta, 'qtimetadata')
    field = ET.SubElement(qmeta, 'qtimetadatafield')
    ET.SubElement(field, 'fieldlabel').text = 'question_type'
    ET.SubElement(field, 'fieldentry').text = 'essay_question'

    # Presentation
    presentation = ET.SubElement(item, 'presentation')
    mat  = ET.SubElement(presentation, 'material')
    mtext = ET.SubElement(mat, 'mattext')
    mtext.set('texttype', 'text/plain')
    mtext.text = question['text']

    resp = ET.SubElement(presentation, 'response_str')
    resp.set('ident',        f"response_{question['number']}")
    resp.set('rcardinality', 'Single')

    render = ET.SubElement(resp, 'render_fib')
    render.set('rows',    '10')
    render.set('columns', '60')

    # Model answer (if provided)
    if question.get('answer'):
        resprocessing = ET.SubElement(item, 'resprocessing')
        outcomes = ET.SubElement(resprocessing, 'outcomes')
        decvar = ET.SubElement(outcomes, 'decvar')
        decvar.set('maxvalue', '100')
        decvar.set('minvalue', '0')
        decvar.set('varname',  'SCORE')
        decvar.set('vartype',  'Decimal')

        # Store model answer as an itemfeedback element
        feedback = ET.SubElement(item, 'itemfeedback')
        feedback.set('ident', 'model_answer')
        flow_mat = ET.SubElement(feedback, 'flow_mat')
        fmat  = ET.SubElement(flow_mat, 'material')
        ftext = ET.SubElement(fmat, 'mattext')
        ftext.set('texttype', 'text/plain')
        ftext.text = question['answer']

    return item


# ── Assessment XML ───────────────────────────────────────────────────────────

def _build_assessment_xml(questions, assessment_id, title):
    root = ET.Element('questestinterop')
    root.set('xmlns', QTI_NS)
    root.set('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')
    root.set('xsi:schemaLocation',
             f'{QTI_NS} http://www.imsglobal.org/xsd/ims_qtiasiv1p2.xsd')

    assessment = ET.SubElement(root, 'assessment')
    assessment.set('ident',  assessment_id)
    assessment.set('title',  title)

    # Assessment metadata
    qmeta = ET.SubElement(assessment, 'qtimetadata')
    for label, entry in [
        ('cc_maxattempts',    '1'),
        ('qmd_assessmenttype', 'Examination'),
    ]:
        field = ET.SubElement(qmeta, 'qtimetadatafield')
        ET.SubElement(field, 'fieldlabel').text = label
        ET.SubElement(field, 'fieldentry').text  = entry

    section = ET.SubElement(assessment, 'section')
    section.set('ident', 'root_section')
    section.set('title', 'Root Section')

    for q in questions:
        if q['type'] == 'mcq':
            section.append(_mcq_item(q))
        else:
            section.append(_essay_item(q))

    return _prettify(root)


# ── Manifest XML ─────────────────────────────────────────────────────────────

def _build_manifest_xml(assessment_id, manifest_id, title):
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<manifest identifier="{manifest_id}"
          xmlns="http://www.imsglobal.org/xsd/imscp_v1p1"
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
          xsi:schemaLocation="http://www.imsglobal.org/xsd/imscp_v1p1 http://www.imsglobal.org/xsd/imscp_v1p1.xsd">
  <metadata>
    <schema>IMS Content</schema>
    <schemaversion>1.1.3</schemaversion>
    <title>{_xml_escape(title)}</title>
  </metadata>
  <organizations/>
  <resources>
    <resource identifier="{assessment_id}" type="imsqti_xmlv1p2">
      <file href="assessment.xml"/>
    </resource>
  </resources>
</manifest>
'''


def _xml_escape(text):
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


# ── Public entry point ───────────────────────────────────────────────────────

def generate_qti_zip(questions, title='Assessment'):
    """
    Generate a QTI 1.2 ZIP package in memory.
    Returns a BytesIO object ready to send as a file download.
    """
    assessment_id = 'assessment_' + uuid.uuid4().hex[:8]
    manifest_id   = 'manifest_'   + uuid.uuid4().hex[:8]

    assessment_xml = _build_assessment_xml(questions, assessment_id, title)
    manifest_xml   = _build_manifest_xml(assessment_id, manifest_id, title)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('imsmanifest.xml', manifest_xml)
        zf.writestr('assessment.xml',  assessment_xml)
    buf.seek(0)
    return buf
