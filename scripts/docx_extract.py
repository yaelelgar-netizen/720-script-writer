# -*- coding: utf-8 -*-
"""
Math-aware text extraction from the "תבנית פיתוח תוכן ליעד" Word document.

WHY THIS EXISTS (construction-checklist.md item 13): feedback cells in the
content doc contain Office Math (OMML, m:oMath) equations. python-docx's
cell.text / paragraph.text silently DROPS them, leaving lines like
"18% מס הם בשקלים: " with the actual computation missing — which previously
led to improvised/recomputed feedback content on slides. That is forbidden:
the feedback is authored pedagogy and must be copied VERBATIM. Always
extract content cells through cell_text()/paragraph_text() here, never
through .text.

OMML handling: fractions m:f become "num/den"; all other math text is
concatenated in document order (which also reunites digits that RTL layout
splits across the math boundary). A minimal RTL-artifact fix reassembles
letter-split Hebrew unit words (e.g. the tokens ח, ", ש render as ש"ח).
Numbers, operators and computation steps are never altered.
"""

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _omml_to_text(el):
    tag = el.tag
    if tag == "{%s}f" % M_NS:  # fraction
        num = el.find("{%s}num" % M_NS)
        den = el.find("{%s}den" % M_NS)
        return "%s/%s" % (_omml_join(num), _omml_join(den))
    if tag == "{%s}t" % M_NS:
        return el.text or ""
    return _omml_join(el)


def _omml_join(el):
    if el is None:
        return ""
    return "".join(_omml_to_text(c) for c in el)


def _fix_rtl_units(text):
    """Reassemble Hebrew unit words that OMML stores letter-by-letter in
    visual (reversed) order — e.g. ח + " + ש  ->  ש"ח. Only whitespace and
    letter order of the UNIT WORD is touched; numbers are never modified."""
    import re
    text = re.sub(r'ח\s*"\s*ש', 'ש"ח', text)
    return re.sub(r"[ \t]{2,}", " ", text)


def _walk(el, out):
    for child in el:
        tag = child.tag
        if tag == "{%s}t" % W_NS:
            out.append(child.text or "")
        elif tag in ("{%s}oMath" % M_NS, "{%s}oMathPara" % M_NS):
            out.append(_fix_rtl_units(_omml_to_text(child)))
        elif tag in ("{%s}br" % W_NS, "{%s}cr" % W_NS):
            out.append("\n")
        elif tag == "{%s}tab" % W_NS:
            out.append("\t")
        else:
            _walk(child, out)


def paragraph_text(p_el):
    """Full text of a w:p element in document order, math included.
    Accepts either a python-docx Paragraph or a raw lxml w:p element."""
    if hasattr(p_el, "_p"):
        p_el = p_el._p
    out = []
    _walk(p_el, out)
    return "".join(out)


def cell_text(cell):
    """Full text of a table cell (direct paragraphs only — nested tables are
    a separate structure and are read separately), math included."""
    paras = cell._tc.findall("{%s}p" % W_NS)
    return "\n".join(paragraph_text(p) for p in paras)


def paragraph_segments(p_el):
    """The paragraph as ordered segments [('t', text), ('m', omml_element)] —
    used to transplant equations into PPTX as NATIVE PowerPoint math
    (pptx_slide_ops.make_para_with_math_xml) instead of linearized text.
    The math elements are returned live (lxml) for verbatim copying."""
    if hasattr(p_el, "_p"):
        p_el = p_el._p
    segs = []

    def walk(el):
        for child in el:
            tag = child.tag
            if tag == "{%s}t" % W_NS:
                if segs and segs[-1][0] == "t":
                    segs[-1] = ("t", segs[-1][1] + (child.text or ""))
                else:
                    segs.append(("t", child.text or ""))
            elif tag == "{%s}oMath" % M_NS:
                segs.append(("m", child))
            elif tag == "{%s}oMathPara" % M_NS:
                for om in child.findall("{%s}oMath" % M_NS):
                    segs.append(("m", om))
            elif tag in ("{%s}br" % W_NS, "{%s}cr" % W_NS):
                segs.append(("t", "\n"))
            else:
                walk(child)

    walk(p_el)
    return segs


def cell_paragraphs(cell):
    """The cell's direct w:p elements, for per-paragraph processing."""
    return cell._tc.findall("{%s}p" % W_NS)


def stage_tables(doc, heading_substring):
    """The content tables that sit under the Heading-1 whose text contains
    heading_substring — the proven heading-pairing extraction (avoids the
    table-index miscounting bug, checklist item on extraction)."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    current, out = None, []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            p = Paragraph(child, doc)
            if p.style.name.startswith("Heading") and p.text.strip():
                current = p.text.strip()
        elif child.tag.endswith("}tbl"):
            if current and heading_substring in current:
                out.append(Table(child, doc))
    return out
