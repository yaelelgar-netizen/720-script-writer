"""
Reusable python-pptx slide-manipulation primitives for the 720-script-writer skill.

python-pptx has no built-in API for: duplicating a slide (same-file or cross-file),
reordering slides, or editing PowerPoint Sections (p14:sectionLst, an OOXML extension).
Every function here was proven against real files during the 4.2 build (see
references/construction-checklist.md for the bugs each one fixes).

Usage pattern:
    from pptx_slide_ops import *
    prs = Presentation(target_path)
    slide, sldId = duplicate_slide(prs, source_index)          # same file
    slide2 = duplicate_slide_cross_file(template_prs, idx, prs) # template -> target
    insert_after(prs, anchor_index, [sldId, ...])
    add_to_section(prs, "תרגול סטנדרטי שאלה 1", [sldId, ...])
    ...
    prs.save(target_path)
"""

import copy
import io
import re

from lxml import etree
from pptx.oxml.ns import qn
from pptx.opc.package import Part
from pptx.util import Emu, Inches, Pt
from pptx.enum.text import PP_ALIGN

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
AONS = {"a": A_NS}


# --------------------------------------------------------------------------
# Slide duplication
# --------------------------------------------------------------------------

def _assign_free_slide_partname(prs, new_slide):
    """python-pptx names new slide parts by COUNT (len(sldIdLst)+1), blind to
    which names are taken. After any delete_slide(), count < max-used-number,
    so the next add_slide() SILENTLY reuses a live slide's partname — the
    saved zip then contains two entries with the same name and PowerPoint
    keeps whichever it likes (construction-checklist.md item 12; surfaced as
    'Duplicate name' UserWarnings during save). Called right after every
    add_slide here: reassigns the new part to max(used)+1, which is always
    free. Renaming is safe pre-save — rels hold object references, partnames
    only matter at serialization."""
    from pptx.opc.packuri import PackURI
    used = []
    for s in prs.slides:
        m = re.match(r"/ppt/slides/slide(\d+)\.xml$", str(s.part.partname))
        if m:
            used.append(int(m.group(1)))
    new_slide.part.partname = PackURI("/ppt/slides/slide%d.xml" % (max(used) + 1))


def duplicate_slide(prs, index):
    """Clone slide at `index` (0-based) within the SAME Presentation.
    Appends at the end of the slide list. Returns (new_slide, new_sldId_element).
    Use insert_after() to move it to the right physical position."""
    source = prs.slides[index]
    dest = prs.slides.add_slide(source.slide_layout)
    _assign_free_slide_partname(prs, dest)
    sldIdLst = prs.slides._sldIdLst
    new_sldId_elem = list(sldIdLst)[-1]

    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)

    rid_map = {}
    for rel_id, rel in source.part.rels.items():
        if rel.reltype.endswith("notesSlide"):
            continue
        if rel.is_external:
            new_rid = dest.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            new_rid = dest.part.relate_to(rel.target_part, rel.reltype)
        rid_map[rel_id] = new_rid

    _copy_shape_tree(source.shapes._spTree, dest.shapes._spTree, rid_map)
    return dest, new_sldId_elem


def duplicate_slide_cross_file(src_prs, src_index, dest_prs, dest_layout=None):
    """Clone a slide from src_prs (e.g. the template library) into dest_prs
    (the working script) — different Presentation objects, different packages.

    Image relationships are copied as raw Parts (not via get_or_add_image_part,
    which pipes through PIL and breaks on non-raster blips like embedded SVG —
    see construction-checklist.md item 5). Table data needs no special handling
    since it's embedded directly in the slide XML, not a separate part.

    dest_layout defaults to the layout of dest_prs's first slide (matches the
    "כותרת ותוכן" blank layout convention seen in every script so far). Pass
    an explicit layout if the target deck uses a different one.
    """
    source = src_prs.slides[src_index]
    if dest_layout is None:
        dest_layout = dest_prs.slides[0].slide_layout
    dest = dest_prs.slides.add_slide(dest_layout)
    _assign_free_slide_partname(dest_prs, dest)

    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)

    rid_map = {}
    for rel_id, rel in source.part.rels.items():
        if rel.reltype.endswith("notesSlide"):
            continue
        if rel.is_external:
            rid_map[rel_id] = dest.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT_IMAGE:
            src_part = rel.target_part
            ext = src_part.partname.ext
            new_partname = dest_prs.part.package.next_image_partname(ext)
            new_part = Part(new_partname, src_part.content_type, dest_prs.part.package, blob=src_part.blob)
            rid_map[rel_id] = dest.part.relate_to(new_part, RT_IMAGE)
        # other reltypes (charts, embedded objects) are not handled yet -
        # none observed in the template library so far; if encountered, extend here.

    _copy_shape_tree(source.shapes._spTree, dest.shapes._spTree, rid_map)
    return dest


def copy_shapes_within_presentation(src_slide, shapes, dest_slide):
    """Copy specific shapes from one slide to another slide of the SAME
    Presentation/package (unlike duplicate_slide_cross_file, the media parts
    are already shared — only relationship IDs need remapping on dest_slide,
    not the underlying image parts themselves).

    Used to merge a subset of one template-state slide's shapes (e.g. a Hint
    slide's hint-panel group) onto another slide that's becoming the single
    consolidated "דפוס הקנייה" slide — see section-structure.md. Returns the
    new shapes, in the same order as `shapes`, so the caller can reposition
    them (e.g. parked off-canvas) individually."""
    rid_map = {}
    for rel_id, rel in src_slide.part.rels.items():
        if rel.reltype.endswith("notesSlide"):
            continue
        if rel.is_external:
            rid_map[rel_id] = dest_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            rid_map[rel_id] = dest_slide.part.relate_to(rel.target_part, rel.reltype)

    new_shapes = []
    for shape in shapes:
        new_elem = copy.deepcopy(shape._element)
        for e in new_elem.iter():
            for attr_name in list(e.attrib.keys()):
                if attr_name.startswith("{" + R_NS + "}"):
                    old_rid = e.attrib[attr_name]
                    if old_rid in rid_map:
                        e.attrib[attr_name] = rid_map[old_rid]
        dest_slide.shapes._spTree.append(new_elem)
        new_shapes.append(next(sh for sh in dest_slide.shapes if sh._element is new_elem))
    return new_shapes


def _copy_shape_tree(src_spTree, dst_spTree, rid_map):
    for elem in list(src_spTree):
        tag = elem.tag
        if tag == qn("p:nvGrpSpPr") or tag == qn("p:grpSpPr"):
            continue
        new_elem = copy.deepcopy(elem)
        for e in new_elem.iter():
            for attr_name in list(e.attrib.keys()):
                if attr_name.startswith("{" + R_NS + "}"):
                    old_rid = e.attrib[attr_name]
                    if old_rid in rid_map:
                        e.attrib[attr_name] = rid_map[old_rid]
        dst_spTree.append(new_elem)


# --------------------------------------------------------------------------
# Slide ordering
# --------------------------------------------------------------------------

def delete_slide(prs, index):
    """Permanently remove the slide at `index` (0-based): drops its part
    relationship (so python-pptx's save walk omits the now-orphaned slide
    part — there is no direct "delete part" API), removes its <p:sldId> from
    the main list, AND removes any matching <p14:sldId> from every
    PowerPoint Section — never leaves a dangling Section membership entry
    pointing at a slide that no longer exists.

    Used to consolidate multiple template-state slides (e.g. Question+Hint+
    Feedback) into one merged slide — see section-structure.md, "דפוס הקנייה"."""
    sldIdLst = prs.slides._sldIdLst
    sldId_elem = list(sldIdLst)[index]
    numeric_id = sldId_elem.get("id")
    r_id = sldId_elem.get(qn("r:id"))
    prs.part.drop_rel(r_id)
    sldIdLst.remove(sldId_elem)

    root = sldIdLst.getroottree().getroot()
    ns = {"p14": P14_NS}
    for section_sldIdLst in root.findall(".//p14:section/p14:sldIdLst", ns):
        for el in list(section_sldIdLst):
            if el.get("id") == numeric_id:
                section_sldIdLst.remove(el)


def insert_after(prs, anchor_index, sldId_elems):
    """Move sldId_elems (as returned by duplicate_slide) to sit immediately
    after the slide currently at anchor_index (0-based), in the given order.
    Uses direct element references, not numeric index arithmetic, so it's
    safe across multiple inserts in one call."""
    sldIdLst = prs.slides._sldIdLst
    anchor_elem = list(sldIdLst)[anchor_index]
    for el in sldId_elems:
        sldIdLst.remove(el)
    anchor_pos = list(sldIdLst).index(anchor_elem)
    for offset, el in enumerate(sldId_elems, start=1):
        sldIdLst.insert(anchor_pos + offset, el)


# --------------------------------------------------------------------------
# PowerPoint Sections (מקטעים)
# --------------------------------------------------------------------------

def add_to_section(prs, section_name, new_sldId_elems):
    """Append sldId elements (already positioned in the main sldIdLst) to the
    named Section's membership list. Call this AFTER insert_after()."""
    root = prs.slides._sldIdLst.getroottree().getroot()
    ns = {"p14": P14_NS}
    sections = root.findall(".//p14:section", ns)
    target = next((s for s in sections if s.get("name") == section_name), None)
    if target is None:
        raise KeyError("section not found: " + section_name)
    sldIdLst_el = target.find("p14:sldIdLst", ns)
    for el in new_sldId_elems:
        new_id_el = etree.SubElement(sldIdLst_el, "{%s}sldId" % P14_NS)
        new_id_el.set("id", el.get("id"))


def list_sections(prs):
    """Return [(name, [1-based slide positions])] for every Section, in document order."""
    root = prs.slides._sldIdLst.getroottree().getroot()
    ns = {"p14": P14_NS}
    all_ids = [el.get("id") for el in prs.slides._sldIdLst]
    out = []
    for s in root.findall(".//p14:section", ns):
        sids = [el.get("id") for el in s.find("p14:sldIdLst", ns)]
        positions = sorted(all_ids.index(sid) + 1 for sid in sids if sid in all_ids)
        out.append((s.get("name"), positions))
    return out


def create_section(prs, section_name, new_sldId_elems, after_section_name=None):
    """Create a brand-new Section (for a whole new question, as opposed to a new
    sub-part of an existing one — use add_to_section for that case). If
    after_section_name is given, the new <p14:section> is inserted right after
    it in the sectionLst XML (keeps the Section panel's display order sane);
    otherwise it's appended at the end.

    IMPORTANT: every real Section has a GUID `id` attribute (e.g.
    id="{4E7D1332-...}"). Omitting it produces XML that python-pptx happily
    reads back (list_sections() reports it fine) but PowerPoint silently does
    NOT show in the Section panel — this bug shipped once already. Always
    generate one."""
    import uuid
    root = prs.slides._sldIdLst.getroottree().getroot()
    ns = {"p14": P14_NS}
    sectionLst = root.find(".//p14:sectionLst", ns)
    section_id = "{%s}" % str(uuid.uuid4()).upper()
    new_section = etree.Element("{%s}section" % P14_NS, name=section_name, id=section_id)
    new_sldIdLst_el = etree.SubElement(new_section, "{%s}sldIdLst" % P14_NS)
    for el in new_sldId_elems:
        child = etree.SubElement(new_sldIdLst_el, "{%s}sldId" % P14_NS)
        child.set("id", el.get("id"))
    if after_section_name is not None:
        anchor = next(s for s in sectionLst.findall("p14:section", ns) if s.get("name") == after_section_name)
        anchor.addnext(new_section)
    else:
        sectionLst.append(new_section)
    return new_section


# --------------------------------------------------------------------------
# High-level orchestration
# --------------------------------------------------------------------------

def assemble_question(src_prs, src_indices, dest_prs, anchor_index, section_name,
                       existing_section=True, after_section_name=None):
    """Clone every slide in src_indices (0-based, in order — typically
    [question, hint, feedback] or [question, hint, feedback_a, feedback_b])
    from src_prs into dest_prs, auto-clean each with clean_template_slide(),
    position them right after anchor_index (0-based), and register them in
    the named PowerPoint Section.

    existing_section=True  -> section_name must already exist (adding a
                               sub-part to a question that already has one).
    existing_section=False -> a brand-new Section is created (a whole new
                               question); after_section_name places it in the
                               Section panel right after that section.

    Returns (slides, cleanup_reports) — the new slide objects in order (for
    the caller to fill with set_scenario_and_instruction/replace_paragraphs/
    table-building etc.) and the clean_template_slide() report per slide.
    """
    slides, sldIds, cleanup_reports = [], [], []
    for idx in src_indices:
        slide = duplicate_slide_cross_file(src_prs, idx, dest_prs)
        sldId = list(dest_prs.slides._sldIdLst)[-1]
        cleanup_reports.append(clean_template_slide(slide))
        slides.append(slide)
        sldIds.append(sldId)

    insert_after(dest_prs, anchor_index, sldIds)

    if existing_section:
        add_to_section(dest_prs, section_name, sldIds)
    else:
        create_section(dest_prs, section_name, sldIds, after_section_name=after_section_name)

    return slides, cleanup_reports


# --------------------------------------------------------------------------
# Shape lookup / removal / cloning
# --------------------------------------------------------------------------

def find_shape_by_pos(slide, left_in, top_in, tol=0.05):
    for sh in slide.shapes:
        try:
            if abs(Emu(sh.left).inches - left_in) < tol and abs(Emu(sh.top).inches - top_in) < tol:
                return sh
        except Exception:
            continue
    raise KeyError((left_in, top_in))


def find_shape_by_text(slide_or_shapes, substring):
    """Search recursively into GROUP shapes. Several templates (SingleChoiceQuestion,
    StatementAssessmentQuestion, SingleChoiceQuestionImage...) nest their real
    content 2+ levels deep inside groups, so a top-level-only scan silently
    finds nothing there. Returns the first match, depth-first."""
    shapes = slide_or_shapes.shapes if hasattr(slide_or_shapes, "shapes") else slide_or_shapes
    for sh in shapes:
        if sh.has_text_frame and substring in sh.text_frame.text:
            return sh
        if sh.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            try:
                return find_shape_by_text(sh, substring)
            except KeyError:
                continue
    raise KeyError(substring)


def find_all_shapes_by_text(slide_or_shapes, substring):
    """Like find_shape_by_text but returns every match (e.g. all 4 'מסיח N'
    choice boxes), recursing into groups. Order is depth-first, document order."""
    shapes = slide_or_shapes.shapes if hasattr(slide_or_shapes, "shapes") else slide_or_shapes
    out = []
    for sh in shapes:
        if sh.has_text_frame and substring in sh.text_frame.text:
            out.append(sh)
        if sh.shape_type == 6:
            out.extend(find_all_shapes_by_text(sh, substring))
    return out


# --------------------------------------------------------------------------
# Generic template cleanup — works across EVERY template in תבנית לתסריט.pptx,
# because the artifacts follow a universal pattern, not a per-template one.
# See construction-checklist.md item 2/6.
# --------------------------------------------------------------------------

_BANNER_IMAGE_POS = (3.07, 0.46, 8.58, 0.61)  # identical placeholder banner across ~all templates
PLACEHOLDER_MARKERS = [
    "למפתחת", "טקסט השאלה", "טקסט רמז", "טקסט משוב",
    "מקום לתמונה", "מקום לגיף", "מקום לוידיאו", "טקסט רץ", "טקסט H1",
    # updated-library feedback placeholders.
    "משוב להצלחה (למשל", "משוב לאי הצלחה (למשל",
    # SingleChoiceQuestion's extra first-attempt box — corrected 09/07/2026:
    # no real script keeps this, clean_template_slide strips it
    # unconditionally (strip_first_attempt_feedback_box); this marker exists
    # only to flag it as residue on slides built before the correction.
    "משוב לאי הצלחה ראשון",
]
# "מסיח N" (a digit) is an unfilled placeholder; "מסיח א'" is real distractor
# content ("distractor" is also normal vocabulary in feedback text) - a plain
# substring match on "מסיח " would false-positive on the latter, so this one
# needs its own regex requiring a following digit.
_PLACEHOLDER_DISTRACTOR_RE = re.compile(r"מסיח\s*\d")


def strip_developer_notes(slide, tol=0.01):
    """Remove every shape whose text starts with 'למפתחת:' — the production-only
    annotations baked into every template. Universal across all template types,
    so this alone prevents the 'leftover note on the hint slide' class of bug.
    Returns the number of shapes removed."""
    removed = 0
    for sh in list(slide.shapes):
        if sh.has_text_frame and sh.text_frame.text.strip().startswith("למפתחת"):
            remove_shape(sh)
            removed += 1
    return removed


def strip_generic_banner_image(slide, tol=0.1):
    """Remove the generic placeholder scenario banner image that appears at an
    identical position/size across nearly every template slide (a leftover
    from template authoring, not a structural element). Safe no-op if absent."""
    l, t, w, h = _BANNER_IMAGE_POS
    for sh in list(slide.shapes):
        try:
            if (abs(Emu(sh.left).inches - l) < tol and abs(Emu(sh.top).inches - t) < tol
                    and abs(Emu(sh.width).inches - w) < tol and abs(Emu(sh.height).inches - h) < tol):
                remove_shape(sh)
                return True
        except Exception:
            continue
    return False


def strip_percent_symbol_demo(slide, tol=0.1):
    """Remove ValueInputQuestion's "אם נדרש סימון מתמטי קבוע בשדה, למשל אחוז"
    demo (a second answer box + a floating '%' text) — irrelevant, and
    visually misleading, for any answer that isn't itself a percentage
    (construction-checklist.md item 9). Geometry per the UPDATED template
    library. Called from build_question, which skips it when the spec asks
    to keep the fixed symbol ('percent_field': True)."""
    removed = 0
    for (l, t) in _VI_PERCENT_DEMO:
        for sh in list(slide.shapes):
            try:
                if abs(Emu(sh.left).inches - l) < tol and abs(Emu(sh.top).inches - t) < tol:
                    remove_shape(sh)
                    removed += 1
                    break
            except Exception:
                continue
    return removed


def strip_optional_image_placeholder(slide):
    """Remove the optional 'מקום לתמונה' image-slot shape present in some
    templates (e.g. StatementAssessmentQuestion). Matched by exact text, not
    position, since its position varies per template. Unlike 'מקום לגיף'
    (a deliberate, deck-wide convention for a production cue left unfilled
    on every transition screen), 'מקום לתמונה' never survives unfilled
    anywhere else in a finished script — so it must be stripped when the
    question has no image, not left as a cue. Returns True if removed."""
    removed = False
    for sh in list(slide.shapes):
        if sh.has_text_frame and sh.text_frame.text.strip() == "מקום לתמונה":
            remove_shape(sh)
            removed = True
    return removed


def strip_first_attempt_feedback_box(slide):
    """Remove SingleChoiceQuestion's extra 'משוב לאי הצלחה ראשון' box (default
    boilerplate 'התשובה אינה נכונה במלואה.'). Originally catalogued as
    template content to leave untouched — corrected by the user (09/07/2026):
    there is no reason for a real script to ship an unfilled/generic
    first-attempt feedback box alongside the real positive/negative ones, so
    it is now stripped unconditionally like the other template-authoring
    artifacts. Matched by exact top-level text, not position (position is
    template-specific). Returns True if removed."""
    removed = False
    for sh in list(slide.shapes):
        if sh.has_text_frame and sh.text_frame.text.strip().startswith("משוב לאי הצלחה ראשון"):
            remove_shape(sh)
            removed = True
    return removed


def clean_template_slide(slide):
    """Run the generic cleanups on a freshly cloned template slide, right
    after duplicate_slide_cross_file(), before filling in real content.
    The percent-symbol demo is NOT handled here — it is conditional
    (build_question strips it unless the spec keeps it). The רמת חשיבה tag
    parked above the canvas is template content and is deliberately kept."""
    return {
        "developer_notes_removed": strip_developer_notes(slide),
        "banner_image_removed": strip_generic_banner_image(slide),
        "image_placeholder_removed": strip_optional_image_placeholder(slide),
        "first_attempt_feedback_removed": strip_first_attempt_feedback_box(slide),
    }


def scan_placeholder_residue(prs, slide_indices):
    """Verification step (construction-checklist.md): scan the given 0-based
    slide indices for any leftover placeholder text (developer notes, unfilled
    'טקסט השאלה'/'מסיח'/'מקום לתמונה' etc.) that should have been replaced with
    real content or stripped. Returns a list of (slide_index_1_based, shape_name,
    matched_marker, text_preview) — empty list means clean. Call this as the
    LAST step before declaring a build done; do not eyeball printed dumps."""
    hits = []

    def scan_shapes(shapes, idx):
        for sh in shapes:
            if sh.shape_type == 6:  # recurse into groups — the hint text
                scan_shapes(sh.shapes, idx)  # lives inside a parked group
                continue
            if not sh.has_text_frame:
                continue
            text = sh.text_frame.text.strip()
            if not text:
                continue
            for marker in PLACEHOLDER_MARKERS:
                if marker in text:
                    hits.append((idx + 1, sh.name, marker, text[:60]))
                    break
            else:
                m = _PLACEHOLDER_DISTRACTOR_RE.search(text)
                if m:
                    hits.append((idx + 1, sh.name, m.group(0), text[:60]))

    for idx in slide_indices:
        scan_shapes(prs.slides[idx].shapes, idx)
    return hits


FONT_NAME = "Assistant"


def style_cell_rtl(cell, text, bold=False, size_pt=18):
    """Set a table cell's text with correct RTL Hebrew rendering and the
    deck's font. python-pptx's plain `cell.text = ...` defaults to LTR
    paragraph direction + the theme's default font (NOT Assistant) — this
    renders Hebrew fine but mangles any bracket/parenthesis pairing (they
    come out mirrored) and looks visually inconsistent with every other
    shape in the deck. ALWAYS use this instead of raw `cell.text = ...` when
    building a new table from scratch (a cloned/existing table already has
    correct paragraph properties inherited, so this isn't needed there).
    See construction-checklist.md item 6."""
    tf = cell.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    p._pPr.set("rtl", "1")
    run = p.add_run()
    run.text = text
    run.font.name = FONT_NAME
    run.font.size = Pt(size_pt)
    run.font.bold = bold


def build_rtl_table(slide, rows_data, left_in, top_in, width_in, height_in, header_bold=True, size_pt=18):
    """Add a new table (add_table) with EVERY cell correctly RTL-styled and
    set to Assistant — the columns are written in the given left-to-right
    index order, so if the source data is in logical RTL reading order
    (rightmost concept first), reverse the row tuples before calling this,
    or pass mirror_columns=True-equivalent by pre-reversing `rows_data`.
    rows_data: list of row-tuples, rows_data[0] is treated as the header row.
    Returns the created table shape."""
    n_rows = len(rows_data)
    n_cols = len(rows_data[0])
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in))
    table = tbl_shape.table
    for r, row in enumerate(rows_data):
        for c, value in enumerate(row):
            style_cell_rtl(table.cell(r, c), value, bold=(header_bold and r == 0), size_pt=size_pt)
    return tbl_shape


def remove_shape(shape):
    shape._element.getparent().remove(shape._element)


def clone_shape(slide, shape, new_left_in, new_top_in):
    """Deep-copy an existing (properly styled) shape within the same slide and
    reposition it. Always prefer this over add_textbox/add_shape for new
    content — see construction-checklist.md item 3."""
    new_el = copy.deepcopy(shape._element)
    slide.shapes._spTree.append(new_el)
    new_shape = next(sh for sh in slide.shapes if sh._element is new_el)
    new_shape.left = Inches(new_left_in)
    new_shape.top = Inches(new_top_in)
    return new_shape


# --------------------------------------------------------------------------
# Text editing — robust to whatever run-structure the source slide has
# --------------------------------------------------------------------------

def set_simple_text(shape, text, bold=None):
    """Collapse a text frame's first paragraph into its first run, clearing
    every other run/paragraph. Keeps the first run's original rPr (only the
    <a:t> text changes), optionally overriding bold. Use for single-line
    labels (buttons, table cells, answer boxes)."""
    txBody = shape.text_frame._txBody
    paras = txBody.findall("a:p", AONS)
    runs0 = paras[0].findall("a:r", AONS)
    runs0[0].find("a:t", AONS).text = text
    if bold is not None:
        rPr = runs0[0].find("a:rPr", AONS)
        rPr.set("b", "1" if bold else "0")
    for r in runs0[1:]:
        r.find("a:t", AONS).text = ""
    for p in paras[1:]:
        for r in p.findall("a:r", AONS):
            r.find("a:t", AONS).text = ""


def set_scenario_and_instruction(shape, scenario_text, instruction_text):
    """For the common 'scenario paragraph + bold/purple instruction sentence'
    text box (see content-style-rules.md #1-#2): run[0] gets the scenario
    text, run[1] gets the full instruction, and any further trailing runs
    (leftover split fragments from manual edits) are cleared. Does NOT assume
    a fixed run count — see construction-checklist.md item 1 for why."""
    txBody = shape.text_frame._txBody
    para0 = txBody.findall("a:p", AONS)[0]
    runs = para0.findall("a:r", AONS)
    runs[0].find("a:t", AONS).text = scenario_text
    runs[1].find("a:t", AONS).text = instruction_text
    for r in runs[2:]:
        r.find("a:t", AONS).text = ""


# --------------------------------------------------------------------------
# Rich multi-paragraph text (feedback boxes with mixed alignment/bold)
# --------------------------------------------------------------------------

_RUN_TMPL = (
    '<a:r xmlns:a="%s"><a:rPr lang="he-IL" sz="%d" %sdirty="0">'
    "<a:solidFill>%s</a:solidFill>"
    '<a:latin typeface="Assistant" pitchFamily="2" charset="-79"/>'
    '<a:cs typeface="Assistant" pitchFamily="2" charset="-79"/></a:rPr>'
    "<a:t>%s</a:t></a:r>"
)


def make_run_xml(text, bold=False, color_hex=None, size_pt=20):
    b_attr = 'b="1" ' if bold else ""
    fill = ('<a:srgbClr val="%s"/>' % color_hex) if color_hex else "<a:schemeClr val=\"tx1\"/>"
    escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return _RUN_TMPL % (A_NS, size_pt * 100, b_attr, fill, escaped)


def make_para_xml(runs, center=False, empty=False, rtl_right=False, size_pt=20):
    """runs: list of (text, bold, color_hex) tuples. center/rtl_right control
    <a:pPr algn=.../>; per content-style-rules.md #8, feedback intro lines are
    centered and detail lines use the default (right, RTL) alignment."""
    algn, extra_ppr = "", ""
    if center:
        algn = ' algn="ctr"'
    elif rtl_right:
        algn = ' algn="r"'
        extra_ppr = ' rtl="1"'
    if empty:
        return (
            '<a:p xmlns:a="%s"><a:pPr%s%s/>'
            '<a:endParaRPr lang="he-IL" sz="%d" dirty="0">'
            '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
            '<a:latin typeface="Assistant" pitchFamily="2" charset="-79"/>'
            '<a:cs typeface="Assistant" pitchFamily="2" charset="-79"/></a:endParaRPr></a:p>'
        ) % (A_NS, algn, extra_ppr, size_pt * 100)
    pPr = ("<a:pPr%s%s/>" % (algn, extra_ppr)) if (algn or extra_ppr) else ""
    runs_xml = "".join(make_run_xml(t, b, c, size_pt) for t, b, c in runs)
    return '<a:p xmlns:a="%s">%s%s</a:p>' % (A_NS, pPr, runs_xml)


A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_W_NS_WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def omml_for_ppt(omml_el, size_pt=20, bold=False):
    """A cleaned copy of a Word m:oMath for embedding in a PPTX text frame:
    WordprocessingML-namespace children (w:rPr inside math runs etc.) are
    stripped — PowerPoint math expects DrawingML properties and treats the
    Word ones as invalid content. Every math run then gets an EXPLICIT
    DrawingML a:rPr with a dark (tx1) fill — without it the equation
    inherits the feedback box's default run color (white) and is invisible
    on the light fills (user correction)."""
    M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    el = copy.deepcopy(omml_el)
    for child in list(el.iter()):
        if child.tag.startswith("{%s}" % _W_NS_WORD):
            child.getparent().remove(child)
    rpr_xml = (
        '<a:rPr xmlns:a="%s" lang="he-IL" sz="%d"%s dirty="0">'
        '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr>'
        % (A_NS, size_pt * 100, ' b="1"' if bold else ""))
    for r in el.iter("{%s}r" % M):
        r.insert(0, etree.fromstring(rpr_xml.encode("utf-8")))
    return el


def make_para_with_math_xml(segments, bold=False, size_pt=20, center=False, rtl_right=False):
    """<a:p> XML mixing text runs with NATIVE PowerPoint equations — the
    m:oMath elements are transplanted VERBATIM from the content Word doc
    (checklist item 13: the equation is authored pedagogy). Each math zone
    is wrapped in mc:AlternateContent with an a14:m choice and a plain-text
    fallback (the linearized equation) so older clients still show content.

    segments: [('t', 'text'), ('m', omml_element_or_xml_string)] in order —
    exactly what docx_extract.paragraph_segments() returns."""
    from docx_extract import _omml_to_text, _fix_rtl_units
    algn, extra = "", ""
    if center:
        algn = ' algn="ctr"'
    elif rtl_right:
        algn = ' algn="r"'
        extra = ' rtl="1"'
    pPr = ("<a:pPr%s%s/>" % (algn, extra)) if (algn or extra) else ""
    parts = []
    for kind, val in segments:
        if kind == "t":
            parts.append(make_run_xml(val, bold, None, size_pt))
        else:
            el = val if not isinstance(val, str) else etree.fromstring(val.encode("utf-8"))
            linear = _fix_rtl_units(_omml_to_text(el))
            om_xml = etree.tostring(omml_for_ppt(el, size_pt=size_pt, bold=bold),
                                    encoding="unicode")
            parts.append(
                '<mc:AlternateContent xmlns:mc="%s">'
                '<mc:Choice xmlns:a14="%s" Requires="a14"><a14:m>%s</a14:m></mc:Choice>'
                "<mc:Fallback>%s</mc:Fallback>"
                "</mc:AlternateContent>"
                % (MC_NS, A14_NS, om_xml, make_run_xml(linear, bold, None, size_pt)))
    return '<a:p xmlns:a="%s">%s%s</a:p>' % (A_NS, pPr, "".join(parts))


def copy_notes_verbatim(src_slide, dest_slide, append_text=None):
    """Copy speaker notes from a template slide as RAW paragraph XML —
    setting .text loses the RTL paragraph direction, which makes mixed
    Hebrew/Latin lines (e.g. 'שם המסך בפיגמה: ValueInputQuestion') render
    with the colon on the wrong side. append_text (verbatim source-doc
    production text ONLY — checklist item 14) is added as RTL paragraphs."""
    dest_body = dest_slide.notes_slide.notes_text_frame._txBody
    for p in dest_body.findall(qn("a:p")):
        dest_body.remove(p)
    if src_slide.has_notes_slide:
        src_body = src_slide.notes_slide.notes_text_frame._txBody
        for p in src_body.findall(qn("a:p")):
            new_p = copy.deepcopy(p)
            # display-only fix (content untouched): the template's notes
            # paragraphs carry no explicit direction, so mixed Hebrew/Latin
            # lines ('שם המסך בפיגמה: ValueInputQuestion') render with the
            # colon jumbled. Force RTL+right so they sit correctly.
            pPr = new_p.find(qn("a:pPr"))
            if pPr is None:
                pPr = etree.SubElement(new_p, qn("a:pPr"))
                new_p.insert(0, pPr)
            pPr.set("rtl", "1")
            pPr.set("algn", "r")
            dest_body.append(new_p)
    if append_text:
        for line in append_text.split("\n"):
            dest_body.append(etree.fromstring(
                make_para_xml([(line, False, None)], rtl_right=True).encode("utf-8")))
    if not dest_body.findall(qn("a:p")):
        dest_body.append(etree.fromstring(
            make_para_xml([("", False, None)]).encode("utf-8")))


def replace_paragraphs(shape, paragraphs_xml):
    """Wipe a text frame's paragraphs and replace with freshly built ones
    (from make_para_xml). Use instead of patching individual runs when the
    target needs a different paragraph count/alignment mix than the source."""
    txBody = shape.text_frame._txBody
    for old_p in txBody.findall("a:p", AONS):
        txBody.remove(old_p)
    for p_xml in paragraphs_xml:
        txBody.append(etree.fromstring(p_xml.encode("utf-8")))


# --------------------------------------------------------------------------
# Workflow layer — declarative question building
#
# Everything below exists so a question build is ~15 lines of content DATA
# instead of ~60 lines of copy-pasted procedural code. The per-script
# helpers (current_anchor, blank_unused_choices, fill_* ...) used to be
# rewritten in every build script, drifting between copies (one script's
# blank_unused_choice took an int, another's took a list). This is the
# single canonical home.
# --------------------------------------------------------------------------

INSTRUCTION_COLOR = "C46BFF"  # bold+purple instruction sentence (content-style-rules.md #2)

# 0-based slide indices in תבנית לתסריט.pptx (the UPDATED 37-slide library,
# 02/07/2026). Each question template is now a SINGLE consolidated slide —
# the library itself adopted the one-slide-per-question format (its slide 20
# announces it), with the hint group and both feedback boxes already parked
# left of the canvas. THIS dict is the machine-readable source of truth for
# code; template-catalog.md is the prose companion. Types marked (unwired)
# are cataloged but their fill-logic is not implemented yet — verify shape
# geometry before first use.
TEMPLATES = {
    "SingleChoiceQuestion": 20,
    "SingleChoiceQuestionImage": 21,   # (unwired)
    "DropdownQuestion": 22,
    "ValueInputQuestion": 23,
    "ImageHotspotQuestion": 24,        # (unwired)
    "StatementAssessmentQuestion": 25,
    "DragAndDropQuestion": 26,         # (unwired)
    "SingleChoice/Bubble": 27,         # (unwired)
    "TransitionScreen": 10,            # generic; stage-specific ones at 11-18
}


def current_anchor(prs, section_name):
    """0-based index of the last slide of the named Section — the anchor
    after which the next insert should land."""
    sections = list_sections(prs)
    sec = next(s for s in sections if s[0] == section_name)
    return sec[1][-1] - 1


def make_scenario_paras(scenario_lines, instruction, size_pt=20):
    """Standard question text box: scenario line(s), blank spacer, then the
    bold+purple instruction sentence (content-style-rules.md #1-#2).
    scenario_lines: list of plain strings, one paragraph each."""
    paras = [make_para_xml([(line, False, None)], rtl_right=True, size_pt=size_pt)
             for line in scenario_lines]
    paras.append(make_para_xml([], rtl_right=True, empty=True, size_pt=size_pt))
    paras.append(make_para_xml([(instruction, True, INSTRUCTION_COLOR)],
                               rtl_right=True, size_pt=size_pt))
    return paras


def make_feedback_paras(intro, detail_lines=(), size_pt=20):
    """Standard feedback box: centered bold intro ('נכון! / טעית, ...'),
    then optionally a spacer + right-aligned detail lines.
    detail_lines: list of (text, bold) tuples — bold only the invariant
    facts/result line (content-style-rules.md #8). A line's text may also be
    a SEGMENTS list ([('t', ...), ('m', omml)], from
    docx_extract.paragraph_segments) — it is rendered with NATIVE PowerPoint
    math via make_para_with_math_xml."""
    paras = [make_para_xml([(intro, True, None)], center=True, size_pt=size_pt)]
    if detail_lines:
        paras.append(make_para_xml([], center=True, empty=True, size_pt=size_pt))
        for text, bold in detail_lines:
            if isinstance(text, (list, tuple)):
                paras.append(make_para_with_math_xml(text, bold, size_pt=size_pt))
            else:
                paras.append(make_para_xml([(text, bold, None)], size_pt=size_pt))
    return paras


def blank_unused_choices(slide, n_used, max_slots=4):
    """Blank the trailing 'מסיח N' slots a choice template ships with but this
    question doesn't use (e.g. 3 real options on a 4-slot template)."""
    for i in range(n_used + 1, max_slots + 1):
        try:
            set_simple_text(find_shape_by_text(slide, "מסיח %d" % i), "", bold=False)
        except KeyError:
            pass


def _fill_choices(slide, choice_texts):
    for i, text in enumerate(choice_texts, start=1):
        set_simple_text(find_shape_by_text(slide, "מסיח %d" % i), text, bold=False)
    blank_unused_choices(slide, len(choice_texts))


def set_shape_line_color(shape, color_hex):
    from pptx.dml.color import RGBColor
    shape.line.color.rgb = RGBColor.from_string(color_hex)


def _row_pill_and_ellipse(slide, row_index_1based):
    """(pill, ellipse) of a choice row in the UPDATED SingleChoice template.
    Rows 1/2/4 nest pill+text in an inner group with the radio ellipse as a
    sibling; row 3's pill sits directly in the row group and its radio is a
    DETACHED top-level ellipse (the template's green-marking demo)."""
    row = find_shape_by_pos(slide, _SC_ROW_LEFT_IN, _SC_ROW_TOPS_IN[row_index_1based - 1])

    def autoshapes(shapes):
        out = []
        for sh in shapes:
            if sh.shape_type == 6:
                out.extend(autoshapes(sh.shapes))
            elif sh.shape_type == 1:
                out.append(sh)
        return out

    shapes = autoshapes(row.shapes)
    pill = next(sh for sh in shapes if Emu(sh.width).inches > 2)
    ellipse = next((sh for sh in shapes if Emu(sh.width).inches < 1), None)
    if ellipse is None:  # row 3: detached radio at top level
        ellipse = find_shape_by_pos(slide, *_SC_LOOSE_ELLIPSE_POS)
    return pill, ellipse


def mark_correct_choices(slide, correct, n_rows=4):
    """Mark the correct choice(s) per the user's convention (slides 154/137):
    green fill on the row's radio ellipse + green outline on the choice pill.
    The updated template ships with DEMO marking colors on its rows (row 2
    red, row 3 green incl. a green-filled detached radio) — every row is
    first normalized back to the default light-blue/empty state, then the
    correct row(s) get the green treatment. `correct` is a 1-based index or
    a list of them (multi-answer)."""
    from pptx.dml.color import RGBColor
    targets = set(correct if isinstance(correct, (list, tuple)) else [correct])
    for i in range(1, n_rows + 1):
        pill, ellipse = _row_pill_and_ellipse(slide, i)
        if i in targets:
            set_shape_line_color(pill, CHOICE_CORRECT_FILL)
            set_shape_fill(ellipse, CHOICE_CORRECT_FILL)
            set_shape_line_color(ellipse, CHOICE_CORRECT_FILL)
        else:
            set_shape_line_color(pill, CHOICE_DEFAULT_LINE)
            ellipse.fill.background()
            set_shape_line_color(ellipse, CHOICE_DEFAULT_LINE)


def _fill_statements(slide, statements):
    """StatementAssessmentQuestion rows. MUST run after the scenario box is
    filled — before that, the scenario box also matches the statement
    placeholder text (construction-checklist.md item 11)."""
    boxes = find_all_shapes_by_text(slide, "טקסט השאלה טקסט השאלה")
    if len(boxes) != len(statements):
        raise AssertionError(
            "expected %d statement boxes, found %d" % (len(statements), len(boxes)))
    for box, stmt in zip(boxes, statements):
        set_simple_text(box, stmt, bold=False)


# The consolidated single-slide format: ONE slide per sub-part. As of the
# updated template library (02/07/2026) the template slides THEMSELVES ship
# in this format — question state on-canvas, hint group + feedback boxes
# already parked left of the canvas with their final fills. The builder
# only fills text into existing shapes; it no longer constructs the parked
# layout. Geometry below is measured from the updated library.
PARK_OFFSET_IN = 15.0  # legacy helper offset (used by _park for ad-hoc parking)
# Correct-answer marking for choice questions (user convention, slides
# 154/137 exemplars): the correct row's radio ellipse is filled green AND
# the choice pill's outline turns green. The updated template demonstrates
# this on its own rows (row 2 red / row 3 green demo) — those demo colors
# are normalized back to the default light-blue before marking.
CHOICE_CORRECT_FILL = "00B050"
CHOICE_DEFAULT_LINE = "00B0F0"
_SC_ROW_LEFT_IN = 5.53
_SC_ROW_TOPS_IN = [3.30, 4.14, 4.98, 5.82]   # tops of the 4 choice-row groups
_SC_LOOSE_ELLIPSE_POS = (12.13, 5.13)  # row 3's radio sits DETACHED at top
                                        # level in the template (green demo)
# ValueInputQuestion geometry (template 23)
_VI_ANSWER_BOX_POS = (7.08, 2.87)      # empty answer box (no placeholder text)
_VI_PERCENT_DEMO = [(6.77, 5.58), (9.28, 5.67)]  # optional fixed-symbol demo
                                        # box + '%' text — strip unless the
                                        # question actually needs the symbol
# DropdownQuestion geometry (template 22): a closed dropdown FIELD (empty box
# + arrow icon) and 4 option cells below it (the open list)
_DD_FIELD_POS = (7.04, 3.56)
_DD_OPTION_LEFT = 7.03                 # cells drift 7.02-7.04 — use tol
_DD_OPTION_TOPS = [4.18, 4.77, 5.37, 5.96]


def set_shape_fill(shape, color_hex):
    from pptx.dml.color import RGBColor
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color_hex)


def top_level_ancestor(slide, shape):
    """The slide-level shape containing `shape` (itself, if not grouped).
    Needed to park a whole feedback GROUP given its nested text box."""
    spTree = slide.shapes._spTree
    el = shape._element
    while el.getparent() is not None and el.getparent() is not spTree:
        el = el.getparent()
    for sh in slide.shapes:
        if sh._element is el:
            return sh
    raise KeyError("top-level ancestor not found")


def copy_shapes_cross_file(src_slide, shapes, dest_slide, dest_prs):
    """Copy specific shapes from a slide in ANOTHER Presentation (e.g. the
    template library) onto dest_slide. Image relationships are copied as raw
    Parts (same technique as duplicate_slide_cross_file — see
    construction-checklist.md item 5); only rels the copied shapes actually
    reference are brought over. Returns the new shapes in order."""
    new_elems = [copy.deepcopy(sh._element) for sh in shapes]
    needed = set()
    for el in new_elems:
        for e in el.iter():
            for attr, val in e.attrib.items():
                if attr.startswith("{" + R_NS + "}"):
                    needed.add(val)
    rid_map = {}
    for rel_id, rel in src_slide.part.rels.items():
        if rel_id not in needed or rel.reltype.endswith("notesSlide"):
            continue
        if rel.is_external:
            rid_map[rel_id] = dest_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT_IMAGE:
            src_part = rel.target_part
            new_partname = dest_prs.part.package.next_image_partname(src_part.partname.ext)
            new_part = Part(new_partname, src_part.content_type, dest_prs.part.package, blob=src_part.blob)
            rid_map[rel_id] = dest_slide.part.relate_to(new_part, RT_IMAGE)
        else:
            rid_map[rel_id] = dest_slide.part.relate_to(rel.target_part, rel.reltype)
    new_shapes = []
    for el in new_elems:
        for e in el.iter():
            for attr in list(e.attrib):
                if attr.startswith("{" + R_NS + "}") and e.attrib[attr] in rid_map:
                    e.attrib[attr] = rid_map[e.attrib[attr]]
        dest_slide.shapes._spTree.append(el)
        new_shapes.append(next(sh for sh in dest_slide.shapes if sh._element is el))
    return new_shapes


def _park(shape):
    shape.left = shape.left - Inches(PARK_OFFSET_IN)


def _first_text_shape(shape_or_group):
    """The shape itself if it holds text, else the first text-bearing shape
    inside the group (depth-first) — e.g. the feedback box within a cloned
    feedback group."""
    if shape_or_group.shape_type != 6:
        return shape_or_group
    for sh in shape_or_group.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            return sh
        if sh.shape_type == 6:
            try:
                return _first_text_shape(sh)
            except KeyError:
                continue
    raise KeyError("no text shape in group")


def _autosize_feedback_box(box, intro, lines):
    """Extend a parked feedback box DOWNWARD so all its content fits (user
    correction — long feedback overflowed the template's 4.88\" box). ~26
    chars wrap per line at 20pt in the 4.56\"-wide box; a native-math
    segment line counts one extra (fractions render tall). Never shrinks."""
    n = 2 + max(1, -(-len(intro) // 26))  # intro (may wrap) + spacer
    for text, _bold in lines:
        if isinstance(text, (list, tuple)):
            joined = "".join(v for k, v in text if k == "t")
            n += max(1, -(-len(joined) // 26)) + 1
        else:
            n += max(1, -(-len(text) // 26))
    needed = n * 0.35 + 0.5
    if needed > Emu(box.height).inches:
        box.height = Inches(needed)


def _direct_child_of_group(grp, shape):
    """The DIRECT child of group `grp` that contains `shape` (possibly shape
    itself). Needed because the hint text / button may be nested another
    level down inside the hint group."""
    el = shape._element
    while el.getparent() is not None and el.getparent() is not grp._element:
        el = el.getparent()
    for k in grp.shapes:
        if k._element is el:
            return k
    raise KeyError("shape is not inside the group")


def _layout_hint_group(slide, hint_shape):
    """User correction: the hint panel must hug its content — panel height
    matches the amount of hint text, and the 'חזרה לשאלה' button sits
    DIRECTLY below the text (no dead gap when the hint is short, no overflow
    when it is long). Works in the group's child space; skipped safely when
    the group is scaled (child-EMU arithmetic would drift)."""
    from pptx.oxml.ns import qn
    grp = top_level_ancestor(slide, hint_shape)
    if grp.shape_type != 6:
        return  # hint not grouped in this template — nothing to lay out
    xfrm = grp._element.find(qn('p:grpSpPr')).find(qn('a:xfrm'))
    ext, ch_off, ch_ext = (xfrm.find(qn('a:ext')), xfrm.find(qn('a:chOff')),
                           xfrm.find(qn('a:chExt')))
    if ch_ext is None or int(ch_ext.get('cy')) == 0:
        return
    scale = int(ext.get('cy')) / int(ch_ext.get('cy'))
    if abs(scale - 1) > 0.02:
        return  # scaled group — skip rather than misplace
    kids = list(grp.shapes)
    text_kid = _direct_child_of_group(grp, hint_shape)
    try:
        btn = _direct_child_of_group(grp, find_shape_by_text(grp, "חזרה לשאלה"))
    except KeyError:
        btn = None
    # panel = the largest remaining child (the group's background rect);
    # small leftovers (icon) are left where they are
    others = [k for k in kids if k is not text_kid and k is not btn]
    panel = max(others, key=lambda s: s.width * s.height) if others else None
    # required text height: same wrap heuristic as _autosize_feedback_box
    # (~26 chars per 4.56" of width per line at 20pt), grow AND shrink
    cpl = max(8, int(Emu(hint_shape.width).inches * (26 / 4.56)))
    n = sum(max(1, -(-len(ln) // cpl))
            for ln in hint_shape.text_frame.text.split("\n"))
    needed = Inches(max(0.4, n * 0.35 + 0.15))
    if hint_shape is not text_kid:
        hint_shape.height = needed
    text_kid.height = needed
    if btn is not None:
        btn.top = text_kid.top + text_kid.height + Inches(0.08)
    if panel is not None:
        content_bottom = (btn.top + btn.height) if btn is not None \
            else (text_kid.top + text_kid.height)
        panel.height = max(Inches(0.5), content_bottom - panel.top + Inches(0.12))
    # keep the group frame hugging its children (scale stays 1:1)
    max_bottom = max(k.top + k.height for k in kids)
    new_cy = max(1, max_bottom - int(ch_off.get('y')))
    ch_ext.set('cy', str(new_cy))
    ext.set('cy', str(new_cy))


def build_question(src_prs, dest_prs, spec):
    """Build a whole question (all sub-parts) from a declarative spec. Each
    sub-part clones ONE consolidated slide from the updated template library
    (question state on-canvas; hint group + feedback boxes already parked
    left of the canvas by the template) and fills in the content.

    CONTENT FIDELITY (checklist item 13): every text in the spec — scenario,
    hint, feedback — must be copied VERBATIM from the content Word doc
    (extracted via docx_extract.py, which reads the embedded equations).
    Never rephrase, recompute or "improve" authored pedagogy.

    NOTES: the built slide's speaker notes are copied VERBATIM from the
    template slide's notes (they carry the official production
    instructions). Nothing is auto-generated and nothing may be added —
    'notes_append' is allowed ONLY for production text that itself comes
    verbatim from the content Word doc.

    spec = {
        'section':       'תרגול מתקדם שאלה 6',
        'after_section': 'תרגול מתקדם שאלה 5',   # present -> NEW Section placed
                                                  # after it; absent -> Section
                                                  # must already exist
        'subparts': [{
            'type':        'ValueInputQuestion',  # a TEMPLATES key (wired types
                                                  # only: ValueInput/SingleChoice/
                                                  # StatementAssessment)
            'scenario':    ['line 1', 'line 2'],  # one paragraph per string
            'instruction': 'א. כמה ... ?',        # gets bold+purple
            'hint':        'טקסט הרמז',
            'feedback_positive': 'נכון!',          # FIXED greeting, not from the
                                                  # Word doc's "משוב:" text
            'feedback_negative': 'טעית, ...',      # an OPENER from feedback-bank.md,
                                                  # not the Word explanation itself
            # NEVER dump the whole Word "משוב:" sentence into
            # feedback_positive/feedback_negative/feedback_intro — it always
            # renders bold (make_feedback_paras bolds the intro paragraph) and
            # skips the required bank opener (checklist item 18). The actual
            # verbatim explanation goes in feedback_lines below, bold only on
            # the invariant-fact/final-result line (content-style-rules #8).
            'feedback_lines': [('detail', False), ('result', True)],  # optional
            # by type:
            'answer':     '140',                  # ValueInputQuestion
            'percent_field': True,                # ValueInput only, optional:
                                                  # KEEP the fixed-symbol (%)
                                                  # field demo (default: strip)
            'choices':    ['א', 'ב', 'ג'],        # SingleChoiceQuestion...
            'correct':    2,                      # ...+ REQUIRED 1-based index
                                                  # (or list): green radio fill
                                                  # + green pill outline
            'statements': ['....', '....'],      # StatementAssessmentQuestion
            # optional extras:
            'table': {'rows': [...], 'pos': (l, t, w, h), 'size_pt': 14},
                # build_rtl_table on the slide. NOTE: write rows in REVERSED
                # column order vs logical RTL reading (checklist item 8)
            'notes_append': '...',                # verbatim source-doc
                                                  # production text ONLY
        }, ...],
    }

    Returns list of the new slides, one per sub-part.
    """
    section = spec["section"]
    after = spec.get("after_section")
    built = []
    for i, part in enumerate(spec["subparts"]):
        ptype = part["type"]
        tpl_idx = TEMPLATES[ptype]
        if i == 0 and after:
            anchor = current_anchor(dest_prs, after)
            slides, _ = assemble_question(src_prs, [tpl_idx], dest_prs, anchor,
                                          section, existing_section=False,
                                          after_section_name=after)
        else:
            anchor = current_anchor(dest_prs, section)
            slides, _ = assemble_question(src_prs, [tpl_idx], dest_prs, anchor,
                                          section, existing_section=True)
        (slide,) = slides

        # question text + answered state, on-canvas. The question font stays
        # at the standard 20pt — no auto-shrinking (user correction); long
        # scenarios simply flow further down the slide.
        replace_paragraphs(find_shape_by_text(slide, "טקסט השאלה"),
                           make_scenario_paras(part["scenario"], part["instruction"]))
        if ptype == "ValueInputQuestion":
            answer_box = find_shape_by_pos(slide, *_VI_ANSWER_BOX_POS)
            replace_paragraphs(answer_box,
                               [make_para_xml([(part["answer"], True, None)], center=True)])
            # the template places the box for a SHORT question text; a long
            # scenario flows past it and swallows the box — drop it below
            # the estimated text bottom (user correction, peak-B slides)
            est_lines = 2  # instruction + spacer
            for s in part["scenario"]:
                est_lines += max(1, -(-len(s) // 55))
            est_bottom = 1.27 + est_lines * 0.34
            answer_box.top = Inches(min(max(2.87, est_bottom + 0.15), 5.55))
            if not part.get("percent_field"):
                strip_percent_symbol_demo(slide)
        elif ptype == "SingleChoiceQuestion":
            _fill_choices(slide, part["choices"])
            mark_correct_choices(slide, part["correct"])
        elif ptype == "StatementAssessmentQuestion":
            _fill_statements(slide, part["statements"])
        elif ptype == "DropdownQuestion":
            # answered state: the closed dropdown FIELD shows the correct
            # option; the open-list cells hold all options, the correct one
            # outlined green (the template's 'לסמן בירוק' instruction)
            options = part["options"]
            field = find_shape_by_pos(slide, *_DD_FIELD_POS, tol=0.08)
            dd_cells = [find_shape_by_pos(slide, _DD_OPTION_LEFT, top, tol=0.08)
                        for top in _DD_OPTION_TOPS]
            replace_paragraphs(field, [make_para_xml(
                [(options[part["correct"] - 1], True, None)], center=True)])
            for j, cell in enumerate(dd_cells):
                if j < len(options):
                    set_simple_text(cell, options[j], bold=False)
                    if j == part["correct"] - 1:
                        set_shape_line_color(cell, CHOICE_CORRECT_FILL)
                else:
                    set_simple_text(cell, "", bold=False)
            # position (user correction): like the ValueInput answer box, the
            # template places the dropdown for a SHORT question text — a long
            # scenario flows past it. Drop the WHOLE stack (field + open-list
            # cells, keeping their relative offsets) below the estimated text
            # bottom. The open list MAY extend past the slide's bottom edge —
            # it is a production annotation of the open state, and the user
            # explicitly allows it to overflow the slide bounds (so never
            # squeeze the cells to force them onto the canvas).
            est_lines = 2  # instruction + spacer
            for s in part["scenario"]:
                est_lines += max(1, -(-len(s) // 55))
            dd_delta = max(0.0, (1.27 + est_lines * 0.34 + 0.15) - _DD_FIELD_POS[1])
            if dd_delta:
                for sh in [field] + dd_cells:
                    sh.top = sh.top + Inches(dd_delta)
        else:
            raise KeyError("template type not wired in build_question: " + ptype)
        if part.get("table"):
            t = part["table"]
            build_rtl_table(slide, t["rows"], *t["pos"],
                            header_bold=t.get("header_bold", True),
                            size_pt=t.get("size_pt", 14))

        # hint: the template's parked hint group already exists — fill, then
        # lay out. NOT bold (user correction): the template placeholder run
        # is bold, so the override matters. _layout_hint_group sizes the
        # panel to the text and puts 'חזרה לשאלה' directly below it
        # (user correction).
        hint_shape = find_shape_by_text(slide, "טקסט רמז")
        set_simple_text(hint_shape, part["hint"], bold=False)
        _layout_hint_group(slide, hint_shape)

        # feedback: the template's parked boxes already exist with their
        # official fills — fill text only. SingleChoice/ValueInput have two
        # separate boxes (positive/negative); StatementAssessment ships ONE
        # combined box using the slash convention. The extra 'משוב לאי
        # הצלחה ראשון' box (first-attempt default) is template content and
        # is left untouched.
        if "feedback_positive" in part:
            pos_intro, neg_intro = part["feedback_positive"], part["feedback_negative"]
        elif " / " in part.get("feedback_intro", ""):
            pos_intro, neg_intro = part["feedback_intro"].split(" / ", 1)
        else:
            pos_intro = neg_intro = part["feedback_intro"]
        lines = part.get("feedback_lines", ())
        try:
            pos_box = find_shape_by_text(slide, "משוב להצלחה (למשל")
        except KeyError:
            pos_box = None
        if pos_box is not None and "משוב לאי הצלחה" in pos_box.text_frame.text:
            # combined single box (StatementAssessmentQuestion)
            combined = "%s / %s" % (pos_intro, neg_intro)
            replace_paragraphs(pos_box, make_feedback_paras(combined, lines))
            _autosize_feedback_box(pos_box, combined, lines)
        elif pos_box is not None:
            replace_paragraphs(pos_box, make_feedback_paras(pos_intro, lines))
            _autosize_feedback_box(pos_box, pos_intro, lines)
            neg_box = find_shape_by_text(slide, "משוב לאי הצלחה (למשל")
            replace_paragraphs(neg_box, make_feedback_paras(neg_intro, lines))
            _autosize_feedback_box(neg_box, neg_intro, lines)

        # speaker notes: VERBATIM copy of the template slide's notes, as RAW
        # XML paragraphs — plain-text copying loses RTL direction and the
        # mixed Hebrew/Latin lines render jumbled (user correction)
        copy_notes_verbatim(src_prs.slides[tpl_idx], slide,
                            append_text=part.get("notes_append"))
        built.append(slide)
    return built


def verify_question_build(path, section_name, before_count, expected_delta):
    """One-call post-build verification (construction-checklist.md 'אימות
    סופי'): zip integrity, slide-count delta, section membership, residue
    scan on the section's slides. The deck-wide 'מקום לגיף' convention
    (deliberately-unfilled gif cue on transition screens) is reported
    separately, not as residue. report['ok'] == True means all clear."""
    import collections
    import zipfile
    from pptx import Presentation as _P
    prs = _P(path)
    sections = list_sections(prs)
    sec = next((s for s in sections if s[0] == section_name), None)
    residue = scan_placeholder_residue(prs, [p - 1 for p in sec[1]]) if sec else []
    zip_names = zipfile.ZipFile(path).namelist()
    report = {
        "zip_ok": verify_zip_integrity(path),
        # two zip entries with the same name = two live parts fighting over
        # one partname; PowerPoint keeps an arbitrary one (checklist item 12)
        "duplicate_zip_entries": [n for n, c in collections.Counter(zip_names).items() if c > 1],
        "slide_count": len(prs.slides),
        "delta_ok": len(prs.slides) - before_count == expected_delta,
        "section_positions": sec[1] if sec else None,
        "residue": [r for r in residue if r[2] != "מקום לגיף"],
        "gif_convention_hits": [r for r in residue if r[2] == "מקום לגיף"],
    }
    report["ok"] = (report["zip_ok"] and report["delta_ok"] and sec is not None
                    and not report["residue"] and not report["duplicate_zip_entries"])
    return report


# --------------------------------------------------------------------------
# Verification (see construction-checklist.md)
# --------------------------------------------------------------------------

def verify_zip_integrity(path):
    import zipfile
    return zipfile.ZipFile(path).testzip() is None


def verify_slide_count_delta(path_before_count, prs_after):
    return len(prs_after.slides) - path_before_count
