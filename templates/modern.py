from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from core.i18n import tr
from core.cv_formatters import skill_strings, language_strings, contact_parts


def _set_layout(doc):
    section = doc.sections[0]
    section.top_margin = Cm(1.3)
    section.bottom_margin = Cm(1.3)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10.2)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")


def _remove_table_borders(table):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "nil")
        borders.append(border)
    tbl_pr.append(borders)


def _header(doc, profile):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _remove_table_borders(table)

    left = table.cell(0, 0)
    right = table.cell(0, 1)

    left.width = Cm(12.0)
    right.width = Cm(5.0)

    p = left.paragraphs[0]
    r = p.add_run(profile.name.upper())
    r.bold = True
    r.font.size = Pt(20)
    p.paragraph_format.space_after = Pt(2)

    if profile.job_title.strip():
        p2 = left.add_paragraph()
        r2 = p2.add_run(profile.job_title.strip())
        r2.font.size = Pt(11.5)
        p2.paragraph_format.space_after = Pt(4)

    parts = contact_parts(profile)
    if parts:
        p3 = left.add_paragraph()
        r3 = p3.add_run(" | ".join(parts))
        r3.font.size = Pt(10)
        p3.paragraph_format.space_after = Pt(8)

    if profile.photo_path:
        try:
            p_img = right.paragraphs[0]
            run_img = p_img.add_run()
            run_img.add_picture(profile.photo_path, width=Cm(3.3))
        except Exception:
            pass


def _heading(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text.upper())
    r.bold = True
    r.font.size = Pt(12.3)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)


def _line(doc, text, bold=False, italic=False, size=10.2, after=2):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(after)


def _bullets(doc, items):
    for item in items:
        if item:
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run(item)
            r.font.size = Pt(10.2)
            p.paragraph_format.space_after = Pt(1)


def _exp(doc, e):
    top = e.title.strip()
    if e.period.strip():
        top = f"{top}  ({e.period.strip()})" if top else e.period.strip()
    if top:
        _line(doc, top, bold=True, size=11, after=1)

    second = " | ".join([x for x in [e.company.strip(), e.location.strip()] if x])
    if second:
        _line(doc, second, italic=True, after=1)

    _bullets(doc, e.details)


def _edu(doc, e):
    top = e.degree.strip()
    if e.period.strip():
        top = f"{top}  ({e.period.strip()})" if top else e.period.strip()
    if top:
        _line(doc, top, bold=True, size=11, after=1)

    second = " | ".join([x for x in [e.school.strip(), e.location.strip()] if x])
    if second:
        _line(doc, second, italic=True, after=1)

    _bullets(doc, e.details)


def build_modern_cv(profile):
    doc = Document()
    _set_layout(doc)

    lang = profile.language
    _header(doc, profile)

    if profile.summary.strip():
        _heading(doc, tr(lang, "summary"))
        _line(doc, profile.summary.strip(), after=3)

    if profile.experience:
        _heading(doc, tr(lang, "experience"))
        for e in profile.experience:
            _exp(doc, e)

    if profile.education:
        _heading(doc, tr(lang, "education"))
        for e in profile.education:
            _edu(doc, e)

    if profile.skills:
        _heading(doc, tr(lang, "skills"))
        _bullets(doc, skill_strings(profile.skills))

    if profile.languages:
        _heading(doc, tr(lang, "languages"))
        _bullets(doc, language_strings(profile.languages))

    if profile.projects:
        _heading(doc, tr(lang, "projects"))
        _bullets(doc, profile.projects)

    if profile.certificates:
        _heading(doc, tr(lang, "certificates"))
        _bullets(doc, profile.certificates)

    return doc