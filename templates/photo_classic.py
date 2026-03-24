from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _set_layout(doc):
    section = doc.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")


def _remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'nil')
        borders.append(border)
    tblPr.append(borders)


def _separator(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    p_pr = p._element.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "4F81BD")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _heading(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(13.5)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)


def _text(doc, text, bold=False, italic=False, size=10.5, after=2):
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
            r.font.size = Pt(10.5)
            p.paragraph_format.space_after = Pt(1)


def _add_header_with_photo(doc, profile):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _remove_table_borders(table)

    left = table.cell(0, 0)
    right = table.cell(0, 1)

    left.width = Cm(12.5)
    right.width = Cm(4.0)

    p = left.paragraphs[0]
    r = p.add_run(profile.name)
    r.bold = True
    r.font.size = Pt(21)
    p.paragraph_format.space_after = Pt(2)

    if profile.job_title.strip():
        p2 = left.add_paragraph()
        r2 = p2.add_run(profile.job_title.strip())
        r2.font.size = Pt(11.5)
        p2.paragraph_format.space_after = Pt(4)

    parts = [profile.email, profile.phone, profile.city, profile.linkedin]
    parts = [x for x in parts if x and str(x).strip()]
    if parts:
        p3 = left.add_paragraph()
        r3 = p3.add_run(" | ".join(parts))
        r3.font.size = Pt(10.2)

    if profile.photo_path:
        try:
            p_img = right.paragraphs[0]
            run_img = p_img.add_run()
            run_img.add_picture(profile.photo_path, width=Cm(3.2))
        except Exception:
            pass


def _exp(doc, e):
    if e.title.strip():
        _text(doc, e.title.strip(), bold=True, size=11.5, after=1)

    second = " | ".join([x for x in [e.company.strip(), e.location.strip()] if x])
    if second:
        _text(doc, second, italic=True, size=10.5, after=1)

    if e.period.strip():
        _text(doc, e.period.strip(), size=10.2, after=2)

    _bullets(doc, e.details)


def _edu(doc, e):
    if e.degree.strip():
        _text(doc, e.degree.strip(), bold=True, size=11.5, after=1)

    second = " | ".join([x for x in [e.school.strip(), e.location.strip()] if x])
    if second:
        _text(doc, second, italic=True, size=10.5, after=1)

    if e.period.strip():
        _text(doc, e.period.strip(), size=10.2, after=2)

    _bullets(doc, e.details)


def build_photo_classic_cv(profile):
    doc = Document()
    _set_layout(doc)

    _add_header_with_photo(doc, profile)
    _separator(doc)

    if profile.summary.strip():
        _heading(doc, "Kurzprofil")
        _text(doc, profile.summary.strip(), after=3)

    if profile.experience:
        _heading(doc, "Berufserfahrung")
        for e in profile.experience:
            _exp(doc, e)

    if profile.education:
        _heading(doc, "Ausbildung")
        for e in profile.education:
            _edu(doc, e)

    if profile.skills:
        _heading(doc, "Kenntnisse")
        _bullets(doc, profile.skills)

    if profile.languages:
        _heading(doc, "Sprachen")
        _bullets(doc, profile.languages)

    if profile.projects:
        _heading(doc, "Projekte")
        _bullets(doc, profile.projects)

    if profile.certificates:
        _heading(doc, "Zertifikate")
        _bullets(doc, profile.certificates)

    return doc