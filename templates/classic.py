from docx import Document
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from core.cv_formatters import skill_strings, language_strings, contact_parts


def _set_layout(doc):
    section = doc.sections[0]
    section.top_margin = Cm(1.7)
    section.bottom_margin = Cm(1.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")


def _add_separator(doc, color="4F81BD"):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)

    p_pr = p._element.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _add_heading(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(13.5)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)


def _add_text(doc, text, bold=False, size=10.5, italic=False, space_after=2):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(space_after)
    return p


def _add_bullets(doc, items):
    for item in items:
        if item:
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run(item)
            r.font.size = Pt(10.5)
            p.paragraph_format.space_after = Pt(1)


def _add_experience(doc, entry):
    line1 = entry.title.strip()
    if line1:
        _add_text(doc, line1, bold=True, size=11.5, space_after=1)

    line2_parts = [entry.company.strip(), entry.location.strip()]
    line2 = " | ".join([x for x in line2_parts if x])
    if line2:
        _add_text(doc, line2, size=10.5, italic=True, space_after=1)

    if entry.period.strip():
        _add_text(doc, entry.period.strip(), size=10.2, italic=False, space_after=2)

    _add_bullets(doc, entry.details)


def _add_education(doc, entry):
    if entry.degree.strip():
        _add_text(doc, entry.degree.strip(), bold=True, size=11.5, space_after=1)

    line2_parts = [entry.school.strip(), entry.location.strip()]
    line2 = " | ".join([x for x in line2_parts if x])
    if line2:
        _add_text(doc, line2, size=10.5, italic=True, space_after=1)

    if entry.period.strip():
        _add_text(doc, entry.period.strip(), size=10.2, italic=False, space_after=2)

    _add_bullets(doc, entry.details)


def build_classic_cv(profile):
    doc = Document()
    _set_layout(doc)

    p = doc.add_paragraph()
    r = p.add_run(profile.name)
    r.bold = True
    r.font.size = Pt(22)
    p.paragraph_format.space_after = Pt(2)

    if profile.job_title.strip():
        _add_text(doc, profile.job_title.strip(), size=12, italic=False, space_after=3)

    parts = contact_parts(profile)
    if parts:
        _add_text(doc, " | ".join(parts), size=10.2, space_after=5)

    _add_separator(doc)

    if profile.summary.strip():
        _add_heading(doc, "Kurzprofil")
        _add_text(doc, profile.summary.strip(), size=10.5, space_after=3)

    if profile.experience:
        _add_heading(doc, "Berufserfahrung")
        for e in profile.experience:
            _add_experience(doc, e)

    if profile.education:
        _add_heading(doc, "Ausbildung")
        for e in profile.education:
            _add_education(doc, e)

    if profile.skills:
        _add_heading(doc, "Kenntnisse")
        _add_bullets(doc, skill_strings(profile.skills))

    if profile.languages:
        _add_heading(doc, "Sprachen")
        _add_bullets(doc, language_strings(profile.languages))

    if profile.projects:
        _add_heading(doc, "Projekte")
        _add_bullets(doc, profile.projects)

    if profile.certificates:
        _add_heading(doc, "Zertifikate")
        _add_bullets(doc, profile.certificates)

    return doc