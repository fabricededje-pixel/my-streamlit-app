from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from core.cv_formatters import skill_strings, language_strings, contact_parts


def _set_layout(doc):
    section = doc.sections[0]
    section.top_margin = Cm(1.1)
    section.bottom_margin = Cm(1.1)
    section.left_margin = Cm(1.3)
    section.right_margin = Cm(1.3)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(9.8)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")


def _p(doc, text, bold=False, size=9.8, italic=False, after=1):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(after)


def _section(doc, title):
    _p(doc, title, bold=True, size=10.8, after=1)


def build_compact_cv(profile):
    doc = Document()
    _set_layout(doc)

    header = profile.name
    if profile.job_title.strip():
        header += f" – {profile.job_title.strip()}"
    _p(doc, header, bold=True, size=16, after=2)

    parts = contact_parts(profile)
    if parts:
        _p(doc, " | ".join(parts), size=9.3, after=3)

    if profile.summary.strip():
        _section(doc, "Kurzprofil")
        _p(doc, profile.summary.strip(), after=2)

    if profile.experience:
        _section(doc, "Berufserfahrung")
        for e in profile.experience:
            line = " | ".join([x for x in [e.title.strip(), e.company.strip(), e.location.strip(), e.period.strip()] if x])
            if line:
                _p(doc, line, bold=True, size=9.9, after=1)
            for d in e.details:
                _p(doc, f"• {d}", size=9.6, after=0)

    if profile.education:
        _section(doc, "Ausbildung")
        for e in profile.education:
            line = " | ".join([x for x in [e.degree.strip(), e.school.strip(), e.location.strip(), e.period.strip()] if x])
            if line:
                _p(doc, line, bold=True, size=9.9, after=1)
            for d in e.details:
                _p(doc, f"• {d}", size=9.6, after=0)

    if profile.skills:
        _section(doc, "Kenntnisse")
        _p(doc, " | ".join(skill_strings(profile.skills)), size=9.7, after=1)

    if profile.languages:
        _section(doc, "Sprachen")
        _p(doc, " | ".join(language_strings(profile.languages)), size=9.7, after=1)

    if profile.projects:
        _section(doc, "Projekte")
        for p in profile.projects:
            _p(doc, f"• {p}", size=9.6, after=0)

    if profile.certificates:
        _section(doc, "Zertifikate")
        for c in profile.certificates:
            _p(doc, f"• {c}", size=9.6, after=0)

    return doc