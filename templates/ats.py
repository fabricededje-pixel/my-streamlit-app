from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from core.cv_formatters import skill_strings, language_strings, contact_parts


def _set_layout(doc):
    section = doc.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")


def _text(doc, text, bold=False, size=11, after=2):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(after)


def _section(doc, title):
    _text(doc, title.upper(), bold=True, size=11.5, after=2)


def build_ats_cv(profile):
    doc = Document()
    _set_layout(doc)

    _text(doc, profile.name, bold=True, size=18, after=1)

    if profile.job_title.strip():
        _text(doc, profile.job_title.strip(), size=11.5, after=2)

    parts = contact_parts(profile)
    if parts:
        _text(doc, " | ".join(parts), size=10.5, after=4)

    if profile.summary.strip():
        _section(doc, "Professional Summary")
        _text(doc, profile.summary.strip(), size=11, after=3)

    if profile.experience:
        _section(doc, "Work Experience")
        for e in profile.experience:
            title_line = " | ".join([x for x in [e.title.strip(), e.company.strip(), e.location.strip()] if x])
            if title_line:
                _text(doc, title_line, bold=True, size=11, after=1)
            if e.period.strip():
                _text(doc, e.period.strip(), size=10.5, after=1)
            for d in e.details:
                _text(doc, f"- {d}", size=10.8, after=0)

    if profile.education:
        _section(doc, "Education")
        for e in profile.education:
            degree_line = " | ".join([x for x in [e.degree.strip(), e.school.strip(), e.location.strip()] if x])
            if degree_line:
                _text(doc, degree_line, bold=True, size=11, after=1)
            if e.period.strip():
                _text(doc, e.period.strip(), size=10.5, after=1)
            for d in e.details:
                _text(doc, f"- {d}", size=10.8, after=0)

    if profile.skills:
        _section(doc, "Skills")
        for s in skill_strings(profile.skills):
            _text(doc, f"- {s}", size=10.8, after=0)

    if profile.languages:
        _section(doc, "Languages")
        for s in language_strings(profile.languages):
            _text(doc, f"- {s}", size=10.8, after=0)

    if profile.projects:
        _section(doc, "Projects")
        for s in profile.projects:
            _text(doc, f"- {s}", size=10.8, after=0)

    if profile.certificates:
        _section(doc, "Certificates")
        for s in profile.certificates:
            _text(doc, f"- {s}", size=10.8, after=0)

    return doc