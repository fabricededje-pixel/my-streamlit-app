from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from core.cv_formatters import skill_strings, language_strings, premium_contact_lines


def _set_layout(doc):
    section = doc.sections[0]
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10)
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


def _add_run(paragraph, text, *, bold=False, italic=False, size=10, color=None):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = "Arial"
    run.font.size = Pt(size)
    if color:
        from docx.shared import RGBColor
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def _header(doc, profile):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _remove_table_borders(table)

    left = table.cell(0, 0)
    right = table.cell(0, 1)

    left.width = Cm(13.5)
    right.width = Cm(3.5)

    parts = profile.name.strip().split()
    first = parts[0] if parts else "Beispiel"
    rest = " ".join(parts[1:]) if len(parts) > 1 else "Name"

    p_name = left.paragraphs[0]
    p_name.paragraph_format.space_after = Pt(5)
    _add_run(p_name, first + " ", size=26, color="374151")
    _add_run(p_name, rest, bold=True, size=26, color="000000")

    if profile.job_title.strip():
        p_job = left.add_paragraph()
        p_job.paragraph_format.space_after = Pt(5)
        _add_run(p_job, profile.job_title.strip(), size=11.5, color="4B5563")

    line1, line2 = premium_contact_lines(profile)

    if line1:
        p1 = left.add_paragraph()
        p1.paragraph_format.space_after = Pt(1)
        _add_run(p1, line1, size=10.5, color="374151")

    if line2:
        p2 = left.add_paragraph()
        p2.paragraph_format.space_after = Pt(0)
        _add_run(p2, line2, size=10.5, color="374151")

    if profile.photo_path:
        try:
            p_img = right.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run_img = p_img.add_run()
            run_img.add_picture(profile.photo_path, width=Cm(3.0))
        except Exception:
            pass


def _section_title(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)

    r = p.add_run(title)
    r.bold = True
    r.font.name = "Arial"
    r.font.size = Pt(15)

    p_pr = p._element.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), "6B7280")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _entry_block(doc, title, subtitle, details, period="", location=""):
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_after = Pt(1)

    left_text = title.strip()
    right_text = " | ".join([x for x in [period.strip(), location.strip()] if x])

    r1 = p1.add_run(left_text)
    r1.bold = True
    r1.font.name = "Arial"
    r1.font.size = Pt(11.5)

    if right_text:
        r2 = p1.add_run(f"    {right_text}")
        r2.italic = True
        r2.font.name = "Arial"
        r2.font.size = Pt(10)

    if subtitle:
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_after = Pt(3)
        r_sub = p2.add_run(subtitle.upper())
        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(9)

    for item in details:
        if str(item).strip():
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(str(item))
            r.font.name = "Arial"
            r.font.size = Pt(10)


def _summary_block(doc, text):
    if not text.strip():
        return
    _section_title(doc, "Kurzprofil")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(10.5)


def _skills_line(doc, label, values):
    if not values:
        return
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)

    r1 = p.add_run(f"{label}: ")
    r1.bold = True
    r1.font.name = "Arial"
    r1.font.size = Pt(10.5)

    r2 = p.add_run(" | ".join(values))
    r2.font.name = "Arial"
    r2.font.size = Pt(10)


def build_german_premium_cv(profile):
    doc = Document()
    _set_layout(doc)

    _header(doc, profile)

    if profile.summary.strip():
        _summary_block(doc, profile.summary.strip())

    if profile.experience:
        _section_title(doc, "Berufliche Erfahrung")
        for e in profile.experience:
            _entry_block(
                doc,
                title=e.title,
                subtitle=e.company or "",
                details=e.details,
                period=e.period,
                location=e.location
            )

    if profile.education:
        _section_title(doc, "Bildungsweg")
        for e in profile.education:
            _entry_block(
                doc,
                title=e.degree,
                subtitle=e.school or "",
                details=e.details,
                period=e.period,
                location=e.location
            )

    if profile.projects:
        _section_title(doc, "Projekte")
        for p in profile.projects:
            _entry_block(doc, str(p), "", [], "", "")

    if profile.certificates:
        _section_title(doc, "Zertifikate")
        for c in profile.certificates:
            _entry_block(doc, str(c), "", [], "", "")

    if profile.languages or profile.skills:
        _section_title(doc, "Fähigkeiten")

        langs = language_strings(profile.languages)
        if langs:
            _skills_line(doc, "Sprachen", langs)

        skills = skill_strings(profile.skills)
        if skills:
            _skills_line(doc, "Techstack", skills)

    return doc