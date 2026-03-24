from html import escape
import base64
import os
from core.cv_formatters import skill_strings, language_strings, contact_parts, premium_contact_lines


def _img_to_base64(path: str) -> str:
    if not path or not os.path.exists(path):
        return ""

    ext = os.path.splitext(path)[1].lower().replace(".", "")
    if ext == "jpg":
        ext = "jpeg"

    try:
        with open(path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
        return f"data:image/{ext};base64,{encoded}"
    except Exception:
        return ""


def _list_items(items):
    if not items:
        return ""
    return "".join(
        f"<li>{escape(str(item))}</li>"
        for item in items
        if str(item).strip()
    )


def _simple_section(title, items):
    if not items:
        return ""
    lis = _list_items(items)
    return f"""
    <div class="cv-section">
        <div class="cv-section-title">{escape(title)}</div>
        <ul class="cv-list">{lis}</ul>
    </div>
    """


def _standard_experience_blocks(experience):
    html = ""
    for e in experience:
        details = _list_items(e.details)
        subtitle = " | ".join([x for x in [e.company, e.location] if x])
        html += f"""
        <div class="cv-item">
            <div class="cv-item-top">
                <div class="cv-item-title">{escape(e.title)}</div>
                <div class="cv-item-period">{escape(e.period)}</div>
            </div>
            <div class="cv-item-sub">{escape(subtitle)}</div>
            {"<ul class='cv-list'>" + details + "</ul>" if details else ""}
        </div>
        """
    return html


def _standard_education_blocks(education):
    html = ""
    for e in education:
        details = _list_items(e.details)
        subtitle = " | ".join([x for x in [e.school, e.location] if x])
        html += f"""
        <div class="cv-item">
            <div class="cv-item-top">
                <div class="cv-item-title">{escape(e.degree)}</div>
                <div class="cv-item-period">{escape(e.period)}</div>
            </div>
            <div class="cv-item-sub">{escape(subtitle)}</div>
            {"<ul class='cv-list'>" + details + "</ul>" if details else ""}
        </div>
        """
    return html


def _premium_entry(title, subtitle, period, location, details):
    detail_html = _list_items(details)
    return f"""
    <div class="gp-entry">
        <div class="gp-left">
            <div class="gp-title">{escape(title)}</div>
            {f'<div class="gp-subtitle">{escape(subtitle.upper())}</div>' if subtitle else ''}
            {"<ul class='gp-list'>" + detail_html + "</ul>" if detail_html else ""}
        </div>
        <div class="gp-right">
            {f'<div class="gp-period">{escape(period)}</div>' if period else ''}
            {f'<div class="gp-location">{escape(location)}</div>' if location else ''}
        </div>
    </div>
    """


def _premium_section(title, body_html):
    if not body_html:
        return ""
    return f"""
    <div class="gp-section">
        <div class="gp-section-head">
            <div class="gp-section-title">{escape(title)}</div>
            <div class="gp-line"></div>
        </div>
        {body_html}
    </div>
    """


def _render_german_premium(profile, theme):
    photo_data = _img_to_base64(profile.photo_path)

    parts = profile.name.strip().split()
    first = parts[0] if parts else "Beispiel"
    rest = " ".join(parts[1:]) if len(parts) > 1 else "Name"

    line1, line2 = premium_contact_lines(profile)

    summary_html = ""
    if profile.summary.strip():
        summary_html = _premium_section(
            "Kurzprofil",
            f'<div class="gp-summary">{escape(profile.summary)}</div>'
        )

    exp_html = ""
    if profile.experience:
        blocks = ""
        for e in profile.experience:
            blocks += _premium_entry(
                e.title,
                e.company,
                e.period,
                e.location,
                e.details
            )
        exp_html = _premium_section("Berufliche Erfahrung", blocks)

    edu_html = ""
    if profile.education:
        blocks = ""
        for e in profile.education:
            blocks += _premium_entry(
                e.degree,
                e.school,
                e.period,
                e.location,
                e.details
            )
        edu_html = _premium_section("Bildungsweg", blocks)

    proj_html = ""
    if profile.projects:
        blocks = "".join(_premium_entry(p, "", "", "", []) for p in profile.projects)
        proj_html = _premium_section("Projekte", blocks)

    cert_html = ""
    if profile.certificates:
        blocks = "".join(_premium_entry(c, "", "", "", []) for c in profile.certificates)
        cert_html = _premium_section("Zertifikate", blocks)

    skills_html = ""
    if profile.languages or profile.skills:
        row_html = ""

        langs = language_strings(profile.languages)
        if langs:
            row_html += f"""
            <div class="gp-skill-row">
                <div class="gp-skill-label">Sprachen</div>
                <div class="gp-skill-value">{escape("  |  ".join(langs))}</div>
            </div>
            """

        skills = skill_strings(profile.skills)
        if skills:
            row_html += f"""
            <div class="gp-skill-row">
                <div class="gp-skill-label">Techstack</div>
                <div class="gp-skill-value">{escape("  |  ".join(skills))}</div>
            </div>
            """

        skills_html = _premium_section("Fähigkeiten", row_html)

    photo_html = f'<img src="{photo_data}" class="gp-photo" alt="Foto">' if photo_data else ""

    return f"""
    <style>
    body {{
        background: {theme["bg"]};
        font-family: Arial, sans-serif;
    }}

    .gp-wrap {{
        max-width: 920px;
        margin: 0 auto;
        background: #ffffff;
        border: 1px solid {theme["border"]};
        border-radius: 20px;
        padding: 28px 34px;
        box-shadow: 0 10px 35px rgba(0,0,0,0.06);
        color: #111827;
    }}

    .gp-header {{
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 24px;
        margin-bottom: 14px;
    }}

    .gp-name {{
        font-size: 32px;
        line-height: 1.1;
        margin-bottom: 8px;
        color: #111827;
    }}

    .gp-name-light {{
        font-weight: 300;
        color: #374151;
    }}

    .gp-name-bold {{
        font-weight: 800;
        color: #000000;
    }}

    .gp-job {{
        font-size: 15px;
        color: #4B5563;
        margin-bottom: 8px;
    }}

    .gp-contact {{
        font-size: 14px;
        color: #374151;
        line-height: 1.55;
    }}

    .gp-photo {{
        width: 120px;
        height: 120px;
        object-fit: cover;
        border-radius: 999px;
        border: 4px solid {theme["accent"]};
        box-shadow: 0 8px 18px rgba(0,0,0,0.08);
    }}

    .gp-section {{
        margin-top: 18px;
    }}

    .gp-section-head {{
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 8px;
    }}

    .gp-section-title {{
        font-size: 16px;
        font-weight: 800;
        color: #111827;
        white-space: nowrap;
    }}

    .gp-line {{
        height: 2px;
        background: #6B7280;
        width: 100%;
        opacity: 0.8;
    }}

    .gp-summary {{
        font-size: 14px;
        line-height: 1.6;
        color: #374151;
        margin-bottom: 8px;
    }}

    .gp-entry {{
        display: flex;
        justify-content: space-between;
        gap: 20px;
        margin-bottom: 14px;
    }}

    .gp-left {{
        width: 76%;
    }}

    .gp-right {{
        width: 24%;
        text-align: right;
    }}

    .gp-title {{
        font-size: 15px;
        font-weight: 800;
        color: #111827;
        margin-bottom: 2px;
    }}

    .gp-subtitle {{
        font-size: 11px;
        color: #6B7280;
        letter-spacing: 0.3px;
        margin-bottom: 4px;
    }}

    .gp-period {{
        font-size: 13px;
        font-style: italic;
        font-weight: 700;
        color: #374151;
        margin-top: 2px;
    }}

    .gp-location {{
        font-size: 12px;
        font-style: italic;
        color: #9CA3AF;
        margin-top: 2px;
    }}

    .gp-list {{
        margin: 0;
        padding-left: 18px;
        color: #374151;
        font-size: 14px;
        line-height: 1.55;
    }}

    .gp-list li {{
        margin-bottom: 3px;
    }}

    .gp-skill-row {{
        display: flex;
        gap: 14px;
        margin-bottom: 6px;
        align-items: flex-start;
    }}

    .gp-skill-label {{
        min-width: 95px;
        font-weight: 800;
        color: #111827;
        font-size: 14px;
    }}

    .gp-skill-value {{
        color: #374151;
        font-size: 14px;
        line-height: 1.5;
    }}
    </style>

    <div class="gp-wrap">
        <div class="gp-header">
            <div>
                <div class="gp-name">
                    <span class="gp-name-light">{escape(first)} </span>
                    <span class="gp-name-bold">{escape(rest)}</span>
                </div>
                {f'<div class="gp-job">{escape(profile.job_title)}</div>' if profile.job_title.strip() else ''}
                {f'<div class="gp-contact">{escape(line1)}</div>' if line1 else ''}
                {f'<div class="gp-contact">{escape(line2)}</div>' if line2 else ''}
            </div>
            {photo_html}
        </div>

        {summary_html}
        {exp_html}
        {edu_html}
        {proj_html}
        {cert_html}
        {skills_html}
    </div>
    """


def render_cv_preview(profile, theme, template_name):
    if template_name == "German Premium":
        return _render_german_premium(profile, theme)

    contact = " | ".join(contact_parts(profile))
    photo_data = _img_to_base64(profile.photo_path)

    photo_html = ""
    if photo_data and template_name in ["Photo Classic", "Modern"]:
        photo_html = f'<img src="{photo_data}" class="cv-photo" alt="Foto">'

    if template_name == "ATS":
        photo_html = ""

    summary_html = ""
    if profile.summary.strip():
        summary_html = f"""
        <div class="cv-section">
            <div class="cv-section-title">Kurzprofil</div>
            <div class="cv-text">{escape(profile.summary)}</div>
        </div>
        """

    exp_html = ""
    if profile.experience:
        exp_html = f"""
        <div class="cv-section">
            <div class="cv-section-title">Berufserfahrung</div>
            {_standard_experience_blocks(profile.experience)}
        </div>
        """

    edu_html = ""
    if profile.education:
        edu_html = f"""
        <div class="cv-section">
            <div class="cv-section-title">Ausbildung</div>
            {_standard_education_blocks(profile.education)}
        </div>
        """

    skills_html = _simple_section("Kenntnisse", skill_strings(profile.skills))
    languages_html = _simple_section("Sprachen", language_strings(profile.languages))
    projects_html = _simple_section("Projekte", profile.projects)
    certs_html = _simple_section("Zertifikate", profile.certificates)

    return f"""
    <style>
    body {{
        background: {theme["bg"]};
    }}

    .cv-preview {{
        background: {theme["card"]};
        color: {theme["text"]};
        border: 1px solid {theme["border"]};
        border-radius: 20px;
        padding: 34px 38px;
        box-shadow: 0 10px 35px rgba(0,0,0,0.06);
        font-family: Arial, sans-serif;
        max-width: 900px;
        margin: 0 auto;
    }}

    .cv-header {{
        display: flex;
        justify-content: space-between;
        gap: 24px;
        align-items: flex-start;
        border-bottom: 3px solid {theme["primary"]};
        padding-bottom: 16px;
        margin-bottom: 20px;
    }}

    .cv-name {{
        font-size: 30px;
        font-weight: 800;
        color: {theme["secondary"]};
        letter-spacing: 0.2px;
    }}

    .cv-role {{
        margin-top: 6px;
        font-size: 16px;
        color: {theme["primary"]};
        font-weight: 600;
    }}

    .cv-contact {{
        margin-top: 10px;
        font-size: 14px;
        color: {theme["muted"]};
        line-height: 1.5;
    }}

    .cv-photo {{
        width: 110px;
        height: 110px;
        object-fit: cover;
        border-radius: 14px;
        border: 2px solid {theme["border"]};
    }}

    .cv-section {{
        margin-top: 18px;
    }}

    .cv-section-title {{
        font-size: 16px;
        font-weight: 800;
        color: {theme["secondary"]};
        margin-bottom: 10px;
        padding: 8px 12px;
        background: {theme["accent"]};
        border-left: 5px solid {theme["primary"]};
        border-radius: 8px;
    }}

    .cv-text {{
        font-size: 14px;
        line-height: 1.65;
        color: {theme["text"]};
    }}

    .cv-item {{
        margin-bottom: 14px;
        padding: 10px 4px 2px 4px;
    }}

    .cv-item-top {{
        display: flex;
        justify-content: space-between;
        gap: 16px;
        flex-wrap: wrap;
    }}

    .cv-item-title {{
        font-size: 15px;
        font-weight: 700;
        color: {theme["secondary"]};
    }}

    .cv-item-period {{
        font-size: 13px;
        color: {theme["primary"]};
        font-weight: 700;
    }}

    .cv-item-sub {{
        font-size: 13px;
        color: {theme["muted"]};
        margin-top: 4px;
        margin-bottom: 6px;
        font-style: italic;
    }}

    .cv-list {{
        margin: 0;
        padding-left: 18px;
        line-height: 1.6;
        font-size: 14px;
    }}

    .cv-list li {{
        margin-bottom: 4px;
    }}
    </style>

    <div class="cv-preview">
        <div class="cv-header">
            <div>
                <div class="cv-name">{escape(profile.name or "Dein Name")}</div>
                <div class="cv-role">{escape(profile.job_title or "Berufsbezeichnung")}</div>
                <div class="cv-contact">{escape(contact)}</div>
            </div>
            {photo_html}
        </div>

        {summary_html}
        {exp_html}
        {edu_html}
        {skills_html}
        {languages_html}
        {projects_html}
        {certs_html}
    </div>
    """