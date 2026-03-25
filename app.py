import json
import os
import re
from dataclasses import asdict
from datetime import datetime

import streamlit as st
import streamlit.components.v1 as components

from core.i18n import tr
from core.preview import render_cv_preview
from core.profile_model import (
    CVProfile,
    CertificateEntry,
    EducationEntry,
    ExperienceEntry,
    LanguageEntry,
    ProjectEntry,
    SkillEntry,
)
from core.themes import THEMES
from core.translator import TRANSLATION_LANGUAGES, language_label
from core.utils import split_lines
from core.validators import validate_profile
from exporters.document_payload import (
    embed_profile_in_docx,
    embed_profile_in_pdf,
    extract_profile_from_docx,
    extract_profile_from_pdf,
)
from exporters.json_store import load_profile_from_json, profile_from_data, profile_to_dict
from exporters.pdf_exporter import convert_docx_to_pdf, is_pdf_export_supported
from templates.ats import build_ats_cv
from templates.classic import build_classic_cv
from templates.compact import build_compact_cv
from templates.german_premium import build_german_premium_cv
from templates.modern import DEFAULT_MODERN_COLORS, build_modern_cv
from templates.photo_classic import build_photo_classic_cv


TEMPLATES = {
    "Classic": build_classic_cv,
    "Modern": build_modern_cv,
    "Compact": build_compact_cv,
    "ATS": build_ats_cv,
    "Photo Classic": build_photo_classic_cv,
    "German Premium": build_german_premium_cv,
}

SKILL_LEVELS = [
    "Grundkenntnisse",
    "Gute Kenntnisse",
    "Sehr gute Kenntnisse",
    "Expertenkenntnisse",
]

LANGUAGE_LEVELS = [
    "Grundkenntnisse",
    "Gute Kenntnisse",
    "Sehr gute Kenntnisse",
    "Fließend",
    "Verhandlungssicher",
    "Muttersprache",
    "C1",
    "C2",
]

UI_LANGUAGES = ["de", "en", "fr"]
TRANSLATION_CODES = list(TRANSLATION_LANGUAGES.keys())
MONTH_OPTIONS = {
    "01": "Januar",
    "02": "Februar",
    "03": "März",
    "04": "April",
    "05": "Mai",
    "06": "Juni",
    "07": "Juli",
    "08": "August",
    "09": "September",
    "10": "Oktober",
    "11": "November",
    "12": "Dezember",
}
YEAR_OPTIONS = [str(year) for year in range(datetime.now().year + 2, 1949, -1)]


def get_dynamic_theme(theme: dict, template_name: str) -> dict:
    custom = dict(theme)

    if template_name == "German Premium":
        custom["bg"] = "#f8fafc"
        custom["accent"] = "#e5e7eb"
    elif template_name == "Modern":
        custom["bg"] = "#f5faff"
    elif template_name == "ATS":
        custom["bg"] = "#fcfcfc"
        custom["accent"] = "#f3f4f6"

    return custom


def inject_custom_css(theme: dict):
    st.markdown(
        f"""
    <style>
    .stApp {{
        background:
            linear-gradient(145deg, rgba(255,255,255,0.94) 0%, rgba(248,250,252,0.98) 38%, rgba(241,245,249,1) 100%);
        background-attachment: fixed;
    }}

    .block-container {{
        padding-top: 2.2rem;
        padding-bottom: 2rem;
        max-width: 1600px;
        position: relative;
        z-index: 1;
    }}

    .stApp::before {{
        content: "";
        position: fixed;
        inset: 0;
        background:
            radial-gradient(circle at 12% 14%, {theme["accent"]}66 0%, transparent 26%),
            radial-gradient(circle at 82% 18%, {theme["primary"]}14 0%, transparent 22%),
            radial-gradient(circle at 78% 82%, {theme["accent"]}55 0%, transparent 24%);
        opacity: 0.95;
        z-index: 0;
        pointer-events: none;
    }}

    .stApp::after {{
        content: "";
        position: fixed;
        inset: 0;
        background-image:
            linear-gradient(rgba(148,163,184,0.06) 1px, transparent 1px),
            linear-gradient(90deg, rgba(148,163,184,0.06) 1px, transparent 1px);
        background-size: 36px 36px;
        mask-image: linear-gradient(to bottom, rgba(0,0,0,0.28), rgba(0,0,0,0.06));
        -webkit-mask-image: linear-gradient(to bottom, rgba(0,0,0,0.28), rgba(0,0,0,0.06));
        opacity: 0.45;
        z-index: 0;
        pointer-events: none;
    }}

    section[data-testid="stSidebar"] {{
        background:
            linear-gradient(180deg, rgba(255,255,255,0.94) 0%, {theme["accent"]} 100%);
        border-right: 1px solid {theme["border"]};
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
    }}

    section[data-testid="stSidebar"] .block-container {{
        padding-top: 1rem;
    }}

    div[data-testid="stExpander"] {{
        border-radius: 16px !important;
        border: 1px solid {theme["border"]} !important;
        background: rgba(255,255,255,0.88) !important;
        overflow: hidden;
        box-shadow: 0 8px 24px rgba(0,0,0,0.05);
    }}

    .form-card {{
        background: rgba(255, 255, 255, 0.82);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        padding: 1rem 1rem 0.45rem 1rem;
        border-radius: 20px;
        border: 1px solid {theme["border"]};
        box-shadow: 0 12px 32px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }}

    .stButton > button {{
        border-radius: 12px !important;
        font-weight: 700 !important;
        padding: 0.58rem 1rem !important;
        border: 1px solid {theme["border"]} !important;
        background: linear-gradient(135deg, {theme["primary"]}, {theme["secondary"]}) !important;
        color: white !important;
        box-shadow: 0 10px 22px rgba(0,0,0,0.08);
        transition: all 0.18s ease-in-out;
    }}

    .stButton > button:hover {{
        transform: translateY(-1px);
        filter: brightness(1.04);
    }}

    .stDownloadButton > button {{
        border-radius: 12px !important;
        font-weight: 700 !important;
    }}

    .stTextInput > div > div > input,
    .stTextArea textarea,
    div[data-baseweb="select"] > div {{
        border-radius: 12px !important;
        border: 1px solid {theme["border"]} !important;
        background: rgba(255,255,255,0.94) !important;
    }}

    .stTextInput label,
    .stTextArea label,
    .stSelectbox label,
    .stFileUploader label {{
        font-weight: 600 !important;
    }}

    .page-title {{
        font-size: 2.1rem;
        font-weight: 800;
        margin-bottom: 0.2rem;
        color: {theme["secondary"]};
        letter-spacing: -0.02em;
        line-height: 1.2;
        padding-top: 0.2rem;
    }}

    .page-subtitle {{
        color: {theme["muted"]};
        margin-bottom: 1rem;
        font-size: 1rem;
    }}

    .preview-title {{
        font-size: 1.15rem;
        font-weight: 700;
        margin-bottom: 0.75rem;
        color: {theme["secondary"]};
    }}
    </style>
    """,
        unsafe_allow_html=True,
    )


def save_uploaded_photo(uploaded_file):
    if uploaded_file is None:
        return st.session_state.get("photo_path", "")

    os.makedirs("temp", exist_ok=True)
    ext = uploaded_file.name.split(".")[-1].lower()
    path = os.path.join("temp", f"uploaded_photo.{ext}")

    with open(path, "wb") as file_obj:
        file_obj.write(uploaded_file.read())

    return path


def parse_period_string(period: str):
    text = (period or "").strip()
    if not text:
        return "", "", "", "", False

    normalized = text.lower().replace("bis heute", "heute").replace("present", "heute").replace("current", "heute")
    matches = re.findall(r"(\d{1,2})\s*/\s*(\d{4})", normalized)

    start_month = ""
    start_year = ""
    end_month = ""
    end_year = ""

    if matches:
        if len(matches) >= 1:
            start_month, start_year = matches[0]
            start_month = start_month.zfill(2)
        if len(matches) >= 2:
            end_month, end_year = matches[1]
            end_month = end_month.zfill(2)

    is_current = "heute" in normalized

    year_matches = re.findall(r"\b(19\d{2}|20\d{2}|21\d{2})\b", normalized)
    if not start_year and year_matches:
        start_year = year_matches[0]
    if not end_year and len(year_matches) > 1:
        end_year = year_matches[1]

    return start_month, start_year, end_month, end_year, is_current


def set_period_state(prefix: str, index: int, period: str):
    start_month, start_year, end_month, end_year, is_current = parse_period_string(period)
    st.session_state[f"{prefix}_start_month_{index}"] = start_month
    st.session_state[f"{prefix}_start_year_{index}"] = start_year
    st.session_state[f"{prefix}_end_month_{index}"] = end_month
    st.session_state[f"{prefix}_end_year_{index}"] = end_year
    st.session_state[f"{prefix}_is_current_{index}"] = is_current


def format_period(prefix: str, index: int) -> str:
    start_month = (st.session_state.get(f"{prefix}_start_month_{index}", "") or "").strip()
    start_year = (st.session_state.get(f"{prefix}_start_year_{index}", "") or "").strip()
    end_month = (st.session_state.get(f"{prefix}_end_month_{index}", "") or "").strip()
    end_year = (st.session_state.get(f"{prefix}_end_year_{index}", "") or "").strip()
    is_current = bool(st.session_state.get(f"{prefix}_is_current_{index}", False))

    start_part = ""
    end_part = ""

    if start_month and start_year:
        start_part = f"{start_month}/{start_year}"
    elif start_year:
        start_part = start_year

    if is_current:
        end_part = "heute"
    elif end_month and end_year:
        end_part = f"{end_month}/{end_year}"
    elif end_year:
        end_part = end_year

    if start_part and end_part:
        return f"{start_part} - {end_part}"
    if start_part:
        return start_part
    return end_part


def render_period_selector(prefix: str, index: int, label: str):
    st.markdown(f"**{label}**")
    month_keys = [""] + list(MONTH_OPTIONS.keys())

    start_col_1, start_col_2 = st.columns(2)
    with start_col_1:
        st.selectbox(
            "Startmonat",
            month_keys,
            key=f"{prefix}_start_month_{index}",
            format_func=lambda value: MONTH_OPTIONS.get(value, "") if value else "",
        )
    with start_col_2:
        st.selectbox("Startjahr", [""] + YEAR_OPTIONS, key=f"{prefix}_start_year_{index}")

    st.checkbox("Aktuell", key=f"{prefix}_is_current_{index}")

    if not st.session_state.get(f"{prefix}_is_current_{index}", False):
        end_col_1, end_col_2 = st.columns(2)
        with end_col_1:
            st.selectbox(
                "Endmonat",
                month_keys,
                key=f"{prefix}_end_month_{index}",
                format_func=lambda value: MONTH_OPTIONS.get(value, "") if value else "",
            )
        with end_col_2:
            st.selectbox("Endjahr", [""] + YEAR_OPTIONS, key=f"{prefix}_end_year_{index}")

    return format_period(prefix, index)


def fill_session_from_profile(profile: CVProfile):
    st.session_state["name"] = profile.name
    st.session_state["job_title"] = profile.job_title
    st.session_state["email"] = profile.email
    st.session_state["phone"] = profile.phone
    st.session_state["city"] = profile.city
    st.session_state["linkedin"] = profile.linkedin
    st.session_state["summary"] = profile.summary
    st.session_state["photo_path"] = profile.photo_path
    st.session_state["cv_language"] = profile.language or "de"

    st.session_state["experience_count"] = max(1, len(profile.experience))
    st.session_state["education_count"] = max(1, len(profile.education))
    st.session_state["skill_count"] = max(1, len(profile.skills))
    st.session_state["language_count"] = max(1, len(profile.languages))
    st.session_state["project_count"] = max(1, len(profile.projects))
    st.session_state["certificate_count"] = max(1, len(profile.certificates))

    for i, skill in enumerate(profile.skills):
        if hasattr(skill, "name"):
            st.session_state[f"skill_name_{i}"] = skill.name
            st.session_state[f"skill_level_{i}"] = skill.level or "Gute Kenntnisse"
        else:
            st.session_state[f"skill_name_{i}"] = str(skill)
            st.session_state[f"skill_level_{i}"] = "Gute Kenntnisse"

    for i in range(len(profile.skills), max(st.session_state["skill_count"], 1)):
        st.session_state.setdefault(f"skill_name_{i}", "")
        st.session_state.setdefault(f"skill_level_{i}", "Gute Kenntnisse")

    for i, language in enumerate(profile.languages):
        if hasattr(language, "name"):
            st.session_state[f"language_name_{i}"] = language.name
            st.session_state[f"language_level_{i}"] = language.level or "Gute Kenntnisse"
        else:
            st.session_state[f"language_name_{i}"] = str(language)
            st.session_state[f"language_level_{i}"] = "Gute Kenntnisse"

    for i in range(len(profile.languages), max(st.session_state["language_count"], 1)):
        st.session_state.setdefault(f"language_name_{i}", "")
        st.session_state.setdefault(f"language_level_{i}", "Gute Kenntnisse")

    for i, exp in enumerate(profile.experience):
        st.session_state[f"exp_title_{i}"] = exp.title
        st.session_state[f"exp_company_{i}"] = exp.company
        st.session_state[f"exp_location_{i}"] = exp.location
        st.session_state[f"exp_period_{i}"] = exp.period
        st.session_state[f"exp_details_{i}"] = "\n".join(exp.details)
        set_period_state("exp", i, exp.period)

    for i, edu in enumerate(profile.education):
        st.session_state[f"edu_degree_{i}"] = edu.degree
        st.session_state[f"edu_school_{i}"] = edu.school
        st.session_state[f"edu_location_{i}"] = edu.location
        st.session_state[f"edu_period_{i}"] = edu.period
        st.session_state[f"edu_details_{i}"] = "\n".join(edu.details)
        set_period_state("edu", i, edu.period)

    for i, project in enumerate(profile.projects):
        if hasattr(project, "title"):
            st.session_state[f"project_title_{i}"] = project.title
            st.session_state[f"project_organization_{i}"] = project.organization
            st.session_state[f"project_summary_{i}"] = project.summary
            set_period_state("project", i, project.period)
        else:
            st.session_state[f"project_title_{i}"] = str(project)
            st.session_state[f"project_organization_{i}"] = ""
            st.session_state[f"project_summary_{i}"] = ""
            set_period_state("project", i, "")

    for i, certificate in enumerate(profile.certificates):
        if hasattr(certificate, "title"):
            st.session_state[f"certificate_title_{i}"] = certificate.title
            st.session_state[f"certificate_issuer_{i}"] = certificate.issuer
            st.session_state[f"certificate_summary_{i}"] = certificate.summary
            set_period_state("certificate", i, certificate.period)
        else:
            st.session_state[f"certificate_title_{i}"] = str(certificate)
            st.session_state[f"certificate_issuer_{i}"] = ""
            st.session_state[f"certificate_summary_{i}"] = ""
            set_period_state("certificate", i, "")

    for i in range(len(profile.experience), max(st.session_state["experience_count"], 1)):
        set_period_state("exp", i, "")

    for i in range(len(profile.education), max(st.session_state["education_count"], 1)):
        set_period_state("edu", i, "")

    for i in range(len(profile.projects), max(st.session_state["project_count"], 1)):
        st.session_state.setdefault(f"project_title_{i}", "")
        st.session_state.setdefault(f"project_organization_{i}", "")
        st.session_state.setdefault(f"project_summary_{i}", "")
        set_period_state("project", i, "")

    for i in range(len(profile.certificates), max(st.session_state["certificate_count"], 1)):
        st.session_state.setdefault(f"certificate_title_{i}", "")
        st.session_state.setdefault(f"certificate_issuer_{i}", "")
        st.session_state.setdefault(f"certificate_summary_{i}", "")
        set_period_state("certificate", i, "")


st.set_page_config(page_title="CV Builder", layout="wide")

if "experience_count" not in st.session_state:
    st.session_state.experience_count = 1
if "education_count" not in st.session_state:
    st.session_state.education_count = 1
if "skill_count" not in st.session_state:
    st.session_state.skill_count = 1
if "language_count" not in st.session_state:
    st.session_state.language_count = 1
if "project_count" not in st.session_state:
    st.session_state.project_count = 1
if "certificate_count" not in st.session_state:
    st.session_state.certificate_count = 1
if "ui_language" not in st.session_state:
    st.session_state.ui_language = "de"
if "cv_language" not in st.session_state:
    st.session_state.cv_language = "de"
if "photo_path" not in st.session_state:
    st.session_state.photo_path = ""
if "modern_header_fill" not in st.session_state:
    st.session_state.modern_header_fill = f"#{DEFAULT_MODERN_COLORS['header_fill']}"
if "modern_sidebar_fill" not in st.session_state:
    st.session_state.modern_sidebar_fill = f"#{DEFAULT_MODERN_COLORS['sidebar_fill']}"
if "modern_border_fill" not in st.session_state:
    st.session_state.modern_border_fill = f"#{DEFAULT_MODERN_COLORS['header_line']}"

if "pending_loaded_profile" in st.session_state:
    pending_profile = profile_from_data(st.session_state.pop("pending_loaded_profile"))
    fill_session_from_profile(pending_profile)

with st.sidebar:
    st.header("Einstellungen")

    ui_language = st.selectbox(
        tr(st.session_state.ui_language, "ui_language"),
        UI_LANGUAGES,
        index=UI_LANGUAGES.index(st.session_state.ui_language),
        key="ui_language",
    )

    cv_language = st.selectbox(
        tr(ui_language, "cv_language"),
        TRANSLATION_CODES,
        index=TRANSLATION_CODES.index(st.session_state.cv_language),
        format_func=language_label,
        key="cv_language",
    )

    template_name = st.selectbox(tr(ui_language, "template"), list(TEMPLATES.keys()))
    theme_name = st.selectbox("Farbschema", list(THEMES.keys()))

    if template_name == "Modern":
        st.markdown("---")
        st.subheader("Modern Farben")
        st.color_picker("Header", key="modern_header_fill")
        st.color_picker("Seitenleiste", key="modern_sidebar_fill")
        st.color_picker("Linie", key="modern_border_fill")

theme = get_dynamic_theme(THEMES[theme_name], template_name)
if template_name == "Modern":
    theme["primary"] = st.session_state.modern_header_fill
    theme["accent"] = st.session_state.modern_sidebar_fill
    theme["border"] = st.session_state.modern_border_fill
inject_custom_css(theme)

st.markdown('<div class="page-title">Lebenslauf / CV Builder</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="page-subtitle">Professionelle Lebenslaeufe mit Live-Vorschau.</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("---")
    st.subheader("Eintraege")

    if st.button("+ Berufserfahrung hinzufügen"):
        st.session_state.experience_count += 1

    if st.session_state.experience_count > 1 and st.button("- Berufserfahrung entfernen"):
        st.session_state.experience_count -= 1

    if st.button("+ Ausbildung hinzufügen"):
        st.session_state.education_count += 1

    if st.session_state.education_count > 1 and st.button("- Ausbildung entfernen"):
        st.session_state.education_count -= 1

    if st.button("+ Kenntnis hinzufügen"):
        st.session_state.skill_count += 1

    if st.session_state.skill_count > 1 and st.button("- Kenntnis entfernen"):
        st.session_state.skill_count -= 1

    if st.button("+ Sprache hinzufügen"):
        st.session_state.language_count += 1

    if st.session_state.language_count > 1 and st.button("- Sprache entfernen"):
        st.session_state.language_count -= 1

    if st.button("+ Projekt hinzufügen"):
        st.session_state.project_count += 1

    if st.session_state.project_count > 1 and st.button("- Projekt entfernen"):
        st.session_state.project_count -= 1

    if st.button("+ Zertifikat hinzufügen"):
        st.session_state.certificate_count += 1

    if st.session_state.certificate_count > 1 and st.button("- Zertifikat entfernen"):
        st.session_state.certificate_count -= 1

    st.markdown("---")

    st.subheader("Profil weiterbearbeiten")
    st.caption("Lade hier eine zuvor in dieser App gespeicherte JSON-, DOCX- oder PDF-Datei hoch.")
    json_upload = st.file_uploader(tr(ui_language, "load_json"), type=["json", "docx", "pdf"])

    if json_upload is not None:
        os.makedirs("temp", exist_ok=True)
        upload_name = json_upload.name or "import_profile"
        upload_ext = os.path.splitext(upload_name)[1].lower() or ".json"
        temp_upload_path = os.path.join("temp", f"import_profile{upload_ext}")
        with open(temp_upload_path, "wb") as file_obj:
            file_obj.write(json_upload.read())

        try:
            if upload_ext == ".json":
                loaded_profile = load_profile_from_json(temp_upload_path)
            elif upload_ext == ".docx":
                loaded_profile = extract_profile_from_docx(temp_upload_path)
            elif upload_ext == ".pdf":
                loaded_profile = extract_profile_from_pdf(temp_upload_path)
            else:
                raise ValueError("Nicht unterstuetztes Dateiformat.")
        except Exception as exc:
            st.error(f"Profil konnte nicht geladen werden: {exc}")
        else:
            st.session_state["pending_loaded_profile"] = profile_to_dict(loaded_profile)
            st.rerun()

    st.info("Tipp: Fuer ATS moeglichst kein Foto verwenden.")

left_col, middle_col, right_col = st.columns([0.95, 1.05, 1.25])

with left_col:
    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    st.subheader(tr(ui_language, "personal_data"))

    name = st.text_input("Name", key="name")
    job_title = st.text_input(tr(ui_language, "job_title"), key="job_title")
    email = st.text_input(tr(ui_language, "email"), key="email")
    phone = st.text_input(tr(ui_language, "phone"), key="phone")
    city = st.text_input(tr(ui_language, "city"), key="city")
    linkedin = st.text_input(tr(ui_language, "linkedin"), key="linkedin")
    summary = st.text_area(tr(ui_language, "summary"), height=130, key="summary")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    st.subheader(tr(ui_language, "photo"))

    uploaded_photo = st.file_uploader("PNG / JPG / JPEG", type=["png", "jpg", "jpeg"])
    st.session_state["photo_path"] = save_uploaded_photo(uploaded_photo)

    if st.session_state["photo_path"]:
        try:
            st.image(st.session_state["photo_path"], width=180)
        except Exception:
            pass

    st.subheader(tr(ui_language, "other_sections"))
    st.markdown("**Kenntnisse mit Niveau**")

    skill_entries = []
    for i in range(st.session_state.skill_count):
        col_1, col_2 = st.columns([1.4, 1])

        with col_1:
            skill_name = st.text_input("Kenntnis", key=f"skill_name_{i}")

        with col_2:
            current_level = st.session_state.get(f"skill_level_{i}", "Gute Kenntnisse")
            level_index = SKILL_LEVELS.index(current_level) if current_level in SKILL_LEVELS else 1
            skill_level = st.selectbox("Niveau", SKILL_LEVELS, index=level_index, key=f"skill_level_{i}")

        if skill_name.strip():
            skill_entries.append(SkillEntry(name=skill_name, level=skill_level))

    st.markdown(f"**{tr(ui_language, 'languages')}**")
    language_entries = []
    for i in range(st.session_state.language_count):
        col_1, col_2 = st.columns([1.4, 1])

        with col_1:
            language_name = st.text_input("Sprache", key=f"language_name_{i}")

        with col_2:
            current_level = st.session_state.get(f"language_level_{i}", "Gute Kenntnisse")
            level_index = LANGUAGE_LEVELS.index(current_level) if current_level in LANGUAGE_LEVELS else 1
            language_level = st.selectbox(
                "Niveau",
                LANGUAGE_LEVELS,
                index=level_index,
                key=f"language_level_{i}",
            )

        if language_name.strip():
            language_entries.append(LanguageEntry(name=language_name, level=language_level))

    st.markdown(f"**{tr(ui_language, 'projects')}**")
    project_entries = []
    for i in range(st.session_state.project_count):
        with st.expander(f"{tr(ui_language, 'projects')} {i + 1}", expanded=(i == 0)):
            project_title = st.text_input("Projektname", key=f"project_title_{i}")
            project_organization = st.text_input("Wo gemacht? / Institution", key=f"project_organization_{i}")
            project_period = render_period_selector("project", i, "Zeitraum / Period")
            project_summary = st.text_area("Kurze Erklaerung", key=f"project_summary_{i}", height=90)

            if (
                project_title.strip()
                or project_organization.strip()
                or project_period.strip()
                or project_summary.strip()
            ):
                project_entries.append(
                    ProjectEntry(
                        title=project_title,
                        organization=project_organization,
                        period=project_period,
                        summary=project_summary.strip(),
                    )
                )

    st.markdown(f"**{tr(ui_language, 'certificates')}**")
    certificate_entries = []
    for i in range(st.session_state.certificate_count):
        with st.expander(f"{tr(ui_language, 'certificates')} {i + 1}", expanded=(i == 0)):
            certificate_title = st.text_input("Zertifikat", key=f"certificate_title_{i}")
            certificate_issuer = st.text_input("Anbieter / Institution", key=f"certificate_issuer_{i}")
            certificate_period = render_period_selector("certificate", i, "Zeitraum / Period")
            certificate_summary = st.text_area("Kurze Erklaerung", key=f"certificate_summary_{i}", height=90)

            if (
                certificate_title.strip()
                or certificate_issuer.strip()
                or certificate_period.strip()
                or certificate_summary.strip()
            ):
                certificate_entries.append(
                    CertificateEntry(
                        title=certificate_title,
                        issuer=certificate_issuer,
                        period=certificate_period,
                        summary=certificate_summary.strip(),
                    )
                )
    st.markdown("</div>", unsafe_allow_html=True)

with middle_col:
    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    st.subheader(tr(ui_language, "experience"))
    experience_entries = []

    for i in range(st.session_state.experience_count):
        with st.expander(f"{tr(ui_language, 'experience')} {i + 1}", expanded=True):
            title = st.text_input("Titel / Title", key=f"exp_title_{i}")
            company = st.text_input("Firma / Company", key=f"exp_company_{i}")
            location = st.text_input("Ort / Location", key=f"exp_location_{i}")
            period = render_period_selector("exp", i, "Zeitraum / Period")
            details_raw = st.text_area("Details", key=f"exp_details_{i}", height=110)

            if title.strip() or company.strip() or period.strip() or details_raw.strip():
                experience_entries.append(
                    ExperienceEntry(
                        title=title,
                        company=company,
                        location=location,
                        period=period,
                        details=split_lines(details_raw),
                    )
                )

    st.subheader(tr(ui_language, "education"))
    education_entries = []

    for i in range(st.session_state.education_count):
        with st.expander(f"{tr(ui_language, 'education')} {i + 1}", expanded=True):
            degree = st.text_input("Abschluss / Degree", key=f"edu_degree_{i}")
            school = st.text_input("Schule / Hochschule", key=f"edu_school_{i}")
            location = st.text_input("Ort / Location", key=f"edu_location_{i}")
            period = render_period_selector("edu", i, "Zeitraum / Period")
            details_raw = st.text_area("Details", key=f"edu_details_{i}", height=95)

            if degree.strip() or school.strip() or period.strip() or details_raw.strip():
                education_entries.append(
                    EducationEntry(
                        degree=degree,
                        school=school,
                        location=location,
                        period=period,
                        details=split_lines(details_raw),
                    )
                )

    st.markdown("</div>", unsafe_allow_html=True)

profile = CVProfile(
    name=name,
    job_title=job_title,
    email=email,
    phone=phone,
    city=city,
    linkedin=linkedin,
    summary=summary,
    photo_path=st.session_state.get("photo_path", ""),
    language=cv_language,
    skills=skill_entries,
    languages=language_entries,
    projects=project_entries,
    certificates=certificate_entries,
    experience=experience_entries,
    education=education_entries,
)
profile_json_data = json.dumps(asdict(profile), ensure_ascii=False, indent=2).encode("utf-8")

with right_col:
    st.markdown('<div class="preview-title">Live-Vorschau</div>', unsafe_allow_html=True)
    preview_html = render_cv_preview(profile, theme, template_name)
    components.html(preview_html, height=1250, scrolling=True)

st.markdown("---")

col_a, col_b, _ = st.columns([1, 1, 2])
generate_clicked = col_a.button(tr(ui_language, "generate"), type="primary")
col_b.download_button(
    tr(ui_language, "save_json"),
    data=profile_json_data,
    file_name="cv_profile.json",
    mime="application/json",
)

if generate_clicked:
    errors = validate_profile(profile)

    if errors:
        for error in errors:
            st.error(error)
    else:
        builder = TEMPLATES[template_name]
        if template_name == "Modern":
            modern_colors = {
                "header_fill": st.session_state.modern_header_fill.lstrip("#").upper(),
                "sidebar_fill": st.session_state.modern_sidebar_fill.lstrip("#").upper(),
                "header_line": st.session_state.modern_border_fill.lstrip("#").upper(),
                "text_dark": DEFAULT_MODERN_COLORS["text_dark"],
                "text_muted": DEFAULT_MODERN_COLORS["text_muted"],
            }
            doc = build_modern_cv(profile, modern_colors)
        else:
            doc = builder(profile)

        os.makedirs("output", exist_ok=True)
        filename = f"lebenslauf_{template_name.lower().replace(' ', '_')}.docx"
        docx_path = os.path.join("output", filename)
        doc.save(docx_path)

        if is_pdf_export_supported():
            try:
                pdf_path = convert_docx_to_pdf(docx_path)
                embed_profile_in_pdf(pdf_path, profile)
                with open(pdf_path, "rb") as file_obj:
                    st.download_button(
                        tr(ui_language, "download_pdf"),
                        data=file_obj,
                        file_name=os.path.basename(pdf_path),
                        mime="application/pdf",
                    )
            except Exception as exc:
                st.warning(f"PDF konnte nicht erstellt werden: {exc}")
        else:
            st.info("PDF-Export ist in der Cloud nicht verfuegbar. Bitte nutze den DOCX-Download.")

        embed_profile_in_docx(docx_path, profile)

        st.success(f"DOCX erstellt: {filename}")

        with open(docx_path, "rb") as file_obj:
            st.download_button(
                tr(ui_language, "download_docx"),
                data=file_obj,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
