import os
import streamlit as st
import streamlit.components.v1 as components

from core.profile_model import CVProfile, ExperienceEntry, EducationEntry, SkillEntry
from core.utils import split_lines
from core.validators import validate_profile
from core.i18n import tr
from core.themes import THEMES
from core.preview import render_cv_preview

from exporters.json_store import save_profile_to_json, load_profile_from_json
from exporters.pdf_exporter import convert_docx_to_pdf

from templates.classic import build_classic_cv
from templates.modern import build_modern_cv
from templates.compact import build_compact_cv
from templates.ats import build_ats_cv
from templates.photo_classic import build_photo_classic_cv
from templates.german_premium import build_german_premium_cv


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
    st.markdown(f"""
    <style>
    .stApp {{
        background:
            radial-gradient(circle at 10% 10%, {theme["accent"]} 0%, transparent 25%),
            radial-gradient(circle at 90% 15%, {theme["primary"]}22 0%, transparent 18%),
            radial-gradient(circle at 80% 85%, {theme["accent"]} 0%, transparent 20%),
            linear-gradient(135deg, {theme["bg"]} 0%, #ffffff 45%, {theme["accent"]}66 100%);
        background-attachment: fixed;
    }}

    .block-container {{
        padding-top: 1.1rem;
        padding-bottom: 2rem;
        max-width: 1600px;
        position: relative;
        z-index: 1;
    }}

    .stApp::before {{
        content: "";
        position: fixed;
        top: -120px;
        left: -120px;
        width: 340px;
        height: 340px;
        background: {theme["primary"]}18;
        filter: blur(70px);
        border-radius: 50%;
        z-index: 0;
        pointer-events: none;
    }}

    .stApp::after {{
        content: "";
        position: fixed;
        bottom: -140px;
        right: -120px;
        width: 380px;
        height: 380px;
        background: {theme["accent"]};
        filter: blur(80px);
        border-radius: 50%;
        opacity: 0.7;
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

    .small-note {{
        color: {theme["muted"]};
        font-size: 0.9rem;
    }}

    hr {{
        border: none;
        border-top: 1px solid {theme["border"]};
        margin-top: 1rem;
        margin-bottom: 1rem;
    }}
    </style>
    """, unsafe_allow_html=True)


def save_uploaded_photo(uploaded_file):
    if uploaded_file is None:
        return st.session_state.get("photo_path", "")

    os.makedirs("temp", exist_ok=True)
    ext = uploaded_file.name.split(".")[-1].lower()
    path = os.path.join("temp", f"uploaded_photo.{ext}")

    with open(path, "wb") as f:
        f.write(uploaded_file.read())

    return path


def fill_session_from_profile(profile: CVProfile):
    st.session_state["name"] = profile.name
    st.session_state["job_title"] = profile.job_title
    st.session_state["email"] = profile.email
    st.session_state["phone"] = profile.phone
    st.session_state["city"] = profile.city
    st.session_state["linkedin"] = profile.linkedin
    st.session_state["summary"] = profile.summary
    st.session_state["languages_raw"] = "\n".join(profile.languages)
    st.session_state["projects_raw"] = "\n".join(profile.projects)
    st.session_state["certificates_raw"] = "\n".join(profile.certificates)
    st.session_state["photo_path"] = profile.photo_path
    st.session_state["language"] = profile.language

    st.session_state["experience_count"] = max(1, len(profile.experience))
    st.session_state["education_count"] = max(1, len(profile.education))
    st.session_state["skill_count"] = max(1, len(profile.skills))

    for i, skill in enumerate(profile.skills):
        if hasattr(skill, "name"):
            st.session_state[f"skill_name_{i}"] = skill.name
            st.session_state[f"skill_level_{i}"] = skill.level if skill.level else "Gute Kenntnisse"
        else:
            st.session_state[f"skill_name_{i}"] = str(skill)
            st.session_state[f"skill_level_{i}"] = "Gute Kenntnisse"

    for i in range(len(profile.skills), max(st.session_state["skill_count"], 1)):
        st.session_state.setdefault(f"skill_name_{i}", "")
        st.session_state.setdefault(f"skill_level_{i}", "Gute Kenntnisse")

    for i, exp in enumerate(profile.experience):
        st.session_state[f"exp_title_{i}"] = exp.title
        st.session_state[f"exp_company_{i}"] = exp.company
        st.session_state[f"exp_location_{i}"] = exp.location
        st.session_state[f"exp_period_{i}"] = exp.period
        st.session_state[f"exp_details_{i}"] = "\n".join(exp.details)

    for i, edu in enumerate(profile.education):
        st.session_state[f"edu_degree_{i}"] = edu.degree
        st.session_state[f"edu_school_{i}"] = edu.school
        st.session_state[f"edu_location_{i}"] = edu.location
        st.session_state[f"edu_period_{i}"] = edu.period
        st.session_state[f"edu_details_{i}"] = "\n".join(edu.details)


st.set_page_config(page_title="CV Builder", layout="wide")

if "experience_count" not in st.session_state:
    st.session_state.experience_count = 1
if "education_count" not in st.session_state:
    st.session_state.education_count = 1
if "skill_count" not in st.session_state:
    st.session_state.skill_count = 1
if "language" not in st.session_state:
    st.session_state.language = "de"
if "photo_path" not in st.session_state:
    st.session_state.photo_path = ""

with st.sidebar:
    st.header("Einstellungen")

    language = st.selectbox(
        "Sprache / Language / Langue",
        ["de", "en", "fr"],
        index=["de", "en", "fr"].index(st.session_state.language),
        key="language"
    )

    template_name = st.selectbox(tr(language, "template"), list(TEMPLATES.keys()))
    theme_name = st.selectbox("Farbschema", list(THEMES.keys()))

theme = THEMES[theme_name]
theme = get_dynamic_theme(theme, template_name)
inject_custom_css(theme)

st.markdown('<div class="page-title">Lebenslauf / CV Builder</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="page-subtitle">professionelle Lebensläufe mit Live-Vorschau.</div>',
    unsafe_allow_html=True
)

with st.sidebar:
    st.markdown("---")
    st.subheader("Einträge")

    if st.button("+ Berufserfahrung hinzufügen"):
        st.session_state.experience_count += 1

    if st.session_state.experience_count > 1:
        if st.button("- Berufserfahrung entfernen"):
            st.session_state.experience_count -= 1

    if st.button("+ Ausbildung hinzufügen"):
        st.session_state.education_count += 1

    if st.session_state.education_count > 1:
        if st.button("- Ausbildung entfernen"):
            st.session_state.education_count -= 1

    if st.button("+ Kenntnis hinzufügen"):
        st.session_state.skill_count += 1

    if st.session_state.skill_count > 1:
        if st.button("- Kenntnis entfernen"):
            st.session_state.skill_count -= 1

    st.markdown("---")

    json_upload = st.file_uploader(
        tr(language, "load_json"),
        type=["json"]
    )

    if json_upload is not None:
        os.makedirs("temp", exist_ok=True)
        temp_json_path = os.path.join("temp", "import_profile.json")
        with open(temp_json_path, "wb") as f:
            f.write(json_upload.read())

        loaded_profile = load_profile_from_json(temp_json_path)
        fill_session_from_profile(loaded_profile)
        st.success("JSON-Profil geladen.")

    st.info("Tipp: Für ATS möglichst kein Foto verwenden.")

left_col, middle_col, right_col = st.columns([0.95, 1.05, 1.25])

with left_col:
    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    st.subheader(tr(language, "personal_data"))

    name = st.text_input("Name", key="name")
    job_title = st.text_input(tr(language, "job_title"), key="job_title")
    email = st.text_input(tr(language, "email"), key="email")
    phone = st.text_input(tr(language, "phone"), key="phone")
    city = st.text_input(tr(language, "city"), key="city")
    linkedin = st.text_input(tr(language, "linkedin"), key="linkedin")
    summary = st.text_area(tr(language, "summary"), height=130, key="summary")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    st.subheader(tr(language, "photo"))

    uploaded_photo = st.file_uploader(
        "PNG / JPG / JPEG",
        type=["png", "jpg", "jpeg"]
    )
    photo_path = save_uploaded_photo(uploaded_photo)
    st.session_state["photo_path"] = photo_path

    if st.session_state["photo_path"]:
        try:
            st.image(st.session_state["photo_path"], width=180)
        except Exception:
            pass

    st.subheader(tr(language, "other_sections"))
    st.markdown("**Kenntnisse mit Niveau**")

    skill_entries = []

    for i in range(st.session_state.skill_count):
        c1, c2 = st.columns([1.4, 1])

        with c1:
            skill_name = st.text_input("Kenntnis", key=f"skill_name_{i}")

        with c2:
            current_level = st.session_state.get(f"skill_level_{i}", "Gute Kenntnisse")
            level_index = SKILL_LEVELS.index(current_level) if current_level in SKILL_LEVELS else 1

            skill_level = st.selectbox(
                "Niveau",
                SKILL_LEVELS,
                index=level_index,
                key=f"skill_level_{i}"
            )

        if skill_name.strip():
            skill_entries.append(SkillEntry(name=skill_name, level=skill_level))

    languages_raw = st.text_area(
        f"{tr(language, 'languages')} (jede Zeile = ein Punkt)",
        height=90,
        key="languages_raw"
    )
    projects_raw = st.text_area(
        f"{tr(language, 'projects')} (jede Zeile = ein Punkt)",
        height=90,
        key="projects_raw"
    )
    certificates_raw = st.text_area(
        f"{tr(language, 'certificates')} (jede Zeile = ein Punkt)",
        height=90,
        key="certificates_raw"
    )
    st.markdown('</div>', unsafe_allow_html=True)

with middle_col:
    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    st.subheader(tr(language, "experience"))
    experience_entries = []

    for i in range(st.session_state.experience_count):
        with st.expander(f"{tr(language, 'experience')} {i + 1}", expanded=True):
            title = st.text_input("Titel / Title", key=f"exp_title_{i}")
            company = st.text_input("Firma / Company", key=f"exp_company_{i}")
            location = st.text_input("Ort / Location", key=f"exp_location_{i}")
            period = st.text_input("Zeitraum / Period", key=f"exp_period_{i}")
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

    st.subheader(tr(language, "education"))
    education_entries = []

    for i in range(st.session_state.education_count):
        with st.expander(f"{tr(language, 'education')} {i + 1}", expanded=True):
            degree = st.text_input("Abschluss / Degree", key=f"edu_degree_{i}")
            school = st.text_input("Schule / Hochschule", key=f"edu_school_{i}")
            location = st.text_input("Ort / Location", key=f"edu_location_{i}")
            period = st.text_input("Zeitraum / Period", key=f"edu_period_{i}")
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

    st.markdown('</div>', unsafe_allow_html=True)

profile = CVProfile(
    name=name,
    job_title=job_title,
    email=email,
    phone=phone,
    city=city,
    linkedin=linkedin,
    summary=summary,
    photo_path=st.session_state.get("photo_path", ""),
    language=language,
    skills=skill_entries,
    languages=split_lines(languages_raw),
    projects=split_lines(projects_raw),
    certificates=split_lines(certificates_raw),
    experience=experience_entries,
    education=education_entries,
)

with right_col:
    st.markdown('<div class="preview-title">Live-Vorschau</div>', unsafe_allow_html=True)
    preview_html = render_cv_preview(profile, theme, template_name)
    components.html(preview_html, height=1250, scrolling=True)

st.markdown("---")

col_a, col_b, col_c = st.columns([1, 1, 2])

generate_clicked = col_a.button(tr(language, "generate"), type="primary")
save_json_clicked = col_b.button(tr(language, "save_json"))

if generate_clicked:
    errors = validate_profile(profile)

    if errors:
        for error in errors:
            st.error(error)
    else:
        builder = TEMPLATES[template_name]
        doc = builder(profile)

        os.makedirs("output", exist_ok=True)
        filename = f"lebenslauf_{template_name.lower().replace(' ', '_')}.docx"
        docx_path = os.path.join("output", filename)
        doc.save(docx_path)

        st.success(f"DOCX erstellt: {filename}")

        with open(docx_path, "rb") as f:
            st.download_button(
                tr(language, "download_docx"),
                data=f,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        try:
            pdf_path = convert_docx_to_pdf(docx_path)
            with open(pdf_path, "rb") as f:
                st.download_button(
                    tr(language, "download_pdf"),
                    data=f,
                    file_name=os.path.basename(pdf_path),
                    mime="application/pdf",
                )
        except Exception as e:
            st.warning(f"PDF konnte nicht erstellt werden: {e}")

if save_json_clicked:
    os.makedirs("output", exist_ok=True)
    json_path = os.path.join("output", "cv_profile.json")
    save_profile_to_json(profile, json_path)

    with open(json_path, "rb") as f:
        st.download_button(
            "JSON herunterladen",
            data=f,
            file_name="cv_profile.json",
            mime="application/json",
        )