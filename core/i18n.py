from functools import lru_cache

from deep_translator import GoogleTranslator


LABELS = {
    "de": {
        "summary": "Kurzprofil",
        "experience": "Berufserfahrung",
        "education": "Ausbildung",
        "skills": "Kenntnisse",
        "languages": "Sprachen",
        "projects": "Projekte",
        "certificates": "Zertifikate",
        "template": "Template waehlen",
        "personal_data": "Persönliche Daten",
        "photo": "Foto",
        "other_sections": "Weitere Bereiche",
        "job_title": "Berufsbezeichnung",
        "email": "E-Mail",
        "phone": "Telefon",
        "city": "Ort",
        "linkedin": "LinkedIn / Website",
        "generate": "Lebenslauf erstellen",
        "download_docx": "DOCX herunterladen",
        "download_pdf": "PDF herunterladen",
        "save_json": "Profil als JSON speichern",
        "load_json": "Profil aus JSON laden",
        "ui_language": "App-Sprache",
        "cv_language": "Lebenslauf-Sprache",
        "translate_to": "Uebersetzen nach",
        "translate_content": "In Zielsprache uebersetzen",
        "translation_done": "Lebenslauf-Inhalte wurden uebersetzt.",
    },
    "en": {
        "summary": "Profile",
        "experience": "Work Experience",
        "education": "Education",
        "skills": "Skills",
        "languages": "Languages",
        "projects": "Projects",
        "certificates": "Certificates",
        "template": "Choose template",
        "personal_data": "Personal information",
        "photo": "Photo",
        "other_sections": "Other sections",
        "job_title": "Job title",
        "email": "Email",
        "phone": "Phone",
        "city": "City",
        "linkedin": "LinkedIn / Website",
        "generate": "Create CV",
        "download_docx": "Download DOCX",
        "download_pdf": "Download PDF",
        "save_json": "Save profile as JSON",
        "load_json": "Load profile from JSON",
        "ui_language": "App language",
        "cv_language": "CV language",
        "translate_to": "Translate to",
        "translate_content": "Translate content",
        "translation_done": "CV content has been translated.",
    },
    "fr": {
        "summary": "Profil",
        "experience": "Experience professionnelle",
        "education": "Formation",
        "skills": "Competences",
        "languages": "Langues",
        "projects": "Projets",
        "certificates": "Certificats",
        "template": "Choisir le modele",
        "personal_data": "Informations personnelles",
        "photo": "Photo",
        "other_sections": "Autres sections",
        "job_title": "Intitule du poste",
        "email": "E-mail",
        "phone": "Telephone",
        "city": "Ville",
        "linkedin": "LinkedIn / Site web",
        "generate": "Creer le CV",
        "download_docx": "Telecharger DOCX",
        "download_pdf": "Telecharger PDF",
        "save_json": "Enregistrer le profil en JSON",
        "load_json": "Charger le profil depuis JSON",
        "ui_language": "Langue de l'application",
        "cv_language": "Langue du CV",
        "translate_to": "Traduire vers",
        "translate_content": "Traduire le contenu",
        "translation_done": "Le contenu du CV a ete traduit.",
    },
}


@lru_cache(maxsize=256)
def _translate_label(lang: str, key: str) -> str:
    base_text = LABELS["en"].get(key, key)
    try:
        return GoogleTranslator(source="en", target=lang).translate(base_text)
    except Exception:
        return base_text


def tr(lang: str, key: str) -> str:
    if lang in LABELS:
        return LABELS.get(lang, LABELS["de"]).get(key, key)
    return _translate_label(lang, key)
