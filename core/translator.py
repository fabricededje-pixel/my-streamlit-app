from dataclasses import replace

from deep_translator import GoogleTranslator

from core.profile_model import CVProfile, EducationEntry, ExperienceEntry, LanguageEntry, SkillEntry


TRANSLATION_LANGUAGES = {
    "de": "Deutsch",
    "en": "English",
    "fr": "Francais",
    "es": "Espanol",
    "it": "Italiano",
    "pt": "Portugues",
    "nl": "Nederlands",
    "pl": "Polski",
    "tr": "Turkce",
    "ro": "Romana",
    "cs": "Cestina",
    "sv": "Svenska",
    "da": "Dansk",
    "fi": "Suomi",
    "no": "Norsk",
    "uk": "Ukrainska",
    "ru": "Russkiy",
    "ar": "Arabic",
    "hi": "Hindi",
    "ja": "Japanese",
    "ko": "Korean",
    "zh-CN": "Chinese (Simplified)",
}


def language_label(code: str) -> str:
    return TRANSLATION_LANGUAGES.get(code, code)


def _translate_text(translator: GoogleTranslator, text: str) -> str:
    if not text or not str(text).strip():
        return text
    return translator.translate(str(text))


def _translate_string_list(translator: GoogleTranslator, items):
    return [_translate_text(translator, item) for item in items if str(item).strip()]


def _translate_language_entries(translator: GoogleTranslator, items):
    translated = []
    for item in items:
        if hasattr(item, "name"):
            translated.append(
                LanguageEntry(
                    name=_translate_text(translator, getattr(item, "name", "")),
                    level=_translate_text(translator, getattr(item, "level", "")),
                )
            )
        else:
            translated.append(_translate_text(translator, item))
    return translated


def translate_profile(profile: CVProfile, target_language: str) -> CVProfile:
    translator = GoogleTranslator(source="auto", target=target_language)

    translated_skills = [
        SkillEntry(
            name=_translate_text(translator, skill.name),
            level=_translate_text(translator, skill.level),
        )
        for skill in profile.skills
    ]

    translated_experience = [
        ExperienceEntry(
            title=_translate_text(translator, entry.title),
            company=entry.company,
            location=entry.location,
            period=entry.period,
            details=_translate_string_list(translator, entry.details),
        )
        for entry in profile.experience
    ]

    translated_education = [
        EducationEntry(
            degree=_translate_text(translator, entry.degree),
            school=entry.school,
            location=entry.location,
            period=entry.period,
            details=_translate_string_list(translator, entry.details),
        )
        for entry in profile.education
    ]

    return replace(
        profile,
        job_title=_translate_text(translator, profile.job_title),
        summary=_translate_text(translator, profile.summary),
        language=target_language,
        skills=translated_skills,
        languages=_translate_language_entries(translator, profile.languages),
        projects=_translate_string_list(translator, profile.projects),
        certificates=_translate_string_list(translator, profile.certificates),
        experience=translated_experience,
        education=translated_education,
    )
