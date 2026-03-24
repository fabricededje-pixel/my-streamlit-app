import json
from dataclasses import asdict
from core.profile_model import (
    CVProfile,
    ExperienceEntry,
    EducationEntry,
    SkillEntry,
    LanguageEntry,
)


def save_profile_to_json(profile: CVProfile, path: str):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(asdict(profile), f, ensure_ascii=False, indent=2)


def load_profile_from_json(path: str) -> CVProfile:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    experience = [ExperienceEntry(**item) for item in data.get("experience", []) if isinstance(item, dict)]
    education = [EducationEntry(**item) for item in data.get("education", []) if isinstance(item, dict)]

    skills = []
    for item in data.get("skills", []):
        if isinstance(item, dict):
            skills.append(SkillEntry(**item))
        else:
            skills.append(SkillEntry(name=str(item), level=""))

    languages = []
    for item in data.get("languages", []):
        if isinstance(item, dict):
            languages.append(LanguageEntry(**item))
        else:
            languages.append(LanguageEntry(name=str(item), level=""))

    data["experience"] = experience
    data["education"] = education
    data["skills"] = skills
    data["languages"] = languages

    return CVProfile(**data)