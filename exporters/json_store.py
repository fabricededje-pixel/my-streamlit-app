import json
from dataclasses import asdict

from core.profile_model import (
    CertificateEntry,
    CVProfile,
    ExperienceEntry,
    EducationEntry,
    ProjectEntry,
    SkillEntry,
    LanguageEntry,
)


def profile_to_dict(profile: CVProfile) -> dict:
    return asdict(profile)


def profile_from_data(data: dict) -> CVProfile:
    experience = [ExperienceEntry(**item) for item in data.get("experience", []) if isinstance(item, dict)]
    education = [EducationEntry(**item) for item in data.get("education", []) if isinstance(item, dict)]
    projects = []
    for item in data.get("projects", []):
        if isinstance(item, dict):
            projects.append(ProjectEntry(**item))
        else:
            text = str(item).strip()
            if text:
                projects.append(ProjectEntry(title=text))

    certificates = []
    for item in data.get("certificates", []):
        if isinstance(item, dict):
            certificates.append(CertificateEntry(**item))
        else:
            text = str(item).strip()
            if text:
                certificates.append(CertificateEntry(title=text))

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
    data["projects"] = projects
    data["certificates"] = certificates
    data["skills"] = skills
    data["languages"] = languages

    return CVProfile(**data)


def save_profile_to_json(profile: CVProfile, path: str):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(profile_to_dict(profile), f, ensure_ascii=False, indent=2)


def load_profile_from_json(path: str) -> CVProfile:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return profile_from_data(data)
