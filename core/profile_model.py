from dataclasses import dataclass, field
from typing import List


@dataclass
class SkillEntry:
    name: str = ""
    level: str = ""


@dataclass
class LanguageEntry:
    name: str = ""
    level: str = ""


@dataclass
class ExperienceEntry:
    title: str = ""
    company: str = ""
    location: str = ""
    period: str = ""
    details: List[str] = field(default_factory=list)


@dataclass
class EducationEntry:
    degree: str = ""
    school: str = ""
    location: str = ""
    period: str = ""
    details: List[str] = field(default_factory=list)


@dataclass
class CVProfile:
    name: str = ""
    job_title: str = ""
    email: str = ""
    phone: str = ""
    city: str = ""
    linkedin: str = ""
    summary: str = ""
    photo_path: str = ""
    language: str = "de"

    skills: List[SkillEntry] = field(default_factory=list)
    languages: List[LanguageEntry] = field(default_factory=list)
    projects: List[str] = field(default_factory=list)
    certificates: List[str] = field(default_factory=list)

    experience: List[ExperienceEntry] = field(default_factory=list)
    education: List[EducationEntry] = field(default_factory=list)