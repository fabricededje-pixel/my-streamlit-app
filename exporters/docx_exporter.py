from dataclasses import dataclass, field
from typing import List


@dataclass
class CVProfile:
    name: str = ""
    job_title: str = ""
    email: str = ""
    phone: str = ""
    city: str = ""
    linkedin: str = ""
    summary: str = ""

    skills: List[str] = field(default_factory=list)
    languages: List[str] = field(default_factory=list)
    experience: List[str] = field(default_factory=list)
    education: List[str] = field(default_factory=list)
    projects: List[str] = field(default_factory=list)
    certificates: List[str] = field(default_factory=list)