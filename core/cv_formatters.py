def skill_strings(skills):
    result = []
    if not skills:
        return result

    for s in skills:
        if hasattr(s, "name"):
            name = (getattr(s, "name", "") or "").strip()
            level = (getattr(s, "level", "") or "").strip()
            if name and level:
                result.append(f"{name} ({level})")
            elif name:
                result.append(name)
        else:
            text = str(s).strip()
            if text:
                result.append(text)
    return result


def language_strings(languages):
    result = []
    if not languages:
        return result

    for lang in languages:
        if hasattr(lang, "name"):
            name = (getattr(lang, "name", "") or "").strip()
            level = (getattr(lang, "level", "") or "").strip()
            if name and level:
                result.append(f"{name} ({level})")
            elif name:
                result.append(name)
        else:
            text = str(lang).strip()
            if text:
                result.append(text)
    return result


def contact_parts(profile):
    parts = []

    if profile.email and str(profile.email).strip():
        parts.append(f"✉ {profile.email}")
    if profile.phone and str(profile.phone).strip():
        parts.append(f"☎ {profile.phone}")
    if profile.city and str(profile.city).strip():
        parts.append(profile.city)
    if profile.linkedin and str(profile.linkedin).strip():
        parts.append(f"🔗 {profile.linkedin}")

    return parts


def premium_contact_lines(profile):
    line1_parts = []
    if profile.city and str(profile.city).strip():
        line1_parts.append(profile.city)
    if profile.phone and str(profile.phone).strip():
        line1_parts.append(f"☎ {profile.phone}")

    line2_parts = []
    if profile.email and str(profile.email).strip():
        line2_parts.append(f"✉ {profile.email}")
    if profile.linkedin and str(profile.linkedin).strip():
        line2_parts.append(f"🔗 {profile.linkedin}")

    return "  |  ".join(line1_parts), "  |  ".join(line2_parts)