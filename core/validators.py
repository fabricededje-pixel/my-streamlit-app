def validate_profile(profile):
    errors = []

    if not profile.name.strip():
        errors.append("Name ist erforderlich.")

    if profile.email and "@" not in profile.email:
        errors.append("Ungültige E-Mail-Adresse.")

    if not profile.summary.strip() and not profile.experience and not profile.education:
        errors.append("Bitte mindestens Kurzprofil, Berufserfahrung oder Ausbildung ausfüllen.")

    return errors