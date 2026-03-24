# CV Builder

Streamlit-App zum Erstellen von Lebenslaeufen als DOCX und optional als PDF.

## Lokal starten

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## Teilen mit Freunden

Es gibt zwei gute Wege:

### 1. GitHub

Deine Freunde klonen das Projekt und starten es lokal:

```powershell
git clone <DEIN-REPO-URL>
cd cv_builder
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

### 2. Direkt im Browser

Wenn deine Freunde nichts installieren sollen, ist ein Deployment besser. Fuer eine Streamlit-App ist `Streamlit Community Cloud` meist der einfachste Weg. Dann verschickst du nur einen Link.

## Wichtige Hinweise

- DOCX-Export funktioniert lokal direkt.
- Der PDF-Export nutzt `docx2pdf` und funktioniert am zuverlaessigsten auf Windows mit installiertem Microsoft Word.
- `output/`, `temp/`, `.venv/` und `.idea/` werden nicht mit ins Repo aufgenommen.
