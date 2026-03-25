# CV Builder

Streamlit-App zum Erstellen von Lebenslaeufen als DOCX und optional als PDF.

## Lokal starten

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## Windows-EXE bauen

```powershell
.\build_exe.ps1
```

Danach liegt die App unter `dist\CVBuilder\CVBuilder.exe`.

## Hinweise

- DOCX-Export funktioniert lokal direkt.
- Der PDF-Export nutzt `docx2pdf` und funktioniert am zuverlaessigsten auf Windows mit installiertem Microsoft Word.
- `output/`, `temp/`, `.venv/` und `.idea/` werden nicht mit ins Repo aufgenommen.
