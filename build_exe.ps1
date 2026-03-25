$ErrorActionPreference = "Stop"

$python = ".\.venv\Scripts\python.exe"

& $python -m PyInstaller --noconfirm cv_builder.spec
