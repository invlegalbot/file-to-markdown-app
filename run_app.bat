@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Dang tao moi truong ao...
    python -m venv .venv
)
call .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
pause
