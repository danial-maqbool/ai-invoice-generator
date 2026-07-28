@echo off
REM One-command setup for the AI Invoice Generator (Windows CMD).
REM   setup.bat

echo ==^> Creating virtual environment (venv\)
python -m venv venv
if errorlevel 1 goto :error

call venv\Scripts\activate.bat

echo ==^> Installing dependencies
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo ==^> Building template.docx
python create_template.py
if errorlevel 1 goto :error

if not exist .env (
    copy .env.example .env >nul
    echo ==^> Created .env - open it and paste in your GITHUB_TOKEN
)

echo.
echo ------------------------------------------------------------
echo Setup complete.
echo.
echo   1. Put your GitHub token in .env
echo   2. venv\Scripts\activate
echo   3. streamlit run app.py
echo ------------------------------------------------------------
goto :eof

:error
echo.
echo Setup failed. Check that Python 3.8+ is installed and on your PATH.
exit /b 1
