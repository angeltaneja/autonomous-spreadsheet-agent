@echo off
echo ===================================================
echo   Autonomous Spreadsheet Agent - Startup Script
echo ===================================================

:: Check for .env file
if not exist .env (
    echo [INFO] .env file not found. Creating a default one from template.
    copy .env.template .env
)

:: Look for Gemini API Key configuration
findstr /C:"GEMINI_API_KEY=your_gemini_api_key_here" .env > nul
if %errorlevel% equ 0 (
    echo [WARNING] Please update GEMINI_API_KEY in the .env file with your actual Gemini API key before running.
)

:: Start browser and uvicorn backend
echo [INFO] Starting FastAPI server on http://127.0.0.1:8000...
start "" http://127.0.0.1:8000
python -m uvicorn backend.main:app --host 127.0.0.1 --port 8000
pause
