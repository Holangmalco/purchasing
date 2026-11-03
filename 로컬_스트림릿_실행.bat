@echo off
set PYTHONIOENCODING=utf-8
chcp 65001 > nul
cd /d "%~dp0"
echo ===================================================
echo   구매추출기 로컬 서버를 시작합니다...
echo   브라우저가 자동으로 열립니다 (http://localhost:8501)
echo ===================================================
python -m streamlit run app.py
pause

