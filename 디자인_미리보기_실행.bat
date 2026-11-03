@echo off
set PYTHONIOENCODING=utf-8
chcp 65001 > nul
cd /d "%~dp0"
echo ===================================================================
echo   [디자인 프리뷰] 토스 & 애플 스타일 스트림릿 데모를 시작합니다...
echo   기존 업무용 앱 및 데이터에 영향을 주지 않는 독립 실행창입니다.
echo   브라우저가 자동으로 열립니다 (http://localhost:8502)
echo ===================================================================
python -m streamlit run 디자인_미리보기.py --server.port 8502
pause

