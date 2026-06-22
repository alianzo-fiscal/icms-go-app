@echo off
taskkill /f /im streamlit.exe >nul 2>&1
taskkill /f /fi "WINDOWTITLE eq streamlit*" >nul 2>&1
wmic process where "commandline like '%%streamlit%%'" delete >nul 2>&1
echo App ICMS/GO encerrado.
timeout /t 2 /nobreak >nul
