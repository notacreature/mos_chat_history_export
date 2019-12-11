@echo off
set /p url=Enter exported chat URL: 

"Scripts\python.exe" "mos_chat_history_export.py" "%url%"

start excel.exe /r "report.xlsx"

pause