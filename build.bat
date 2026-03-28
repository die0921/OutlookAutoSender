@echo off
echo Building OutlookAutoSender...
call venv\Scripts\activate
pyinstaller OutlookAutoSender.spec --clean
echo Build complete. Check dist\OutlookAutoSender.exe
pause
