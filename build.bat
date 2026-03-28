@echo off
echo Building OutlookAutoSender...
call venv\Scripts\activate
pyinstaller --onefile ^
            --windowed ^
            --name "OutlookAutoSender" ^
            --add-data "config.yaml;." ^
            --add-data "templates.yaml;." ^
            --add-data "resources;resources" ^
            main.py
echo Build complete. Check dist\OutlookAutoSender.exe
pause
