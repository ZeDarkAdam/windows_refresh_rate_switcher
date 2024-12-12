@echo off
pyinstaller --onefile --windowed --add-data "icons/icon_light.png;." --icon=icons/icon_light.png WRRS.py
pause
