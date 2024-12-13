@echo off
pyinstaller --onefile --windowed --add-data "icons/icon_light.png;." --icon=icons/monitor_hz_icon.png WRRS.py
pause
