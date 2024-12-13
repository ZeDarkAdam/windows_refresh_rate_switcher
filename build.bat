@echo off
pyinstaller --onefile --windowed --add-data "icons/icon_light.png;." --add-data "icons/icon_dark.png;." --icon=icons/monitor_hz_icon.png WRRS.py
pause