@echo off
pyinstaller --onefile --windowed --add-data "icons/icon_light.ico;." --add-data "icons/icon_dark.ico;." --add-data "icons/icon_color.ico;." --icon=icons/icon_color.ico WRRS.py
pause