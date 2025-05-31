@echo off
pyinstaller --onefile --windowed --icon=favicon.ico --add-data "config;config" --add-data "logs;logs" main.py
pause