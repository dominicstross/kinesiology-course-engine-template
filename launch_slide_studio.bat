@echo off
cd /d "%~dp0"
python scripts\slide_studio_gui.py
if %errorlevel% neq 0 pause
