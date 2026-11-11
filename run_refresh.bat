@echo off
cd /d "%~dp0"
python refresh.py
timeout /t 5
