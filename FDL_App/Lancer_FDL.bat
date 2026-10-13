@echo off
cd /d "%~dp0"
py fdl_app.py
if errorlevel 1 pause
