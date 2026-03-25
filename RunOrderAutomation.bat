@echo off
REM — Change into this script’s directory —
cd /d "%~dp0"

REM — Run the Python script —
py OrderAutomation.py

REM — Pause so you can read any output or errors —
pause
