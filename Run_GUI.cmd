@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "App\MAS_GUI.ps1"
exit
