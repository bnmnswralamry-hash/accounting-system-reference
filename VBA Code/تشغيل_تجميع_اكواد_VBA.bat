@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0تجميع_اكواد_VBA_الى_TXT.ps1"
