@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ==========================================
echo   تجميع ملفات VBA
echo ==========================================
echo.
echo المجلد الحالي:
echo %cd%
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0تجميع_اكواد_VBA_الى_TXT.ps1"

echo.
echo ==========================================
echo انتهى التنفيذ.
echo ابحث عن الملف: VBA_Code_All.txt
echo ==========================================
echo.

pause