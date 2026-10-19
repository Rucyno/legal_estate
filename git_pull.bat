@echo off
chcp 65001 >nul
echo Получение свежих обновлений legal_estate из GitHub...
"D:\Git\cmd\git.exe" pull origin main
echo.
echo ========================================================
echo Обновление завершено!
echo ========================================================
pause
