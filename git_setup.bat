@echo off
chcp 65001 >nul
echo ========================================================
echo Настройка Git и первая отправка legal_estate на GitHub...
echo ========================================================
echo.

"D:\Git\cmd\git.exe" config --global user.name "Rucyno"
"D:\Git\cmd\git.exe" config --global user.email "isfrolov58@gmail.com"

"D:\Git\cmd\git.exe" init
"D:\Git\cmd\git.exe" branch -M main
"D:\Git\cmd\git.exe" remote add origin https://github.com/Rucyno/legal_estate.git

"D:\Git\cmd\git.exe" add .
"D:\Git\cmd\git.exe" commit -m "Initial commit: LegalEstate deals and contract system"
"D:\Git\cmd\git.exe" push -u origin main

echo.
echo ========================================================
echo ГОТОВО! Репозиторий legal_estate отправлен на GitHub.
echo Если появилось окно авторизации GitHub - подтверди вход.
echo ========================================================
pause
