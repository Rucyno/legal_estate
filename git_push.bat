@echo off
chcp 65001 >nul
echo Отправка изменений legal_estate на GitHub...
"D:\Git\cmd\git.exe" add .
"D:\Git\cmd\git.exe" commit -m "Update legal estate documents and context"
"D:\Git\cmd\git.exe" push origin main
echo.
echo ========================================================
echo Изменения успешно отправлены на GitHub!
echo ========================================================
pause
