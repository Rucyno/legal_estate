@echo off
chcp 65001 > nul
title Автономный Конструктор Договоров
echo ========================================================
echo Запуск Автономного Экспертного Конструктора Сделок...
echo Адрес в браузере: http://localhost:8501
echo ========================================================
"C:\Users\Asus\AppData\Local\Programs\Python\Python312\python.exe" -m streamlit run "d:\Антигравити\legal_estate\app.py" --server.port 8501 --server.headless false
pause
