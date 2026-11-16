Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "d:\Антигравити\legal_estate"
WshShell.Run """C:\Users\Asus\AppData\Local\Programs\Python\Python312\python.exe"" -m streamlit run app.py --server.port 8501 --server.headless true", 0, False
