Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "d:\Антигравити\legal_estate"
WshShell.Run "cmd /c start_app.bat", 0, False
