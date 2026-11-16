import os

startup = os.path.expandvars(r'%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup')
target_vbs = r'd:\Антигравити\legal_estate\start_app_hidden.vbs'
lnk_path = os.path.join(startup, 'LegalEstate_Autostart.lnk')

vbs_content = f'''Set oWS = CreateObject("WScript.Shell")
Set oLink = oWS.CreateShortcut("{lnk_path}")
oLink.TargetPath = "wscript.exe"
oLink.Arguments = chr(34) & "{target_vbs}" & chr(34)
oLink.WorkingDirectory = "d:\\Антигравити\\legal_estate"
oLink.WindowStyle = 0
oLink.Description = "LegalEstate Автозапуск"
oLink.Save
'''

temp_vbs = r'd:\Антигравити\legal_estate\make_shortcut.vbs'
with open(temp_vbs, 'w', encoding='cp1251') as f:
    f.write(vbs_content)

res = os.system(f'cscript //nologo "{temp_vbs}"')
if os.path.exists(temp_vbs):
    os.remove(temp_vbs)

print("Result code:", res)
print("Shortcut exists:", os.path.exists(lnk_path))
