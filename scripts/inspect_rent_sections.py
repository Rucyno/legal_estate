# -*- coding: utf-8 -*-
import docx, os, sys
sys.stdout.reconfigure(encoding="utf-8")

base = r"D:\Антигравити\legal_estate\deals\10_Аренда_Греченкова_Артамонов\готовые"
f = os.path.join(base, "Договор_аренды_Греченкова_Артамонов_УСИЛЕННЫЙ_ПРОЕКТ.docx")
doc = docx.Document(f)

print("--- Paragraphs 0 to 30 ---")
for i in range(0, 30):
    print(f"P{i}: {doc.paragraphs[i].text}")
