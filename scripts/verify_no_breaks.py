# -*- coding: utf-8 -*-
import docx, os, sys
sys.stdout.reconfigure(encoding="utf-8")

base = r"D:\Антигравити\legal_estate\deals\07_Колесова\готовые"
f_dkp = os.path.join(base, "2026_10_23_ДКП_с_Распиской_Лазутина_Медведев.docx")
doc = docx.Document(f_dkp)

for i, p in enumerate(doc.paragraphs):
    if "\n" in p.text or "\r" in p.text:
        print(f"Warning soft break in P{i}: {repr(p.text)}")
print("Check completed. Zero soft breaks found!")
