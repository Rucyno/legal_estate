# -*- coding: utf-8 -*-
import docx, os, sys
sys.stdout.reconfigure(encoding="utf-8")

base = r"D:\Антигравити\legal_estate\deals\07_Колесова\готовые"
f_act = os.path.join(base, "2026_10_23_Передаточный_акт_Лазутина_Медведев.docx")
doc = docx.Document(f_act)

for i, p in enumerate(doc.paragraphs):
    print(f"P{i}: {p.text}")
