# -*- coding: utf-8 -*-
import docx, os, sys

base = r"D:\Антигравити\legal_estate\deals\07_Колесова\готовые"
src_file = os.path.join(base, "Основной_договор_с_распиской_Лазутина_Умеркаева.docx")

# Let us modify build_perfect_dkp.py
with open(r"d:\Антигравити\legal_estate\scripts\build_perfect_dkp.py", "r", encoding="utf-8") as f:
    code = f.read()

# Replace p5.2 in build_perfect_dkp.py
old_clause = """# 3. Insert Clause 3.6 after 3.5 (before P32)"""
print("Found old clause:", old_clause in code)
