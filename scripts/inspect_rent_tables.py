# -*- coding: utf-8 -*-
import docx, os, sys
sys.stdout.reconfigure(encoding="utf-8")

base = r"D:\Антигравити\legal_estate\deals\10_Аренда_Греченкова_Артамонов\готовые"
f = os.path.join(base, "Договор_аренды_Греченкова_Артамонов_УСИЛЕННЫЙ_ПРОЕКТ.docx")
doc = docx.Document(f)

for idx, t in enumerate(doc.tables):
    print(f"Table {idx} (rows={len(t.rows)}, cols={len(t.columns)}):")
    for r_i, r in enumerate(t.rows):
        for c_i, c in enumerate(r.cells):
            t_prev = c.text.strip().replace("\n", " ")[:60]
            print(f"  ({r_i}, {c_i}): {t_prev}")
