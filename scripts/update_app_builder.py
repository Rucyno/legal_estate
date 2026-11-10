with open(r'd:\Антигравити\legal_estate\app.py', 'r', encoding='utf-8') as f:
    text = f.read()

target = "generated_files = generate_deal_package(deal_payload, templates_dir, out_dir)"
replacement = """from benchmark_pdkp_builder import build_benchmark_pdkp
        pdkp_master_file = os.path.join(out_dir, "1_ПДКП_с_встроенной_распиской.docx")
        build_benchmark_pdkp(deal_payload, pdkp_master_file)
        generated_files = generate_deal_package(deal_payload, templates_dir, out_dir)
        generated_files["pdkp"] = pdkp_master_file"""

text = text.replace(target, replacement)

with open(r'd:\Антигравити\legal_estate\app.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Updated app.py cleanly!")
