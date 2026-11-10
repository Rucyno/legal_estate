import docx
import re
import os
import datetime

def num2words_ru(n):
    millions = n // 1000000
    thousands = (n % 1000000) // 1000
    
    digits_m = {1: "Один миллион", 2: "Два миллиона", 3: "Три миллиона", 4: "Четыре миллиона", 5: "Пять миллионов",
                6: "Шесть миллионов", 7: "Семь миллионов", 8: "Восемь миллионов", 9: "Девять миллионов", 10: "Десять миллионов"}
    digits_th = {50: "пятьдесят тысяч", 100: "сто тысяч", 150: "сто пятьдесят тысяч", 200: "двести тысяч",
                 300: "триста тысяч", 500: "пятьсот тысяч", 800: "восемьсот тысяч", 900: "девятьсот тысяч"}
    
    res = []
    if millions in digits_m:
        res.append(digits_m[millions])
    elif millions > 0:
        res.append(f"{millions} миллионов")
        
    if thousands in digits_th:
        res.append(digits_th[thousands])
    elif thousands > 0:
        res.append(f"{thousands} тысяч")
        
    if not res:
        return f"{n} рублей 00 копеек"
    return " ".join(res) + " рублей 00 копеек"

def replace_text_in_doc(doc, replacements):
    """
    Заменяет плейсхолдеры [Ключ] во всех абзацах и таблицах документа,
    сохраняя оригинальный стиль и форматирование run.
    """
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                # Replace in full paragraph text while preserving runs if single run
                # Or replace across runs
                for r in p.runs:
                    if key in r.text:
                        r.text = r.text.replace(key, str(val))
                # Fallback if split across runs
                if key in p.text:
                    p.text = p.text.replace(key, str(val))

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in replacements.items():
                        if key in p.text:
                            for r in p.runs:
                                if key in r.text:
                                    r.text = r.text.replace(key, str(val))
                            if key in p.text:
                                p.text = p.text.replace(key, str(val))

def generate_deal_package(data, templates_dir, output_dir):
    """
    Генерирует полный комплект документов на базе эталонных шаблонов:
    1. ПДКП с встроенной распиской и запретами
    2. ОДКП (Основной договор купли-продажи)
    3. Передаточный акт
    4. СОПД (Согласие на обработку персданных)
    """
    os.makedirs(output_dir, exist_ok=True)
    results = {}

    sellers = data.get("sellers", [])
    buyers = data.get("buyers", [])
    s1 = sellers[0] if sellers else {}
    s2 = sellers[1] if len(sellers) > 1 else {}
    b1 = buyers[0] if buyers else {}

    s1_fio = s1.get("fio", "Иванов Иван Иванович")
    s2_fio = s2.get("fio", "")
    b1_fio = b1.get("fio", "Петров Петр Петрович")
    
    price = int(data.get("price", 3000000))
    deposit = int(data.get("deposit", 100000))
    cadastral_num = data.get("cadastral_num", "58:29:1005001:1425")
    address = data.get("address", "г. Пенза, ул. Тернопольская, д. 4, кв. 28")
    area = data.get("area", "42.5")
    floor = data.get("floor", "3")
    deal_date = data.get("deal_date", "19.10.2026 г.")

    # 1. ПДКП С ВСТРОЕННОЙ РАСПИСКОЙ
    pdkp_tpl = os.path.join(templates_dir, "ШАБЛОН_ПДКП_Этажи_с_запретами_и_распиской.docx")
    if os.path.exists(pdkp_tpl):
        doc = docx.Document(pdkp_tpl)
        repl = {
            "[Кадастровый номер]": cadastral_num,
            "[Адрес объекта]": address,
            "[ФИО Продавца 1]": s1_fio,
            "[ФИО Продавца 2]": s2_fio if s2_fio else s1_fio,
            "[ФИО Покупателя]": b1_fio,
            "[Сумма задатка]": f"{deposit:,}".replace(",", " "),
            "[СНИЛС]": "036-006-943 27"
        }
        replace_text_in_doc(doc, repl)
        pdkp_out = os.path.join(output_dir, "1_ПДКП_с_встроенной_распиской.docx")
        doc.save(pdkp_out)
        results["pdkp"] = pdkp_out

    # 2. ОСНОВНОЙ ДКП (ОДКП)
    odkp_tpl = os.path.join(templates_dir, "ШАБЛОН_ДКП_Сбербанк_Аккредитив_с_трансферацией.docx")
    if os.path.exists(odkp_tpl):
        doc2 = docx.Document(odkp_tpl)
        repl2 = {
            "[ФИО Продавца]": s1_fio,
            "[ФИО Покупателя]": b1_fio,
            "[Кадастровый номер]": cadastral_num,
            "[Субъект РФ, город, улица, дом, квартира]": address,
            "[Сумма цифрами]": f"{price:,}".replace(",", " "),
            "[Сумма прописью]": num2words_ru(price),
            "[Сумма до сделки цифрами]": f"{deposit:,}".replace(",", " "),
            "[Сумма аккредитива цифрами]": f"{price - deposit:,}".replace(",", " "),
            "[ДД.ММ.ГГГГ]": deal_date,
            "[____]": area,
            "[__]": floor
        }
        replace_text_in_doc(doc2, repl2)
        odkp_out = os.path.join(output_dir, "2_Основной_ДКП.docx")
        doc2.save(odkp_out)
        results["odkp"] = odkp_out

    # 3. ПЕРЕДАТОЧНЫЙ АКТ
    akt_tpl = os.path.join(templates_dir, "ШАБЛОН_Передаточный_акт_Этажи_квартира.docx")
    if os.path.exists(akt_tpl):
        doc3 = docx.Document(akt_tpl)
        repl3 = {
            "[ФИО Продавца]": s1_fio,
            "[ФИО Продавца полностью]": s1_fio,
            "[ФИО Покупателя]": b1_fio,
            "[ФИО Покупателя полностью]": b1_fio,
            "[Кадастровый номер]": cadastral_num,
            "[Адрес объекта]": address,
            "[Площадь]": area,
            "[Этаж]": floor,
            "[дата рождения]": s1.get("birth", "10.05.1985"),
            "[серия]": "56 12",
            "[номер]": "123456",
            "[кем выдан]": "УМВД России по Пензенской области",
            "[код]": "580-001",
            "[дата выдачи]": "15.06.2015",
            "[адрес регистрации]": s1.get("address", address)
        }
        replace_text_in_doc(doc3, repl3)
        akt_out = os.path.join(output_dir, "3_Передаточный_акт.docx")
        doc3.save(akt_out)
        results["akt"] = akt_out

    # 4. СОГЛАСИЕ СОПД
    sopd_tpl = os.path.join(templates_dir, "ШАБЛОН_СОПД_Этажи_с_логотипом.docx")
    if os.path.exists(sopd_tpl):
        doc4 = docx.Document(sopd_tpl)
        repl4 = {
            "[ФИО Покупателя]": b1_fio
        }
        replace_text_in_doc(doc4, repl4)
        sopd_out = os.path.join(output_dir, "4_СОПД_клиента.docx")
        doc4.save(sopd_out)
        results["sopd"] = sopd_out

    return results
