import re
import pdfplumber
import os

def parse_package_files(file_paths):
    """
    Парсит весь загруженный пакет документов (PDF выписки, сканы, текстовые файлы)
    и собирает единый профиль сделки.
    """
    deal_info = {
        "cadastral_num": "",
        "address": "",
        "area": "",
        "floor": "",
        "rooms": "1",
        "basis_doc": "Договор купли-продажи",
        "sellers": [],
        "buyers": [],
        "encumbrances": [],
        "bank_encumbrance": "Нет",
        "fssp_debt": 0
    }

    all_texts = []

    for fpath in file_paths:
        ext = os.path.splitext(fpath)[1].lower()
        if ext == ".pdf":
            try:
                with pdfplumber.open(fpath) as pdf:
                    pdf_text = ""
                    for page in pdf.pages:
                        pdf_text += (page.extract_text() or "") + "\n"
                    all_texts.append(("pdf", fpath, pdf_text))
            except Exception as e:
                pass
        elif ext in [".txt", ".doc", ".docx"]:
            if ext == ".docx":
                try:
                    import docx
                    d = docx.Document(fpath)
                    dtext = "\n".join([p.text for p in d.paragraphs])
                    all_texts.append(("docx", fpath, dtext))
                except:
                    pass
        elif ext in [".jpg", ".jpeg", ".png"]:
            # Image OCR placeholder or fast scan
            pass

    for file_type, fpath, text in all_texts:
        # 1. ЕГРН: Кадастровый номер
        if not deal_info["cadastral_num"]:
            cad_match = re.search(r'Кадастровый номер\s*[:\s]?\s*([0-9]{2}:[0-9]{2}:[0-9]{6,7}:[0-9]+)', text)
            if cad_match:
                deal_info["cadastral_num"] = cad_match.group(1).strip()

        # 2. Площадь
        if not deal_info["area"]:
            area_match = re.search(r'Площадь(?:,\s*м2)?\s*[:\s]?\s*([0-9]+(?:[\.,][0-9]+)?)', text, re.IGNORECASE)
            if area_match:
                deal_info["area"] = area_match.group(1).replace(',', '.').strip()

        # 3. Этаж
        if not deal_info["floor"]:
            floor_match = re.search(r'(?:Номер(?:,\s*тип)?\s*этажа|Этаж)\s*[:\s]?\s*([0-9]+[^\n\r]*)', text, re.IGNORECASE)
            if floor_match:
                deal_info["floor"] = floor_match.group(1).strip()

        # 4. Адрес
        if not deal_info["address"]:
            addr_match = re.search(r'(?:Местоположение|Адрес)\s*[:\s]?\s*([^\n\r]+(?:\n[^\n\r]+)?)', text, re.IGNORECASE)
            if addr_match:
                raw_addr = addr_match.group(1).replace('\n', ' ')
                raw_addr = re.sub(r'Российская Федерация,?\s*', '', raw_addr)
                raw_addr = re.sub(r'\s{2,}', ' ', raw_addr)
                deal_info["address"] = raw_addr.split('Назначение')[0].split('Площадь')[0].strip()

        # 5. Собственники (ФИО, паспорт, дата рождения)
        owner_pattern = re.compile(
            r'(?:[0-9]+\.[0-9]+\s+)?([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+),\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4})'
            r'(?:.*?Паспорт гражданина[^:]*:\s*([^,\n\r]+(?:,\s*выдан:[^\n\r]+)?))?',
            re.DOTALL
        )
        for m in owner_pattern.finditer(text):
            fio = m.group(1).strip()
            birth = m.group(2).strip()
            passport = m.group(3).strip() if m.group(3) else ""
            if "Российская" not in fio and "Федерация" not in fio:
                pos = m.end()
                share_chunk = text[pos:pos+300]
                share_match = re.search(r'(?:доля в праве|Собственность|долевая собственность)[^\n\r]*([0-9]/[0-9]|совместная собственность)?', share_chunk, re.IGNORECASE)
                share = share_match.group(1) if (share_match and share_match.group(1)) else "1/1"
                
                # Check if already added
                if not any(s['fio'] == fio for s in deal_info["sellers"]):
                    deal_info["sellers"].append({
                        "fio": fio,
                        "birth": birth,
                        "passport": passport,
                        "share": share
                    })

        # 6. Обременения
        if "запрещение регистрации" in text.lower() or "запрет" in text.lower() or "роспо" in text.lower():
            ip_match = re.search(r'№\s*([0-9]+/[0-9]+/[0-9]+-ИП)', text)
            ip_str = f" (ИП {ip_match.group(1)})" if ip_match else ""
            deal_info["encumbrances"].append(f"Запрет на совершение регистрационных действий ФССП{ip_str}")
            if not deal_info["fssp_debt"]:
                deal_info["fssp_debt"] = 150000 # базовый дефолт для проработки
        
        if "ипотека" in text.lower() or "залог" in text.lower():
            bank = "Банк (залог)"
            if "сбербанк" in text.lower():
                bank = "ПАО Сбербанк"
            elif "втб" in text.lower():
                bank = "Банк ВТБ"
            elif "дом.рф" in text.lower():
                bank = "АО «Банк ДОМ.РФ»"
            deal_info["encumbrances"].append(f"Ипотека в силу закона ({bank})")
            deal_info["bank_encumbrance"] = bank

        # 7. Основание
        if "приватизац" in text.lower() or ("передач" in text.lower() and "собственность граждан" in text.lower()):
            deal_info["basis_doc"] = "Приватизация"
        elif "участия в долевом" in text.lower() or "дду" in text.lower():
            deal_info["basis_doc"] = "Договор долевого участия (ДДУ)"
        elif "наследств" in text.lower():
            deal_info["basis_doc"] = "Наследство по закону/завещанию"
        elif "дарен" in text.lower():
            deal_info["basis_doc"] = "Договор дарения"

    return deal_info
