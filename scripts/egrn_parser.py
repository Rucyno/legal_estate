import re
import pdfplumber

def parse_egrn_pdf(pdf_stream_or_path):
    text = ""
    try:
        with pdfplumber.open(pdf_stream_or_path) as pdf:
            for page in pdf.pages:
                text += (page.extract_text() or "") + "\n"
    except Exception as e:
        return {"error": f"Ошибка чтения PDF: {str(e)}"}

    data = {
        "cadastral_num": "",
        "address": "",
        "area": "",
        "floor": "",
        "object_type": "Квартира",
        "owners": [],
        "encumbrances": [],
        "basis_doc": "Договор купли-продажи"
    }

    # 1. Кадастровый номер
    cad_match = re.search(r'Кадастровый номер\s*[:\s]?\s*([0-9]{2}:[0-9]{2}:[0-9]{6,7}:[0-9]+)', text)
    if cad_match:
        data["cadastral_num"] = cad_match.group(1).strip()

    # 2. Площадь
    area_match = re.search(r'Площадь(?:,\s*м2)?\s*[:\s]?\s*([0-9]+(?:[\.,][0-9]+)?)', text, re.IGNORECASE)
    if area_match:
        data["area"] = area_match.group(1).replace(',', '.').strip()

    # 3. Этаж
    floor_match = re.search(r'(?:Номер(?:,\s*тип)?\s*этажа|Этаж)\s*[:\s]?\s*([0-9]+[^\n\r]*)', text, re.IGNORECASE)
    if floor_match:
        data["floor"] = floor_match.group(1).strip()

    # 4. Адрес
    addr_match = re.search(r'(?:Местоположение|Адрес)\s*[:\s]?\s*([^\n\r]+(?:\n[^\n\r]+)?)', text, re.IGNORECASE)
    if addr_match:
        raw_addr = addr_match.group(1).replace('\n', ' ')
        raw_addr = re.sub(r'Российская Федерация,?\s*', '', raw_addr)
        raw_addr = re.sub(r'\s{2,}', ' ', raw_addr)
        data["address"] = raw_addr.split('Назначение')[0].split('Площадь')[0].strip()

    # 5. Правообладатели (ФИО, дата рождения, паспорт)
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
            
            data["owners"].append({
                "fio": fio,
                "birth": birth,
                "passport": passport,
                "share": share
            })

    seen = set()
    unique_owners = []
    for o in data["owners"]:
        if o["fio"] not in seen:
            seen.add(o["fio"])
            unique_owners.append(o)
    data["owners"] = unique_owners

    # 6. Обременения (Ипотека, запреты ФССП, аресты)
    if "запрещение регистрации" in text.lower() or "запрет" in text.lower() or "роспо" in text.lower():
        ip_match = re.search(r'№\s*([0-9]+/[0-9]+/[0-9]+-ИП)', text)
        ip_str = f" (ИП {ip_match.group(1)})" if ip_match else ""
        data["encumbrances"].append(f"Запрет на совершение регистрационных действий ФССП{ip_str}")
    
    if "ипотека" in text.lower() or "залог" in text.lower():
        bank = "Банк (залог)"
        if "сбербанк" in text.lower():
            bank = "ПАО Сбербанк"
        elif "втб" in text.lower():
            bank = "Банк ВТБ"
        data["encumbrances"].append(f"Ипотека в силу закона ({bank})")

    # 7. Основание
    if "приватизац" in text.lower() or ("передач" in text.lower() and "собственность граждан" in text.lower()):
        data["basis_doc"] = "Приватизация"
    elif "участия в долевом" in text.lower() or "дду" in text.lower():
        data["basis_doc"] = "Договор долевого участия (ДДУ)"
    elif "наследств" in text.lower():
        data["basis_doc"] = "Наследство по закону/завещанию"
    elif "дарен" in text.lower():
        data["basis_doc"] = "Договор дарения"
    else:
        data["basis_doc"] = "Договор купли-продажи"

    return data
