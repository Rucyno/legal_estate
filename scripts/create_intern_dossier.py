# -*- coding: utf-8 -*-
"""
Модуль создания папки личного дела стажера (HR & СБ) и генерации документов.
Стандарт структуры личного дела (как в 'D:\\Этажи\\ДОУ Акты\\внутренние договоры\\СБ\\{ФИО}'):
  - Договор о сотрудничестве.docx (Соглашение о партнерстве + СОПД на отд. листе)
  - {ИНН}_{ГГГГММДДЧЧММСС}.pdf (Выписка ЕГРН/ЕГРНП из ФНС с ЭЦП)
  - паспорт.png / паспорт.jpg (Скан или фото разворота паспорта)
  - фссп.png (Скриншот проверки на сайте ФССП по ИП/долгам)
  - банкротство.png (Скриншот реестра банкротов Федресурс)
  - арбитраж.png (Скриншот КАД Арбитр)
  - {Суд} гр суд.png (Скриншот районного суда по гражданским делам)
  - {Суд} уг суд.png (Скриншот районного суда по уголовным делам)
  - рсфн.png / росфинмониторинг.png (Скриншот реестра экстремистов/террористов Росфинмониторинга)
"""

import os
import sys
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta
import docx

MONTHS_RU = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля",
    5: "мая", 6: "июня", 7: "июля", 8: "августа",
    9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"
}

def format_date_ru(dt: datetime) -> str:
    return f"«{dt.day:02d}» {MONTHS_RU[dt.month]} {dt.year} г."

def set_para_clean(p, text, align=None, is_bold=False, font_size_pt=10.0):
    p.text = text
    if align is not None:
        p.alignment = align
    for r in p.runs:
        r.font.name = "Times New Roman"
        r.font.size = docx.shared.Pt(font_size_pt)
        r.bold = is_bold

def generate_intern_contract(
    template_path: str,
    output_path: str,
    fio: str,
    birth_date: str,
    passport_series: str,
    passport_number: str,
    passport_issued_date: str,
    passport_issued_by: str,
    passport_department_code: str,
    reg_address: str,
    deal_date: datetime = None,
    email: str = "",
    phone: str = ""
):
    if deal_date is None:
        deal_date = datetime.now()

    end_date = deal_date + relativedelta(months=2)
    start_date_str = format_date_ru(deal_date)
    end_date_str = format_date_ru(end_date)

    doc = docx.Document(template_path)

    # 1. Шапка Таблица 0 (Дата заключения)
    t0 = doc.tables[0]
    cell_date = t0.rows[0].cells[1]
    cell_date.paragraphs[0].text = start_date_str
    if cell_date.paragraphs[0].runs:
        cell_date.paragraphs[0].runs[0].font.name = "Times New Roman"
        cell_date.paragraphs[0].runs[0].font.size = docx.shared.Pt(10)

    # 2. Преамбула Соглашения (Параграф 9)
    # Полный текст преамбулы стажера
    p9 = doc.paragraphs[9]
    p9.text = ""
    r_lead = p9.add_run("Гр. ")
    r_lead.font.name = "Times New Roman"
    r_lead.font.size = docx.shared.Pt(10)
    r_lead.bold = True

    r_fio = p9.add_run(fio)
    r_fio.font.name = "Times New Roman"
    r_fio.font.size = docx.shared.Pt(10)
    r_fio.bold = True

    tail_preamble = (
        f", {birth_date} г.р., паспорт гражданина Российской Федерации {passport_series} {passport_number} "
        f"выдан {passport_issued_date} г. {passport_issued_by}, код подразделения {passport_department_code}, "
        f"адрес регистрации: {reg_address}, именуемый в дальнейшем «Партнер», с другой стороны, "
        f"совместно именуемые «Стороны», заключили настоящее Соглашение о нижеследующем:"
    )
    r_tail = p9.add_run(tail_preamble)
    r_tail.font.name = "Times New Roman"
    r_tail.font.size = docx.shared.Pt(10)
    r_tail.bold = False

    # 3. Пункт 1.5 — Срок стажировки
    p17 = doc.paragraphs[17]
    p17.text = f"1.5. Срок стажировки устанавливается с {start_date_str} по {end_date_str}."
    for r in p17.runs:
        r.font.name = "Times New Roman"
        r.font.size = docx.shared.Pt(10)

    # 4. СОПД Преамбула (Параграф 145)
    p145 = doc.paragraphs[145]
    p145.text = ""
    r_ya = p145.add_run("Я, Гр. ")
    r_ya.font.name = "Times New Roman"
    r_ya.font.size = docx.shared.Pt(9)
    r_ya.bold = True

    r_fio_sopd = p145.add_run(fio)
    r_fio_sopd.font.name = "Times New Roman"
    r_fio_sopd.font.size = docx.shared.Pt(9)
    r_fio_sopd.bold = True

    sopd_tail = (
        f", {birth_date} г.р., паспорт гражданина Российской Федерации {passport_series} {passport_number} "
        f"выдан {passport_issued_date} г. {passport_issued_by}, код подразделения {passport_department_code}, "
        f"адрес регистрации: {reg_address}, именуемый, в соответствии со ст. 9 ФЗ от 27.07.2006 N 152-ФЗ "
        f"\"О персональных данных\",  действуя свободно, своей волей и в своем интересе, а также подтверждая свою "
        f"дееспособность даю настоящее конкретное, предметное, информированное, сознательное и однозначное Согласие "
        f"на обработку персональных данных (далее - Согласие) оператору персональных данных ООО «Межрегиональный Центр Недвижимости» "
        f"(далее - Оператор), местом нахождения которого является: г. Пенза, пр-кт Строителей, д. 152, кв. 19, на следующих условиях:"
    )
    r_sopd_tail = p145.add_run(sopd_tail)
    r_sopd_tail.font.name = "Times New Roman"
    r_sopd_tail.font.size = docx.shared.Pt(9)
    r_sopd_tail.bold = False

    # 5. Таблица 1 — Реквизиты и подписи сторон
    t1 = doc.tables[1]
    cell_p = t1.rows[1].cells[0]
    cell_p.text = (
        f"Гр. {fio}, \n"
        f"{birth_date} г.р., паспорт гражданина Российской Федерации {passport_series} {passport_number} "
        f"выдан {passport_issued_date} г. {passport_issued_by}, код подразделения {passport_department_code}, "
        f"адрес регистрации: {reg_address}, именуемый\n\n"
        f"E-mail: {email if email else '___________________________________'}\n\n"
        f"Телефон: {phone if phone else '_________________________________'}"
    )
    for p in cell_p.paragraphs:
        for r in p.runs:
            r.font.name = "Times New Roman"
            r.font.size = docx.shared.Pt(9)

    # Инициалы для подписи
    parts = fio.strip().split()
    short_name = f"{parts[0]} {parts[1][0]}.{parts[2][0]}." if len(parts) >= 3 else fio
    cell_sign_p = t1.rows[2].cells[0]
    cell_sign_p.text = f"________________/{short_name}"
    for p in cell_sign_p.paragraphs:
        for r in p.runs:
            r.font.name = "Times New Roman"
            r.font.size = docx.shared.Pt(10)

    # 6. Таблица 2 — Подпись под СОПД
    t2 = doc.tables[2]
    # Ячейка даты
    c_date = t2.rows[0].cells[0]
    c_date.paragraphs[0].text = deal_date.strftime("%d.%m.%Y")
    c_date.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    if c_date.paragraphs[0].runs:
        c_date.paragraphs[0].runs[0].font.name = "Times New Roman"
        c_date.paragraphs[0].runs[0].font.size = docx.shared.Pt(9)

    # Ячейка расшифровки
    c_fio = t2.rows[0].cells[4]
    c_fio.paragraphs[0].text = short_name
    c_fio.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    if c_fio.paragraphs[0].runs:
        c_fio.paragraphs[0].runs[0].font.name = "Times New Roman"
        c_fio.paragraphs[0].runs[0].font.size = docx.shared.Pt(9)

    doc.save(output_path)
    return output_path

def create_intern_dossier(
    fio: str,
    birth_date: str,
    passport_series: str,
    passport_number: str,
    passport_issued_date: str,
    passport_issued_by: str,
    passport_department_code: str,
    reg_address: str,
    deal_date: datetime = None,
    email: str = "",
    phone: str = "",
    passport_scan_path: str = None,
    inn: str = None
) -> dict:
    """
    Создает личное дело стажера в 'legal_estate/interns/{ФИО}' по корпоративному стандарту.
    """
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    interns_dir = os.path.join(base_dir, "interns")
    person_dir = os.path.join(interns_dir, fio.strip())
    os.makedirs(person_dir, exist_ok=True)

    template_path = os.path.join(base_dir, "templates", "ШАБЛОН_Соглашение_о_партнерстве_стажер_Этажи_с_СОПД.docx")
    contract_filename = "Договор о сотрудничестве.docx"
    contract_path = os.path.join(person_dir, contract_filename)

    generate_intern_contract(
        template_path=template_path,
        output_path=contract_path,
        fio=fio,
        birth_date=birth_date,
        passport_series=passport_series,
        passport_number=passport_number,
        passport_issued_date=passport_issued_date,
        passport_issued_by=passport_issued_by,
        passport_department_code=passport_department_code,
        reg_address=reg_address,
        deal_date=deal_date,
        email=email,
        phone=phone
    )

    result = {
        "folder": person_dir,
        "contract": contract_path,
        "fio": fio,
        "files_expected": [
            "Договор о сотрудничестве.docx",
            "паспорт.png (или .jpg)",
            "фссп.png",
            "банкротство.png",
            "арбитраж.png",
            "росфинмониторинг.png (рсфн.png)",
            "суд_гр.png",
            "суд_уг.png",
            "{ИНН}_{дата}.pdf (выписка ФНС)"
        ]
    }
    return result

if __name__ == "__main__":
    print("Модуль create_intern_dossier загружен успешно.")
