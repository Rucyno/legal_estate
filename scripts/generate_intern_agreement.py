# -*- coding: utf-8 -*-
"""
Генератор Соглашения о партнерстве со стажером и согласия на обработку персональных данных (СОПД).
Соблюдает строгие правила:
- Точное сохранение оригинального форматирования, шрифтов и стилей исходного шаблона.
- Дата заключения = дата диалога (или передаваемая дата, по умолчанию сегодня).
- Срок действия (п. 1.5) = дата заключения + 2 календарных месяца.
- Согласие СОПД оформляется на отдельном листе.
- Замена данных стажера в преамбуле, п. 10 (реквизиты) и в тексте СОПД.
"""

import os
import sys
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

def replace_text_in_paragraph(paragraph, old_text, new_text):
    if old_text in paragraph.text:
        # Пытаемся заменить внутри одного run, чтобы не ломать форматирование
        replaced = False
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                replaced = True
                break
        if not replaced:
            # Если текст разбит между run'ами:
            full_text = paragraph.text.replace(old_text, new_text)
            if paragraph.runs:
                first_run = paragraph.runs[0]
                first_run.text = full_text
                for r in paragraph.runs[1:]:
                    r.text = ""
            else:
                paragraph.text = full_text

def replace_text_in_table(table, old_text, new_text):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                replace_text_in_paragraph(p, old_text, new_text)

def generate_intern_agreement(
    fio: str,
    birth_date: str,
    passport_series_num: str,
    passport_issued_date: str,
    passport_issued_by: str,
    passport_subdivision_code: str,
    reg_address: str,
    email: str = "",
    phone: str = "",
    deal_date: datetime = None,
    output_path: str = None,
    template_path: str = None
) -> str:
    """
    Заполняет соглашение о партнерстве со стажером.
    """
    if deal_date is None:
        deal_date = datetime.now()

    end_date = deal_date + relativedelta(months=2)

    date_start_str = format_date_ru(deal_date)
    date_end_str = format_date_ru(end_date)

    if template_path is None:
        template_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "templates",
            "ШАБЛОН_Соглашение_о_партнерстве_стажер_Этажи_с_СОПД.docx"
        )

    doc = docx.Document(template_path)

    # 1. Замена даты в Таблице 0 (г. Пенза | «21» сентября 2026 г.)
    table0 = doc.tables[0]
    replace_text_in_table(table0, "«21» сентября 2026 г.", date_start_str)

    # 2. Формирование строк стажера
    # Оригинальный шаблон:
    # "Гр. Звыков Даниил Дмитриевич, 25.10.1999 г.р., паспорт гражданина Российской Федерации 5619 542212 выдан 06.11.2019 г. УМВД России по Пензенской области, код подразделения 580-001, адрес регистрации: Пензенская обл., г. Пенза, ул. Толстого, д. 5, кв. 60, именуемый в дальнейшем «Партнер», с другой стороны,"
    
    partner_preamble_old = "Гр. Звыков Даниил Дмитриевич, 25.10.1999 г.р., паспорт гражданина Российской Федерации 5619 542212 выдан 06.11.2019 г. УМВД России по Пензенской области, код подразделения 580-001, адрес регистрации: Пензенская обл., г. Пенза, ул. Толстого, д. 5, кв. 60, именуемый в дальнейшем «Партнер», с другой стороны,"
    partner_preamble_new = f"Гр. {fio}, {birth_date} г.р., паспорт гражданина Российской Федерации {passport_series_num} выдан {passport_issued_date} г. {passport_issued_by}, код подразделения {passport_subdivision_code}, адрес регистрации: {reg_address}, именуемый в дальнейшем «Партнер», с другой стороны,"

    # 3. Срок стажировки в п. 1.5
    # Оригинал: "1.5. Срок стажировки устанавливается с «21» сентября 2026 г. по «21» ноября 2026 г."
    term_old = "1.5. Срок стажировки устанавливается с «21» сентября 2026 г. по «21» ноября 2026 г."
    term_new = f"1.5. Срок стажировки устанавливается с {date_start_str} по {date_end_str}."

    # 4. СОПД преамбула
    # "Я, Гр. Звыков Даниил Дмитриевич, 25.10.1999 г.р., паспорт гражданина Российской Федерации 5619 542212 выдан 06.11.2019 г. УМВД России по Пензенской области, код подразделения 580-001, адрес регистрации: Пензенская обл., г. Пенза, ул. Толстого, д. 5, кв. 60, именуемый, в соответствии со ст. 9 ФЗ от 27.07.2006 N 152-ФЗ \"О персональных данных\","
    sopd_preamble_old = "Я, Гр. Звыков Даниил Дмитриевич, 25.10.1999 г.р., паспорт гражданина Российской Федерации 5619 542212 выдан 06.11.2019 г. УМВД России по Пензенской области, код подразделения 580-001, адрес регистрации: Пензенская обл., г. Пенза, ул. Толстого, д. 5, кв. 60, именуемый,"
    sopd_preamble_new = f"Я, Гр. {fio}, {birth_date} г.р., паспорт гражданина Российской Федерации {passport_series_num} выдан {passport_issued_date} г. {passport_issued_by}, код подразделения {passport_subdivision_code}, адрес регистрации: {reg_address}, именуемый,"

    # Проход по абзацам документа
    for p in doc.paragraphs:
        if partner_preamble_old in p.text:
            replace_text_in_paragraph(p, partner_preamble_old, partner_preamble_new)
        elif term_old in p.text:
            replace_text_in_paragraph(p, term_old, term_new)
        elif sopd_preamble_old in p.text:
            replace_text_in_paragraph(p, sopd_preamble_old, sopd_preamble_new)

    # 5. Замена в реквизитах сторон (Таблица 1)
    table1 = doc.tables[1]
    # В ячейке стажера заменяем текст
    cell_partner = table1.rows[1].cells[0]
    partner_details_old = "Гр. Звыков Даниил Дмитриевич, \n25.10.1999 г.р., паспорт гражданина Российской Федерации 5619 542212 выдан 06.11.2019 г. УМВД России по Пензенской области, код подразделения 580-001, адрес регистрации: Пензенская обл., г. Пенза, ул. Толстого, д. 5, кв. 60, именуемый"
    partner_details_new = f"Гр. {fio}, \n{birth_date} г.р., паспорт гражданина Российской Федерации {passport_series_num} выдан {passport_issued_date} г. {passport_issued_by}, код подразделения {passport_subdivision_code}, адрес регистрации: {reg_address}, именуемый"
    
    replace_text_in_table(table1, partner_details_old, partner_details_new)
    if email:
        replace_text_in_table(table1, "E-mail: ___________________________________", f"E-mail: {email}")
    if phone:
        replace_text_in_table(table1, "Телефон: _________________________________", f"Телефон: {phone}")

    # Подпись стажера: "________________/_________________________" -> можно подставить фамилию с инициалами
    parts = fio.strip().split()
    if len(parts) >= 3:
        short_name = f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    else:
        short_name = fio
    replace_text_in_table(table1, "________________/_________________________", f"________________/{short_name}")

    # 6. В таблице 2 (подпись под СОПД) можно подставить дату
    table2 = doc.tables[2]
    # Первая строка таблицы 2 содержит пустые ячейки для подписи, вторая — заголовки 'Дата', 'Подпись', 'Расшифровка подписи'
    # В строке 0 ячейка 0 — дата:
    cell_date = table2.rows[0].cells[0]
    if not cell_date.text.strip():
        p_date = cell_date.paragraphs[0]
        p_date.text = deal_date.strftime("%d.%m.%Y")
        if p_date.runs:
            p_date.runs[0].font.name = "Times New Roman"
            p_date.runs[0].font.size = docx.shared.Pt(9)

    # В строке 0 ячейка 4 — расшифровка:
    cell_name = table2.rows[0].cells[4]
    if not cell_name.text.strip():
        p_name = cell_name.paragraphs[0]
        p_name.text = short_name
        if p_name.runs:
            p_name.runs[0].font.name = "Times New Roman"
            p_name.runs[0].font.size = docx.shared.Pt(9)

    if output_path is None:
        clean_fio = "_".join(parts[:2]) if len(parts) >= 2 else "Стажер"
        output_path = f"Соглашение_о_партнерстве_стажер_{clean_fio}_{deal_date.strftime('%Y_%m_%d')}.docx"

    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    print("Модуль intern_agreement готов к работе.")
