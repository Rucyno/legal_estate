import os
import re
import docx
from typing import Dict, Any

def replace_placeholders_in_doc(doc: docx.Document, context: Dict[str, str]) -> docx.Document:
    """
    Надёжная замена плейсхолдеров вида {{KEY}} в python-docx.
    Обрабатывает случай разбиения плейсхолдера на несколько run'ов Word'ом.
    """
    pattern = re.compile(r'\{\{([A-Za-z0-9_]+)\}\}')
    
    for p in doc.paragraphs:
        full_text = "".join(run.text for run in p.runs)
        if "{{" in full_text and "}}" in full_text:
            new_text = full_text
            for match in pattern.finditer(full_text):
                key = match.group(1)
                placeholder = f"{{{{{key}}}}}"
                if placeholder in new_text:
                    val = str(context.get(key, placeholder))
                    new_text = new_text.replace(placeholder, val)
            
            if new_text != full_text and len(p.runs) > 0:
                p.runs[0].text = new_text
                for r in p.runs[1:]:
                    r.text = ""

    # Также обход таблиц (если есть реквизиты в таблицах)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    full_text = "".join(run.text for run in p.runs)
                    if "{{" in full_text and "}}" in full_text:
                        new_text = full_text
                        for match in pattern.finditer(full_text):
                            key = match.group(1)
                            placeholder = f"{{{{{key}}}}}"
                            if placeholder in new_text:
                                val = str(context.get(key, placeholder))
                                new_text = new_text.replace(placeholder, val)
                        if new_text != full_text and len(p.runs) > 0:
                            p.runs[0].text = new_text
                            for r in p.runs[1:]:
                                r.text = ""
                                
    return doc

# Пример тестирования заполнения
if __name__ == "__main__":
    template_path = r"d:\Антигравити\legal_estate\templates\ШАБЛОН_ДКП_Сбербанк_Этажи_с_плейсхолдерами.docx"
    doc = docx.Document(template_path)
    
    context = {
        "GOROD_SDELKI": "Москва",
        "DEN_SDELKI": "14",
        "MESYAC_SDELKI": "ноября",
        "GOD_SDELKI": "2026",
        
        "SELLER_FIO": "Петров Пётр Петрович",
        "SELLER_BIRTH_DATE": "12.03.1982",
        "SELLER_PASSPORT_SERIES": "4510",
        "SELLER_PASSPORT_NUMBER": "123456",
        "SELLER_PASSPORT_ISSUED_BY": "ОВД Пресненского района г. Москвы",
        "SELLER_PASSPORT_DATE": "15.04.2002",
        "SELLER_PASSPORT_CODE": "770-001",
        "SELLER_REG_ADDRESS": "г. Москва, ул. Красная Пресня, д. 10, кв. 25",
        "SELLER_SNILS": "123-456-789 00",
        
        "BUYER_FIO": "Иванов Иван Иванович",
        "BUYER_BIRTH_DATE": "25.07.1990",
        "BUYER_PASSPORT_SERIES": "4518",
        "BUYER_PASSPORT_NUMBER": "654321",
        "BUYER_PASSPORT_ISSUED_BY": "ГУ МВД России по г. Москве",
        "BUYER_PASSPORT_DATE": "10.08.2010",
        "BUYER_PASSPORT_CODE": "770-025",
        "BUYER_REG_ADDRESS": "г. Москва, Ленинский пр-т, д. 45, кв. 112",
        "BUYER_SNILS": "987-654-321 99",
        
        "OBJECT_AREA": "62.4",
        "OBJECT_FLOOR": "7",
        "OBJECT_CADASTRAL_NUMBER": "77:01:0001001:1234",
        "OBJECT_ADDRESS": "г. Москва, ул. Примерная, д. 1, кв. 1",
        "TITLE_DOCUMENTS": "Договора купли-продажи квартиры от 10.02.2018 г.",
        "RECORD_DATE": "20.02.2018",
        "RECORD_NUMBER": "77:01:0001001:1234-77/001/2018-1",
        
        "SUMMA_TOTAL": "7 730 000",
        "SUMMA_TOTAL_PROPIS": "Семь миллионов семьсот тридцать тысяч",
        
        "SUMMA_ZADATKA": "100 000",
        "SUMMA_ZADATKA_PROPIS": "Сто тысяч",
        "DATA_ZADATKA": "01.11.2026",
        
        "SUMMA_NALICHNYE_DO_SDELKI": "500 000",
        "SUMMA_NALICHNYE_DO_SDELKI_PROPIS": "Пятьсот тысяч",
        
        "SUMMA_SBR": "1 500 000",
        "SUMMA_SBR_PROPIS": "Один миллион пятьсот тысяч",
        
        "SUMMA_AKKREDITIV": "0",
        "SUMMA_AKKREDITIV_PROPIS": "Ноль",
        "AKKREDITIV_DAYS": "60",
        
        "SUMMA_MSK": "630 000",
        "SUMMA_MSK_PROPIS": "Шестьсот тридцать тысяч",
        "MSK_CERT_SERIES": "III-АА",
        "MSK_CERT_NUMBER": "123456",
        "MSK_ISSUED_BY": "ГУ - Отделение СФР по г. Москве",
        "MSK_CERT_DATE": "15.05.2020",
        "MSK_DECISION_NUMBER": "123",
        "MSK_HOLDER_NAME": "Иванова Мария Ивановна",
        
        "SUMMA_IPOTEKA": "3 000 000",
        "SUMMA_IPOTEKA_PROPIS": "Три миллиона",
        "MORTGAGE_CONTRACT_NUMBER": "КД-2026-998811",
        "MORTGAGE_CONTRACT_DATE": "10.11.2026",
        "MORTGAGE_CONTRACT_CITY": "Москва",
        
        "SELLER_BANK_RECIPIENT": "Петров Пётр Петрович",
        "SELLER_BANK_ACCOUNT": "40817810000000000000",
        "SELLER_BANK_NAME": "ПАО Сбербанк",
        "SELLER_BANK_BIK": "044525225",
        "SELLER_BANK_CORR_ACCOUNT": "30101810400000000225",
        "SELLER_BANK_INN": "7700000000",
        "SELLER_BANK_KPP": "770001001",
        
        "DATA_PEREDACHI_OBJECTA": "25",
        "MESYAC_PEREDACHI_OBJECTA": "ноября",
        "GOD_PEREDACHI_OBJECTA": "2026",
        
        "SUMMA_RECEIPT": "600 000",
        "SUMMA_RECEIPT_PROPIS": "Шестьсот тысяч"
    }
    
    filled_doc = replace_placeholders_in_doc(doc, context)
    test_out = r"C:\Users\PC\Desktop\ТЕСТ_Заполненный_ДКП_Сбербанк.docx"
    filled_doc.save(test_out)
    print(f"Тестовый документ успешно заполнен и сохранен: {test_out}")
