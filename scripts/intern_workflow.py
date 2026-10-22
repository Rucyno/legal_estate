# -*- coding: utf-8 -*-
"""
Автоматизированный комплекс формирования личного дела стажера (HR & СБ «Этажи»):
1. Генерация «Договор о сотрудничестве.docx» (Соглашение + СОПД на отд. листе)
2. Получение ИНН и выписки из ФНС с ЭЦП
3. Автоматическое копирование готовой папки прямо в:
   'D:\\Этажи\\ДОУ Акты\\внутренние договоры\\СБ\\{ФИО}'
   без необходимости ручного вмешательства пользователя!
"""

import os
import sys
import shutil
import urllib.parse
import time
from datetime import datetime
from dateutil.relativedelta import relativedelta
import docx
import requests
from dotenv import load_dotenv

# Подгружаем переменные окружения из корня legal_estate или osint_checker
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))
if not os.getenv("DADATA_API_KEY"):
    load_dotenv(r"d:\Антигравити\osint_checker_v3\.env")

DADATA_API_KEY = os.getenv("DADATA_API_KEY")
DADATA_SECRET_KEY = os.getenv("DADATA_SECRET_KEY")

SB_TARGET_DIR = r"D:\Этажи\ДОУ Акты\внутренние договоры\СБ"

MONTHS_RU = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля",
    5: "мая", 6: "июня", 7: "июля", 8: "августа",
    9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"
}

def format_date_ru(dt: datetime) -> str:
    return f"«{dt.day:02d}» {MONTHS_RU[dt.month]} {dt.year} г."

def clean_address_dadata(raw_address: str) -> dict:
    """
    Нормализует адрес и возвращает каноническую строку, район города (для подсудности) и ОКТМО.
    """
    if not raw_address or not DADATA_API_KEY or not DADATA_SECRET_KEY:
        return {"result": raw_address, "district": "", "court": ""}
    try:
        url = "https://cleaner.dadata.ru/api/v1/clean/address"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Token {DADATA_API_KEY}",
            "X-Secret": DADATA_SECRET_KEY
        }
        res = requests.post(url, headers=headers, json=[raw_address], timeout=10)
        if res.status_code == 200:
            d = res.json()[0]
            district = d.get("city_district") or d.get("area") or ""
            court_name = ""
            if "железнодорожн" in district.lower():
                court_name = "Железнодорожный"
            elif "ленинск" in district.lower():
                court_name = "Ленинский"
            elif "октябрьск" in district.lower():
                court_name = "Октябрьский"
            elif "первомайск" in district.lower():
                court_name = "Первомайский"
            elif district:
                court_name = district.replace("р-н", "").strip().capitalize()

            return {
                "result": d.get("result", raw_address),
                "district": district,
                "court": court_name,
                "fias_id": d.get("fias_id")
            }
    except Exception as e:
        print(f"Ошибка DaData: {e}")
    return {"result": raw_address, "district": "", "court": ""}

def fetch_inn_and_cert(
    fio: str,
    birth_date: str,
    passport_series: str,
    passport_number: str,
    passport_issued_date: str,
    output_dir: str
) -> dict:
    """
    Получает ИНН из ФНС (service.nalog.ru) и скачивает официальную выписку с квалифицированной ЭЦП.
    """
    parts = fio.strip().split()
    last = parts[0]
    first = parts[1] if len(parts) > 1 else ""
    mid = parts[2] if len(parts) > 2 else ""

    series_clean = passport_series.replace(" ", "")
    number_clean = passport_number.replace(" ", "")
    doc_no = f"{series_clean[:2]} {series_clean[2:]} {number_clean}" if len(series_clean) == 4 else f"{series_clean} {number_clean}"

    INN_BASE = "https://service.nalog.ru"
    INN_PROC = INN_BASE + "/inn-new-proc.json"
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0",
        "Accept-Language": "ru-RU,ru;q=0.9",
        "Origin": INN_BASE
    })

    try:
        # Шаг 1: инициализация сессии
        params = urllib.parse.urlencode({
            "fam": last, "nam": first, "otch": mid, "bdate": birth_date,
            "bplace": "", "doctype": "21", "docno": doc_no
        })
        from_path = f"/inn.do?{params}"
        s.get(f"{INN_BASE}/static/personal-data.html?svc=inn&from={urllib.parse.quote(from_path)}", timeout=12)
        s.post(f"{INN_BASE}/static/personal-data-proc.json", data={"from": from_path, "svc": "inn", "personalData": "1"}, timeout=12)
        s.get(f"{INN_BASE}{from_path}", timeout=12)

        # Шаг 2: find
        r_find = s.post(INN_PROC, data={
            "c": "find", "captcha": "", "captchaToken": "",
            "fam": last, "nam": first, "otch": mid, "bdate": birth_date,
            "doctype": "21", "docno": doc_no, "docdt": passport_issued_date
        }, timeout=15)
        d_find = r_find.json()
        req_id = d_find.get("requestId")
        if not req_id:
            return {"inn": None, "cert_path": None}

        # Шаг 3: опрос get
        inn_val = None
        for _ in range(8):
            time.sleep(2)
            r_get = s.post(INN_PROC, data={"c": "get", "requestId": req_id}, timeout=15)
            d_get = r_get.json()
            inn_val = d_get.get("inn")
            if inn_val:
                break
            if d_get.get("state") == 0:
                break

        if not inn_val:
            return {"inn": None, "cert_path": None}

        # Шаг 4: скачивание выписки с ЭЦП
        now_str = datetime.now().strftime("%Y%m%d%H%M%S")
        cert_filename = f"{inn_val}_{now_str}.pdf"
        cert_path = os.path.join(output_dir, cert_filename)

        r_cert = s.get(f"{INN_BASE}/inn-my-cert.do?inn={inn_val}", timeout=20)
        if r_cert.content[:4] == b"%PDF":
            with open(cert_path, "wb") as f:
                f.write(r_cert.content)
            return {"inn": inn_val, "cert_path": cert_path}

        return {"inn": inn_val, "cert_path": None}
    except Exception as e:
        print(f"Ошибка получения ИНН ФНС: {e}")
        return {"inn": None, "cert_path": None}

def build_full_intern_dossier(
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
    passport_image_source: str = None,
    auto_transfer_to_sb: bool = True
) -> dict:
    """
    Полный автономный цикл:
    - Создает папку 'legal_estate/interns/{ФИО}'
    - Нормализует адрес
    - Генерирует 'Договор о сотрудничестве.docx' со строгим корпоративным стилем
    - Подтягивает ИНН и PDF выписку с ЭЦП из ФНС
    - Копирует фото паспорта в 'паспорт.png' (если передано)
    - Автоматически синхронизирует и перекладывает всё в 'D:\\Этажи\\ДОУ Акты\\внутренние договоры\\СБ\\{ФИО}'
    """
    if deal_date is None:
        deal_date = datetime.now()

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    local_person_dir = os.path.join(base_dir, "interns", fio.strip())
    os.makedirs(local_person_dir, exist_ok=True)

    # Нормализация адреса через DaData (чистая строка без мусора)
    clean_addr_info = clean_address_dadata(reg_address)
    final_address = clean_addr_info["result"] if clean_addr_info["result"] else reg_address

    # 1. Генерация договора
    from scripts.create_intern_dossier import generate_intern_contract
    template_path = os.path.join(base_dir, "templates", "ШАБЛОН_Соглашение_о_партнерстве_стажер_Этажи_с_СОПД.docx")
    contract_path = os.path.join(local_person_dir, "Договор о сотрудничестве.docx")

    generate_intern_contract(
        template_path=template_path,
        output_path=contract_path,
        fio=fio.strip(),
        birth_date=birth_date.strip(),
        passport_series=passport_series.strip(),
        passport_number=passport_number.strip(),
        passport_issued_date=passport_issued_date.strip(),
        passport_issued_by=passport_issued_by.strip(),
        passport_department_code=passport_department_code.strip(),
        reg_address=final_address,
        deal_date=deal_date,
        email=email,
        phone=phone
    )

    # 2. Копирование паспорта
    if passport_image_source and os.path.exists(passport_image_source):
        ext = os.path.splitext(passport_image_source)[1].lower()
        if not ext: ext = ".png"
        target_passport = os.path.join(local_person_dir, f"паспорт{ext}")
        shutil.copy2(passport_image_source, target_passport)

    # 3. Доставание ИНН и выписки с ЭЦП
    inn_res = fetch_inn_and_cert(
        fio=fio,
        birth_date=birth_date,
        passport_series=passport_series,
        passport_number=passport_number,
        passport_issued_date=passport_issued_date,
        output_dir=local_person_dir
    )
    inn_val = inn_res.get("inn")

    # 4. Автоматическое снятие всех скриншотов СБ (Банкротство, КАД, РСФН, ФССП, Суды)
    try:
        from scripts.take_sb_screenshots import take_sb_screenshots
        print("  -> Запуск автоматических скриншотов проверок СБ...")
        take_sb_screenshots(
            fio=fio,
            birth_date=birth_date,
            output_dir=local_person_dir,
            inn=inn_val,
            court_name=clean_addr_info.get("court")
        )
    except Exception as e:
        print(f"  [-] Ошибка при снятии скриншотов СБ: {e}")

    # 5. Автоматическое перекладывание в СБ
    sb_person_dir = os.path.join(SB_TARGET_DIR, fio.strip())
    transferred = False
    if auto_transfer_to_sb and os.path.exists(SB_TARGET_DIR):
        os.makedirs(sb_person_dir, exist_ok=True)
        # Копируем все созданные файлы в СБ
        for item in os.listdir(local_person_dir):
            s_file = os.path.join(local_person_dir, item)
            d_file = os.path.join(sb_person_dir, item)
            if os.path.isfile(s_file):
                shutil.copy2(s_file, d_file)
        transferred = True

    return {
        "fio": fio,
        "local_folder": local_person_dir,
        "sb_folder": sb_person_dir if transferred else None,
        "contract": contract_path,
        "inn": inn_val,
        "cert_path": inn_res.get("cert_path"),
        "court_jurisdiction": clean_addr_info.get("court"),
        "transferred_to_sb": transferred
    }

if __name__ == "__main__":
    print("Модуль intern_workflow готов к работе.")
