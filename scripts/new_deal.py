# -*- coding: utf-8 -*-
"""
Автоматический генератор сделок LegalEstate (Стандарт «Этажи»).
Возможности:
1. Создание рабочей папки сделки в deals/{YYYY_MM_DD}_{client}.
2. Подбор и копирование эталонного корпоративного шаблона (D:\\ШАБЛОНЫ_ДОГОВОРОВ_ЭТАЖИ).
3. Интеллектуальное AI Vision OCR распознавание фото паспорта клиента (--passport).
4. Автоматическая генерация банковского QR-кода оплаты госпошлины Росреестра по ГОСТ Р 56042-2014 (--duty).
"""

import os
import sys
import argparse
import shutil
import json
from datetime import datetime

# Обеспечиваем безопасный вывод в Windows консоль без падений кодировки
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from scripts.generate_state_duty_qr import create_duty_qr_for_deal

def get_templates():
    tpl_dir = r"D:\ШАБЛОНЫ_ДОГОВОРОВ_ЭТАЖИ"
    if not os.path.exists(tpl_dir):
        return {}
    files = [f for f in os.listdir(tpl_dir) if f.endswith(".docx")]
    return {f: os.path.join(tpl_dir, f) for f in files}

def run_passport_ocr(image_path: str) -> dict:
    """
    Распознает фото паспорта через модуль passport_ocr из estate_checker или estate_agency_bot.
    """
    try:
        checker_scripts = r"D:\Антигравити\estate_checker\scripts"
        if checker_scripts not in sys.path:
            sys.path.insert(0, checker_scripts)
        from passport_ocr import recognize_passport_file
        return recognize_passport_file(image_path)
    except Exception as e:
        return {"error": f"Ошибка запуска OCR: {e}"}

def main():
    parser = argparse.ArgumentParser(description="Автоматический генератор сделок LegalEstate (Стандарт «Этажи»)")
    parser.add_argument("--client", "-c", required=True, help="ФИО или краткое название клиента/сделки (например: 'Колесова_Лазутина')")
    parser.add_argument("--template", "-t", help="Номер или ключевое слово шаблона (например: 'дом', 'наличка', 'опека', 'коммерция')")
    parser.add_argument("--list", "-l", action="store_true", help="Показать доступные корпоративные шаблоны")
    parser.add_argument("--passport", "-p", default="", help="Путь к фото паспорта клиента для авто-OCR распознавания")
    parser.add_argument("--duty", action="store_true", help="Сгенерировать платежный QR-код госпошлины Росреестра (ГОСТ Р 56042-2014)")
    parser.add_argument("--duty-sum", type=float, default=2000.0, help="Сумма госпошлины (по умолчанию 2000 руб.)")

    args = parser.parse_args()
    templates = get_templates()

    if args.list or not templates:
        print("Доступные эталонные шаблоны (D:\\ШАБЛОНЫ_ДОГОВОРОВ_ЭТАЖИ):")
        for idx, name in enumerate(templates.keys(), 1):
            print(f"[{idx}] {name}")
        return

    # Поиск шаблона
    chosen_file = None
    if args.template:
        t_query = args.template.lower()
        for name, path in templates.items():
            if t_query in name.lower():
                chosen_file = path
                break

    if not chosen_file:
        # По умолчанию берем чистую наличку с осмотром и распиской
        for name, path in templates.items():
            if "чистая_наличка" in name.lower():
                chosen_file = path
                break

    if not chosen_file and templates:
        chosen_file = list(templates.values())[0]

    # Создание папки сделки в deals
    date_prefix = datetime.now().strftime("%Y_%m_%d")
    clean_client = args.client.replace(" ", "_").strip()
    deal_folder_name = f"{date_prefix}_{clean_client}"
    target_dir = os.path.join(r"D:\Антигравити\legal_estate\deals", deal_folder_name)
    os.makedirs(target_dir, exist_ok=True)

    # Копирование шаблона
    dest_filename = f"{deal_folder_name}_ДОГОВОР_ПРОЕКТ.docx"
    dest_path = os.path.join(target_dir, dest_filename)
    shutil.copy2(chosen_file, dest_path)

    print("=" * 60)
    print("💼 СДЕЛКА УСПЕШНО СОЗДАНА В ЮРИДИЧЕСКОМ МОДУЛЕ")
    print("=" * 60)
    print(f"📁 Папка сделки: {target_dir}")
    print(f"📄 Базовый шаблон: {os.path.basename(chosen_file)}")
    print(f"📝 Проект договора: {dest_path}")

    payer_fio = clean_client.replace("_", " ")
    payer_address = ""

    # Если передан паспорт — запускаем AI Vision OCR
    if args.passport:
        print(f"\n🔍 [AI OCR] Сканирую паспорт: {args.passport}...")
        ocr_data = run_passport_ocr(args.passport)
        if not ocr_data.get("error"):
            passport_json_path = os.path.join(target_dir, "паспортные_данные_клиента.json")
            with open(passport_json_path, "w", encoding="utf-8") as f:
                json.dump(ocr_data, f, ensure_ascii=False, indent=2)

            fio = ocr_data.get("fio")
            if fio:
                payer_fio = fio
            payer_address = ocr_data.get("address", "")

            print(f"  ✅ Паспорт распознан: {payer_fio}")
            print(f"  📄 Данные сохранены: {passport_json_path}")
        else:
            print(f"  ⚠️ OCR предупреждение: {ocr_data.get('error')}")

    # Генерация QR-кода госпошлины Росреестра
    if args.duty:
        print(f"\n💳 [ГОСТ QR] Формирую QR-код госпошлины Росреестра ({args.duty_sum:.0f} руб.)...")
        qr_info = create_duty_qr_for_deal(
            output_dir=target_dir,
            payer_fio=payer_fio,
            payer_address=payer_address,
            amount_rub=args.duty_sum
        )
        print(f"  ✅ QR-код готов: {qr_info['qr_path']}")
        print(f"  📲 Плательщик может отсканировать в Сбер / Т-Банк для мгновенной оплаты!")

    print("=" * 60)

if __name__ == "__main__":
    main()
