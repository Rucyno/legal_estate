# -*- coding: utf-8 -*-
"""
Генератор платежных QR-кодов по ГОСТ Р 56042-2014 для оплаты госпошлин Росреестра
и реквизитов сделок агентства недвижимости «Этажи».

Совместим со всеми банковскими мобильными приложениями РФ:
- СберБанк Онлайн
- Т-Банк (Тинькофф)
- ВТБ Онлайн
- Альфа-Банк
- Райффайзен и др.
"""

import os
import sys
import argparse
import qrcode
from PIL import Image, ImageDraw, ImageFont

# Обеспечиваем безопасный вывод в Windows консоль
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Реквизиты Управления Росреестра по Пензенской области (стандарт для сделок в регионе)
ROSREESTR_PENZA_REQUISITES = {
    "Name": "УФК по Пензенской области (Управление Росреестра по Пензенской области)",
    "PersonalAcc": "03100643000000015800",
    "BankName": "ОТДЕЛЕНИЕ ПЕНЗА БАНКА РОССИИ//УФК по Пензенской области г. Пенза",
    "BIC": "015655003",
    "CorrespAcc": "40102810045370000047",
    "PayeeINN": "5835058778",
    "KPP": "583501001",
    "CBC": "32110807020011000110",  # КБК: Госпошлина за госрегистрацию прав на недвижимое имущество
    "OKTMO": "56701000",            # г. Пенза
    "Purpose": "Госпошлина за государственную регистрацию прав на объект недвижимости"
}

def build_gost_payment_string(
    requisites: dict,
    payer_fio: str = "",
    payer_address: str = "",
    payer_inn: str = "",
    amount_rub: float = 2000.0,
    purpose: str = None
) -> str:
    """
    Формирует строку двумерного штрихкода в стандарте ГОСТ Р 56042-2014.
    """
    fields = [
        "ST00012",  # Идентификатор стандарта (UTF-8)
        f"Name={requisites.get('Name', '')}",
        f"PersonalAcc={requisites.get('PersonalAcc', '')}",
        f"BankName={requisites.get('BankName', '')}",
        f"BIC={requisites.get('BIC', '')}",
        f"CorrespAcc={requisites.get('CorrespAcc', '')}",
        f"PayeeINN={requisites.get('PayeeINN', '')}",
        f"KPP={requisites.get('KPP', '')}",
        f"CBC={requisites.get('CBC', '')}",
        f"OKTMO={requisites.get('OKTMO', '')}",
    ]

    # Назначение платежа
    cur_purpose = purpose if purpose else requisites.get("Purpose", "Оплата госпошлины")
    fields.append(f"Purpose={cur_purpose}")

    # Сумма в копейках
    if amount_rub > 0:
        amount_kopecks = int(round(amount_rub * 100))
        fields.append(f"Sum={amount_kopecks}")

    # Данные плательщика (ФИО разбивается на Last, First, Middle)
    if payer_fio:
        parts = payer_fio.strip().split()
        if len(parts) >= 1:
            fields.append(f"LastName={parts[0]}")
        if len(parts) >= 2:
            fields.append(f"FirstName={parts[1]}")
        if len(parts) >= 3:
            fields.append(f"MiddleName={' '.join(parts[2:])}")

    if payer_address:
        fields.append(f"PayerAddress={payer_address.strip()}")

    if payer_inn:
        fields.append(f"PayerINN={payer_inn.strip()}")

    return "|".join(fields)

def generate_qr_image(
    payment_str: str,
    output_path: str,
    title_text: str = "Госпошлина за регистрацию перехода права",
    amount_text: str = "2 000 руб."
) -> str:
    """
    Генерирует аккуратную брендированную карточку с QR-кодом и поясняющим заголовком.
    """
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2
    )
    qr.add_data(payment_str.encode("utf-8"))
    qr.make(fit=True)

    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    qr_w, qr_h = qr_img.size

    # Добавляем верхнюю плашку с названием и суммой
    header_h = 100
    footer_h = 60
    total_w = qr_w + 40
    total_h = qr_h + header_h + footer_h

    card = Image.new("RGB", (total_w, total_h), color="#FFFFFF")
    draw = ImageDraw.Draw(card)

    # Верхняя красная полоска в корпоративном стиле «Этажи» (#C8102E)
    draw.rectangle([0, 0, total_w, 8], fill="#C8102E")

    # Текст заголовка (используем стандартный шрифт или дефолтный)
    try:
        font_title = ImageFont.truetype("arial.ttf", 18)
        font_sub = ImageFont.truetype("arial.ttf", 15)
        font_footer = ImageFont.truetype("arial.ttf", 13)
    except Exception:
        font_title = ImageFont.load_default()
        font_sub = font_title
        font_footer = font_title

    draw.text((20, 20), title_text, fill="#1E293B", font=font_title)
    draw.text((20, 48), f"К оплате: {amount_text} (Сканируйте в приложении любого банка)", fill="#C8102E", font=font_sub)

    # Вставляем сам QR-код
    card.paste(qr_img, (20, header_h))

    # Нижняя подпись
    draw.text((20, total_h - 45), "ГОСТ Р 56042-2014 • Автозаполнение всех реквизитов", fill="#64748B", font=font_footer)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    card.save(output_path, "PNG", quality=95)
    return output_path

def create_duty_qr_for_deal(
    output_dir: str,
    payer_fio: str = "",
    payer_address: str = "",
    payer_inn: str = "",
    amount_rub: float = 2000.0,
    purpose: str = None
) -> dict:
    """
    Создает QR-код для сделки и сохраняет в output_dir.
    """
    pay_str = build_gost_payment_string(
        requisites=ROSREESTR_PENZA_REQUISITES,
        payer_fio=payer_fio,
        payer_address=payer_address,
        payer_inn=payer_inn,
        amount_rub=amount_rub,
        purpose=purpose
    )

    clean_fio = payer_fio.replace(" ", "_") if payer_fio else "Плательщик"
    qr_filename = f"QR_Госпошлина_Росреестр_{clean_fio}.png"
    qr_path = os.path.join(output_dir, qr_filename)

    sum_str = f"{int(amount_rub):,} руб.".replace(",", " ")
    generate_qr_image(
        payment_str=pay_str,
        output_path=qr_path,
        title_text="Госпошлина за регистрацию перехода прав (Росреестр)",
        amount_text=sum_str
    )

    return {
        "qr_path": qr_path,
        "payment_string": pay_str,
        "sum_rub": amount_rub,
        "payer": payer_fio
    }

def main():
    parser = argparse.ArgumentParser(description="Генератор QR-кода оплаты госпошлины Росреестра (ГОСТ Р 56042-2014)")
    parser.add_argument("--fio", "-f", default="", help="ФИО плательщика/покупателя")
    parser.add_argument("--address", "-a", default="", help="Адрес регистрации плательщика")
    parser.add_argument("--inn", default="", help="ИНН плательщика (опционально)")
    parser.add_argument("--sum", "-s", type=float, default=2000.0, help="Сумма госпошлины в рублях (по умолчанию 2000 руб.)")
    parser.add_argument("--out", "-o", default="", help="Путь сохранения PNG-файла QR-кода")

    args = parser.parse_args()

    out_file = args.out
    if not out_file:
        out_dir = r"D:\Антигравити\legal_estate\deals"
        clean_fio = args.fio.replace(" ", "_") if args.fio else "клиент"
        out_file = os.path.join(out_dir, f"QR_Госпошлина_{clean_fio}.png")

    res = create_duty_qr_for_deal(
        output_dir=os.path.dirname(os.path.abspath(out_file)),
        payer_fio=args.fio,
        payer_address=args.address,
        payer_inn=args.inn,
        amount_rub=args.sum
    )

    print("=" * 60)
    print("💳 QR-КОД ОПЛАТЫ ГОСПОШЛИНЫ УСПЕШНО СФОРМИРОВАН")
    print("=" * 60)
    print(f"📄 Плательщик: {args.fio if args.fio else 'Не указан'}")
    print(f"💰 Сумма: {args.sum:.2f} руб.")
    print(f"🏛️ Получатель: {ROSREESTR_PENZA_REQUISITES['Name']}")
    print(f"🎯 Назначение: {ROSREESTR_PENZA_REQUISITES['Purpose']}")
    print(f"🖼️ Файл QR-кода: {res['qr_path']}")
    print("=" * 60)

if __name__ == "__main__":
    main()
