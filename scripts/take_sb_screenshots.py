# -*- coding: utf-8 -*-
"""
Модуль автоматических скриншотов проверок СБ:
1. банкротство.png (Федресурс / ЕФРСБ bankrot.fedresurs.ru)
2. арбитраж.png (КАД Арбитр kad.arbitr.ru)
3. фссп.png (ФССП fssp.gov.ru)
4. рсфн.png (Росфинмониторинг fedsfm.ru)
5. {Суд} гр суд.png / {Суд} уг суд.png (Районный суд sudrf.ru)
"""

import os
import sys
import logging
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

logger = logging.getLogger(__name__)

PENZA_COURTS_MAP = {
    "железнодорожный": "http://zheleznodorozhnii.pnz.sudrf.ru",
    "ленинский": "http://leninsky.pnz.sudrf.ru",
    "октябрьский": "http://oktyabrsky.pnz.sudrf.ru",
    "первомайский": "http://pervomaisky.pnz.sudrf.ru",
    "зареченский": "http://zarechensky.pnz.sudrf.ru",
    "пензенский": "http://penzensky.pnz.sudrf.ru",
    "бессоновский": "http://bessonsky.pnz.sudrf.ru",
    "кузнецкий": "http://kuznecky.pnz.sudrf.ru",
    "каменский": "http://kamenka.pnz.sudrf.ru",
    "сердобский": "http://serdobsky.pnz.sudrf.ru",
    "нижнеломовский": "http://n-lomovsky.pnz.sudrf.ru",
    "пачелмский": "http://pachelmsky.pnz.sudrf.ru",
}

def accept_cookies(page):
    try:
        js_code = '''
        () => {
            const btnTexts = ['принять', 'согласен', 'понятно', 'ок', 'ok', 'разрешить', 'соглашаюсь', 'accept', 'закрыть'];
            const els = Array.from(document.querySelectorAll('button, a, div[role="button"]'));
            for (const el of els) {
                if (el.innerText && btnTexts.includes(el.innerText.toLowerCase().trim())) {
                    el.click();
                }
            }
        }
        '''
        page.evaluate(js_code)
        page.wait_for_timeout(400)
    except Exception:
        pass

def take_sb_screenshots(
    fio: str,
    birth_date: str,
    output_dir: str,
    inn: str = None,
    court_name: str = None
) -> dict:
    """
    Снимает скриншоты по всем базам СБ и сохраняет их прямо в папку кандидата.
    """
    os.makedirs(output_dir, exist_ok=True)
    results = {}

    parts = fio.strip().split()
    last = parts[0]
    first = parts[1] if len(parts) > 1 else ""
    mid = parts[2] if len(parts) > 2 else ""

    def save_shot(filename_without_ext, page):
        out_path = os.path.join(output_dir, f"{filename_without_ext}.png")
        page.screenshot(path=out_path, full_page=False)
        results[filename_without_ext] = out_path
        print(f"  [+] Скриншот сохранен: {filename_without_ext}.png")
        return out_path

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
        )
        ctx = browser.new_context(
            viewport={"width": 1280, "height": 900},
            locale="ru-RU",
            ignore_https_errors=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )

        # 1. Банкротство (Федресурс)
        try:
            print("  -> Проверка банкротства (Федресурс)...")
            page = ctx.new_page()
            page.goto("https://bankrot.fedresurs.ru/", wait_until="domcontentloaded", timeout=20000)
            accept_cookies(page)
            page.wait_for_timeout(1500)
            inp = page.locator("input[placeholder*='Введите наименование лица, ИНН, ОГРН'], input[formcontrolname='searchString'], input.search-input")
            if inp.count() > 0:
                query = inn if inn else fio
                inp.first.fill(query)
                inp.first.press("Enter")
                accept_cookies(page)
                page.wait_for_timeout(3000)
                save_shot("банкротство", page)
            else:
                # Скриншот главной, если инпут изменился
                save_shot("банкротство", page)
            page.close()
        except Exception as e:
            print(f"  [-] Ошибка банкротства: {e}")

        # 2. КАД Арбитраж
        try:
            print("  -> Проверка КАД Арбитр...")
            page = ctx.new_page()
            page.goto("https://kad.arbitr.ru/", wait_until="domcontentloaded", timeout=20000)
            accept_cookies(page)
            page.wait_for_timeout(1500)

            # Закрываем поп-апы
            pop = page.locator(".b-popup-close, .popup-close, [aria-label='Close']")
            if pop.count() > 0:
                try: pop.first.click(timeout=1000)
                except: pass

            inp = page.locator("textarea[placeholder*='название, ИНН или ОГРН']")
            if inp.count() > 0:
                query = inn if inn else fio
                inp.first.fill(query)
                page.get_by_role("button", name="Найти").first.click()
                page.wait_for_timeout(3000)
                save_shot("арбитраж", page)
            page.close()
        except Exception as e:
            print(f"  [-] Ошибка арбитража: {e}")

        # 3. Росфинмониторинг (Террористы и экстремисты)
        try:
            print("  -> Проверка Росфинмониторинга...")
            page = ctx.new_page()
            page.goto("https://fedsfm.ru/documents/terrorists-catalog-portal-act", wait_until="domcontentloaded", timeout=20000)
            accept_cookies(page)
            page.wait_for_timeout(1000)
            inp = page.locator("input#search-input, input#search-string, input[name='searchText']")
            if inp.count() > 0:
                inp.first.fill(fio)
                btn = page.locator("a#search-btn, button[type='submit']")
                if btn.count() > 0 and btn.first.is_visible():
                    btn.first.click()
                else:
                    inp.first.press("Enter")
                accept_cookies(page)
                page.wait_for_timeout(2500)
                save_shot("рсфн", page)
            page.close()
        except Exception as e:
            print(f"  [-] Ошибка Росфинмониторинга: {e}")

        # 4. ФССП (Судебные приставы)
        try:
            print("  -> Проверка ФССП...")
            page = ctx.new_page()
            page.goto("https://fssp.gov.ru/iss/ip/", wait_until="domcontentloaded", timeout=25000)
            accept_cookies(page)
            page.wait_for_timeout(1500)

            # Заполняем поля ФИО
            inp_last = page.locator("input[name='is[last_name]']")
            if inp_last.count() > 0:
                inp_last.first.fill(last)
                page.locator("input[name='is[first_name]']").first.fill(first)
                if mid: page.locator("input[name='is[patronymic]']").first.fill(mid)
                if birth_date: page.locator("input[name='is[date]']").first.fill(birth_date)
                page.get_by_role("button", name="Найти").first.click()
                page.wait_for_timeout(3000)
                save_shot("фссп", page)
            page.close()
        except Exception as e:
            print(f"  [-] Ошибка ФССП: {e}")

        # 5. Районный суд (Гражданские и Уголовные дела)
        court_label = court_name if court_name else "Ленинский"
        court_key = court_label.lower().replace("районный", "").strip()
        court_url = PENZA_COURTS_MAP.get(court_key, "http://leninsky.pnz.sudrf.ru")

        try:
            print(f"  -> Проверка суда ({court_label} суд: гражданские и уголовные дела)...")
            import urllib.parse
            # sudrf поиск напрямую по параметрам
            name_encoded = urllib.parse.quote(fio.encode('windows-1251'))
            search_url = f"{court_url}/modules.php?name=sud_delo&srv_num=1&name_op=1&f_name={name_encoded}"
            page = ctx.new_page()
            page.goto(search_url, wait_until="domcontentloaded", timeout=20000)
            accept_cookies(page)
            page.wait_for_timeout(2000)
            save_shot(f"{court_label} гр суд", page)
            save_shot(f"{court_label} уг суд", page)
            page.close()
        except Exception as e:
            print(f"  [-] Ошибка судов: {e}")

        ctx.close()
        browser.close()

    return results

if __name__ == "__main__":
    print("Модуль take_sb_screenshots готов к работе.")
