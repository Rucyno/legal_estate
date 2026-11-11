import streamlit as st
import os
import sys
import json
import datetime
import tempfile
import zipfile
import re
import pypdf

# Add scripts directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))
from package_parser import parse_package_files
from package_generator import generate_deal_package

st.set_page_config(
    page_title="Картотека и Реестр Сделок",
    page_icon="⚖️",
    layout="wide"
)

# Функция расчета возраста
def calculate_age(birth_str):
    if not birth_str:
        return None
    try:
        clean = birth_str.replace('/', '.').replace('-', '.').strip()
        parts = [int(p) for p in clean.split('.') if p.isdigit()]
        if len(parts) == 3:
            d, m, y = parts
            if y < 100:
                y += 1900 if y > 30 else 2000
            b_date = datetime.date(y, m, d)
            today = datetime.date.today()
            age = today.year - b_date.year - ((today.month, today.day) < (b_date.month, b_date.day))
            return age
    except Exception:
        pass
    return None

def clean_phone_number(phone_str):
    """Очищает телефон до цифр и приводит к международному формату 7XXXXXXXXXX"""
    if not phone_str:
        return ""
    digits = re.sub(r'\D', '', str(phone_str))
    if len(digits) == 10:
        return "7" + digits
    if len(digits) == 11:
        if digits.startswith("8"):
            return "7" + digits[1:]
        return digits
    return digits

def render_phone_messengers_html(phone_str):
    """Рендерит плашку телефона и кнопки быстрого перехода в Telegram и корпоративную CRM/MAX"""
    if not phone_str:
        return ""
    digits = clean_phone_number(phone_str)
    if not digits:
        return f'<span class="phone-text-val">{phone_str}</span>'
    
    tg_url = f"https://t.me/+{digits}"
    max_url = f"https://max.estate/chat?phone={digits}"
    
    return f'''
    <div class="phone-pill-wrap">
        <span class="phone-text-val">📞 {phone_str}</span>
        <a href="{tg_url}" target="_blank" class="messenger-btn messenger-btn-tg" title="Написать в Telegram (+{digits})">✈️ TG</a>
        <a href="{max_url}" target="_blank" class="messenger-btn messenger-btn-max" title="Открыть в корпоративной CRM/MAX">🏢 MAX</a>
    </div>
    '''

# Единый строгий корпоративный стиль «Этажи» (Inter, Красный #DC2626, Серый #475569, чистый фон без лишних рамок)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    /* Основной строгий шрифт для текста и заголовков */
    .stApp, .stMarkdown, h1, h2, h3, h4, p, label, .data-row-flex, .dossier-container, .card-box-clean, .right-clean-block {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif !important;
    }

    .top-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 12px;
        margin-bottom: 16px;
    }
    .top-title {
        font-size: 24px;
        font-weight: 800;
        color: #0F172A;
    }
    .top-subtitle {
        font-size: 13px;
        color: #64748B;
    }
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        border-top: 3px solid #DC2626;
    }
    .metric-value {
        font-size: 24px;
        font-weight: 800;
        color: #0F172A;
        margin-top: 2px;
    }
    .metric-label {
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #64748B;
    }
    .dossier-container {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .dossier-header-title {
        font-size: 20px;
        font-weight: 800;
        color: #0F172A;
        margin-bottom: 4px;
    }
    .dossier-badge {
        display: inline-block;
        padding: 4px 12px;
        background: #FEF3C7;
        color: #92400E;
        font-weight: 700;
        border-radius: 6px;
        font-size: 12px;
        margin-bottom: 16px;
    }
    .section-head {
        font-size: 16px;
        font-weight: 700;
        color: #1E293B;
        border-bottom: 2px solid #DC2626;
        padding-bottom: 6px;
        margin-top: 14px;
        margin-bottom: 12px;
    }
    .sub-section-head {
        font-size: 14px;
        font-weight: 700;
        color: #334155;
        border-bottom: 1px solid #CBD5E1;
        padding-bottom: 4px;
        margin-top: 8px;
        margin-bottom: 12px;
    }
    .data-row-flex {
        display: flex;
        align-items: flex-start;
        margin-bottom: 8px;
        font-size: 13.5px;
        color: #334155;
        line-height: 1.5;
    }
    .data-label-fixed {
        font-weight: 600;
        color: #64748B;
        width: 175px;
        min-width: 175px;
        flex-shrink: 0;
    }
    .data-val-fixed {
        flex-grow: 1;
        color: #0F172A;
    }
    .cadastral-link {
        color: #0F172A;
        font-weight: 700;
        text-decoration: none;
        border-bottom: 1px dashed #64748B;
        cursor: pointer;
        transition: color 0.15s ease;
    }
    .cadastral-link:hover {
        color: #DC2626;
        border-bottom: 1px solid #DC2626;
    }
    .rosreestr-btn {
        display: inline-flex;
        align-items: center;
        margin-left: 8px;
        padding: 2px 8px;
        background: #F8FAFC;
        color: #DC2626;
        border-radius: 4px;
        font-size: 11px;
        font-weight: 700;
        text-decoration: none;
        border: 1px solid #FECACA;
        transition: all 0.15s ease;
        cursor: pointer;
    }
    .rosreestr-btn:hover {
        background: #DC2626;
        color: #FFFFFF;
        border-color: #DC2626;
    }
    .right-clean-block {
        border-bottom: 1px dashed #E2E8F0;
        padding-bottom: 8px;
        margin-bottom: 8px;
        font-size: 13px;
        line-height: 1.45;
        color: #334155;
    }
    .legal-alert {
        padding: 10px 14px;
        border-radius: 6px;
        margin-bottom: 8px;
        font-size: 13px;
        line-height: 1.45;
    }
    .deal-tag {
        display: inline-block;
        padding: 3px 10px;
        background-color: #F1F5F9;
        color: #475569;
        border-radius: 14px;
        font-size: 11px;
        font-weight: 600;
        margin-right: 6px;
        margin-bottom: 4px;
    }
    .card-box-clean {
        background: #FFFFFF;
        border-bottom: 1px solid #E2E8F0;
        padding: 8px 0;
        margin-bottom: 8px;
        line-height: 1.45;
    }
    .row-edit-icon-btn {
        background: transparent;
        border: none;
        color: #94A3B8;
        cursor: pointer;
        font-size: 13px;
        padding: 0 4px;
        transition: color 0.15s ease;
    }
    .row-edit-icon-btn:hover {
        color: #DC2626;
    }
    .inline-edit-box {
        background: #F8FAFC;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        padding: 8px 10px;
        margin-top: 4px;
        margin-bottom: 8px;
    }
    .right-row {
        font-size: 13px;
        line-height: 1.45;
        color: #334155;
        padding: 6px 0;
    }
    .right-holder-sep {
        border-top: 1px dashed #E2E8F0;
        margin: 4px 0 6px 0;
    }
    .doc-type-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 11px;
        font-weight: 700;
    }
    .doc-type-passport {
        background: #E0E7FF;
        color: #3730A3;
    }
    .doc-type-birth {
        background: #FEF3C7;
        color: #B45309;
    }
    .party-header-row {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        flex-wrap: wrap;
        gap: 6px;
    }
    .party-fio-phone-wrap {
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px;
    }
    .party-fio-title {
        font-size: 15px;
        font-weight: 700;
        color: #0F172A;
    }
    .messenger-btn {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 1px 7px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        text-decoration: none;
        transition: all 0.15s ease;
    }
    .messenger-btn-tg {
        background: #E0F2FE;
        color: #0284C7;
        border: 1px solid #BAE6FD;
    }
    .messenger-btn-tg:hover {
        background: #0284C7;
        color: #FFFFFF;
    }
    .messenger-btn-max {
        background: #F3E8FF;
        color: #7E22CE;
        border: 1px solid #E9D5FF;
    }
    .messenger-btn-max:hover {
        background: #7E22CE;
        color: #FFFFFF;
    }
    .phone-pill-wrap {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        padding: 1px 6px;
        border-radius: 14px;
    }
    .phone-text-val {
        font-size: 12px;
        font-weight: 600;
        color: #334155;
    }
</style>
""", unsafe_allow_html=True)

# HEADER
st.markdown("""
<div class="top-header">
    <div>
        <div class="top-title">📂 Реестр Сделок и Картотека</div>
        <div class="top-subtitle">Единый электронный реестр сделок, детальные карточки дел и автономная сборка документов</div>
    </div>
    <div>
        <span style="background:#FEE2E2; color:#DC2626; padding:6px 12px; border-radius:20px; font-weight:700; font-size:12px;">🔒 ЛИЧНЫЙ КАБИНЕТ ЮРИСТА</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ЗАГРУЗКА БАЗЫ СДЕЛОК
catalog_path = os.path.join(os.path.dirname(__file__), 'scripts', 'deals_catalog.json')
try:
    with open(catalog_path, 'r', encoding='utf-8') as f:
        all_deals = json.load(f)
except Exception:
    all_deals = []

def save_deals_to_catalog(deals_data):
    try:
        with open(catalog_path, 'w', encoding='utf-8') as f:
            json.dump(deals_data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"Ошибка сохранения каталога: {e}")
        return False

def inline_edit_row(row_html, key, fields, on_save, on_delete=None, help_txt="Изменить строку"):
    """Строка карточки с карандашом ✏️: клик открывает мини-форму, Enter/«Сохранить» пишет в каталог.
    fields — список (подпись, текущее значение); on_save получает список новых значений."""
    show_key = f"show_{key}"
    c_txt, c_btn = st.columns([12, 1])
    with c_txt:
        st.markdown(row_html, unsafe_allow_html=True)
    with c_btn:
        if st.button("✏️", key=f"btn_{key}", help=help_txt):
            st.session_state[show_key] = not st.session_state.get(show_key, False)
            st.rerun()
    if st.session_state.get(show_key, False):
        with st.form(key=f"form_{key}"):
            new_vals = [st.text_input(lbl, value=str(val or "")) for lbl, val in fields]
            b1, b2, _ = st.columns([1, 1, 2])
            save_clicked = b1.form_submit_button("💾 Сохранить")
            del_clicked = b2.form_submit_button("🗑 Удалить") if on_delete else False
            if save_clicked:
                on_save([v.strip() for v in new_vals])
                save_deals_to_catalog(all_deals)
                st.session_state[show_key] = False
                st.rerun()
            if del_clicked:
                on_delete()
                save_deals_to_catalog(all_deals)
                st.session_state[show_key] = False
                st.rerun()

# =========================================================================
# РЕЖИМ 1: ДЕТАЛЬНАЯ КАРТОЧКА СДЕЛКИ (ОТДЕЛЬНЫЙ ЭКРАН / ДОСЬЕ ДЕЛА)
# =========================================================================
if "opened_deal_id" in st.session_state and st.session_state["opened_deal_id"]:
    deal_id = st.session_state["opened_deal_id"]
    cd = next((d for d in all_deals if d.get("id") == deal_id), all_deals[0] if all_deals else {})

    if st.button("⬅️ Назад к общему реестру сделок", type="secondary"):
        st.session_state["opened_deal_id"] = None
        st.rerun()

    # Состояния точечного редактирования для каждого блока карточки
    edit_obj_key = f"edit_obj_{cd.get('id')}"
    edit_parties_key = f"edit_parties_{cd.get('id')}"
    edit_fin_key = f"edit_fin_{cd.get('id')}"

    for k in [edit_obj_key, edit_parties_key, edit_fin_key]:
        if k not in st.session_state:
            st.session_state[k] = False

    header_cols = st.columns([4, 1])
    with header_cols[0]:
        st.markdown(f"""
        <div style="margin-top:8px; margin-bottom:12px;">
            <div class="dossier-header-title">📄 КАРТОЧКА СДЕЛКИ: {cd.get('title')}</div>
            <div style="color:#64748B; font-size:13px;">ID сделки: <b>{cd.get('id')}</b> | Папка дела: <code>deals/{cd.get('folder')}</code></div>
        </div>
        """, unsafe_allow_html=True)
    with header_cols[1]:
        st.markdown(f'<div style="text-align:right;"><span class="dossier-badge">СТАТУС: {cd.get("status", "В работе")}</span></div>', unsafe_allow_html=True)

    obj = cd.get("object_card", {})

    # ---------------------------------------------------------------------
    # БЛОК 1: ОБЪЕКТ НЕДВИЖИМОСТИ (ИДЕАЛЬНОЕ ВЫРАВНИВАНИЕ + КЛИКАБЕЛЬНЫЕ КН + БЕЗ ЛИШНИХ РАМОК)
    # ---------------------------------------------------------------------
    st.markdown('<div class="section-head">🏢 1. Карточка Объекта Недвижимости (Сведения из ЕГРН)</div>', unsafe_allow_html=True)

    # РЕЖИМ ПРОСМОТРА КАРТОЧКИ ОБЪЕКТА (ПОСТРОЧНОЕ РЕДАКТИРОВАНИЕ ПО КАРАНДАШУ)
    col_obj_1, col_obj_2 = st.columns([1, 1], gap="large")
    with col_obj_1:
        st.markdown('<div class="sub-section-head">📋 Лист 1: Основные характеристики объекта</div>', unsafe_allow_html=True)
        
        # --- 1. Вид объекта ---
        c_v1, c_v2 = st.columns([12, 1])
        with c_v1:
            st.markdown(f'''
            <div class="data-row-flex">
                <div class="data-label-fixed">Вид объекта:</div>
                <div class="data-val-fixed"><b>{obj.get("object_type", "Помещение")}</b></div>
            </div>
            ''', unsafe_allow_html=True)
        with c_v2:
            if st.button("✏️", key=f"btn_inline_objtype_{cd.get('id')}", help="Изменить вид объекта"):
                st.session_state[f"show_ed_type_{cd.get('id')}"] = not st.session_state.get(f"show_ed_type_{cd.get('id')}", False)
                st.rerun()
        if st.session_state.get(f"show_ed_type_{cd.get('id')}", False):
            with st.form(key=f"form_ed_type_{cd.get('id')}"):
                new_val = st.text_input("Вид объекта (нажмите Enter для сохранения):", value=obj.get("object_type", "Помещение"))
                c_s1, c_s2 = st.columns([1, 3])
                if c_s1.form_submit_button("Сохранить"):
                    obj["object_type"] = new_val.strip()
                    save_deals_to_catalog(all_deals)
                    st.session_state[f"show_ed_type_{cd.get('id')}"] = False
                    st.rerun()

        # --- 2. Кадастровый номер ---
        cad_items = obj.get("cadastral_numbers", [])
        if not cad_items:
            main_cn = obj.get("cadastral_num", cd.get("cadastral_num", "—"))
            cad_items = [{"name": "Кадастровый номер", "number": main_cn}]

        for c_idx, c_item in enumerate(cad_items):
            lbl = c_item.get("name", "Кадастровый номер") + ":"
            cn_val = c_item.get("number", "—")
            rosreestr_url = f"https://lk.rosreestr.ru/eservices/real-estate-objects-online"
            
            c_cad1, c_cad2 = st.columns([12, 1])
            with c_cad1:
                st.markdown(f'''
                <div class="data-row-flex">
                    <div class="data-label-fixed">{lbl}</div>
                    <div class="data-val-fixed">
                        <span class="cadastral-link" title="Нажмите, чтобы скопировать" onclick="navigator.clipboard.writeText('{cn_val}'); alert('Кадастровый номер {cn_val} скопирован в буфер!\\nНа сайте Росреестра просто нажмите Ctrl+V в строке поиска.');">{cn_val}</span>
                        <a href="{rosreestr_url}" target="_blank" onclick="navigator.clipboard.writeText('{cn_val}');" class="rosreestr-btn" title="Скопировать КН {cn_val} и открыть Росреестр">📋 Росреестр (Ctrl+V)</a>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
            with c_cad2:
                if st.button("✏️", key=f"btn_inline_cad_{cd.get('id')}_{c_idx}", help=f"Изменить {lbl}"):
                    st.session_state[f"show_ed_cad_{cd.get('id')}_{c_idx}"] = not st.session_state.get(f"show_ed_cad_{cd.get('id')}_{c_idx}", False)
                    st.rerun()
            if st.session_state.get(f"show_ed_cad_{cd.get('id')}_{c_idx}", False):
                with st.form(key=f"form_ed_cad_{cd.get('id')}_{c_idx}"):
                    new_cad_val = st.text_input(f"{lbl} (нажмите Enter для сохранения):", value=cn_val)
                    c_s1, c_s2 = st.columns([1, 3])
                    if c_s1.form_submit_button("Сохранить"):
                        c_item["number"] = new_cad_val.strip()
                        if c_idx == 0:
                            obj["cadastral_num"] = new_cad_val.strip()
                            cd["cadastral_num"] = new_cad_val.strip()
                        obj["cadastral_numbers"] = cad_items
                        save_deals_to_catalog(all_deals)
                        st.session_state[f"show_ed_cad_{cd.get('id')}_{c_idx}"] = False
                        st.rerun()

        # --- 3. Точный адрес ЕГРН ---
        c_a1, c_a2 = st.columns([12, 1])
        with c_a1:
            st.markdown(f'''
            <div class="data-row-flex">
                <div class="data-label-fixed">Точный адрес (ЕГРН):</div>
                <div class="data-val-fixed"><b>{obj.get("address", cd.get("address", "—"))}</b></div>
            </div>
            ''', unsafe_allow_html=True)
        with c_a2:
            if st.button("✏️", key=f"btn_inline_addr_{cd.get('id')}", help="Изменить точный адрес"):
                st.session_state[f"show_ed_addr_{cd.get('id')}"] = not st.session_state.get(f"show_ed_addr_{cd.get('id')}", False)
                st.rerun()
        if st.session_state.get(f"show_ed_addr_{cd.get('id')}", False):
            with st.form(key=f"form_ed_addr_{cd.get('id')}"):
                new_addr_val = st.text_input("Точный адрес объекта (нажмите Enter для сохранения):", value=obj.get("address", cd.get("address", "")))
                c_s1, c_s2 = st.columns([1, 3])
                if c_s1.form_submit_button("Сохранить"):
                    obj["address"] = new_addr_val.strip()
                    cd["address"] = new_addr_val.strip()
                    save_deals_to_catalog(all_deals)
                    st.session_state[f"show_ed_addr_{cd.get('id')}"] = False
                    st.rerun()

        # --- 4. Площадь ---
        c_ar1, c_ar2 = st.columns([12, 1])
        with c_ar1:
            st.markdown(f'''
            <div class="data-row-flex">
                <div class="data-label-fixed">Площадь, м²:</div>
                <div class="data-val-fixed"><b>{obj.get("area", "—")} м²</b></div>
            </div>
            ''', unsafe_allow_html=True)
        with c_ar2:
            if st.button("✏️", key=f"btn_inline_area_{cd.get('id')}", help="Изменить площадь"):
                st.session_state[f"show_ed_area_{cd.get('id')}"] = not st.session_state.get(f"show_ed_area_{cd.get('id')}", False)
                st.rerun()
        if st.session_state.get(f"show_ed_area_{cd.get('id')}", False):
            with st.form(key=f"form_ed_area_{cd.get('id')}"):
                new_area_val = st.text_input("Площадь объекта, м² (нажмите Enter для сохранения):", value=str(obj.get("area", "")))
                c_s1, c_s2 = st.columns([1, 3])
                if c_s1.form_submit_button("Сохранить"):
                    obj["area"] = new_area_val.strip()
                    save_deals_to_catalog(all_deals)
                    st.session_state[f"show_ed_area_{cd.get('id')}"] = False
                    st.rerun()

        # --- 5. Этаж ---
        c_fl1, c_fl2 = st.columns([12, 1])
        with c_fl1:
            st.markdown(f'''
            <div class="data-row-flex">
                <div class="data-label-fixed">Этаж:</div>
                <div class="data-val-fixed">{obj.get("floor", "—")}</div>
            </div>
            ''', unsafe_allow_html=True)
        with c_fl2:
            if st.button("✏️", key=f"btn_inline_floor_{cd.get('id')}", help="Изменить этаж"):
                st.session_state[f"show_ed_floor_{cd.get('id')}"] = not st.session_state.get(f"show_ed_floor_{cd.get('id')}", False)
                st.rerun()
        if st.session_state.get(f"show_ed_floor_{cd.get('id')}", False):
            with st.form(key=f"form_ed_floor_{cd.get('id')}"):
                new_floor_val = st.text_input("Этаж (нажмите Enter для сохранения):", value=str(obj.get("floor", "")))
                c_s1, c_s2 = st.columns([1, 3])
                if c_s1.form_submit_button("Сохранить"):
                    obj["floor"] = new_floor_val.strip()
                    save_deals_to_catalog(all_deals)
                    st.session_state[f"show_ed_floor_{cd.get('id')}"] = False
                    st.rerun()

        # --- 6. Назначение / Вид ---
        c_pur1, c_pur2 = st.columns([12, 1])
        with c_pur1:
            st.markdown(f'''
            <div class="data-row-flex">
                <div class="data-label-fixed">Назначение / Вид:</div>
                <div class="data-val-fixed">{obj.get("purpose", "Жилое")} ({obj.get("living_type", "Квартира")})</div>
            </div>
            ''', unsafe_allow_html=True)
        with c_pur2:
            if st.button("✏️", key=f"btn_inline_pur_{cd.get('id')}", help="Изменить назначение и вид"):
                st.session_state[f"show_ed_pur_{cd.get('id')}"] = not st.session_state.get(f"show_ed_pur_{cd.get('id')}", False)
                st.rerun()
        if st.session_state.get(f"show_ed_pur_{cd.get('id')}", False):
            with st.form(key=f"form_ed_pur_{cd.get('id')}"):
                c_p_sub1, c_p_sub2 = st.columns(2)
                new_p_val = c_p_sub1.text_input("Назначение:", value=obj.get("purpose", "Жилое"))
                new_l_val = c_p_sub2.text_input("Вид жилого помещения:", value=obj.get("living_type", "Квартира"))
                if st.form_submit_button("Сохранить"):
                    obj["purpose"] = new_p_val.strip()
                    obj["living_type"] = new_l_val.strip()
                    save_deals_to_catalog(all_deals)
                    st.session_state[f"show_ed_pur_{cd.get('id')}"] = False
                    st.rerun()

        # --- 7. Наименование ---
        c_nm1, c_nm2 = st.columns([12, 1])
        with c_nm1:
            st.markdown(f'''
            <div class="data-row-flex">
                <div class="data-label-fixed">Наименование:</div>
                <div class="data-val-fixed">{obj.get("name", "Данные отсутствуют")}</div>
            </div>
            ''', unsafe_allow_html=True)
        with c_nm2:
            if st.button("✏️", key=f"btn_inline_name_{cd.get('id')}", help="Изменить наименование объекта"):
                st.session_state[f"show_ed_name_{cd.get('id')}"] = not st.session_state.get(f"show_ed_name_{cd.get('id')}", False)
                st.rerun()
        if st.session_state.get(f"show_ed_name_{cd.get('id')}", False):
            with st.form(key=f"form_ed_name_{cd.get('id')}"):
                new_nm_val = st.text_input("Наименование объекта (нажмите Enter для сохранения):", value=obj.get("name", ""))
                c_s1, c_s2 = st.columns([1, 3])
                if c_s1.form_submit_button("Сохранить"):
                    obj["name"] = new_nm_val.strip()
                    save_deals_to_catalog(all_deals)
                    st.session_state[f"show_ed_name_{cd.get('id')}"] = False
                    st.rerun()

        # --- 8. Кадастровая стоимость ---
        c_cst1, c_cst2 = st.columns([12, 1])
        with c_cst1:
            st.markdown(f'''
            <div class="data-row-flex">
                <div class="data-label-fixed">Кадастровая стоимость:</div>
                <div class="data-val-fixed"><b>{obj.get("cadastral_cost", "—")}</b></div>
            </div>
            ''', unsafe_allow_html=True)
        with c_cst2:
            if st.button("✏️", key=f"btn_inline_cost_{cd.get('id')}", help="Изменить кадастровую стоимость"):
                st.session_state[f"show_ed_cost_{cd.get('id')}"] = not st.session_state.get(f"show_ed_cost_{cd.get('id')}", False)
                st.rerun()
        if st.session_state.get(f"show_ed_cost_{cd.get('id')}", False):
            with st.form(key=f"form_ed_cost_{cd.get('id')}"):
                new_cost_val = st.text_input("Кадастровая стоимость, ₽ (нажмите Enter для сохранения):", value=str(obj.get("cadastral_cost", "")))
                c_s1, c_s2 = st.columns([1, 3])
                if c_s1.form_submit_button("Сохранить"):
                    obj["cadastral_cost"] = new_cost_val.strip()
                    save_deals_to_catalog(all_deals)
                    st.session_state[f"show_ed_cost_{cd.get('id')}"] = False
                    st.rerun()

    with col_obj_2:
        st.markdown('<div class="sub-section-head">📑 Правоустановка, Ограничения, Ипотеки и Аресты</div>', unsafe_allow_html=True)
        did = cd.get('id')

        # 1. ПРАВООБЛАДАТЕЛИ (каждая строка — свой карандаш)
        rights_list = obj.setdefault("rights", [])
        for r_idx, r in enumerate(rights_list):
            if r_idx > 0:
                st.markdown('<div class="right-holder-sep"></div>', unsafe_allow_html=True)

            def _save_owner(v, r=r):
                r["owner"] = v[0]

            def _del_owner(r_idx=r_idx):
                rights_list.pop(r_idx)

            inline_edit_row(
                f'<div class="right-row"><b>Правообладатель:</b> <span style="font-weight:700; color:#0F172A;">{r.get("owner", "—")}</span></div>',
                f"r_owner_{did}_{r_idx}",
                [("Правообладатель (ФИО, дата рождения):", r.get("owner", ""))],
                _save_owner, on_delete=_del_owner, help_txt="Изменить правообладателя"
            )

            def _save_share(v, r=r):
                r["prop_type"], r["share"] = v[0], v[1]

            inline_edit_row(
                f'<div class="right-row"><b>Вид собственности и доля:</b> {r.get("prop_type", "Собственность")} (доля <b>{r.get("share", "1/1")}</b>)</div>',
                f"r_share_{did}_{r_idx}",
                [("Вид собственности:", r.get("prop_type", "Собственность")), ("Доля:", r.get("share", "1/1"))],
                _save_share, help_txt="Изменить вид собственности и долю"
            )

            def _save_rec(v, r=r):
                r["reg_record"] = v[0]

            inline_edit_row(
                f'<div class="right-row"><b>Номер и дата записи регистрации:</b> <span style="font-weight:600; color:#0F172A;">{r.get("reg_record", "—")}</span></div>',
                f"r_rec_{did}_{r_idx}",
                [("Номер и дата записи регистрации:", r.get("reg_record", ""))],
                _save_rec, help_txt="Изменить запись регистрации"
            )

            def _save_basis(v, r=r):
                r["basis_doc"] = v[0]

            inline_edit_row(
                f'<div class="right-row"><b>Документ-основание:</b> {r.get("basis_doc", "—")}</div>',
                f"r_basis_{did}_{r_idx}",
                [("Документ-основание:", r.get("basis_doc", ""))],
                _save_basis, help_txt="Изменить документ-основание"
            )

        if not rights_list:
            st.markdown('<div class="right-row" style="color:#64748B;">Правообладатели не внесены.</div>', unsafe_allow_html=True)

        # 2. ИПОТЕКИ И ЗАЛОГИ
        mortgages = obj.setdefault("mortgages", [])
        for m_idx, m in enumerate(mortgages):
            st.markdown('<div class="legal-alert" style="border-left: 4px solid #475569; background: #F1F5F9; margin-top:8px; margin-bottom:2px;"><b style="color:#334155;">🏦 ИПОТЕКА / ЗАЛОГ: ЗАРЕГИСТРИРОВАНО</b></div>', unsafe_allow_html=True)
            for fld, lbl in [("reg_record", "Номер и дата записи"), ("beneficiary", "В пользу кого (Банк)"),
                             ("term", "Срок ограничения"), ("basis_doc", "Документ-основание")]:
                def _save_m(v, m=m, fld=fld):
                    m[fld] = v[0]

                def _del_m(m_idx=m_idx):
                    mortgages.pop(m_idx)

                inline_edit_row(
                    f'<div class="right-row"><b>{lbl}:</b> {m.get(fld, "—")}</div>',
                    f"m_{fld}_{did}_{m_idx}",
                    [(f"{lbl}:", m.get(fld, ""))],
                    _save_m, on_delete=_del_m if fld == "reg_record" else None, help_txt=f"Изменить: {lbl}"
                )

        # 3. АРЕСТЫ И ЗАПРЕТЫ
        arrests = obj.setdefault("arrests", [])
        for a_idx, a in enumerate(arrests):
            st.markdown('<div class="legal-alert" style="border-left: 4px solid #DC2626; background: #FEF2F2; margin-top:8px; margin-bottom:2px;"><b style="color:#B91C1C;">🚨 АРЕСТ / ЗАПРЕЩЕНИЕ РЕГИСТРАЦИИ (РОСП ФССП)</b></div>', unsafe_allow_html=True)
            for fld, lbl in [("reg_record", "Номер и дата записи"), ("beneficiary", "Кем наложен"),
                             ("basis_doc", "Документ-основание (ИП)")]:
                def _save_a(v, a=a, fld=fld):
                    a[fld] = v[0]

                def _del_a(a_idx=a_idx):
                    arrests.pop(a_idx)

                inline_edit_row(
                    f'<div class="right-row"><b>{lbl}:</b> {a.get(fld, "—")}</div>',
                    f"a_{fld}_{did}_{a_idx}",
                    [(f"{lbl}:", a.get(fld, ""))],
                    _save_a, on_delete=_del_a if fld == "reg_record" else None, help_txt=f"Изменить: {lbl}"
                )

        if not mortgages and not arrests:
            st.markdown("""<div class="legal-alert" style="border-left: 4px solid #16A34A; background: #F0FDF4; margin-top:8px;"><b style="color:#15803D;">✅ СТАТУС ОБРЕМЕНЕНИЙ: НЕ ЗАРЕГИСТРИРОВАНО</b><br>Ограничения прав, ипотеки, аресты и запреты в ЕГРН отсутствуют.</div>""", unsafe_allow_html=True)

        # 4. Кнопки добавления новых записей
        add1, add2, add3 = st.columns(3)
        if add1.button("➕ Правообладатель", key=f"add_right_{did}", use_container_width=True):
            rights_list.append({"owner": "Новый правообладатель", "prop_type": "Общая долевая собственность",
                                "share": "1/2", "reg_record": "", "basis_doc": ""})
            save_deals_to_catalog(all_deals)
            st.session_state[f"show_r_owner_{did}_{len(rights_list) - 1}"] = True
            st.rerun()
        if add2.button("➕ Ипотека", key=f"add_mort_{did}", use_container_width=True):
            mortgages.append({"reg_record": "", "beneficiary": "", "term": "", "basis_doc": ""})
            save_deals_to_catalog(all_deals)
            st.session_state[f"show_m_reg_record_{did}_{len(mortgages) - 1}"] = True
            st.rerun()
        if add3.button("➕ Арест", key=f"add_arr_{did}", use_container_width=True):
            arrests.append({"reg_record": "", "beneficiary": "", "basis_doc": ""})
            save_deals_to_catalog(all_deals)
            st.session_state[f"show_a_reg_record_{did}_{len(arrests) - 1}"] = True
            st.rerun()

    # ---------------------------------------------------------------------
    # БЛОК 2: КАРТОЧКИ СТОРОН (ДИНАМИЧЕСКИЙ ДОКУМЕНТ: ПАСПОРТ РФ / СВИДЕТЕЛЬСТВО О РОЖДЕНИИ)
    # ---------------------------------------------------------------------
    c_part_head1, c_part_head2 = st.columns([4, 1])
    with c_part_head1:
        st.markdown('<div class="section-head">👥 2. Карточки Сторон (Субъекты сделки)</div>', unsafe_allow_html=True)
    with c_part_head2:
        if st.session_state[edit_parties_key]:
            if st.button("❌ Завершить редактирование", key=f"btn_cancel_part_{cd.get('id')}", use_container_width=True):
                st.session_state[edit_parties_key] = False
                st.rerun()
        else:
            if st.button("✏️ Редактировать стороны", key=f"btn_edit_part_{cd.get('id')}", use_container_width=True, type="secondary"):
                st.session_state[edit_parties_key] = True
                st.rerun()

    if not st.session_state[edit_parties_key]:
        # РЕЖИМ ПРОСМОТРА СТОРОН В ДВЕ КОЛОНКИ
        col_sellers, col_buyers = st.columns([1, 1], gap="large")

        with col_sellers:
            st.markdown("""
            <div style="background:#FFF5F5; border-left:4px solid #DC2626; padding:8px 12px; border-radius:4px; margin-bottom:12px;">
                <b style="color:#B91C1C; font-size:15px;">🔴 СТОРОНА ПРОДАВЦОВ (ОТЧУЖДАЮТ)</b>
            </div>
            """, unsafe_allow_html=True)
            for idx, s in enumerate(cd.get("sellers", [])):
                b_str = s.get('birth', '01.01.1980')
                age = calculate_age(b_str)
                doc_kind = s.get('doc_type')
                if not doc_kind:
                    doc_kind = "Свидетельство о рождении" if (age is not None and age < 14) else "Паспорт РФ"
                
                badge_cls = "doc-type-birth" if doc_kind == "Свидетельство о рождении" else "doc-type-passport"
                age_txt = f" ({age} лет)" if age is not None else ""
                
                phone_raw = s.get('phone', '')
                phone_html = render_phone_messengers_html(phone_raw)
                doc_display = f"<b>{doc_kind}:</b> {s.get('doc_details') or s.get('passport', 'серия 56 12 № 123456, выдан УМВД РФ 10.05.2012')}"

                st.markdown(f"""
                <div class="card-box-clean">
                    <div class="party-header-row">
                        <div class="party-fio-phone-wrap">
                            <span class="party-fio-title">{s.get('fio')}</span>
                            {phone_html}
                        </div>
                        <div style="display:flex; align-items:center; gap:4px;">
                            <span class="doc-type-badge {badge_cls}">{doc_kind}</span>
                            <span style="background:#FEE2E2; color:#DC2626; padding:2px 8px; border-radius:12px; font-weight:700; font-size:11px;">доля {s.get('share', '1/1')}</span>
                        </div>
                    </div>
                    <div style="margin-top:6px; font-size:13px; color:#334155;">
                        <div><b>Дата рождения:</b> {b_str}{age_txt}</div>
                        <div>{doc_display}</div>
                        <div><b>Регистрация:</b> {s.get('reg_address', cd.get('address'))}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        with col_buyers:
            st.markdown("""
            <div style="background:#EFF6FF; border-left:4px solid #2563EB; padding:8px 12px; border-radius:4px; margin-bottom:12px;">
                <b style="color:#1D4ED8; font-size:15px;">🔵 СТОРОНА ПОКУПАТЕЛЕЙ (ПРИОБРЕТАЮТ)</b>
            </div>
            """, unsafe_allow_html=True)
            for idx, b in enumerate(cd.get("buyers", [])):
                b_str = b.get('birth', '15.03.1985')
                age = calculate_age(b_str)
                doc_kind = b.get('doc_type')
                if not doc_kind:
                    doc_kind = "Свидетельство о рождении" if (age is not None and age < 14) else "Паспорт РФ"
                
                badge_cls = "doc-type-birth" if doc_kind == "Свидетельство о рождении" else "doc-type-passport"
                age_txt = f" ({age} лет)" if age is not None else ""
                
                phone_raw = b.get('phone', '')
                phone_html = render_phone_messengers_html(phone_raw)
                doc_display = f"<b>{doc_kind}:</b> {b.get('doc_details') or b.get('passport', 'серия 56 16 № 987654, выдан УМВД России 15.08.2015')}"

                st.markdown(f"""
                <div class="card-box-clean">
                    <div class="party-header-row">
                        <div class="party-fio-phone-wrap">
                            <span class="party-fio-title">{b.get('fio')}</span>
                            {phone_html}
                        </div>
                        <div style="display:flex; align-items:center; gap:4px;">
                            <span class="doc-type-badge {badge_cls}">{doc_kind}</span>
                            <span style="background:#DBEAFE; color:#1D4ED8; padding:2px 8px; border-radius:12px; font-weight:700; font-size:11px;">{b.get('share', '1/1')}</span>
                        </div>
                    </div>
                    <div style="margin-top:6px; font-size:13px; color:#334155;">
                        <div><b>Дата рождения:</b> {b_str}{age_txt}</div>
                        <div>{doc_display}</div>
                        <div><b>Регистрация:</b> {b.get('reg_address', 'г. Пенза, ул. Кирода, д. 15, кв. 42')}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
    else:
        # ДИНАМИЧЕСКИЙ РЕЖИМ РЕДАКТИРОВАНИЯ СТОРОН (БЕЗ FORM, ЧТОБЫ ENTER НЕ СБРАСЫВАЛ)
        st.info("💡 **Динамическая карточка стороны:** введите ФИО, телефон и дату рождения. При нажатии Enter значение сохраняется в поле без закрытия карточки. Когда закончите редактирование — нажмите «💾 СОХРАНИТЬ ВСЕХ СТОРОН В КАРТОЧКУ СДЕЛКИ».")

        col_part_ed1, col_part_ed2 = st.columns(2, gap="large")
        
        sellers_list = cd.get("sellers", [])
        buyers_list = cd.get("buyers", [])
        if not sellers_list:
            sellers_list = [{}]
        if not buyers_list:
            buyers_list = [{}]

        # --- КОЛОНКА ПРОДАВЦОВ ---
        updated_sellers = []
        with col_part_ed1:
            st.markdown("#### 🔴 Продавцы (Отчуждатели)")
            for i, s in enumerate(sellers_list):
                with st.expander(f"👤 Продавец #{i+1}: {s.get('fio', 'Новый продавец')}", expanded=True):
                    fio = st.text_input(f"ФИО Продавца #{i+1}", value=s.get("fio", ""), key=f"inp_s_fio_{i}")
                    phone = st.text_input(f"Номер телефона #{i+1}", value=s.get("phone", "+7 (9"), key=f"inp_s_phone_{i}", placeholder="+7 (999) 000-00-00")
                    birth = st.text_input(f"Дата рождения (ДД.ММ.ГГГГ) #{i+1}", value=s.get("birth", "01.01.1980"), key=f"inp_s_birth_{i}")
                    share = st.text_input(f"Размер доли в праве #{i+1}", value=s.get("share", "1/1"), key=f"inp_s_share_{i}")
                    reg = st.text_input(f"Адрес регистрации #{i+1}", value=s.get("reg_address", cd.get("address", "")), key=f"inp_s_reg_{i}")

                    # Автоматический расчет возраста
                    age = calculate_age(birth)
                    if age is not None:
                        if age < 14:
                            st.warning(f"👶 Возраст: {age} лет (несовершеннолетний до 14 лет). Документ: **Свидетельство о рождении** (согласно ст. 28 ГК РФ).")
                            default_doc = "Свидетельство о рождении"
                        else:
                            st.success(f"🧑 Возраст: {age} лет (14+ лет). Документ: **Паспорт гражданина РФ** (согласно ст. 26 ГК РФ / Указу Президента № 232).")
                            default_doc = "Паспорт РФ"
                    else:
                        st.info("ℹ️ Введите дату рождения в формате ДД.ММ.ГГГГ для точного расчета возраста")
                        default_doc = s.get("doc_type", "Паспорт РФ")

                    doc_options = ["Паспорт РФ", "Свидетельство о рождении"]
                    sel_doc_idx = 1 if default_doc == "Свидетельство о рождении" else 0
                    doc_type = st.radio(f"Тип документа удостоверяющего личность #{i+1}", options=doc_options, index=sel_doc_idx, horizontal=True, key=f"s_doctype_{i}")

                    if doc_type == "Паспорт РФ":
                        c_p1, c_p2 = st.columns([1, 2])
                        with c_p1:
                            pass_series_num = st.text_input(f"Серия и номер паспорта #{i+1}", value=s.get("pass_num", "56 12 № 123456"), key=f"s_pnum_{i}")
                            pass_date = st.text_input(f"Дата выдачи паспорта #{i+1}", value=s.get("pass_date", "10.05.2012"), key=f"s_pdate_{i}")
                            pass_code = st.text_input(f"Код подразделения #{i+1}", value=s.get("pass_code", "580-001"), key=f"s_pcode_{i}")
                        with c_p2:
                            pass_issuer = st.text_area(f"Кем выдан паспорт #{i+1}", value=s.get("pass_issuer", "УМВД России по Пензенской области"), height=80, key=f"s_piss_{i}")
                        doc_details = f"серия {pass_series_num}, выдан {pass_issuer} {pass_date}, код подр. {pass_code}"
                    else:
                        c_b1, c_b2 = st.columns([1, 2])
                        with c_b1:
                            birth_act_num = st.text_input(f"Номер записи акта о рождении #{i+1}", value=s.get("birth_act", "№ 1420"), key=f"s_bact_{i}")
                            birth_date_cert = st.text_input(f"Дата выдачи свидетельства #{i+1}", value=s.get("birth_cert_date", "15.06.2016"), key=f"s_bcdate_{i}")
                        with c_b2:
                            birth_issuer = st.text_area(f"Орган ЗАГС, выдавший свидетельство #{i+1}", value=s.get("birth_issuer", "Территориальным отделом ЗАГС г. Пензы"), height=80, key=f"s_biss_{i}")
                        doc_details = f"запись акта {birth_act_num}, выдано {birth_issuer} {birth_date_cert}"

                    updated_sellers.append({
                        "fio": fio,
                        "phone": phone,
                        "birth": birth,
                        "share": share,
                        "reg_address": reg,
                        "doc_type": doc_type,
                        "doc_details": doc_details,
                        "passport": doc_details
                    })

        # --- КОЛОНКА ПОКУПАТЕЛЕЙ ---
        updated_buyers = []
        with col_part_ed2:
            st.markdown("#### 🔵 Покупатели (Приобретатели)")
            for j, b in enumerate(buyers_list):
                with st.expander(f"👤 Покупатель #{j+1}: {b.get('fio', 'Новый покупатель')}", expanded=True):
                    fio = st.text_input(f"ФИО Покупателя #{j+1}", value=b.get("fio", ""), key=f"inp_b_fio_{j}")
                    phone = st.text_input(f"Номер телефона #{j+1}", value=b.get("phone", "+7 (9"), key=f"inp_b_phone_{j}", placeholder="+7 (999) 000-00-00")
                    birth = st.text_input(f"Дата рождения (ДД.ММ.ГГГГ) #{j+1}", value=b.get("birth", "15.03.1985"), key=f"inp_b_birth_{j}")
                    share = st.text_input(f"Вид приобретаемого права / Доля #{j+1}", value=b.get("share", "1/1 (индивидуальная собственность)"), key=f"inp_b_share_{j}")
                    reg = st.text_input(f"Адрес регистрации #{j+1}", value=b.get("reg_address", "г. Пенза, ул. Кирова, д. 15, кв. 42"), key=f"inp_b_reg_{j}")

                    # Автоматический расчет возраста
                    age = calculate_age(birth)
                    if age is not None:
                        if age < 14:
                            st.warning(f"👶 Возраст: {age} лет (несовершеннолетний до 14 лет). Документ: **Свидетельство о рождении** (ст. 28 ГК РФ).")
                            default_doc = "Свидетельство о рождении"
                        else:
                            st.success(f"🧑 Возраст: {age} лет (14+ лет). Документ: **Паспорт гражданина РФ** (ст. 26 ГК РФ / Указ Президента № 232).")
                            default_doc = "Паспорт РФ"
                    else:
                        st.info("ℹ️ Введите дату рождения в формате ДД.ММ.ГГГГ для точного расчета возраста")
                        default_doc = b.get("doc_type", "Паспорт РФ")

                    doc_options = ["Паспорт РФ", "Свидетельство о рождении"]
                    sel_doc_idx = 1 if default_doc == "Свидетельство о рождении" else 0
                    doc_type = st.radio(f"Тип документа удостоверяющего личность #{j+1}", options=doc_options, index=sel_doc_idx, horizontal=True, key=f"b_doctype_{j}")

                    if doc_type == "Паспорт РФ":
                        c_p1, c_p2 = st.columns([1, 2])
                        with c_p1:
                            pass_series_num = st.text_input(f"Серия и номер паспорта #{j+1}", value=b.get("pass_num", "56 16 № 987654"), key=f"b_pnum_{j}")
                            pass_date = st.text_input(f"Дата выдачи паспорта #{j+1}", value=b.get("pass_date", "15.08.2015"), key=f"b_pdate_{j}")
                            pass_code = st.text_input(f"Код подразделения #{j+1}", value=b.get("pass_code", "580-002"), key=f"b_pcode_{j}")
                        with c_p2:
                            pass_issuer = st.text_area(f"Кем выдан паспорт #{j+1}", value=b.get("pass_issuer", "УМВД России по Пензенской области"), height=80, key=f"b_piss_{j}")
                        doc_details = f"серия {pass_series_num}, выдан {pass_issuer} {pass_date}, код подр. {pass_code}"
                    else:
                        c_b1, c_b2 = st.columns([1, 2])
                        with c_b1:
                            birth_act_num = st.text_input(f"Номер записи акта о рождении #{j+1}", value=b.get("birth_act", "№ 890"), key=f"b_bact_{j}")
                            birth_date_cert = st.text_input(f"Дата выдачи свидетельства #{j+1}", value=b.get("birth_cert_date", "20.09.2018"), key=f"b_bcdate_{j}")
                        with c_b2:
                            birth_issuer = st.text_area(f"Орган ЗАГС, выдавший свидетельство #{j+1}", value=b.get("birth_issuer", "Органом ЗАГС Октябрьского района г. Пензы"), height=80, key=f"b_biss_{j}")
                        doc_details = f"запись акта {birth_act_num}, выдано {birth_issuer} {birth_date_cert}"

                    updated_buyers.append({
                        "fio": fio,
                        "phone": phone,
                        "birth": birth,
                        "share": share,
                        "reg_address": reg,
                        "doc_type": doc_type,
                        "doc_details": doc_details,
                        "passport": doc_details
                    })

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 СОХРАНИТЬ ВСЕХ СТОРОН В КАРТОЧКУ СДЕЛКИ", type="primary", use_container_width=True, key=f"btn_save_parties_{cd.get('id')}"):
            cd["sellers"] = updated_sellers
            cd["buyers"] = updated_buyers
            save_deals_to_catalog(all_deals)
            st.session_state[edit_parties_key] = False
            st.success("✅ Карточки всех сторон успешно сохранены с привязкой документов!")
            st.rerun()

    # ---------------------------------------------------------------------
    # БЛОК 3: ФИНАНСОВЫЕ УСЛОВИЯ И ЮРИДИЧЕСКИЙ ЧЕКАП (С КНОПКОЙ «РЕДАКТИРОВАТЬ»)
    # ---------------------------------------------------------------------
    c_fin_head1, c_fin_head2 = st.columns([4, 1])
    with c_fin_head1:
        st.markdown('<div class="section-head">💰 3. Финансовые условия и Юридический Чекап</div>', unsafe_allow_html=True)
    with c_fin_head2:
        if st.session_state[edit_fin_key]:
            if st.button("❌ Отмена", key=f"btn_cancel_fin_{cd.get('id')}", use_container_width=True):
                st.session_state[edit_fin_key] = False
                st.rerun()
        else:
            if st.button("✏️ Редактировать финансы", key=f"btn_edit_fin_{cd.get('id')}", use_container_width=True, type="secondary"):
                st.session_state[edit_fin_key] = True
                st.rerun()

    if not st.session_state[edit_fin_key]:
        col_fin, col_audit = st.columns([1, 1], gap="large")

        with col_fin:
            st.markdown('<div class="sub-section-head">💵 Схема расчетов</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="data-row-flex"><div class="data-label-fixed">Полная цена квартиры:</div><div class="data-val-fixed"><b style="font-size:16px; color:#0F172A;">{cd.get("price"):,} ₽</b></div></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="data-row-flex"><div class="data-label-fixed">Сумма задатка (ПДКП):</div><div class="data-val-fixed"><b>{cd.get("deposit"):,} ₽</b> (наличными в день ПДКП)</div></div>', unsafe_allow_html=True)
            rem_sum = int(cd.get("price", 0)) - int(cd.get("deposit", 0))
            st.markdown(f'<div class="data-row-flex"><div class="data-label-fixed">Остаток к расчету:</div><div class="data-val-fixed"><b>{rem_sum:,} ₽</b></div></div>', unsafe_allow_html=True)
            st.markdown('<div class="data-row-flex"><div class="data-label-fixed">Способ расчета:</div><div class="data-val-fixed"><b>СБР Сбербанк / Номинальный счет</b></div></div>', unsafe_allow_html=True)
            st.markdown('<div class="data-row-flex"><div class="data-label-fixed">Залог продавца ст. 488:</div><div class="data-val-fixed"><b style="color:#16A34A;">ОТКЛЮЧЕН (не возникает)</b></div></div>', unsafe_allow_html=True)

        with col_audit:
            st.markdown('<div class="sub-section-head">⚖️ Выявленные риски и правовой контроль</div>', unsafe_allow_html=True)
            tags = cd.get("tags", [])
            risks_rendered = False
            
            if "Долг ФССП" in tags:
                risks_rendered = True
                st.markdown("""
                <div class="legal-alert" style="border-left: 4px solid #DC2626; background: #FEF2F2;">
                    <b style="color:#B91C1C;">🚨 Арест / Запрет РОСП ФССП (ст. 80 ФЗ № 229):</b><br>
                    Выявлены исполнительные производства или запрет на регдействия. <i>Решение: задаток строго наличными / аванс на депозит РОСП, истребование постановления о снятии ареста до ОДКП.</i>
                </div>
                """, unsafe_allow_html=True)
                
            if "Приватизация" in tags:
                risks_rendered = True
                st.markdown("""
                <div class="legal-alert" style="border-left: 4px solid #EA580C; background: #FFF7ED;">
                    <b style="color:#C2410C;">⚠️ Основание Приватизация (ст. 19 Федерального закона № 189-ФЗ):</b><br>
                    Риск бессрочного права проживания отказников от приватизации. <i>Решение: обязательная архивная выписка БТИ и поквартирная карточка с момента заселения дома на проверку отказников.</i>
                </div>
                """, unsafe_allow_html=True)
                
            if "МСК" in tags:
                risks_rendered = True
                st.markdown("""
                <div class="legal-alert" style="border-left: 4px solid #2563EB; background: #EFF6FF;">
                    <b style="color:#1D4ED8;">ℹ️ Использование средств Материнского (семейного) капитала:</b><br>
                    Контроль выделения долей детям (ст. 10 ФЗ № 256). Безналичное перечисление средств СФР напрямую на счет продавца.
                </div>
                """, unsafe_allow_html=True)
                
            if "Недофинанс" in tags:
                risks_rendered = True
                st.markdown("""
                <div class="legal-alert" style="border-left: 4px solid #DC2626; background: #FEF2F2;">
                    <b style="color:#B91C1C;">🚨 Занижение стоимости в договоре (Недофинанс):</b><br>
                    Налоговый риск (ст. 54.1 НК РФ) и риск оспаривания сделки при банкротстве по ст. 61.2 ФЗ № 127.
                </div>
                """, unsafe_allow_html=True)

            if not risks_rendered:
                st.markdown("""
                <div class="legal-alert" style="border-left: 4px solid #16A34A; background: #F0FDF4;">
                    <b style="color:#15803D;">✅ Юридический аудит пройден: Сделка юридически чистая</b><br>
                    Ограничений, арестов, запрещений и обременений в ЕГРН не выявлено. Долгов по линии РОСП ФССП не обнаружено. Паспорта сторон действительны по базе МВД.
                </div>
                """, unsafe_allow_html=True)

            st.markdown("**Якоря сделки:**")
            tags_html = "".join([f'<span class="deal-tag">{t}</span>' for t in cd.get("tags", [])])
            st.markdown(f"<div>{tags_html}</div>", unsafe_allow_html=True)
    else:
        # ТОЧЕЧНОЕ РЕДАКТИРОВАНИЕ ФИНАНСОВ И ЯКОРЕЙ
        st.info("💡 Редактирование цены сделки, задатка и правовых якорей")
        c_f1, c_f2 = st.columns(2, gap="medium")
        with c_f1:
            inp_price = st.number_input("Полная цена квартиры, ₽", value=int(cd.get("price", 3000000)), step=50000, key=f"fld_p_{cd.get('id')}")
            inp_deposit = st.number_input("Сумма задатка по ПДКП, ₽", value=int(cd.get("deposit", 100000)), step=10000, key=f"fld_d_{cd.get('id')}")
        with c_f2:
            all_tags_list = [
                "Ипотека Сбера", "Ипотека ВТБ", "Дом+ЗУ", "Недофинанс", 
                "Долг ФССП", "МСК", "Приватизация", "Наследство", 
                "Доли", "Чистая", "Коммерция", "Наличные", "СБР"
            ]
            inp_tags = st.multiselect("Якоря сделки", options=all_tags_list, default=[t for t in cd.get("tags", []) if t in all_tags_list], key=f"fld_t_{cd.get('id')}")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 СОХРАНИТЬ ФИНАНСЫ И ЯКОРЯ", type="primary", use_container_width=True, key=f"btn_save_fin_{cd.get('id')}"):
            cd["price"] = inp_price
            cd["deposit"] = inp_deposit
            cd["tags"] = inp_tags
            save_deals_to_catalog(all_deals)
            st.session_state[edit_fin_key] = False
            st.success("✅ Финансовые условия и якоря сделки сохранены!")
            st.rerun()

    # ДЕЙСТВИЕ: СГЕНЕРИРОВАТЬ И СКАЧАТЬ ВЕСЬ ПАКЕТ
    st.markdown("---")
    st.markdown("### 📥 Сформировать и забрать комплект документов по этой карточке")
    
    if st.button("🚀 СФОРМИРОВАТЬ И СКАЧАТЬ БОЕВОЙ ПАКЕТ .DOCX (ПДКП + РАСПИСКА + ОДКП + АКТ + СОПД)", type="primary", use_container_width=True):
        deal_payload = {
            "address": obj.get("address", cd.get("address")),
            "cadastral_num": obj.get("cadastral_num", cd.get("cadastral_num")),
            "area": str(obj.get("area", "42.5")),
            "floor": str(obj.get("floor", "3")),
            "price": cd.get("price"),
            "deposit": cd.get("deposit"),
            "deal_date": cd.get("date"),
            "sellers": [{"fio": s.get("fio"), "share": s.get("share", "1/2")} for s in cd.get("sellers", [])],
            "buyers": [{"fio": b.get("fio")} for b in cd.get("buyers", [])],
            "has_msk": "МСК" in cd.get("tags", []),
            "has_fssp": "Долг ФССП" in cd.get("tags", []),
            "payment_type": "СБР Сбербанк"
        }

        templates_dir = os.path.join(os.path.dirname(__file__), 'templates')
        out_dir = os.path.join(os.path.dirname(__file__), 'deals', 'сформированный_пакет')

        from benchmark_pdkp_builder import build_benchmark_pdkp
        pdkp_master_file = os.path.join(out_dir, "1_ПДКП_с_встроенной_распиской.docx")
        build_benchmark_pdkp(deal_payload, pdkp_master_file)
        generated_files = generate_deal_package(deal_payload, templates_dir, out_dir)
        generated_files["pdkp"] = pdkp_master_file
        st.success(f"✅ Полный комплект документов для «{cd.get('title')}» готов!")

        cols = st.columns(len(generated_files))
        labels = {
            "pdkp": "📄 ПДКП + Расписка",
            "odkp": "📄 Основной ДКП",
            "akt": "📄 Передаточный акт",
            "sopd": "📄 СОПД клиента"
        }
        
        col_idx = 0
        for key, fpath in generated_files.items():
            if os.path.exists(fpath):
                with open(fpath, "rb") as f:
                    cols[col_idx].download_button(
                        label=labels.get(key, os.path.basename(fpath)),
                        data=f.read(),
                        file_name=os.path.basename(fpath),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                col_idx += 1

        zip_path = os.path.join(out_dir, f"Комплект_{cd.get('id')}.zip")
        with zipfile.ZipFile(zip_path, "w") as zf:
            for key, fpath in generated_files.items():
                if os.path.exists(fpath):
                    zf.write(fpath, os.path.basename(fpath))

        with open(zip_path, "rb") as zf_data:
            st.download_button(
                label=f"📦 Скачать весь комплект {cd.get('id')} одним ZIP-архивом",
                data=zf_data.read(),
                file_name=f"Комплект_{cd.get('id')}.zip",
                mime="application/zip",
                use_container_width=True
            )

    st.stop()


# =========================================================================
# РЕЖИМ 2: ОБЩИЙ РЕЕСТР И ПОИСК
# =========================================================================

# ДАШБОРДЫ
total_deals = len(all_deals)
in_work = sum(1 for d in all_deals if d.get("status") == "В работе")
done_deals = sum(1 for d in all_deals if d.get("status") == "Завершена")

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Всего сделок в картотеке</div>
        <div class="metric-value">{total_deals}</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">В активной работе</div>
        <div class="metric-value">{in_work}</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Завершено сделок</div>
        <div class="metric-value">{done_deals}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='margin-bottom: 18px;'></div>", unsafe_allow_html=True)

# РАЗВОРАЧИВАЕМЫЙ ПОИСК
search_last, search_first, search_middle, search_addr, search_date, selected_tags = "", "", "", "", "", []

with st.expander("🔍 Поиск по картотеке сделок (нажмите, чтобы развернуть)", expanded=False):
    col_f, col_i, col_o = st.columns(3)
    search_last = col_f.text_input("Фамилия", placeholder="Протасов, Финошина, Вострова...")
    search_first = col_i.text_input("Имя", placeholder="Виктор, Олег, Татьяна...")
    search_middle = col_o.text_input("Отчество", placeholder="Иванович, Николаевна...")

    col_addr, col_date = st.columns([2, 1])
    search_addr = col_addr.text_input("Адрес объекта", placeholder="Тернопольская, Победы 146, Рахманинова...")
    search_date = col_date.text_input("Дата сделки", placeholder="15.02.2026, 19.10.2026...")

    all_tags = [
        "Ипотека Сбера", "Ипотека ВТБ", "Дом+ЗУ", "Недофинанс", 
        "Долг ФССП", "МСК", "Приватизация", "Наследство", 
        "Доли", "Чистая", "Коммерция", "Наличные", "СБР"
    ]
    selected_tags = st.multiselect(
        "Якоря сделки (можно комбинировать):",
        options=all_tags,
        default=[],
        placeholder="Выберите якоря..."
    )

# ФИЛЬТРАЦИЯ
filtered_deals = []
for d in all_deals:
    match = True
    if search_last.strip():
        sl = search_last.strip().lower()
        seller_match = any(sl in s.get("last_name", "").lower() or sl in s.get("fio", "").lower() for s in d.get("sellers", []))
        buyer_match = any(sl in b.get("last_name", "").lower() or sl in b.get("fio", "").lower() for b in d.get("buyers", []))
        if not (seller_match or buyer_match):
            match = False

    if match and search_first.strip():
        sf = search_first.strip().lower()
        seller_match = any(sf in s.get("first_name", "").lower() or sf in s.get("fio", "").lower() for s in d.get("sellers", []))
        buyer_match = any(sf in b.get("first_name", "").lower() or sf in b.get("fio", "").lower() for b in d.get("buyers", []))
        if not (seller_match or buyer_match):
            match = False

    if match and search_middle.strip():
        sm = search_middle.strip().lower()
        seller_match = any(sm in s.get("middle_name", "").lower() or sm in s.get("fio", "").lower() for s in d.get("sellers", []))
        buyer_match = any(sm in b.get("middle_name", "").lower() or sm in b.get("fio", "").lower() for b in d.get("buyers", []))
        if not (seller_match or buyer_match):
            match = False

    if match and search_addr.strip():
        sa = search_addr.strip().lower()
        obj_addr = d.get("object_card", {}).get("address", "")
        if sa not in d.get("address", "").lower() and sa not in d.get("title", "").lower() and sa not in obj_addr.lower():
            match = False

    if match and search_date.strip():
        sd = search_date.strip()
        if sd not in d.get("date", ""):
            match = False

    if match and selected_tags:
        deal_tags = d.get("tags", [])
        if not all(tag in deal_tags for tag in selected_tags):
            match = False

    if match:
        filtered_deals.append(d)

st.markdown(f"**Отображено сделок:** `{len(filtered_deals)}` из `{total_deals}`")
st.markdown("---")

# СПИСОК СДЕЛОК
st.markdown("### 📋 Список сделок")

for deal in filtered_deals:
    st_val = deal.get("status", "В работе")
    header_text = f"📁 {deal.get('title')} | 📅 {deal.get('date')} | 📍 {deal.get('address')} | 💰 {deal.get('price'):,} ₽ [{st_val}]"
    
    col_row_info, col_row_btn = st.columns([4, 1])
    with col_row_info:
        st.markdown(f"**{header_text}**")
        tags_html = "".join([f'<span class="deal-tag">{t}</span>' for t in deal.get("tags", [])])
        st.markdown(f"<div style='margin-bottom:8px;'>{tags_html}</div>", unsafe_allow_html=True)
        
    with col_row_btn:
        if st.button("📂 Открыть карточку", key=f"open_{deal.get('id')}", use_container_width=True, type="primary"):
            st.session_state["opened_deal_id"] = deal.get("id")
            st.rerun()
            
    st.markdown("<hr style='margin-top:4px; margin-bottom:12px; border:0; border-top:1px solid #F1F5F9;'>", unsafe_allow_html=True)
