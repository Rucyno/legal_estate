import os
import sys
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def generate_pdkp(
    font_size=9.5,
    line_spacing=1.02,
    space_after=0.8,
    deal_date="«15» февраля 2026 года",
    out_path=r"d:\Антигравити\legal_estate\deals\ПДКП_Протасовы_Фирсова.docx"
):
    doc = docx.Document()

    # Поля документа «Этажи»
    for s in doc.sections:
        s.top_margin = Inches(0.71)    # 18 mm (51.1 pt)
        s.bottom_margin = Inches(0.71) # 18 mm (51.1 pt)
        s.left_margin = Inches(0.79)   # 20 mm (56.9 pt)
        s.right_margin = Inches(0.59)  # 15 mm (42.5 pt)
        s.page_width = Inches(8.27)    # 210 mm (595.45 pt)
        s.page_height = Inches(11.69)  # 297 mm (841.7 pt)

    BODY_WIDTH = s.page_width - s.left_margin - s.right_margin # ~496 pt / 9921 dxa
    INDENT = Pt(25.0)  # 0.88 см

    def set_p_format(p, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=INDENT):
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(s_after)
        p.paragraph_format.line_spacing = l_spacing
        p.paragraph_format.left_indent = Pt(0)
        p.paragraph_format.right_indent = Pt(0)
        p.alignment = align
        if first_indent is not None:
            p.paragraph_format.first_line_indent = first_indent
        else:
            p.paragraph_format.first_line_indent = Pt(0)

    def add_h(doc, text):
        p = doc.add_paragraph()
        set_p_format(p, f_size=font_size, space_before=3.0, s_after=1.5, l_spacing=line_spacing, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size)
        return p

    def add_p(doc, text="", bold_prefix="", s_after=space_after, space_before=0, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=INDENT):
        p = doc.add_paragraph()
        set_p_format(p, f_size=font_size, space_before=space_before, s_after=s_after, l_spacing=line_spacing, align=align, first_indent=first_indent)
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.bold = True
            rb.font.name = "Times New Roman"
            rb.font.size = Pt(font_size)
        if text:
            r = p.add_run(text)
            r.font.name = "Times New Roman"
            r.font.size = Pt(font_size)
        return p

    def add_city_date_tab(doc, city_text="г. Пенза", date_text=deal_date):
        """Город и дата отдельным чистым параграфом с выравниванием даты по правому краю через Right Tab."""
        p = doc.add_paragraph()
        set_p_format(p, f_size=font_size, space_before=0, s_after=6.0, l_spacing=line_spacing, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
        p.paragraph_format.tab_stops.add_tab_stop(BODY_WIDTH, WD_TAB_ALIGNMENT.RIGHT)

        r_city = p.add_run(city_text)
        r_city.bold = True
        r_city.font.name = "Times New Roman"
        r_city.font.size = Pt(font_size)

        p.add_run("\t")

        r_date = p.add_run(date_text)
        r_date.bold = True
        r_date.font.name = "Times New Roman"
        r_date.font.size = Pt(font_size)
        return p

    def add_fullwidth_signature(doc, role_title, printed_name, underline_space_before=0, underline_space_after=1.0):
        """Оформление подписанта на всю ширину листа (17 см / 9921 dxa) по стандарту «Этажи»."""
        # 1. Заголовок стороны
        p_title = doc.add_paragraph()
        set_p_format(p_title, f_size=font_size, space_before=2.5, s_after=1.0, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
        p_title.paragraph_format.keep_with_next = True
        r_t = p_title.add_run(role_title)
        r_t.bold = True
        r_t.font.name = "Times New Roman"
        r_t.font.size = Pt(font_size)

        # 2. Черта для подписи на всю ширину листа
        p_line = doc.add_paragraph()
        set_p_format(p_line, f_size=font_size, space_before=underline_space_before, s_after=underline_space_after, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
        p_line.paragraph_format.keep_with_next = True
        pPr = p_line._p.get_or_add_pPr()
        pBdr = parse_xml(r"""
            <w:pBdr %s >
                <w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>
            </w:pBdr>
        """ % nsdecls("w"))
        pPr.append(pBdr)

        # 3. Подстрочник с выравниванием по правому краю
        p_sub = doc.add_paragraph()
        set_p_format(p_sub, f_size=8.0, space_before=0, s_after=1.0, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
        p_sub.paragraph_format.keep_with_next = True
        p_sub.paragraph_format.tab_stops.add_tab_stop(BODY_WIDTH, WD_TAB_ALIGNMENT.RIGHT)

        r_sub1 = p_sub.add_run("ФИО полностью")
        r_sub1.font.name = "Times New Roman"
        r_sub1.font.size = Pt(8.0)
        r_sub1.italic = True

        p_sub.add_run("\t")

        r_sub2 = p_sub.add_run("подпись")
        r_sub2.font.name = "Times New Roman"
        r_sub2.font.size = Pt(8.0)
        r_sub2.italic = True

        # 4. Печатное имя стороны
        p_name = doc.add_paragraph()
        set_p_format(p_name, f_size=font_size, space_before=0, s_after=3.0, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
        p_name.paragraph_format.keep_with_next = True
        r_n = p_name.add_run(printed_name)
        r_n.font.name = "Times New Roman"
        r_n.font.size = Pt(font_size)
        return p_name

    # =========================================================================
    # ШАПКА И ПРЕАМБУЛА
    # =========================================================================
    # Заголовок договора (12 pt, полужирный, по центру)
    p_title = doc.add_paragraph()
    set_p_format(p_title, f_size=12.0, space_before=0, s_after=2.0, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    r_t = p_title.add_run("ПРЕДВАРИТЕЛЬНЫЙ ДОГОВОР КУПЛИ-ПРОДАЖИ КВАРТИРЫ")
    r_t.bold = True
    r_t.font.name = "Times New Roman"
    r_t.font.size = Pt(12.0)

    # Подзаголовок (10 pt, курсив, по центру)
    p_sub = doc.add_paragraph()
    set_p_format(p_sub, f_size=10.0, space_before=0, s_after=6.0, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    r_s = p_sub.add_run("(с условием обеспечения обязательства задатком)")
    r_s.italic = True
    r_s.font.name = "Times New Roman"
    r_s.font.size = Pt(10.0)

    # Город и дата (параграф с табулятором по правому полю, полужирный, 10 pt)
    add_city_date_tab(doc, "г. Пенза", deal_date)

    # Вводная строка преамбулы
    p_intro = doc.add_paragraph()
    set_p_format(p_intro, f_size=font_size, space_before=0, s_after=1.2, l_spacing=line_spacing, first_indent=INDENT)
    r_in = p_intro.add_run("Мы, нижеподписавшиеся:")
    r_in.font.name = "Times New Roman"
    r_in.font.size = Pt(font_size)

    # Продавец 1
    p_pre1 = doc.add_paragraph()
    set_p_format(p_pre1, f_size=font_size, space_before=0, s_after=1.2, l_spacing=line_spacing, first_indent=INDENT)
    p_pre1.add_run("гр. ").font.name = "Times New Roman"
    r = p_pre1.add_run("Протасов Олег Иванович")
    r.bold = True
    r.font.name = "Times New Roman"
    p_pre1.add_run(", 15.04.1988 года рождения, пол: мужской, место рождения: гор. Пенза, паспорт гражданина РФ: серия 5612 № 345678, выдан Отделом УФМС России по Пензенской области в Ленинском р-не г. Пензы 20.04.2012 г., код подразделения: 580-001, адрес регистрации: Пензенская область, г. Пенза, ул. Суворова, д. 145, кв. 12, и").font.name = "Times New Roman"

    # Продавец 2
    p_pre2 = doc.add_paragraph()
    set_p_format(p_pre2, f_size=font_size, space_before=0, s_after=1.2, l_spacing=line_spacing, first_indent=INDENT)
    p_pre2.add_run("гр. ").font.name = "Times New Roman"
    r = p_pre2.add_run("Протасов Виктор Иванович")
    r.bold = True
    r.font.name = "Times New Roman"
    p_pre2.add_run(", 22.09.1963 года рождения, пол: мужской, место рождения: гор. Пенза, паспорт гражданина РФ: серия 5608 № 901234, выдан Отделом внутренних дел Октябрьского р-на г. Пензы 05.10.2008 г., код подразделения: 580-002, адрес регистрации: Пензенская область, г. Пенза, ул. Тернопольская, д. 7, кв. 34, именуемые в дальнейшем совместно ").font.name = "Times New Roman"
    r_pr = p_pre2.add_run("«Продавцы»")
    r_pr.bold = True
    r_pr.font.name = "Times New Roman"
    p_pre2.add_run(", с одной стороны, и").font.name = "Times New Roman"

    # Покупатель
    p_pre3 = doc.add_paragraph()
    set_p_format(p_pre3, f_size=font_size, space_before=0, s_after=1.2, l_spacing=line_spacing, first_indent=INDENT)
    p_pre3.add_run("гр. ").font.name = "Times New Roman"
    r = p_pre3.add_run("Фирсова Ангелина Сергеевна")
    r.bold = True
    r.font.name = "Times New Roman"
    p_pre3.add_run(", 10.08.1995 года рождения, пол: женский, место рождения: гор. Пенза, паспорт гражданина РФ: серия 5615 № 123456, выдан Отделом УФМС России по Пензенской области в Первомайском р-не г. Пензы 15.08.2015 г., код подразделения: 580-003, адрес регистрации: Пензенская область, г. Пенза, ул. Мира, д. 40, кв. 89, именуемая в дальнейшем ").font.name = "Times New Roman"
    r_by = p_pre3.add_run("«Покупатель»")
    r_by.bold = True
    r_by.font.name = "Times New Roman"
    p_pre3.add_run(", с другой стороны,").font.name = "Times New Roman"

    # Преамбула заключение
    add_p(doc, "при совместном упоминании именуемые «Стороны», а по отдельности — «Сторона», находясь в здравом уме и ясной памяти, действуя добровольно, заключили настоящий Предварительный договор (далее — «Договор») о нижеследующем:", s_after=2.0)

    # =========================================================================
    # РАЗДЕЛ 1. ПРЕДМЕТ ПРЕДВАРИТЕЛЬНОГО ДОГОВОРА
    # =========================================================================
    add_h(doc, "1. Предмет предварительного договора")

    p11 = doc.add_paragraph()
    set_p_format(p11, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    r11 = p11.add_run("1.1. ")
    r11.bold = True
    r11.font.name = "Times New Roman"
    p11.add_run("Стороны обязуются заключить в будущем Основной договор купли-продажи квартиры (далее — «Основной договор») на согласованных в настоящем Договоре условиях в срок не позднее ").font.name = "Times New Roman"
    r_dt = p11.add_run("«15» апреля 2026 года")
    r_dt.bold = True
    r_dt.font.name = "Times New Roman"
    p11.add_run(" включительно.").font.name = "Times New Roman"

    add_p(doc, "По Основному договору Продавцы обязуются передать в собственность Покупателя, а Покупатель обязуется принять и оплатить объект недвижимого имущества — 1-комнатную квартиру, расположенную на 4 этаже жилого дома по адресу: Пензенская область, г. Пенза, ул. Тернопольская, д. 7, кв. 34, кадастровый номер: 58:29:1005008:1234, общей площадью 38,5 (тридцать восемь целых пять десятых) кв. м, жилой площадью 19,2 кв. м (далее — «Объект»).", bold_prefix="1.2. ")

    add_p(doc, "Указанный Объект принадлежит Продавцам на праве общей долевой собственности на основании следующих правоустанавливающих документов:", bold_prefix="1.3. ")

    # Доля Олега
    p_oleg = doc.add_paragraph()
    set_p_format(p_oleg, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p_oleg.add_run("— ").font.name = "Times New Roman"
    r_fio1 = p_oleg.add_run("Протасову Олегу Ивановичу")
    r_fio1.bold = True
    r_fio1.font.name = "Times New Roman"
    p_oleg.add_run(" принадлежит ").font.name = "Times New Roman"
    r_sh1 = p_oleg.add_run("1/3 (одна третья)")
    r_sh1.bold = True
    r_sh1.font.name = "Times New Roman"
    p_oleg.add_run(" доля в праве на основании Свидетельства о праве на наследство по закону от 14.11.2024 года, выданного нотариусом г. Пензы Ивановой М.А., реестровый номер: 58/12-н/58-2024-4-567, о чем в Едином государственном реестре недвижимости сделана запись регистрации № 58:29:1005008:1234-58/001/2024-3 от 22.11.2024 года;").font.name = "Times New Roman"

    # Доля Виктора
    p_vik = doc.add_paragraph()
    set_p_format(p_vik, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p_vik.add_run("— ").font.name = "Times New Roman"
    r_fio2 = p_vik.add_run("Протасову Виктору Ивановичу")
    r_fio2.bold = True
    r_fio2.font.name = "Times New Roman"
    p_vik.add_run(" принадлежит ").font.name = "Times New Roman"
    r_sh2 = p_vik.add_run("2/3 (две третьих)")
    r_sh2.bold = True
    r_sh2.font.name = "Times New Roman"
    p_vik.add_run(" доли в праве, в том числе: 1/3 доля на основании Договора на передачу квартир в собственность граждан (приватизации) № 4512 от 12.03.1998 г., зарегистрированного БТИ г. Пензы 18.03.1998 г. за № 3-4512, о чем в ЕГРН сделана запись № 58:29:1005008:1234-58/001/2024-1 от 18.04.2024 г., и 1/3 доля на основании Свидетельства о праве на наследство по закону от 14.11.2024 г., выданного нотариусом г. Пензы Ивановой М.А., реестровый номер: 58/12-н/58-2024-4-568, о чем в ЕГРН сделана запись № 58:29:1005008:1234-58/001/2024-2 от 22.11.2024 г.").font.name = "Times New Roman"

    # Пункт 1.4: Обременение, снятие ареста приставов, гарантия чистоты (структурировано без стенописи)
    p14_1 = doc.add_paragraph()
    set_p_format(p14_1, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    r14_b = p14_1.add_run("1.4. ")
    r14_b.bold = True
    r14_b.font.name = "Times New Roman"
    p14_1.add_run("По данным выписки из Единого государственного реестра недвижимости от ").font.name = "Times New Roman"
    r_egrn = p14_1.add_run("15.02.2026 года")
    r_egrn.bold = True
    r_egrn.font.name = "Times New Roman"
    p14_1.add_run(" в отношении 1/3 доли Протасова Олега Ивановича зарегистрировано ограничение прав и обременение объекта недвижимости: запрет на совершение действий по регистрации на основании постановления судебного пристава-исполнителя Первомайского РОСП г. Пензы от 10.01.2026 года по исполнительному производству № 12345/26/58043-ИП.").font.name = "Times New Roman"

    # Обязательство Олега по снятию ареста
    p14_2 = doc.add_paragraph()
    set_p_format(p14_2, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p14_2.add_run("Продавец ").font.name = "Times New Roman"
    r_ol_ob = p14_2.add_run("Протасов Олег Иванович")
    r_ol_ob.bold = True
    r_ol_ob.font.name = "Times New Roman"
    p14_2.add_run(" обязуется за счет собственных денежных средств погасить задолженность по указанному исполнительному производству в полном объеме и обеспечить государственную регистрацию снятия запрета на совершение регистрационных действий в ЕГРН в срок ").font.name = "Times New Roman"
    r_term = p14_2.add_run("не позднее 5 (пяти) рабочих дней")
    r_term.bold = True
    r_term.font.name = "Times New Roman"
    p14_2.add_run(" до даты заключения Основного договора купли-продажи.").font.name = "Times New Roman"

    # Гарантия отсутствия иных обременений
    p14_3 = doc.add_paragraph()
    set_p_format(p14_3, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p14_3.add_run("Продавцы гарантируют, что иных ограничений прав и обременений в отношении Объекта (включая ипотеку, залог, аренду, наем, права третьих лиц) не зарегистрировано, Объект в споре и под арестом не состоит.").font.name = "Times New Roman"

    # Пункт 1.5: В пользу третьего лица (ст. 430 ГК РФ)
    add_p(doc, "Основной договор купли-продажи заключается в пользу третьего лица (ст. 430 ГК РФ) — совершеннолетнего гражданина РФ, назначенного Покупателем, либо непосредственно с Покупателем. Покупатель уведомляет Продавцов о данных лица, в пользу которого оформляется Объект, не позднее чем за 3 (три) рабочих дня до дня заключения Основного договора.", bold_prefix="1.5. ")

    # =========================================================================
    # РАЗДЕЛ 2. ЦЕНА ОБЪЕКТА И ПОРЯДОК РАСЧЕТОВ
    # =========================================================================
    add_h(doc, "2. Цена Объекта и порядок расчетов")

    # Пункт 2.1: Цена объекта
    p21 = doc.add_paragraph()
    set_p_format(p21, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    r21_b = p21.add_run("2.1. ")
    r21_b.bold = True
    r21_b.font.name = "Times New Roman"
    p21.add_run("Цена Объекта недвижимости по соглашению Сторон составляет ").font.name = "Times New Roman"
    r_price = p21.add_run("7 000 000 (Семь миллионов) рублей 00 копеек")
    r_price.bold = True
    r_price.font.name = "Times New Roman"
    p21.add_run(".").font.name = "Times New Roman"

    # Пункт 2.2: Вводная строка траншей
    add_p(doc, "Оплата цены Объекта производится Покупателем в следующем порядке и в следующие согласованные Сторонами сроки:", bold_prefix="2.2. ")

    # Задаток
    p_zadatok = doc.add_paragraph()
    set_p_format(p_zadatok, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p_zadatok.add_run("— ").font.name = "Times New Roman"
    r_sum1 = p_zadatok.add_run("100 000 (Сто тысяч) рублей 00 копеек")
    r_sum1.bold = True
    r_sum1.font.name = "Times New Roman"
    p_zadatok.add_run(" выплачиваются Покупателем Продавцам в день подписания настоящего Договора в качестве единого задатка в счет причитающихся платежей по покупке Объекта наличными денежными средствами, о чем составляется Расписка (Приложение № 1 к настоящему Договору). В соответствии со ст. 312 ГК РФ денежные средства задатка передаются Продавцу ").font.name = "Times New Roman"
    r_v_zad = p_zadatok.add_run("Протасову Виктору Ивановичу")
    r_v_zad.bold = True
    r_v_zad.font.name = "Times New Roman"
    p_zadatok.add_run(", действующему от своего имени, а также с прямого согласия и по поручению Продавца ").font.name = "Times New Roman"
    r_o_zad = p_zadatok.add_run("Протасова Олега Ивановича")
    r_o_zad.bold = True
    r_o_zad.font.name = "Times New Roman"
    p_zadatok.add_run(", что признается надлежащим исполнением обязательства Покупателя перед ").font.name = "Times New Roman"
    r_ob = p_zadatok.add_run("обоими")
    r_ob.bold = True
    r_ob.font.name = "Times New Roman"
    p_zadatok.add_run(" Продавцами в полном объеме и влечет их солидарную ответственность по ст. 322, 381 ГК РФ;").font.name = "Times New Roman"

    # СБР
    p_sbr = doc.add_paragraph()
    set_p_format(p_sbr, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p_sbr.add_run("— ").font.name = "Times New Roman"
    r_sum2 = p_sbr.add_run("5 000 000 (Пять миллионов) рублей 00 копеек")
    r_sum2.bold = True
    r_sum2.font.name = "Times New Roman"
    p_sbr.add_run(" уплачиваются Покупателем за счет собственных денежных средств с использованием Сервиса безопасных расчетов (СБР) либо безотзывного покрытого аккредитива ПАО Сбербанк. Перечисление денежных средств Продавцам осуществляется в срок ").font.name = "Times New Roman"
    r_term_sbr = p_sbr.add_run("от 1 до 5 рабочих дней")
    r_term_sbr.bold = True
    r_term_sbr.font.name = "Times New Roman"
    p_sbr.add_run(" с момента государственной регистрации перехода права собственности на Объект к Покупателю в Едином государственном реестре недвижимости;").font.name = "Times New Roman"

    # Ипотека
    p_ipoteka = doc.add_paragraph()
    set_p_format(p_ipoteka, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p_ipoteka.add_run("— ").font.name = "Times New Roman"
    r_sum3 = p_ipoteka.add_run("1 000 000 (Один миллион) рублей 00 копеек")
    r_sum3.bold = True
    r_sum3.font.name = "Times New Roman"
    p_ipoteka.add_run(" уплачиваются Покупателем за счет кредитных средств, предоставляемых Банком-кредитором (ПАО Сбербанк) по кредитному договору. Перечисление кредитных средств Продавцам осуществляется в срок ").font.name = "Times New Roman"
    r_term_ip = p_ipoteka.add_run("от 1 до 5 рабочих дней")
    r_term_ip.bold = True
    r_term_ip.font.name = "Times New Roman"
    p_ipoteka.add_run(" после государственной регистрации перехода права собственности на Объект и ипотеки в силу закона в пользу Банка-кредитора в Едином государственном реестре недвижимости;").font.name = "Times New Roman"

    # Маткапитал
    p_msk = doc.add_paragraph()
    set_p_format(p_msk, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p_msk.add_run("— ").font.name = "Times New Roman"
    r_sum4 = p_msk.add_run("900 000 (Девятьсот тысяч) рублей 00 копеек")
    r_sum4.bold = True
    r_sum4.font.name = "Times New Roman"
    p_msk.add_run(" уплачиваются за счет средств материнского (семейного) капитала путем безналичного перечисления из Фонда пенсионного и социального страхования РФ (СФР) на банковский счет Продавцов в срок ").font.name = "Times New Roman"
    r_term_msk = p_msk.add_run("не позднее 21 (двадцати одного) календарного дня")
    r_term_msk.bold = True
    r_term_msk.font.name = "Times New Roman"
    p_msk.add_run(" со дня подачи заявления о распоряжении средствами МСК. В случае использования иных государственных жилищных сертификатов перечисление средств осуществляется в срок ").font.name = "Times New Roman"
    r_term_cert = p_msk.add_run("до 30 (тридцати) календарных дней")
    r_term_cert.bold = True
    r_term_cert.font.name = "Times New Roman"
    p_msk.add_run(" с момента подачи заявления на перечисление.").font.name = "Times New Roman"

    # Пункт 2.3: Распределение средств между сособственниками
    p23 = doc.add_paragraph()
    set_p_format(p23, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    r23_b = p23.add_run("2.3. ")
    r23_b.bold = True
    r23_b.font.name = "Times New Roman"
    p23.add_run("Денежные средства, подлежащие перечислению Продавцам по Основному договору (за вычетом задатка), распределяются между Продавцами пропорционально их долям в праве общей долевой собственности: Продавцу ").font.name = "Times New Roman"
    r_ol_fio = p23.add_run("Протасову Олегу Ивановичу")
    r_ol_fio.bold = True
    r_ol_fio.font.name = "Times New Roman"
    p23.add_run(" (1/3 доля) — ").font.name = "Times New Roman"
    r_ol_s = p23.add_run("2 300 000 (Два миллиона триста тысяч) рублей 00 копеек")
    r_ol_s.bold = True
    r_ol_s.font.name = "Times New Roman"
    p23.add_run(", Продавцу ").font.name = "Times New Roman"
    r_vk_fio = p23.add_run("Протасову Виктору Ивановичу")
    r_vk_fio.bold = True
    r_vk_fio.font.name = "Times New Roman"
    p23.add_run(" (2/3 доли) — ").font.name = "Times New Roman"
    r_vk_s = p23.add_run("4 600 000 (Четыре миллиона шестьсот тысяч) рублей 00 копеек")
    r_vk_s.bold = True
    r_vk_s.font.name = "Times New Roman"
    p23.add_run(".").font.name = "Times New Roman"

    # =========================================================================
    # РАЗДЕЛ 3. СРОКИ И УСЛОВИЯ ПЕРЕДАЧИ ОБЪЕКТА
    # =========================================================================
    add_h(doc, "3. Сроки и условия передачи Объекта")

    add_p(doc, "Передача Объекта Продавцами и принятие его Покупателем оформляется Сторонами путем составления и подписания Передаточного акта в течение 3 (трех) рабочих дней со дня поступления денежных средств на банковские счета Продавцов в полном объеме.", bold_prefix="3.1. ")

    # Пункт 3.2: Броня приватизации (выписка СТРОГО ДО заключения ОДКП + адресная справка МВД)
    p32 = doc.add_paragraph()
    set_p_format(p32, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    r32_b = p32.add_run("3.2. ")
    r32_b.bold = True
    r32_b.font.name = "Times New Roman"
    p32.add_run("В Объекте зарегистрирован по месту жительства Продавец Протасов Виктор Иванович. С учетом возникновения права собственности на основании договора передачи (приватизации) жилого помещения Продавец ").font.name = "Times New Roman"
    r_v_reg = p32.add_run("Протасов Виктор Иванович")
    r_v_reg.bold = True
    r_v_reg.font.name = "Times New Roman"
    p32.add_run(" обязуется сняться с регистрационного учета по месту жительства из Объекта ").font.name = "Times New Roman"
    r_v_term = p32.add_run("СТРОГО ДО заключения Основного договора купли-продажи")
    r_v_term.bold = True
    r_v_term.font.name = "Times New Roman"
    p32.add_run(" и в день подписания Основного договора предоставить Покупателю ").font.name = "Times New Roman"
    r_spravka = p32.add_run("адресную справку (справку органа регистрационного учета МВД России) об отсутствии зарегистрированных лиц в Объекте")
    r_spravka.bold = True
    r_spravka.font.name = "Times New Roman"
    p32.add_run(". Иных лиц, зарегистрированных либо сохраняющих право пользования Объектом на момент подписания настоящего Договора, не имеется.").font.name = "Times New Roman"

    add_p(doc, "Продавцы обязуются передать Объект в пригодном для постоянного проживания состоянии, укомплектованным исправным санитарно-техническим, электротехническим и газовым оборудованием, свободным от прав третьих лиц, долгов по коммунальным платежам и капитальному ремонту.", bold_prefix="3.3. ")

    # =========================================================================
    # РАЗДЕЛ 4. ОТВЕТСТВЕННОСТЬ СТОРОН И СОГЛАШЕНИЕ О ЗАДАТКЕ
    # =========================================================================
    add_h(doc, "4. Ответственность Сторон и соглашение о задатке")

    add_p(doc, "Выданная Покупателем сумма в размере 100 000 (Сто тысяч) рублей 00 копеек является задатком и передается в доказательство заключения настоящего Договора и в обеспечение исполнения всех обязательств по заключению Основного договора купли-продажи (ст. 380, 381, 429 ГК РФ).", bold_prefix="4.1. ")

    add_p(doc, "В случае если за неисполнение настоящего Договора либо незаключение Основного договора в установленный срок ответственен Покупатель (включая немотивированный отказ от покупки либо непредоставление документов), сумма задатка остается у Продавцов в полном объеме.", bold_prefix="4.2. ")

    # Пункт 4.3: Возврат двойного задатка Продавцами (жирным)
    p43 = doc.add_paragraph()
    set_p_format(p43, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    r43_b = p43.add_run("4.3. ")
    r43_b.bold = True
    r43_b.font.name = "Times New Roman"
    p43.add_run("В случае если за неисполнение настоящего Договора либо незаключение Основного договора ответственны Продавцы (включая немотивированный отказ от продажи, неснятие запрета на регистрационные действия в ЕГРН Продавцом Протасовым О.И. либо неснятие Протасова В.И. с регистрационного учета до заключения Основного договора), Продавцы обязаны ").font.name = "Times New Roman"
    r_sol = p43.add_run("солидарно уплатить Покупателю двойную сумму задатка")
    r_sol.bold = True
    r_sol.font.name = "Times New Roman"
    p43.add_run(" в размере ").font.name = "Times New Roman"
    r_2sum = p43.add_run("200 000 (Двести тысяч) рублей 00 копеек")
    r_2sum.bold = True
    r_2sum.font.name = "Times New Roman"
    p43.add_run(" в течение 3 (трех) рабочих дней со дня истечения срока, установленного п. 1.1 настоящего Договора.").font.name = "Times New Roman"

    add_p(doc, "В случае если Основной договор не может быть заключен вследствие объективных обстоятельств, за которые ни одна из Сторон не отвечает (включая обоснованный отказ Банка-кредитора в согласовании Объекта либо непреодолимую силу), сумма задатка возвращается Покупателю в одинарном размере (100 000 рублей 00 копеек) в течение 3 (трех) рабочих дней со дня наступления таких обстоятельств без применения штрафных санкций.", bold_prefix="4.4. ")

    add_p(doc, "При уклонении одной из Сторон от заключения Основного договора другая Сторона вправе обратиться в суд с требованием о понуждении заключить договор на согласованных условиях в соответствии со ст. 429, 445 ГК РФ и о возмещении причиненных убытков.", bold_prefix="4.5. ")

    # =========================================================================
    # РАЗДЕЛ 5. ОБЯЗАННОСТИ СТОРОН ПО ПОДГОТОВКЕ СДЕЛКИ
    # =========================================================================
    add_h(doc, "5. Обязанности Сторон по подготовке сделки")

    add_p(doc, "В целях заключения Основного договора и согласования сделки в Банке-кредиторе Продавцы обязуются предоставить Покупателю и Банку-кредитору в срок не позднее 10 (десяти) календарных дней до даты заключения Основного договора следующие документы:", bold_prefix="5.1. ")

    # Документы Олега
    p_doc_ol = doc.add_paragraph()
    set_p_format(p_doc_ol, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p_doc_ol.add_run("— ").font.name = "Times New Roman"
    r_dol = p_doc_ol.add_run("Протасов Олег Иванович")
    r_dol.bold = True
    r_dol.font.name = "Times New Roman"
    p_doc_ol.add_run(" — оригинал паспорта гражданина РФ, оригинал Свидетельства о праве на наследство по закону, страховое свидетельство (СНИЛС), свидетельство ИНН, адрес электронной почты, справку из психоневрологического диспансера об отсутствии диспансерного учета, справку из наркологического диспансера об отсутствии диспансерного учета, а также постановление судебного пристава-исполнителя и выписку из ЕГРН, подтверждающие полное снятие запрета на совершение регистрационных действий;").font.name = "Times New Roman"

    # Документы Виктора
    p_doc_vk = doc.add_paragraph()
    set_p_format(p_doc_vk, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p_doc_vk.add_run("— ").font.name = "Times New Roman"
    r_dvk = p_doc_vk.add_run("Протасов Виктор Иванович")
    r_dvk.bold = True
    r_dvk.font.name = "Times New Roman"
    p_doc_vk.add_run(" — оригинал паспорта гражданина РФ, оригинал Договора передачи квартиры в собственность граждан (приватизации), оригинал Свидетельства о праве на наследство по закону, страховое свидетельство (СНИЛС), свидетельство ИНН, адрес электронной почты, справку из психоневрологического диспансера об отсутствии диспансерного учета, справку из наркологического диспансера об отсутствии диспансерного учета, а также адресную справку (справку о зарегистрированных лицах) органа регистрационного учета МВД России, подтверждающую снятие с регистрационного учета;").font.name = "Times New Roman"

    # Общие документы
    p_doc_all = doc.add_paragraph()
    set_p_format(p_doc_all, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    p_doc_all.add_run("— ").font.name = "Times New Roman"
    r_dall = p_doc_all.add_run("Продавцы совместно")
    r_dall.bold = True
    r_dall.font.name = "Times New Roman"
    p_doc_all.add_run(" — справки об отсутствии задолженности по оплате коммунальных услуг, взносов на капитальный ремонт и энергоресурсов, технический паспорт / поэтажный план Объекта (по требованию Банка).").font.name = "Times New Roman"

    # Документы Покупателя
    p_doc_buy = doc.add_paragraph()
    set_p_format(p_doc_buy, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    r52_b = p_doc_buy.add_run("5.2. ")
    r52_b.bold = True
    r52_b.font.name = "Times New Roman"
    r_buy_fio = p_doc_buy.add_run("Покупатель")
    r_buy_fio.bold = True
    r_buy_fio.font.name = "Times New Roman"
    p_doc_buy.add_run(" обязуется предоставить: оригинал паспорта гражданина РФ, страховое свидетельство (СНИЛС), свидетельство ИНН, адрес электронной почты, свидетельства о рождении детей, документы о регистрации детей по месту жительства (форма № 8), государственный сертификат на материнский (семейный) капитал и справку СФР об остатке средств материнского капитала.").font.name = "Times New Roman"

    # =========================================================================
    # РАЗДЕЛ 6. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ И ПОДПИСИ СТОРОН
    # =========================================================================
    add_h(doc, "6. Заключительные положения")

    add_p(doc, "Настоящий Договор вступает в силу с момента его подписания всеми Сторонами и действует до полного исполнения Сторонами всех принятых на себя обязательств.", bold_prefix="6.1. ")

    add_p(doc, "Все изменения и дополнения к настоящему Договору действительны лишь в том случае, если они совершены в письменной форме и подписаны уполномоченными представителями Сторон.", bold_prefix="6.2. ")

    add_p(doc, "Настоящий Договор составлен и подписан в 3 (трех) подлинных экземплярах, имеющих равную юридическую силу: по одному экземпляру — для каждого из Продавцов, и один экземпляр — для Покупателя.", bold_prefix="6.3. ", s_after=2.0)

    # Заголовок блока подписей договора
    p_sig_title = doc.add_paragraph()
    set_p_format(p_sig_title, f_size=font_size, space_before=3.0, s_after=1.5, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    p_sig_title.paragraph_format.keep_with_next = True
    r_st = p_sig_title.add_run("Подписи Сторон:")
    r_st.bold = True
    r_st.font.name = "Times New Roman"

    # Подписант 1: Протасов Олег Иванович
    add_fullwidth_signature(doc, "Продавец 1:", "Протасов Олег Иванович", underline_space_before=0, underline_space_after=1.0)

    # Подписант 2: Протасов Виктор Иванович
    add_fullwidth_signature(doc, "Продавец 2:", "Протасов Виктор Иванович", underline_space_before=0, underline_space_after=1.0)

    # Подписант 3: Фирсова Ангелина Сергеевна
    add_fullwidth_signature(doc, "Покупатель:", "Фирсова Ангелина Сергеевна", underline_space_before=0, underline_space_after=1.0)

    # =========================================================================
    # ПРИЛОЖЕНИЕ № 1: РАСПИСКА В ПОЛУЧЕНИИ ЗАДАТКА (СТРОГО СТР. 5)
    # =========================================================================
    p_app_title = doc.add_paragraph()
    set_p_format(p_app_title, f_size=10.0, space_before=0, s_after=1.5, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    p_app_title.paragraph_format.page_break_before = True  # Ровно на 5-й лист
    r_apt = p_app_title.add_run("ПРИЛОЖЕНИЕ № 1 К ПРЕДВАРИТЕЛЬНОМУ ДОГОВОРУ КУПЛИ-ПРОДАЖИ КВАРТИРЫ")
    r_apt.bold = True
    r_apt.font.name = "Times New Roman"

    p_rec_h = doc.add_paragraph()
    set_p_format(p_rec_h, f_size=11.0, space_before=1.5, s_after=3.0, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    r_rh = p_rec_h.add_run("РАСПИСКА В ПОЛУЧЕНИИ И ПЕРЕДАЧЕ ЗАДАТКА")
    r_rh.bold = True
    r_rh.font.name = "Times New Roman"

    # Шапка расписки (Параграф с табулятором по правому полю)
    add_city_date_tab(doc, "г. Пенза", deal_date)

    # Преамбула расписки
    add_p(doc, "Мы, нижеподписавшиеся Стороны: Продавцы гр. Протасов Олег Иванович, 15.04.1988 года рождения, паспорт гражданина РФ: серия 5612 № 345678, и гр. Протасов Виктор Иванович, 22.09.1963 года рождения, паспорт гражданина РФ: серия 5608 № 901234, с одной стороны, и Покупатель гр. Фирсова Ангелина Сергеевна, 10.08.1995 года рождения, паспорт гражданина РФ: серия 5615 № 123456, с другой стороны, составили и подписали настоящую расписку о нижеследующем:", s_after=space_after)

    # Пункт 1 расписки: фиксация получения задатка Протасовым В.И.
    p_r1 = doc.add_paragraph()
    set_p_format(p_r1, f_size=font_size, space_before=0, s_after=space_after, l_spacing=line_spacing, first_indent=INDENT)
    r_r1_b = p_r1.add_run("1. ")
    r_r1_b.bold = True
    r_r1_b.font.name = "Times New Roman"
    p_r1.add_run("В соответствии с п. 2.2 Предварительного договора купли-продажи квартиры от ").font.name = "Times New Roman"
    p_r1.add_run(deal_date).font.name = "Times New Roman"
    p_r1.add_run(" Покупатель Фирсова Ангелина Сергеевна передала, а Продавец Протасов Виктор Иванович, действуя от своего имени, а также с прямого согласия и по поручению Продавца Протасова Олега Ивановича, принял наличными денежными средствами в качестве единого ").font.name = "Times New Roman"
    r_rzad = p_r1.add_run("ЗАДАТКА")
    r_rzad.bold = True
    r_rzad.font.name = "Times New Roman"
    p_r1.add_run(" в счет обеспечения исполнения обязательств по продаже 1-комнатной квартиры по адресу: Пензенская область, г. Пенза, ул. Тернопольская, д. 7, кв. 34, денежную сумму в размере ").font.name = "Times New Roman"
    r_rsum = p_r1.add_run("100 000 (Сто тысяч) рублей 00 копеек")
    r_rsum.bold = True
    r_rsum.font.name = "Times New Roman"
    p_r1.add_run(".").font.name = "Times New Roman"

    # Пункт 2 расписки: правовые последствия, ст. 312, 322, 381 ГК РФ
    add_p(doc, "Денежные средства переданы Покупателем в полном объеме и приняты Продавцами лично наличными купюрами в момент подписания Предварительного договора. Продавцы подтверждают, что факт передачи задатка Протасову В.И. признается надлежащим исполнением обязательства Покупателя перед обоими Продавцами в полном объеме (ст. 312 ГК РФ), порождает солидарную ответственность Продавцов по ст. 322, 381 ГК РФ и засчитывается в счет общей стоимости Объекта. Взаимных финансовых, материальных либо иных претензий по сумме, срокам и порядку передачи указанной суммы задатка Стороны друг к другу не имеют.", bold_prefix="2. ")

    # Пункт 3 расписки: экземпляры
    add_p(doc, "Настоящая расписка составлена и подписана Сторонами в 2 (двух) подлинных экземплярах, имеющих равную юридическую силу: один экземпляр — стороне Продавцов, один экземпляр — Покупателю Фирсовой А.С., и является неотъемлемой частью Предварительного договора купли-продажи квартиры.", bold_prefix="3. ", s_after=2.0)

    # Подписи сторон расписки
    p_rsig_t = doc.add_paragraph()
    set_p_format(p_rsig_t, f_size=font_size, space_before=3.5, s_after=1.5, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    p_rsig_t.paragraph_format.keep_with_next = True
    r_rst = p_rsig_t.add_run("Подписи Сторон расписки:")
    r_rst.bold = True
    r_rst.font.name = "Times New Roman"

    # 1. Принял Виктор
    add_fullwidth_signature(doc, "Денежные средства по поручению обоих Продавцов принял (Продавец):", "Протасов Виктор Иванович", underline_space_before=0, underline_space_after=1.0)

    # 2. Одобрил Олег
    add_fullwidth_signature(doc, "Полномочия подтверждаю, согласие на получение задатка даю (Продавец):", "Протасов Олег Иванович", underline_space_before=0, underline_space_after=1.0)

    # 3. Передала Фирсова
    add_fullwidth_signature(doc, "Денежные средства передала (Покупатель):", "Фирсова Ангелина Сергеевна", underline_space_before=0, underline_space_after=1.0)

    # Сохранение документа
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print(f"ПДКП успешно сгенерирован и сохранен в: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_pdkp()
