import os
import sys
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_document_defaults(doc, font_name="Times New Roman", font_size_pt=10.0):
    """Устанавливает Times New Roman и базовый кегль как системный дефолт всего документа (docDefaults и стиль Normal)."""
    # 1. Стиль Normal
    style = doc.styles['Normal']
    style.font.name = font_name
    style.font.size = Pt(font_size_pt)

    # 2. docDefaults в styles.xml (чтобы любые спецсимволы, табы и пустые строки наследовали Times New Roman)
    rPrDefault = doc.styles.element.xpath('w:docDefaults/w:rPrDefault/w:rPr')
    if rPrDefault:
        sz_val = str(int(round(font_size_pt * 2)))
        rFonts = parse_xml(r'<w:rFonts %s w:ascii="%s" w:hAnsi="%s" w:cs="%s" w:eastAsia="%s"/>' % (nsdecls('w'), font_name, font_name, font_name, font_name))
        rPrDefault[0].append(rFonts)
        sz = parse_xml(r'<w:sz %s w:val="%s"/>' % (nsdecls('w'), sz_val))
        szCs = parse_xml(r'<w:szCs %s w:val="%s"/>' % (nsdecls('w'), sz_val))
        rPrDefault[0].append(sz)
        rPrDefault[0].append(szCs)
        lang = parse_xml(r'<w:lang %s w:val="ru-RU"/>' % nsdecls('w'))
        rPrDefault[0].append(lang)

def add_styled_run(p, text, font_size=10.0, bold=False, italic=False, font_name="Times New Roman"):
    """Добавляет run со СТРОГО фиксированным шрифтом, кеглем и языком для всех скриптов (ASCII, Cyrillic CS, EastAsia)."""
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.name = font_name
    r.font.size = Pt(font_size)

    # Жесткая фиксация в XML (защита от сброса на Calibri в Word)
    sz_val = str(int(round(font_size * 2)))
    rPr = r._r.get_or_add_rPr()

    for tag in ['w:rFonts', 'w:sz', 'w:szCs', 'w:lang']:
        for el in rPr.xpath(tag):
            rPr.remove(el)

    rFonts = parse_xml(r'<w:rFonts %s w:ascii="%s" w:hAnsi="%s" w:cs="%s" w:eastAsia="%s"/>' % (nsdecls('w'), font_name, font_name, font_name, font_name))
    sz = parse_xml(r'<w:sz %s w:val="%s"/>' % (nsdecls('w'), sz_val))
    szCs = parse_xml(r'<w:szCs %s w:val="%s"/>' % (nsdecls('w'), sz_val))
    lang = parse_xml(r'<w:lang %s w:val="ru-RU"/>' % nsdecls('w'))

    rPr.append(rFonts)
    rPr.append(sz)
    rPr.append(szCs)
    rPr.append(lang)
    return r

def generate_pdkp(
    font_size=10.0,
    line_spacing=1.06,
    space_after=1.2,
    deal_date="«15» февраля 2026 года",
    out_path=r"d:\Антигравити\legal_estate\deals\ПДКП_Протасовы_Фирсова.docx"
):
    doc = docx.Document()
    set_document_defaults(doc, font_name="Times New Roman", font_size_pt=font_size)

    s = doc.sections[0]
    s.top_margin = Inches(0.71)    # 18 mm (51.1 pt)
    s.bottom_margin = Inches(0.71) # 18 mm (51.1 pt)
    s.left_margin = Inches(0.79)   # 20 mm (56.9 pt)
    s.right_margin = Inches(0.59)  # 15 mm (42.5 pt)
    s.page_width = Inches(8.27)    # 210 mm
    s.page_height = Inches(11.69)  # 297 mm

    BODY_WIDTH = s.page_width - s.left_margin - s.right_margin
    INDENT = Pt(25.0)

    def set_p(p, space_before=0, s_after=space_after, l_spacing=line_spacing, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=INDENT):
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(s_after)
        p.paragraph_format.line_spacing = l_spacing
        p.paragraph_format.left_indent = Pt(0)
        p.paragraph_format.right_indent = Pt(0)
        p.alignment = align
        p.paragraph_format.first_line_indent = first_indent if first_indent is not None else Pt(0)

    def add_p(text="", bold_prefix="", s_after=space_after, space_before=0, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=INDENT):
        p = doc.add_paragraph()
        set_p(p, space_before=space_before, s_after=s_after, l_spacing=line_spacing, align=align, first_indent=first_indent)
        if bold_prefix:
            add_styled_run(p, bold_prefix, font_size=font_size, bold=True)
        if text:
            add_styled_run(p, text, font_size=font_size, bold=False)
        return p

    def add_h(text):
        p = doc.add_paragraph()
        set_p(p, space_before=3.0, s_after=1.5, l_spacing=line_spacing, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
        p.paragraph_format.keep_with_next = True
        add_styled_run(p, text, font_size=font_size, bold=True)
        return p

    def add_sig(role_title, printed_name):
        p_t = doc.add_paragraph()
        set_p(p_t, space_before=2.5, s_after=1.0, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
        p_t.paragraph_format.keep_with_next = True
        add_styled_run(p_t, role_title, font_size=font_size, bold=True)

        p_line = doc.add_paragraph()
        set_p(p_line, space_before=0, s_after=1.0, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
        p_line.paragraph_format.keep_with_next = True
        pPr = p_line._p.get_or_add_pPr()
        pBdr = parse_xml(r"""
            <w:pBdr %s >
                <w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>
            </w:pBdr>
        """ % nsdecls("w"))
        pPr.append(pBdr)

        p_sub = doc.add_paragraph()
        set_p(p_sub, space_before=0, s_after=1.0, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
        p_sub.paragraph_format.keep_with_next = True
        p_sub.paragraph_format.tab_stops.add_tab_stop(BODY_WIDTH, WD_TAB_ALIGNMENT.RIGHT)
        add_styled_run(p_sub, "ФИО полностью", font_size=8.0, italic=True)
        add_styled_run(p_sub, "\t", font_size=8.0)
        add_styled_run(p_sub, "подпись", font_size=8.0, italic=True)

        p_name = doc.add_paragraph()
        set_p(p_name, space_before=0, s_after=3.0, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
        p_name.paragraph_format.keep_with_next = True
        add_styled_run(p_name, printed_name, font_size=font_size, bold=False)
        return p_name

    # =========================================================================
    # ШАПКА И ПРЕАМБУЛА
    # =========================================================================
    p_title = doc.add_paragraph()
    set_p(p_title, space_before=0, s_after=2.0, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    add_styled_run(p_title, "ПРЕДВАРИТЕЛЬНЫЙ ДОГОВОР КУПЛИ-ПРОДАЖИ КВАРТИРЫ", font_size=12.0, bold=True)

    p_sub = doc.add_paragraph()
    set_p(p_sub, space_before=0, s_after=6.0, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    add_styled_run(p_sub, "(с условием обеспечения обязательства задатком)", font_size=font_size, italic=True)

    # Город и дата
    p_cd = doc.add_paragraph()
    set_p(p_cd, space_before=0, s_after=6.0, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
    p_cd.paragraph_format.tab_stops.add_tab_stop(BODY_WIDTH, WD_TAB_ALIGNMENT.RIGHT)
    add_styled_run(p_cd, "г. Пенза", font_size=font_size, bold=True)
    add_styled_run(p_cd, "\t", font_size=font_size)
    add_styled_run(p_cd, deal_date, font_size=font_size, bold=True)

    # Вводная преамбулы
    p_in = doc.add_paragraph()
    set_p(p_in, space_before=0, s_after=1.2, first_indent=INDENT)
    add_styled_run(p_in, "Мы, нижеподписавшиеся:", font_size=font_size)

    # Продавец 1
    p_pr1 = doc.add_paragraph()
    set_p(p_pr1, space_before=0, s_after=1.2, first_indent=INDENT)
    add_styled_run(p_pr1, "гр. ", font_size=font_size)
    add_styled_run(p_pr1, "Протасов Олег Иванович", font_size=font_size, bold=True)
    add_styled_run(p_pr1, ", 15.04.1988 года рождения, пол: мужской, место рождения: гор. Пенза, паспорт гражданина РФ: серия 5612 № 345678, выдан Отделом УФМС России по Пензенской области в Ленинском р-не г. Пензы 20.04.2012 г., код подразделения: 580-001, адрес регистрации: Пензенская область, г. Пенза, ул. Суворова, д. 145, кв. 12, и", font_size=font_size)

    # Продавец 2
    p_pr2 = doc.add_paragraph()
    set_p(p_pr2, space_before=0, s_after=1.2, first_indent=INDENT)
    add_styled_run(p_pr2, "гр. ", font_size=font_size)
    add_styled_run(p_pr2, "Протасов Виктор Иванович", font_size=font_size, bold=True)
    add_styled_run(p_pr2, ", 22.09.1963 года рождения, пол: мужской, место рождения: гор. Пенза, паспорт гражданина РФ: серия 5608 № 901234, выдан Отделом внутренних дел Октябрьского р-на г. Пензы 05.10.2008 г., код подразделения: 580-002, адрес регистрации: Пензенская область, г. Пенза, ул. Тернопольская, д. 7, кв. 34, именуемые в дальнейшем совместно ", font_size=font_size)
    add_styled_run(p_pr2, "«Продавцы»", font_size=font_size, bold=True)
    add_styled_run(p_pr2, ", с одной стороны, и", font_size=font_size)

    # Покупатель
    p_by = doc.add_paragraph()
    set_p(p_by, space_before=0, s_after=1.2, first_indent=INDENT)
    add_styled_run(p_by, "гр. ", font_size=font_size)
    add_styled_run(p_by, "Фирсова Ангелина Сергеевна", font_size=font_size, bold=True)
    add_styled_run(p_by, ", 10.08.1995 года рождения, пол: женский, место рождения: гор. Пенза, паспорт гражданина РФ: серия 5615 № 123456, выдан Отделом УФМС России по Пензенской области в Первомайском р-не г. Пензы 15.08.2015 г., код подразделения: 580-003, адрес регистрации: Пензенская область, г. Пенза, ул. Мира, д. 40, кв. 89, именуемая в дальнейшем ", font_size=font_size)
    add_styled_run(p_by, "«Покупатель»", font_size=font_size, bold=True)
    add_styled_run(p_by, ", с другой стороны,", font_size=font_size)

    add_p("при совместном упоминании именуемые «Стороны», а по отдельности — «Сторона», находясь в здравом уме и ясной памяти, действуя добровольно, заключили настоящий Предварительный договор (далее — «Договор») о нижеследующем:", s_after=2.0)

    # =========================================================================
    # РАЗДЕЛ 1. ПРЕДМЕТ ПРЕДВАРИТЕЛЬНОГО ДОГОВОРА
    # =========================================================================
    add_h("1. Предмет предварительного договора")

    p11 = doc.add_paragraph()
    set_p(p11, first_indent=INDENT)
    add_styled_run(p11, "1.1. ", font_size=font_size, bold=True)
    add_styled_run(p11, "Стороны обязуются заключить в будущем Основной договор купли-продажи квартиры (далее — «Основной договор») на согласованных в настоящем Договоре условиях в срок не позднее ", font_size=font_size)
    add_styled_run(p11, "«15» апреля 2026 года включительно", font_size=font_size, bold=True)
    add_styled_run(p11, ".", font_size=font_size)

    add_p("По Основному договору Продавцы обязуются передать в собственность Покупателя, а Покупатель обязуется принять и оплатить объект недвижимого имущества — 1-комнатную квартиру, расположенную на 4 этаже жилого дома по адресу: Пензенская область, г. Пенза, ул. Тернопольская, д. 7, кв. 34, кадастровый номер: 58:29:1005008:1234, общей площадью 38,5 (тридцать восемь целых пять десятых) кв. м, жилой площадью 19,2 кв. м (далее — «Объект»).", bold_prefix="1.2. ")

    add_p("Указанный Объект принадлежит Продавцам на праве общей долевой собственности на основании следующих правоустанавливающих документов:", bold_prefix="1.3. ")

    p_dol1 = doc.add_paragraph()
    set_p(p_dol1, first_indent=INDENT)
    add_styled_run(p_dol1, "— ", font_size=font_size)
    add_styled_run(p_dol1, "Протасову Олегу Ивановичу", font_size=font_size, bold=True)
    add_styled_run(p_dol1, " принадлежит ", font_size=font_size)
    add_styled_run(p_dol1, "1/3 (одна третья)", font_size=font_size, bold=True)
    add_styled_run(p_dol1, " доля в праве на основании Свидетельства о праве на наследство по закону от 14.11.2024 года, выданного нотариусом г. Пензы Ивановой М.А., реестровый номер: 58/12-н/58-2024-4-567, о чем в Едином государственном реестре недвижимости сделана запись регистрации № 58:29:1005008:1234-58/001/2024-3 от 22.11.2024 года;", font_size=font_size)

    p_dol2 = doc.add_paragraph()
    set_p(p_dol2, first_indent=INDENT)
    add_styled_run(p_dol2, "— ", font_size=font_size)
    add_styled_run(p_dol2, "Протасову Виктору Ивановичу", font_size=font_size, bold=True)
    add_styled_run(p_dol2, " принадлежит ", font_size=font_size)
    add_styled_run(p_dol2, "2/3 (две третьих)", font_size=font_size, bold=True)
    add_styled_run(p_dol2, " доли в праве, в том числе: 1/3 доля на основании Договора на передачу квартир в собственность граждан (приватизации) № 4512 от 12.03.1998 г., зарегистрированного БТИ г. Пензы 18.03.1998 г. за № 3-4512, о чем в ЕГРН сделана запись № 58:29:1005008:1234-58/001/2024-1 от 18.04.2024 г., и 1/3 доля на основании Свидетельства о праве на наследство по закону от 14.11.2024 г., выданного нотариусом г. Пензы Ивановой М.А., реестровый номер: 58/12-н/58-2024-4-568, о чем в ЕГРН сделана запись № 58:29:1005008:1234-58/001/2024-2 от 22.11.2024 г.", font_size=font_size)

    # 1.4: без жирного на сроке приставов и дате выписки
    p14_1 = doc.add_paragraph()
    set_p(p14_1, first_indent=INDENT)
    add_styled_run(p14_1, "1.4. ", font_size=font_size, bold=True)
    add_styled_run(p14_1, "По данным выписки из Единого государственного реестра недвижимости от 15.02.2026 года в отношении 1/3 доли Протасова Олега Ивановича зарегистрировано ограничение прав и обременение объекта недвижимости: запрет на совершение действий по регистрации на основании постановления судебного пристава-исполнителя Первомайского РОСП г. Пензы от 10.01.2026 года по исполнительному производству № 12345/26/58043-ИП.", font_size=font_size)

    p14_2 = doc.add_paragraph()
    set_p(p14_2, first_indent=INDENT)
    add_styled_run(p14_2, "Продавец ", font_size=font_size)
    add_styled_run(p14_2, "Протасов Олег Иванович", font_size=font_size, bold=True)
    add_styled_run(p14_2, " обязуется за счет собственных денежных средств погасить задолженность по указанному исполнительному производству в полном объеме и обеспечить государственную регистрацию снятия запрета на совершение регистрационных действий в ЕГРН в срок не позднее 5 (пяти) рабочих дней до даты заключения Основного договора купли-продажи.", font_size=font_size)

    add_p("Продавцы гарантируют, что иных ограничений прав и обременений в отношении Объекта (включая ипотеку, залог, аренду, наем, права третьих лиц) не зарегистрировано, Объект в споре и под арестом не состоит.")

    add_p("Основной договор купли-продажи заключается в пользу третьего лица (ст. 430 ГК РФ) — совершеннолетнего гражданина РФ, назначенного Покупателем, либо непосредственно с Покупателем. Покупатель уведомляет Продавцов о данных лица, в пользу которого оформляется Объект, не позднее чем за 3 (три) рабочих дня до дня заключения Основного договора.", bold_prefix="1.5. ")

    # =========================================================================
    # РАЗДЕЛ 2. ЦЕНА ОБЪЕКТА И ПОРЯДОК РАСЧЕТОВ
    # =========================================================================
    add_h("2. Цена Объекта и порядок расчетов")

    p21 = doc.add_paragraph()
    set_p(p21, first_indent=INDENT)
    add_styled_run(p21, "2.1. ", font_size=font_size, bold=True)
    add_styled_run(p21, "Цена Объекта недвижимости по соглашению Сторон составляет ", font_size=font_size)
    add_styled_run(p21, "7 000 000 (Семь миллионов) рублей 00 копеек", font_size=font_size, bold=True)
    add_styled_run(p21, ".", font_size=font_size)

    add_p("Оплата цены Объекта производится Покупателем в следующем порядке и в следующие согласованные Сторонами сроки:", bold_prefix="2.2. ")

    p_zad = doc.add_paragraph()
    set_p(p_zad, first_indent=INDENT)
    add_styled_run(p_zad, "— ", font_size=font_size)
    add_styled_run(p_zad, "100 000 (Сто тысяч) рублей 00 копеек", font_size=font_size, bold=True)
    add_styled_run(p_zad, " выплачиваются Покупателем Продавцам в день подписания настоящего Договора в качестве задатка в счет причитающихся платежей по покупке Объекта наличными денежными средствами, о чем составляется Расписка (Приложение № 1 к настоящему Договору). В соответствии со ст. 312 ГК РФ денежные средства задатка передаются Продавцу ", font_size=font_size)
    add_styled_run(p_zad, "Протасову Виктору Ивановичу", font_size=font_size, bold=True)
    add_styled_run(p_zad, ", действующему от своего имени, а также с прямого согласия и по поручению Продавца ", font_size=font_size)
    add_styled_run(p_zad, "Протасова Олега Ивановича", font_size=font_size, bold=True)
    add_styled_run(p_zad, ", что признается надлежащим исполнением обязательства Покупателя перед ", font_size=font_size)
    add_styled_run(p_zad, "обоими", font_size=font_size, bold=True)
    add_styled_run(p_zad, " Продавцами в полном объеме и влечет их солидарную ответственность по ст. 322, 381 ГК РФ;", font_size=font_size)

    p_sbr = doc.add_paragraph()
    set_p(p_sbr, first_indent=INDENT)
    add_styled_run(p_sbr, "— ", font_size=font_size)
    add_styled_run(p_sbr, "5 000 000 (Пять миллионов) рублей 00 копеек", font_size=font_size, bold=True)
    add_styled_run(p_sbr, " уплачиваются Покупателем за счет собственных денежных средств с использованием Сервиса безопасных расчетов (СБР) либо безотзывного покрытого аккредитива ПАО Сбербанк. Перечисление денежных средств Продавцам осуществляется в срок от 1 до 5 рабочих дней с момента государственной регистрации перехода права собственности на Объект к Покупателю в Едином государственном реестре недвижимости;", font_size=font_size)

    p_ip = doc.add_paragraph()
    set_p(p_ip, first_indent=INDENT)
    add_styled_run(p_ip, "— ", font_size=font_size)
    add_styled_run(p_ip, "1 000 000 (Один миллион) рублей 00 копеек", font_size=font_size, bold=True)
    add_styled_run(p_ip, " уплачиваются Покупателем за счет кредитных средств, предоставляемых Банком-кредитором (ПАО Сбербанк) по кредитному договору. Перечисление кредитных средств Продавцам осуществляется в срок от 1 до 5 рабочих дней после государственной регистрации перехода права собственности на Объект и ипотеки в силу закона в пользу Банка-кредитора в Едином государственном реестре недвижимости;", font_size=font_size)

    p_msk = doc.add_paragraph()
    set_p(p_msk, first_indent=INDENT)
    add_styled_run(p_msk, "— ", font_size=font_size)
    add_styled_run(p_msk, "900 000 (Девятьсот тысяч) рублей 00 копеек", font_size=font_size, bold=True)
    add_styled_run(p_msk, " уплачиваются за счет средств материнского (семейного) капитала путем безналичного перечисления из Фонда пенсионного и социального страхования РФ (СФР) на банковский счет Продавцов в срок не позднее 21 (двадцати одного) календарного дня со дня подачи заявления о распоряжении средствами МСК. В случае использования иных государственных жилищных сертификатов перечисление средств осуществляется в срок до 30 (тридцати) календарных дней с момента подачи заявления на перечисление.", font_size=font_size)

    p23 = doc.add_paragraph()
    set_p(p23, first_indent=INDENT)
    add_styled_run(p23, "2.3. ", font_size=font_size, bold=True)
    add_styled_run(p23, "Денежные средства, подлежащие перечислению Продавцам по Основному договору (за вычетом задатка), распределяются между Продавцами пропорционально их долям в праве общей долевой собственности: Продавцу ", font_size=font_size)
    add_styled_run(p23, "Протасову Олегу Ивановичу", font_size=font_size, bold=True)
    add_styled_run(p23, " (1/3 доля) — ", font_size=font_size)
    add_styled_run(p23, "2 300 000 (Два миллиона триста тысяч) рублей 00 копеек", font_size=font_size, bold=True)
    add_styled_run(p23, ", Продавцу ", font_size=font_size)
    add_styled_run(p23, "Протасову Виктору Ивановичу", font_size=font_size, bold=True)
    add_styled_run(p23, " (2/3 доли) — ", font_size=font_size)
    add_styled_run(p23, "4 600 000 (Четыре миллиона шестьсот тысяч) рублей 00 копеек", font_size=font_size, bold=True)
    add_styled_run(p23, ".", font_size=font_size)

    # =========================================================================
    # РАЗДЕЛ 3. СРОКИ И УСЛОВИЯ ПЕРЕДАЧИ ОБЪЕКТА
    # =========================================================================
    add_h("3. Сроки и условия передачи Объекта")

    p31 = doc.add_paragraph()
    set_p(p31, first_indent=INDENT)
    add_styled_run(p31, "3.1. ", font_size=font_size, bold=True)
    add_styled_run(p31, "Передача Объекта Продавцами и принятие его Покупателем оформляется Сторонами путем составления и подписания Передаточного акта ", font_size=font_size)
    add_styled_run(p31, "в течение 3 (трех) рабочих дней", font_size=font_size, bold=True)
    add_styled_run(p31, " со дня поступления денежных средств на банковские счета Продавцов в полном объеме.", font_size=font_size)

    # 3.2: Чистая формулировка договора, срок снятия жирным, адресная только в Разделе 5
    p32 = doc.add_paragraph()
    set_p(p32, first_indent=INDENT)
    add_styled_run(p32, "3.2. ", font_size=font_size, bold=True)
    add_styled_run(p32, "В Объекте зарегистрирован по месту жительства Продавец Протасов Виктор Иванович. Продавец ", font_size=font_size)
    add_styled_run(p32, "Протасов Виктор Иванович", font_size=font_size, bold=True)
    add_styled_run(p32, " обязуется сняться с регистрационного учета по месту жительства из Объекта ", font_size=font_size)
    add_styled_run(p32, "до заключения Основного договора купли-продажи", font_size=font_size, bold=True)
    add_styled_run(p32, ". Иных лиц, зарегистрированных либо сохраняющих право пользования Объектом на момент подписания настоящего Договора, не имеется.", font_size=font_size)

    add_p("Продавцы обязуются передать Объект в пригодном для постоянного проживания состоянии, укомплектованным исправным санитарно-техническим, электротехническим и газовым оборудованием, свободным от прав третьих лиц, долгов по коммунальным платежам и капитальному ремонту.", bold_prefix="3.3. ")

    # =========================================================================
    # РАЗДЕЛ 4. ОТВЕТСТВЕННОСТЬ СТОРОН И СОГЛАШЕНИЕ О ЗАДАТКЕ
    # =========================================================================
    add_h("4. Ответственность Сторон и соглашение о задатке")

    add_p("Выданная Покупателем сумма в размере 100 000 (Сто тысяч) рублей 00 копеек является задатком и передается в доказательство заключения настоящего Договора и в обеспечение исполнения всех обязательств по заключению Основного договора купли-продажи (ст. 380, 381, 429 ГК РФ).", bold_prefix="4.1. ")

    add_p("В случае если за неисполнение настоящего Договора либо незаключение Основного договора в установленный срок ответственен Покупатель (включая немотивированный отказ от покупки либо непредоставление документов), сумма задатка остается у Продавцов в полном объеме.", bold_prefix="4.2. ")

    p43 = doc.add_paragraph()
    set_p(p43, first_indent=INDENT)
    add_styled_run(p43, "4.3. ", font_size=font_size, bold=True)
    add_styled_run(p43, "В случае если за неисполнение настоящего Договора либо незаключение Основного договора ответственны Продавцы (включая немотивированный отказ от продажи, неснятие запрета на регистрационные действия в ЕГРН Продавцом Протасовым О.И. либо неснятие Протасова В.И. с регистрационного учета до заключения Основного договора), Продавцы обязаны солидарно уплатить Покупателю двойную сумму задатка в размере ", font_size=font_size)
    add_styled_run(p43, "200 000 (Двести тысяч) рублей 00 копеек", font_size=font_size, bold=True)
    add_styled_run(p43, " в течение 3 (трех) рабочих дней со дня истечения срока, установленного п. 1.1 настоящего Договора.", font_size=font_size)

    add_p("В случае если Основной договор не может быть заключен вследствие объективных обстоятельств, за которые ни одна из Сторон не отвечает (включая обоснованный отказ Банка-кредитора в согласовании Объекта либо непреодолимую силу), сумма задатка возвращается Покупателю в одинарном размере (100 000 рублей 00 копеек) в течение 3 (трех) рабочих дней со дня наступления таких обстоятельств без применения штрафных санкций.", bold_prefix="4.4. ")

    add_p("При уклонении одной из Сторон от заключения Основного договора другая Сторона вправе обратиться в суд с требованием о понуждении заключить договор на согласованных условиях в соответствии со ст. 429, 445 ГК РФ и о возмещении причиненных убытков.", bold_prefix="4.5. ")

    # =========================================================================
    # РАЗДЕЛ 5. ОБЯЗАННОСТИ СТОРОН ПО ПОДГОТОВКЕ СДЕЛКИ
    # =========================================================================
    add_h("5. Обязанности Сторон по подготовке сделки")

    add_p("В целях заключения Основного договора и согласования сделки в Банке-кредиторе Продавцы обязуются предоставить Покупателю и Банку-кредитору в срок не позднее 10 (десяти) календарных дней до даты заключения Основного договора следующие документы:", bold_prefix="5.1. ")

    p_dol = doc.add_paragraph()
    set_p(p_dol, first_indent=INDENT)
    add_styled_run(p_dol, "— ", font_size=font_size)
    add_styled_run(p_dol, "Протасов Олег Иванович", font_size=font_size, bold=True)
    add_styled_run(p_dol, " — оригинал паспорта гражданина РФ, оригинал Свидетельства о праве на наследство по закону, страховое свидетельство (СНИЛС), свидетельство ИНН, адрес электронной почты, справку из психоневрологического диспансера об отсутствии диспансерного учета, справку из наркологического диспансера об отсутствии диспансерного учета, а также постановление судебного пристава-исполнителя и выписку из ЕГРН, подтверждающие полное снятие запрета на совершение регистрационных действий;", font_size=font_size)

    p_dvk = doc.add_paragraph()
    set_p(p_dvk, first_indent=INDENT)
    add_styled_run(p_dvk, "— ", font_size=font_size)
    add_styled_run(p_dvk, "Протасов Виктор Иванович", font_size=font_size, bold=True)
    add_styled_run(p_dvk, " — оригинал паспорта гражданина РФ, оригинал Договора передачи квартиры в собственность граждан (приватизации), оригинал Свидетельства о праве на наследство по закону, страховое свидетельство (СНИЛС), свидетельство ИНН, адрес электронной почты, справку из психоневрологического диспансера об отсутствии диспансерного учета, справку из наркологического диспансера об отсутствии диспансерного учета, а также адресную справку (справку о зарегистрированных лицах) органа регистрационного учета МВД России, подтверждающую снятие с регистрационного учета;", font_size=font_size)

    p_dall = doc.add_paragraph()
    set_p(p_dall, first_indent=INDENT)
    add_styled_run(p_dall, "— ", font_size=font_size)
    add_styled_run(p_dall, "Продавцы совместно", font_size=font_size, bold=True)
    add_styled_run(p_dall, " — справки об отсутствии задолженности по оплате коммунальных услуг, взносов на капитальный ремонт и энергоресурсов, технический паспорт / поэтажный план Объекта (по требованию Банка).", font_size=font_size)

    p_dbuy = doc.add_paragraph()
    set_p(p_dbuy, first_indent=INDENT)
    add_styled_run(p_dbuy, "5.2. ", font_size=font_size, bold=True)
    add_styled_run(p_dbuy, "Покупатель", font_size=font_size, bold=True)
    add_styled_run(p_dbuy, " обязуется предоставить: оригинал паспорта гражданина РФ, страховое свидетельство (СНИЛС), свидетельство ИНН, адрес электронной почты, свидетельства о рождении детей, документы о регистрации детей по месту жительства (форма № 8), государственный сертификат на материнский (семейный) капитал и справку СФР об остатке средств материнского капитала.", font_size=font_size)

    # =========================================================================
    # РАЗДЕЛ 6. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ И ПОДПИСИ СТОРОН
    # =========================================================================
    add_h("6. Заключительные положения")

    add_p("Настоящий Договор вступает в силу с момента его подписания всеми Сторонами и действует до полного исполнения Сторонами всех принятых на себя обязательств.", bold_prefix="6.1. ")

    add_p("Все изменения и дополнения к настоящему Договору действительны лишь в том случае, если они совершены в письменной форме и подписаны уполномоченными представителями Сторон.", bold_prefix="6.2. ")

    add_p("Настоящий Договор составлен и подписан в 3 (трех) подлинных экземплярах, имеющих равную юридическую силу: по одному экземпляру — для каждого из Продавцов, и один экземпляр — для Покупателя.", bold_prefix="6.3. ", s_after=2.0)

    p_sigt = doc.add_paragraph()
    set_p(p_sigt, space_before=3.0, s_after=1.5, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    p_sigt.paragraph_format.keep_with_next = True
    add_styled_run(p_sigt, "Подписи Сторон:", font_size=font_size, bold=True)

    add_sig("Продавец 1:", "Протасов Олег Иванович")
    add_sig("Продавец 2:", "Протасов Виктор Иванович")
    add_sig("Покупатель:", "Фирсова Ангелина Сергеевна")

    # =========================================================================
    # ПРИЛОЖЕНИЕ № 1: РАСПИСКА (СТРОГО НА ОТДЕЛЬНОМ ЛИСТЕ)
    # =========================================================================
    p_app = doc.add_paragraph()
    set_p(p_app, space_before=0, s_after=1.5, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    p_app.paragraph_format.page_break_before = True
    add_styled_run(p_app, "ПРИЛОЖЕНИЕ № 1 К ПРЕДВАРИТЕЛЬНОМУ ДОГОВОРУ КУПЛИ-ПРОДАЖИ КВАРТИРЫ", font_size=font_size, bold=True)

    p_rech = doc.add_paragraph()
    set_p(p_rech, space_before=1.5, s_after=3.0, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    add_styled_run(p_rech, "РАСПИСКА В ПОЛУЧЕНИИ И ПЕРЕДАЧЕ ЗАДАТКА", font_size=11.0, bold=True)

    p_rcd = doc.add_paragraph()
    set_p(p_rcd, space_before=0, s_after=6.0, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None)
    p_rcd.paragraph_format.tab_stops.add_tab_stop(BODY_WIDTH, WD_TAB_ALIGNMENT.RIGHT)
    add_styled_run(p_rcd, "г. Пенза", font_size=font_size, bold=True)
    add_styled_run(p_rcd, "\t", font_size=font_size)
    add_styled_run(p_rcd, deal_date, font_size=font_size, bold=True)

    add_p("Мы, нижеподписавшиеся Стороны: Продавцы гр. Протасов Олег Иванович, 15.04.1988 года рождения, паспорт гражданина РФ: серия 5612 № 345678, и гр. Протасов Виктор Иванович, 22.09.1963 года рождения, паспорт гражданина РФ: серия 5608 № 901234, с одной стороны, и Покупатель гр. Фирсова Ангелина Сергеевна, 10.08.1995 года рождения, паспорт гражданина РФ: серия 5615 № 123456, с другой стороны, составили и подписали настоящую расписку о нижеследующем:")

    # п. 1 Расписки: "в качестве задатка" строчными без выделения
    p_r1 = doc.add_paragraph()
    set_p(p_r1, first_indent=INDENT)
    add_styled_run(p_r1, "1. ", font_size=font_size, bold=True)
    add_styled_run(p_r1, "В соответствии с п. 2.2 Предварительного договора купли-продажи квартиры от ", font_size=font_size)
    add_styled_run(p_r1, deal_date, font_size=font_size)
    add_styled_run(p_r1, " Покупатель Фирсова Ангелина Сергеевна передала, а Продавец Протасов Виктор Иванович, действуя от своего имени, а также с прямого согласия и по поручению Продавца Протасова Олега Ивановича, принял наличными денежными средствами в качестве задатка в счет обеспечения исполнения обязательств по продаже 1-комнатной квартиры по адресу: Пензенская область, г. Пенза, ул. Тернопольская, д. 7, кв. 34, денежную сумму в размере ", font_size=font_size)
    add_styled_run(p_r1, "100 000 (Сто тысяч) рублей 00 копеек", font_size=font_size, bold=True)
    add_styled_run(p_r1, ".", font_size=font_size)

    add_p("Денежные средства переданы Покупателем в полном объеме и приняты Продавцами лично наличными купюрами в момент подписания Предварительного договора. Продавцы подтверждают, что факт передачи задатка Протасову В.И. признается надлежащим исполнением обязательства Покупателя перед обоими Продавцами в полном объеме (ст. 312 ГК РФ), порождает солидарную ответственность Продавцов по ст. 322, 381 ГК РФ и засчитывается в счет общей стоимости Объекта. Взаимных финансовых, материальных либо иных претензий по сумме, срокам и порядку передачи указанной суммы задатка Стороны друг к другу не имеют.", bold_prefix="2. ")

    add_p("Настоящая расписка составлена и подписана Сторонами в 2 (двух) подлинных экземплярах, имеющих равную юридическую силу: один экземпляр — стороне Продавцов, один экземпляр — Покупателю Фирсовой А.С., и является неотъемлемой частью Предварительного договора купли-продажи квартиры.", bold_prefix="3. ", s_after=2.0)

    p_rsigt = doc.add_paragraph()
    set_p(p_rsigt, space_before=3.5, s_after=1.5, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None)
    p_rsigt.paragraph_format.keep_with_next = True
    add_styled_run(p_rsigt, "Подписи Сторон расписки:", font_size=font_size, bold=True)

    add_sig("Денежные средства по поручению обоих Продавцов принял (Продавец):", "Протасов Виктор Иванович")
    add_sig("Полномочия подтверждаю, согласие на получение задатка даю (Продавец):", "Протасов Олег Иванович")
    add_sig("Денежные средства передала (Покупатель):", "Фирсова Ангелина Сергеевна")

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print(f"ПДКП успешно сгенерирован и сохранен в: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_pdkp()
