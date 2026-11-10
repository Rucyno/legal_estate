import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import datetime

def num2words_ru(n):
    # Упрощенная сумма прописью для частых цен
    millions = n // 1000000
    thousands = (n % 1000000) // 1000
    remainder = n % 1000
    
    digits_m = {1: "Один миллион", 2: "Два миллиона", 3: "Три миллиона", 4: "Четыре миллиона", 5: "Пять миллионов",
                6: "Шесть миллионов", 7: "Семь миллионов", 8: "Восемь миллионов", 9: "Девять миллионов", 10: "Десять миллионов"}
    
    digits_th = {50: "пятьдесят тысяч", 100: "сто тысяч", 150: "сто пятьдесят тысяч", 200: "двести тысяч",
                 300: "триста тысяч", 500: "пятьсот тысяч", 800: "восемьсот тысяч", 900: "девятьсот тысяч"}
    
    res = []
    if millions in digits_m:
        res.append(digits_m[millions])
    elif millions > 0:
        res.append(f"{millions} миллионов")
        
    if thousands in digits_th:
        res.append(digits_th[thousands])
    elif thousands > 0:
        res.append(f"{thousands} тысяч")
        
    if not res:
        return f"{n} рублей 00 копеек"
    return " ".join(res) + " рублей 00 копеек"

def set_document_defaults(doc, font_name="Times New Roman", font_size_pt=10.0):
    style = doc.styles['Normal']
    style.font.name = font_name
    style.font.size = Pt(font_size_pt)

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
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.name = font_name
    r.font.size = Pt(font_size)

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

def add_paragraph(doc, text="", font_size=10.0, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=1.2, line_spacing=1.06, keep_with_next=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if keep_with_next:
        p.paragraph_format.keep_with_next = True
    if text:
        add_styled_run(p, text, font_size=font_size, bold=bold, italic=italic)
    return p

def add_signature_block(doc, title, subtitle_hint="", keep_with_next=True):
    p_title = add_paragraph(doc, title, font_size=10.0, bold=True, space_after=4.0, keep_with_next=True)
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(0)
    p_line.paragraph_format.space_after = Pt(2.0)
    p_line.paragraph_format.keep_with_next = True
    pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>' % nsdecls('w'))
    p_line._p.get_or_add_pPr().append(pBdr)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(8.0)
    p_sub.paragraph_format.keep_with_next = keep_with_next
    tabs = parse_xml(r'<w:tabs %s><w:tab w:val="right" w:pos="9638"/></w:tabs>' % nsdecls('w'))
    p_sub._p.get_or_add_pPr().append(tabs)
    add_styled_run(p_sub, "ФИО полностью", font_size=8.0, italic=True)
    add_styled_run(p_sub, "\tподпись", font_size=8.0, italic=True)

def generate_deal_documents(data, output_pdkp_path, output_receipt_path=None):
    city = data.get("city", "г. Пенза")
    deal_date = data.get("deal_date", "«___» _________ 2026 года")
    sellers = data.get("sellers", [])
    buyers = data.get("buyers", [])
    address = data.get("address", "г. Пенза, ул. ________________, д. ___, кв. ___")
    cadastral_num = data.get("cadastral_num", "58:29:____________:____")
    area = data.get("area", "____")
    floor = data.get("floor", "__")
    rooms = data.get("rooms", "1")
    
    price = int(data.get("price", 3000000))
    deposit = int(data.get("deposit", 100000))
    fssp_debt = int(data.get("fssp_debt", 0))
    fssp_fee = int(fssp_debt * 0.1) if fssp_debt > 0 else 0
    fssp_total = fssp_debt + fssp_fee
    
    basis_doc = data.get("basis_doc", "Договор купли-продажи")
    has_msc = data.get("has_msc", False)
    is_underage_kids = data.get("is_underage_kids", False)
    
    bank_encumbrance = data.get("bank_encumbrance", "Нет")
    bank_debt_amount = int(data.get("bank_debt_amount", 0))
    buyer_down_payment = int(data.get("buyer_down_payment", 0))
    
    is_undervalued = data.get("is_undervalued", False)
    declared_price = int(data.get("declared_price", price))
    
    reg_discharge = data.get("reg_discharge", "before")
    keys_handover = data.get("keys_handover", "sbr_3days")
    third_party_risk = data.get("third_party_risk", True)

    # 1. BUILD PROFESSIONAL MASTER PDKP
    doc = docx.Document()
    set_document_defaults(doc, font_name="Times New Roman", font_size_pt=10.0)
    s = doc.sections[0]
    s.top_margin = Inches(0.71)
    s.bottom_margin = Inches(0.71)
    s.left_margin = Inches(0.79)
    s.right_margin = Inches(0.59)

    add_paragraph(doc, "ПРЕДВАРИТЕЛЬНЫЙ ДОГОВОР КУПЛИ-ПРОДАЖИ КВАРТИРЫ", font_size=11.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1.0)
    add_paragraph(doc, "(с условием обеспечения обязательства задатком)", font_size=10.0, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=3.0)
    add_paragraph(doc, f"{city}                                                                                                    {deal_date}", font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6.0)

    # Preambles
    add_paragraph(doc, "Мы, нижеподписавшиеся:", bold=True, space_after=2.0)
    
    if sellers:
        for idx, s_item in enumerate(sellers):
            fio = s_item.get('fio', '________________')
            birth = s_item.get('birth', '_________')
            passport = s_item.get('passport', '________________________________')
            addr = s_item.get('address', address)
            share_text = f", доля в праве: {s_item.get('share', '1/1')}" if len(sellers) > 1 else ""
            add_paragraph(doc, f"гр. {fio}, {birth} года рождения, паспорт гражданина РФ: {passport}, зарегистрированный(ая) по адресу: {addr}{share_text}, именуемый(ая) в дальнейшем «Продавец {idx+1 if len(sellers) > 1 else ''}»,", space_after=2.0)
    else:
        add_paragraph(doc, "гр. ________________________________________, именуемый(ая) в дальнейшем «Продавец»,", space_after=2.0)

    add_paragraph(doc, "с одной стороны, и", italic=True, space_after=2.0)

    if buyers:
        for idx, b_item in enumerate(buyers):
            fio = b_item.get('fio', '________________')
            birth = b_item.get('birth', '_________')
            passport = b_item.get('passport', '________________________________')
            addr = b_item.get('address', '________________________________')
            add_paragraph(doc, f"гр. {fio}, {birth} года рождения, паспорт гражданина РФ: {passport}, зарегистрированный(ая) по адресу: {addr}, именуемый(ая) в дальнейшем «Покупатель {idx+1 if len(buyers) > 1 else ''}»,", space_after=2.0)
    else:
        add_paragraph(doc, "гр. ________________________________________, именуемый(ая) в дальнейшем «Покупатель»,", space_after=2.0)

    add_paragraph(doc, "при совместном упоминании именуемые «Стороны», а по отдельности — «Сторона», находясь в здравом уме и ясной памяти, действуя добровольно, заключили настоящий Предварительный договор (далее — «Договор») о нижеследующем:", space_after=5.0)

    # 1. Subject
    add_paragraph(doc, "1. Предмет предварительного договора", font_size=10.0, bold=True, space_after=2.0, keep_with_next=True)
    add_paragraph(doc, f"1.1. Стороны обязуются заключить в будущем Основной договор купли-продажи (далее — «Основной договор») квартиры, назначение: жилое, общей площадью {area} кв. м, этаж: {floor}, кадастровый номер {cadastral_num}, расположенной по адресу: {address} (далее — «Объект»).", space_after=2.0)
    add_paragraph(doc, f"1.2. По Основному договору Продавец обязуется передать в собственность Покупателя, а Покупатель обязуется принять и оплатить Объект по согласованной цене.", space_after=2.0)
    add_paragraph(doc, f"1.3. Объект принадлежит Продавцу на праве собственности на основании: {basis_doc}.", space_after=2.0)
    
    if basis_doc == "Приватизация":
        add_paragraph(doc, "1.3.1. До подписания Основного договора Продавец обязуется предоставить архивную справку из органа БТИ (приватизационное дело) и расширенную поквартирную карточку (выписку из домовой книги), подтверждающие отсутствие лиц, сохранивших бессрочное право пользования Объектом на основании ст. 19 Федерального закона от 29.12.2004 № 189-ФЗ.", space_after=2.0)
    elif basis_doc == "Наследство по закону/завещанию":
        add_paragraph(doc, "1.3.1. В соответствии со ст. 431.2 ГК РФ Продавец заверяет Покупателя об отсутствии иных наследников любой очереди по закону и завещанию. В случае предъявления претензий третьими лицами Продавец обязуется удовлетворить их самостоятельно и за свой счет без привлечения Покупателя.", space_after=2.0)
    
    if has_msc and is_underage_kids:
        add_paragraph(doc, "1.4. При приобретении Объекта Продавцом использовались средства материнского (семейного) капитала. Продавец обязуется в срок до заключения Основного договора заключить соглашение об определении долей в праве общей собственности с наделением несовершеннолетних детей, зарегистрировать право общей долевой собственности в Росреестре и получить приказ органа опеки и попечительства на отчуждение Объекта.", space_after=2.0)

    # FSSP & Mortgages
    if fssp_debt > 0:
        add_paragraph(doc, f"1.5. В отношении Объекта в ЕГРН зарегистрирован запрет на совершение регистрационных действий судебным приставом-исполнителем РОСП УФССП по Пензенской области. Продавец гарантирует полное погашение задолженности по исполнительным производствам на сумму {fssp_total:,} рублей в порядке, предусмотренном разделом 2 настоящего Договора, и предоставление постановления о снятии запрета до даты совершения Основного договора.", space_after=2.0)
    elif bank_encumbrance != "Нет":
        add_paragraph(doc, f"1.5. Объект обременен ипотекой в силу закона в пользу {bank_encumbrance}. Продавец обязуется обеспечить снятие записи об ипотеке в порядке расчетов по сделке.", space_after=2.0)
    else:
        add_paragraph(doc, "1.5. Продавец гарантирует, что на момент заключения настоящего Договора Объект никому не отчужден, не заложен, в споре и под арестом (запрещением) не состоит, правами третьих лиц не обременен.", space_after=2.0)

    # 2. Price and calculations
    add_paragraph(doc, "2. Цена Объекта и порядок расчетов", font_size=10.0, bold=True, space_after=2.0, keep_with_next=True)
    if is_undervalued:
        add_paragraph(doc, f"2.1. Полная согласованная цена Объекта составляет {price:,} ({num2words_ru(price)}). Стороны согласовали указать в Основном договоре цену {declared_price:,} ({num2words_ru(declared_price)}), а разницу в размере {price - declared_price:,} рублей зафиксировать Дополнительным соглашением на условиях строгой конфиденциальности.", space_after=2.0)
    else:
        add_paragraph(doc, f"2.1. Цена Объекта недвижимости по соглашению Сторон составляет {price:,} ({num2words_ru(price)}). Указанная цена является окончательной и изменению в одностороннем порядке не подлежит.", space_after=2.0)

    # Breakdown of payments
    add_paragraph(doc, "2.2. Оплата цены Объекта производится Покупателем в следующем порядке и согласованные сроки:", space_after=2.0)
    
    actual_deposit = min(deposit, 100000) if fssp_debt > 0 else deposit
    add_paragraph(doc, f"— {actual_deposit:,} ({num2words_ru(actual_deposit)}) выплачиваются Покупателем Продавцу в день подписания настоящего Договора в качестве задатка в обеспечение исполнения обязательства по заключению Основного договора;", space_after=2.0)

    if fssp_debt > 0:
        add_paragraph(doc, f"— {fssp_total:,} ({num2words_ru(fssp_total)}) перечисляются Покупателем в счет оплаты цены Объекта непосредственно на депозитный счет структурного подразделения РОСП УФССП по Пензенской области по квитанции в счет погашения задолженности Продавца по исполнительным производствам с целью снятия ареста и запрета регистрации;", space_after=2.0)

    remaining_sum = price - actual_deposit - (fssp_total if fssp_debt > 0 else 0)
    add_paragraph(doc, f"— {remaining_sum:,} ({num2words_ru(remaining_sum)}) уплачиваются Покупателем за счет собственных и/или целевых кредитных средств с использованием Сервиса безопасных расчетов (СБР) / безотзывного аккредитива с зачислением Продавцу после государственной регистрации перехода права собственности.", space_after=2.0)

    # 3. Terms, Discharge & Handover
    add_paragraph(doc, "3. Сроки заключения Основного договора и передача Объекта", font_size=10.0, bold=True, space_after=2.0, keep_with_next=True)
    term_days = 30 if (fssp_debt > 0 or has_msc) else 14
    add_paragraph(doc, f"3.1. Стороны обязуются заключить Основной договор купли-продажи в срок не позднее {term_days} (тридцати) календарных дней с момента подписания настоящего Договора.", space_after=2.0)
    
    if reg_discharge == "before":
        add_paragraph(doc, "3.2. Продавец обязуется обеспечить снятие с регистрационного учета всех лиц, зарегистрированных в Объекте, до момента заключения Основного договора и предоставить соответствующую адресную справку органа МВД.", space_after=2.0)
    else:
        add_paragraph(doc, "3.2. Продавец обязуется обеспечить снятие всех зарегистрированных лиц с регистрационного учета в срок не позднее 14 (четырнадцати) календарных дней со дня государственной регистрации перехода права собственности на Объект.", space_after=2.0)

    if keys_handover == "at_signing_cash":
        add_paragraph(doc, "3.3. Фактическая передача Объекта, ключей и подписание Передаточного акта производятся Сторонами непосредственно на Объекте до подписания Основного договора и подачи документов на государственную регистрацию.", space_after=2.0)
    else:
        add_paragraph(doc, "3.3. Фактическая передача Объекта, комплектов ключей и подписание Передаточного акта производятся Сторонами в срок не позднее 3 (трех) рабочих дней с момента полного фактического зачисления денежных средств на счет Продавца по результатам раскрытия СБР / аккредитива.", space_after=2.0)

    # 4. Liability & Third party risk
    add_paragraph(doc, "4. Ответственность Сторон и обеспечение обязательств", font_size=10.0, bold=True, space_after=2.0, keep_with_next=True)
    add_paragraph(doc, "4.1. В случае отказа или уклонения Покупателя от заключения Основного договора сумма внесенного задатка остается у Продавца в соответствии с п. 2 ст. 381 ГК РФ.", space_after=2.0)
    add_paragraph(doc, "4.2. В случае отказа или уклонения Продавца от заключения Основного договора Продавец обязан возвратить Покупателю двойную сумму задатка в соответствии с п. 2 ст. 381 ГК РФ в течение 3 (трех) банковских дней.", space_after=2.0)
    
    if third_party_risk:
        add_paragraph(doc, "4.3. (Оговорка о риске третьих лиц) В случае невозможности заключения Основного договора вследствие обстоятельств, не зависящих от воли Сторон (в том числе: отказ банка-кредитора в согласовании Объекта/заемщика, отказ органа опеки и попечительства в выдаче приказа на отчуждение, задержка снятия ареста РОСП ФССП свыше установленного срока), сумма задатка признается авансом и возвращается Покупателю в однократном размере в течение 3 (трех) рабочих дней без применения штрафных санкций.", space_after=2.0)

    # 5. Signatures
    add_paragraph(doc, "5. Подписи Сторон", font_size=10.0, bold=True, space_after=4.0, keep_with_next=True)
    
    if sellers:
        for idx, s_item in enumerate(sellers):
            fio = s_item.get('fio', '________________')
            add_signature_block(doc, f"Продавец {idx+1 if len(sellers) > 1 else ''}: {fio}", keep_with_next=True)
    else:
        add_signature_block(doc, "Продавец:", keep_with_next=True)

    if buyers:
        for idx, b_item in enumerate(buyers):
            fio = b_item.get('fio', '________________')
            is_last = (idx == len(buyers) - 1)
            add_signature_block(doc, f"Покупатель {idx+1 if len(buyers) > 1 else ''}: {fio}", keep_with_next=not is_last)
    else:
        add_signature_block(doc, "Покупатель:", keep_with_next=False)

    doc.save(output_pdkp_path)

    # 2. BUILD RECEIPT (Расписка)
    if output_receipt_path:
        rdoc = docx.Document()
        set_document_defaults(rdoc, font_name="Times New Roman", font_size_pt=10.0)
        rs = rdoc.sections[0]
        rs.top_margin = Inches(0.71)
        rs.bottom_margin = Inches(0.71)
        rs.left_margin = Inches(0.79)
        rs.right_margin = Inches(0.59)

        add_paragraph(rdoc, "РАСПИСКА В ПОЛУЧЕНИИ ДЕНЕЖНЫХ СРЕДСТВ (ЗАДАТКА)", font_size=11.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2.0)
        add_paragraph(rdoc, f"{city}                                                                                                    {deal_date}", font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6.0)

        sellers_p = ", ".join([f"гр. {s.get('fio')} (паспорт: {s.get('passport', '__________')})" for s in sellers]) if sellers else "гр. ____________________"
        buyers_p = ", ".join([f"гр. {b.get('fio')} (паспорт: {b.get('passport', '__________')})" for b in buyers]) if buyers else "гр. ____________________"

        actual_rec_deposit = min(deposit, 100000) if fssp_debt > 0 else deposit
        add_paragraph(rdoc, f"Я (мы), {sellers_p},", space_after=2.0)
        add_paragraph(rdoc, f"настоящей распиской подтверждаю(ем), что получил(и) от {buyers_p}", space_after=2.0)
        add_paragraph(rdoc, f"денежные средства в размере {actual_rec_deposit:,} ({num2words_ru(actual_rec_deposit)}) в качестве задатка во исполнение условий Предварительного договора купли-продажи квартиры от {deal_date} по адресу: {address}.", space_after=3.0)
        add_paragraph(rdoc, "Денежные средства переданы лично наличными купюрами в полном объеме. Претензий по сумме, срокам и порядку передачи денежных средств не имею(ем).", space_after=5.0)

        if sellers:
            for idx, s_item in enumerate(sellers):
                add_signature_block(rdoc, f"Деньги в качестве задатка принял (Продавец {idx+1 if len(sellers)>1 else ''} - {s_item.get('fio')}):", keep_with_next=True)
        else:
            add_signature_block(rdoc, "Деньги в качестве задатка принял (Продавец):", keep_with_next=True)

        if buyers:
            for idx, b_item in enumerate(buyers):
                is_last = (idx == len(buyers) - 1)
                add_signature_block(rdoc, f"Деньги передал (Покупатель {idx+1 if len(buyers)>1 else ''} - {b_item.get('fio')}):", keep_with_next=not is_last)
        else:
            add_signature_block(rdoc, "Деньги передал (Покупатель):", keep_with_next=False)

        rdoc.save(output_receipt_path)

    return True
