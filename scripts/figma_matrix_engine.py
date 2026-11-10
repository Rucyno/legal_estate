"""
figma_matrix_engine.py
Модуль трансляции полной матрицы расчетов и архитектуры сделок из Figma / FigJam (docs/deal_matrix.md, pdkp_tree.mermaid, odkp_tree.mermaid)
в точные юридические формулировки и правила сборки договоров.
"""

import re

# 8 фундаментальных сценариев FigJam
SCENARIOS = {
    "1.1": {
        "title": "Объект без залога ➔ 100% Собственные средства (Наличные / СБР / Аккредитив)",
        "has_seller_pledge": False,
        "has_buyer_mortgage": False,
        "has_msk": False
    },
    "1.2": {
        "title": "Объект без залога ➔ Собственные средства + МСК (Материнский капитал)",
        "has_seller_pledge": False,
        "has_buyer_mortgage": False,
        "has_msk": True
    },
    "1.3": {
        "title": "Объект без залога ➔ Собственные средства + Ипотека покупателя",
        "has_seller_pledge": False,
        "has_buyer_mortgage": True,
        "has_msk": False
    },
    "1.4": {
        "title": "Объект без залога ➔ Собственные средства + Ипотека + МСК (Комбинированный)",
        "has_seller_pledge": False,
        "has_buyer_mortgage": True,
        "has_msk": True
    },
    "2.1": {
        "title": "Залог продавца (Ипотека) ➔ Собственные средства покупателя (СБР / Трансферация)",
        "has_seller_pledge": True,
        "has_buyer_mortgage": False,
        "has_msk": False
    },
    "2.2": {
        "title": "Залог продавца (Ипотека) ➔ Ипотека покупателя (Ипотека на ипотеку)",
        "has_seller_pledge": True,
        "has_buyer_mortgage": True,
        "has_msk": False
    },
    "2.3": {
        "title": "Залог продавца (Ипотека) ➔ Собственные средства + МСК (Запрет МСК на чужой залог)",
        "has_seller_pledge": True,
        "has_buyer_mortgage": False,
        "has_msk": True
    },
    "2.4": {
        "title": "Залог продавца (Ипотека) ➔ Ипотека покупателя + МСК (Максимальная сложность)",
        "has_seller_pledge": True,
        "has_buyer_mortgage": True,
        "has_msk": True
    }
}

def determine_scenario(has_seller_pledge, has_buyer_mortgage, has_msk):
    """
    Автоматически определяет сценарий из матрицы FigJam
    """
    branch = "2" if has_seller_pledge else "1"
    if not has_buyer_mortgage and not has_msk:
        sub = ".1"
    elif not has_buyer_mortgage and has_msk:
        sub = ".2" if branch == "1" else ".3"
    elif has_buyer_mortgage and not has_msk:
        sub = ".3" if branch == "1" else ".2"
    else:
        sub = ".4"
    return f"{branch}{sub}"

def generate_settlement_clauses(scenario_id, data):
    """
    Генерирует юридически выверенные пункты договора по сценарию матрицы Figma/FigJam.
    """
    price = int(data.get("price", 3000000))
    deposit = int(data.get("deposit", 100000))
    buyer_mortgage_bank = data.get("buyer_mortgage_bank", "ПАО Сбербанк")
    mortgage_amount = int(data.get("mortgage_amount", price * 0.7))
    msk_amount = int(data.get("msk_amount", 900000))
    settlement_tool = data.get("settlement_tool", "СБР") # Наличные, СБР, Аккредитив, Ячейка
    
    clauses = []
    
    # 1. Задаток
    clauses.append(f"1. Сумма в размере {deposit:,} руб. передана Покупателем Продавцу в день заключения Предварительного договора в качестве задатка в обеспечение исполнения обязательств.")
    
    # 2. Сценарии
    if scenario_id == "1.1":
        remaining = price - deposit
        if settlement_tool == "Наличные":
            clauses.append(f"2. Сумма в размере {remaining:,} руб. передается Покупателем Продавцу за счет собственных средств наличными денежными средствами до подписания Основного договора.")
            clauses.append("3. В соответствии с п. 5 ст. 488 ГК РФ право залога у Продавца на Объект НЕ возникает.")
        elif settlement_tool == "СБР":
            clauses.append(f"2. Сумма в размере {remaining:,} руб. вносится Покупателем на номинальный счет Сервиса безопасных расчетов (СБР) и перечисляется Продавцу в течение 1–5 рабочих дней после государственной регистрации перехода права собственности.")
            clauses.append("3. В соответствии с п. 5 ст. 488 ГК РФ залог в пользу Продавца не возникает.")
        elif settlement_tool == "Аккредитив":
            clauses.append(f"2. Сумма в размере {remaining:,} руб. депонируется на безотзывном покрытом аккредитиве и раскрывается в пользу Продавца после регистрации перехода права собственности в Росреестре.")
            clauses.append("3. Залог в силу закона в пользу Продавца (ст. 488 ГК РФ) не возникает.")
            
    elif scenario_id in ["1.3", "2.2"]:
        own_funds = price - deposit - mortgage_amount
        clauses.append(f"2. Сумма в размере {own_funds:,} руб. уплачивается Покупателем за счет собственных денежных средств через сервис безопасных расчетов / аккредитив.")
        clauses.append(f"3. Сумма в размере {mortgage_amount:,} руб. оплачивается за счет целевых кредитных средств, предоставляемых Банком {buyer_mortgage_bank} по кредитному договору.")
        clauses.append(f"4. Объект считается находящимся в залоге (ипотеке) у {buyer_mortgage_bank} в силу закона (ст. 77 Федерального закона № 102-ФЗ «Об ипотеке»). Право залога Продавца по п. 5 ст. 488 ГК РФ НЕ возникает.")
        
    elif scenario_id in ["1.4", "2.4"]:
        own_funds = price - deposit - mortgage_amount - msk_amount
        if own_funds > 0:
            clauses.append(f"2. Сумма в размере {own_funds:,} руб. уплачивается за счет собственных средств Покупателя.")
        clauses.append(f"3. Сумма в размере {mortgage_amount:,} руб. уплачивается за счет кредитных средств Банка {buyer_mortgage_bank}.")
        clauses.append(f"4. Сумма в размере {msk_amount:,} руб. оплачивается за счет средств материнского (семейного) капитала безналичным перечислением из Социального фонда РФ (СФР) на счет Продавца в течение установленного законом срока.")
        clauses.append(f"5. Объект находится в залоге у {buyer_mortgage_bank}. Стороны прямо подтверждают, что залог в пользу Продавца по ст. 488 ГК РФ не возникает.")

    elif scenario_id == "2.1":
        # Залог продавца + Нал/СБР
        seller_debt = int(data.get("seller_bank_debt", 1500000))
        delta = price - deposit - seller_debt
        clauses.append(f"2. Часть стоимости в размере {seller_debt:,} руб. перечисляется со счета расчетов непосредственно в счет полного погашения ипотечной задолженности Продавца для снятия обременения банка.")
        clauses.append(f"3. Оставшаяся сумма в размере {delta:,} руб. зачисляется на расчетный счет Продавца после снятия обременения и государственной регистрации перехода права собственности.")
        clauses.append("4. Залог в пользу Продавца на основании п. 5 ст. 488 ГК РФ не устанавливается.")
        
    elif scenario_id == "2.3":
        # Внимание: СФР не перечисляет МСК на счет банка продавца!
        clauses.append("2. ВНИМАНИЕ: Прямое направление средств МСК на погашение залога Продавца запрещено регламентом СФР. Расчеты проводятся после полного снятия залога Продавца либо МСК направляется на остаток стоимости после погашения долга собственными средствами.")

    return clauses
