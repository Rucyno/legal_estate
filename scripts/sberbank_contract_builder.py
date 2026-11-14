import sys
import io
from typing import Dict, Any

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def num_to_words_ru(n: int) -> str:
    """Встроенный конвертер целых чисел в сумму прописью на русском языке (с заглавной буквы)."""
    if n == 0:
        return "Ноль"

    units = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
    units_fem = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
    teens = [
        "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать",
        "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"
    ]
    tens = [
        "", "десять", "двадцать", "тридцать", "сорок", "пятьдесят",
        "шестьдесят", "семьдесят", "восемьдесят", "девяносто"
    ]
    hundreds = [
        "", "сто", "двести", "триста", "четыреста", "пятьсот",
        "шестьсот", "семьсот", "восемьсот", "девятьсот"
    ]

    orders = [
        ("", "", "", False),
        ("тысяча", "тысячи", "тысяч", True),
        ("миллион", "миллиона", "миллионов", False),
        ("миллиард", "миллиарда", "миллиардов", False),
    ]

    def morph(val: int, f1: str, f2: str, f5: str) -> str:
        val = abs(val) % 100
        if 10 < val < 20:
            return f5
        val_last = val % 10
        if val_last == 1:
            return f1
        if 2 <= val_last <= 4:
            return f2
        return f5

    def triad_to_words(val: int, is_fem: bool) -> list[str]:
        res = []
        h = val // 100
        t_u = val % 100
        t = t_u // 10
        u = t_u % 10
        if h > 0:
            res.append(hundreds[h])
        if 10 <= t_u <= 19:
            res.append(teens[t_u - 10])
        else:
            if t > 1:
                res.append(tens[t])
            if u > 0:
                res.append(units_fem[u] if is_fem else units[u])
        return res

    words = []
    temp = n
    order_idx = 0
    while temp > 0:
        triad = temp % 1000
        if triad > 0:
            f1, f2, f5, is_fem = orders[order_idx]
            triad_words = triad_to_words(triad, is_fem)
            if f1:
                triad_words.append(morph(triad, f1, f2, f5))
            words = triad_words + words
        temp //= 1000
        order_idx += 1

    result = " ".join(words)
    return result.capitalize()


class DKP_Payment_Builder:
    """
    Нейтральный генератор раздела "Цена и порядок расчетов" 
    для ПДКП (кратко) и ОДКП (полно).
    Поддерживает прямую сборку через методы и автоматическую сборку из JSON (Neutral DKP Payment Terms Schema).
    """
    def __init__(self, total_price: int):
        self.total_price = total_price
        self.zadatok_amount = 0
        self.blocks_odkp = []  # Полные формулировки для ОДКП
        self.blocks_pdkp = []  # Краткие формулировки для ПДКП
        self.sum_check = 0     # Для контроля баланса сумм

    def _add_sum(self, amount: int):
        self.sum_check += amount

    def _fmt(self, amount: int) -> str:
        """Форматирование чисел с пробелом-разделителем тысяч (8 000 000)."""
        return f"{amount:,}".replace(",", " ")

    @classmethod
    def from_json(cls, payload: Dict[str, Any]) -> "DKP_Payment_Builder":
        """Создает и наполняет экземпляр DKP_Payment_Builder напрямую из JSON-объекта по схеме."""
        total_price = payload["total_price"]
        builder = cls(total_price=total_price)
        blocks = payload.get("blocks", {})

        # 1. Задаток
        z = blocks.get("zadatok", {})
        if z.get("enabled") and z.get("amount", 0) > 0:
            builder.add_zadatok(
                amount=z["amount"],
                date=z.get("date", "__.__.____"),
                pdkp_number=z.get("pdkp_number", "б/н"),
                pdkp_date=z.get("pdkp_date", "__.__.____")
            )

        # 2. Наличные до сделки
        c = blocks.get("cash_before_deal", {})
        if c.get("enabled") and c.get("amount", 0) > 0:
            builder.add_cash_before_deal(amount=c["amount"])

        # 3. СБР
        s = blocks.get("sbr", {})
        if s.get("enabled") and s.get("amount", 0) > 0:
            builder.add_sbr(
                amount=s["amount"],
                punkt=s.get("punkt", "1.1"),
                recipient=s.get("recipient", "_______________"),
                rs=s.get("rs", "_______________"),
                bank_name=s.get("bank_name", "_______________"),
                bik=s.get("bik", "044525225"),
                ks=s.get("ks", "30101810400000000225"),
                inn=s.get("inn", "_______________")
            )

        # 4. Аккредитив
        a = blocks.get("accreditive", {})
        if a.get("enabled") and a.get("amount", 0) > 0:
            builder.add_accreditive(
                amount=a["amount"],
                punkt=a.get("punkt", "1.1"),
                owner_name=a.get("owner_name", "_______________")
            )

        # 5. Ипотека Сбербанк
        m = blocks.get("mortgage_sber", {})
        if m.get("enabled") and m.get("amount", 0) > 0:
            builder.add_mortgage_sber(
                amount=m["amount"],
                credit_date=m.get("contract_date", "__.__.____"),
                city=m.get("contract_city", "_______________")
            )

        # 6. Материнский капитал
        msk = blocks.get("msk", {})
        if msk.get("enabled") and msk.get("amount", 0) > 0:
            has_mortgage = msk.get("is_mortgage_related", m.get("enabled", False))
            builder.add_msk(amount=msk["amount"], has_mortgage=has_mortgage)

        # 7. Аккредитив с трансферацией (Сбербанк)
        tr = blocks.get("transfer_sber", {})
        if tr.get("enabled") and tr.get("amount", 0) > 0:
            builder.add_transfer_sber(
                amount=tr["amount"],
                seller_credit_num=tr.get("seller_credit_number", "__"),
                seller_credit_date=tr.get("seller_credit_date", "__.__.____")
            )

        # 8. Перезалог чужого банка
        po = blocks.get("payoff_other_bank", {})
        if po.get("enabled") and po.get("total_mortgage_amount", 0) > 0:
            builder.add_payoff_other_bank(
                amount=po["total_mortgage_amount"],
                payoff_amount=po.get("payoff_amount", 0),
                other_bank_name=po.get("other_bank_name", "_______________"),
                seller_credit_num=po.get("seller_credit_number", "__"),
                seller_credit_date=po.get("seller_credit_date", "__.__.____")
            )

        return builder

    # ================= БАЗОВЫЕ БЛОКИ =================
    
    def add_zadatok(self, amount: int, date: str, pdkp_number: str = "б/н", pdkp_date: str = "__.__.____"):
        """Задаток (всегда есть, если есть ПДКП)"""
        self.zadatok_amount = amount
        self._add_sum(amount)
        # ПДКП
        self.blocks_pdkp.append(
            f"В обеспечение заключения настоящего Договора и будущего Договора купли-продажи "
            f"Покупатель передает, а Продавец принимает задаток в размере {self._fmt(amount)} ({num_to_words_ru(amount)}) рублей "
            f"в день подписания настоящего Договора."
        )
        # ОДКП
        self.blocks_odkp.append(
            f"Часть стоимости Объекта недвижимости в размере {self._fmt(amount)} ({num_to_words_ru(amount)}) рублей оплачена {date} "
            f"в качестве задатка по Предварительному договору купли-продажи № {pdkp_number} от {pdkp_date}."
        )

    def add_cash_before_deal(self, amount: int):
        """Наличные до сделки (опционально: комиссия, остаток ПВ)"""
        if amount > 0:
            self._add_sum(amount)
            self.blocks_pdkp.append(
                f"{self._fmt(amount)} ({num_to_words_ru(amount)}) руб. – за счет собственных наличных денежных средств в день подписания основного Договора купли-продажи."
            )
            self.blocks_odkp.append(
                f"Часть стоимости Объекта недвижимости в размере {self._fmt(amount)} ({num_to_words_ru(amount)}) рублей оплачена за счет собственных "
                f"наличных денежных средств до заключения настоящего договора."
            )

    # ================= БАНКОВСКИЕ ИНСТРУМЕНТЫ (СБЕРБАНК) =================

    def add_sbr(self, amount: int, punkt: str = "1.1", recipient: str = "_______________", rs: str = "_______________", bank_name: str = "_______________", bik: str = "044525225", ks: str = "30101810400000000225", inn: str = "_______________"):
        """Сервис безопасных расчетов (ООО «Домклик»)"""
        self._add_sum(amount)
        self.blocks_pdkp.append(
            f"{self._fmt(amount)} ({num_to_words_ru(amount)}) руб. – с использованием сервиса безопасных расчетов (номинального счета ООО «Домклик») "
            f"после государственной регистрации перехода права собственности."
        )
        self.blocks_odkp.append(
            f"Часть стоимости Объекта недвижимости в размере {self._fmt(amount)} ({num_to_words_ru(amount)}) рублей оплачивается с использованием "
            f"номинального счета Общества с ограниченной ответственностью «Домклик» (ООО «Домклик»), "
            f"ИНН 7736249247, открытого в Операционном управлении Московского банка ПАО Сбербанк г. Москва, "
            f"к/счет 30101810400000000225, БИК 044525225. "
            f"Бенефициаром в отношении денежных средств, размещаемых на номинальном счёте, является Покупатель "
            f"до момента исполнения условий перечисления денежных средств Продавцу.\n"
            f"Перечисление денежных средств Продавцу осуществляется ООО «Домклик» по реквизитам: "
            f"получатель – {recipient}, р/с {rs}, открытый в {bank_name}, БИК: {bik}, Кор. счет: {ks}, ИНН: {inn}, "
            f"по поручению Покупателя после государственной регистрации перехода права собственности на Объект "
            f"недвижимости к Покупателю. Передача осуществляется в течение от 1 (одного) до 5 (пяти) рабочих дней "
            f"с момента получения ООО «Домклик» информации от органа госрегистрации о переходе права собственности."
        )

    def add_accreditive(self, amount: int, punkt: str = "1.1", owner_name: str = "_______________"):
        """Аккредитив ПАО Сбербанк"""
        self._add_sum(amount)
        self.blocks_pdkp.append(
            f"{self._fmt(amount)} ({num_to_words_ru(amount)}) руб. – путем открытия безотзывного покрытого аккредитива в ПАО «Сбербанк» сроком на 60 дней, "
            f"с условием раскрытия после государственной регистрации перехода права собственности."
        )
        self.blocks_odkp.append(
            f"Часть стоимости Объекта недвижимости в размере {self._fmt(amount)} ({num_to_words_ru(amount)}) рублей оплачивается Покупателем за счет "
            f"собственных денежных средств с использованием безотзывного покрытого аккредитива, открытого в "
            f"Операционном управлении Московского банка ПАО Сбербанк г. Москва, к/счет 30101810400000000225, БИК 044525225 "
            f"до подписания настоящего Договора на следующих условиях:\n"
            f"– Сумма аккредитива: {self._fmt(amount)} руб.\n"
            f"– Срок действия аккредитива: 60 (Шестьдесят) дней.\n"
            f"– Условие исполнения (раскрытия): предоставление в ПАО Сбербанк выписки из ЕГРН на Объект недвижимости, "
            f"указанный в п. {punkt}, где в качестве правообладателя указан(а): {owner_name}.\n"
            f"Перечисление осуществляется ПАО Сбербанк после госрегистрации перехода права собственности в течение "
            f"от 1 до 5 рабочих дней с момента получения Банком документов, подтверждающих переход права."
        )

    def add_mortgage_sber(self, amount: int, credit_date: str = "__.__.____", city: str = "_______________"):
        """Ипотека ПАО Сбербанк через СБР"""
        self._add_sum(amount)
        self.blocks_pdkp.append(
            f"{self._fmt(amount)} ({num_to_words_ru(amount)}) руб. – за счет целевых кредитных средств ПАО «Сбербанк» по кредитному договору, "
            f"который будет заключен Покупателем. Перечисление после госрегистрации перехода права и ипотеки в силу закона."
        )
        self.blocks_odkp.append(
            f"Часть стоимости Объекта недвижимости в размере {self._fmt(amount)} ({num_to_words_ru(amount)}) рублей оплачивается за счёт целевых кредитных "
            f"денежных средств, предоставляемых Покупателю ПАО «Сбербанк» (ИНН 7707083893, ОГРН 1027700132195, "
            f"БИК 044525225, к/с 30101810400000000225, г. Москва, ул. Вавилова д. 19) в соответствии с Кредитным "
            f"договором, заключенным {credit_date} в г. {city}.\n"
            f"Расчеты производятся с использованием номинального счета ООО «Домклик» (ИНН 7736249247, к/с 30101810400000000225, БИК 044525225). "
            f"Бенефициаром является Покупатель до момента исполнения условий перечисления.\n"
            f"Перечисление осуществляется ООО «Домклик» после государственной регистрации перехода права собственности "
            f"к Заемщику, а также государственной регистрации ипотеки Объекта недвижимости в силу закона в пользу Банка."
        )

    def add_msk(self, amount: int, has_mortgage: bool):
        """Материнский (семейный) капитал"""
        self._add_sum(amount)
        note = "" if has_mortgage else " Перечисление возможно только в случае, если детям покупателя на момент заключения договора 3 или более лет."
        
        self.blocks_pdkp.append(
            f"{self._fmt(amount)} ({num_to_words_ru(amount)}) руб. – за счет средств материнского (семейного) капитала путем перечисления Социальным фондом "
            f"России (СФР) на счет Продавца после государственной регистрации перехода права собственности."
        )
        self.blocks_odkp.append(
            f"Часть стоимости Объекта недвижимости в размере {self._fmt(amount)} ({num_to_words_ru(amount)}) рублей будет оплачена не ранее государственной "
            f"регистрации перехода права собственности за счет средств Федерального бюджета в соответствии с "
            f"Федеральным законом № 256-ФЗ от 29.12.2006 г. на основании Государственного сертификата на "
            f"материнский (семейный) капитал, путем перечисления Территориальным органом СФР на лицевой счет Продавца."
            f"{note}"
        )

    # ================= СПЕЦИАЛЬНЫЕ СЦЕНАРИИ (ПЕРЕЗАЛОГ) =================

    def add_transfer_sber(self, amount: int, seller_credit_num: str, seller_credit_date: str):
        """Аккредитив с трансферацией (ипотека продавца в Сбербанке)"""
        self._add_sum(amount)
        self.blocks_pdkp.append(
            f"{self._fmt(amount)} ({num_to_words_ru(amount)}) руб. – через аккредитив в ПАО «Сбербанк» с условием трансферации части средств в счет "
            f"погашения ипотеки Продавца по кредитному договору № {seller_credit_num} от {seller_credit_date}."
        )
        self.blocks_odkp.append(
            f"Часть стоимости в размере {self._fmt(amount)} ({num_to_words_ru(amount)}) рублей оплачивается с использованием безотзывного покрытого "
            f"аккредитива в ПАО «Сбербанк». Для исполнения аккредитива предоставляется выписка из ЕГРН с двумя "
            f"записями об ипотеке в пользу ПАО «Сбербанк» (в обеспечение обязательств Продавца по кредиту "
            f"№ {seller_credit_num} от {seller_credit_date} и в обеспечение обязательств Покупателя), а также "
            f"Заявление Продавца на трансферацию аккредитива в пользу Банка на сумму погашения задолженности. "
            f"Продавец обязуется направить часть средств на погашение ипотеки и снять обременение в течение 7 дней."
        )

    def add_payoff_other_bank(self, amount: int, payoff_amount: int, other_bank_name: str, seller_credit_num: str, seller_credit_date: str):
        """Перезалог: часть ипотечных средств идет на погашение кредита продавца в другом банке"""
        self._add_sum(amount)
        self.blocks_pdkp.append(
            f"{self._fmt(amount)} ({num_to_words_ru(amount)}) руб. – за счет кредитных средств ПАО «Сбербанк», часть из которых (в размере {self._fmt(payoff_amount)} руб.) "
            f"направляется на погашение ипотеки Продавца в {other_bank_name} для снятия обременения."
        )
        self.blocks_odkp.append(
            f"Часть стоимости в размере {self._fmt(amount)} ({num_to_words_ru(amount)}) рублей оплачивается за счет целевых кредитных средств ПАО «Сбербанк» "
            f"с использованием номинального счета ООО «Домклик». Из указанной суммы денежные средства в размере "
            f"{self._fmt(payoff_amount)} ({num_to_words_ru(payoff_amount)}) рублей направляются на погашение задолженности Продавца по кредитному договору "
            f"№ {seller_credit_num} от {seller_credit_date}, заключенному с {other_bank_name}. "
            f"Продавец обязуется направить данные средства на полное погашение задолженности для снятия обременения "
            f"в день заключения настоящего договора. В случае недостаточности средств Продавец обязуется оплатить разницу самостоятельно."
        )

    # ================= СБОРКА И ВАЛИДАЦИЯ =================

    def _validate(self):
        if self.sum_check != self.total_price:
            raise ValueError(f"ОШИБКА БАЛАНСА: Сумма блоков ({self._fmt(self.sum_check)}) не равна цене договора ({self._fmt(self.total_price)})!")

    def generate_pdkp(self) -> str:
        self._validate()
        pdkp_text = [
            f"3.1. Цена Объекта недвижимости составляет {self._fmt(self.total_price)} ({self._num_to_words(self.total_price)}) рублей.",
        ]
        if self.zadatok_amount > 0:
            pdkp_text.append(f"3.2. {self.blocks_pdkp[0]}")
            rem = self.total_price - self.zadatok_amount
            other_blocks = self.blocks_pdkp[1:]
        else:
            rem = self.total_price
            other_blocks = self.blocks_pdkp

        if other_blocks:
            pdkp_text.append(f"3.3. Оставшаяся часть цены в размере {self._fmt(rem)} ({self._num_to_words(rem)}) рублей оплачивается в следующем порядке:")
            for block in other_blocks:
                pdkp_text.append(f"– {block}")
        return "\n".join(pdkp_text)

    def generate_odkp(self) -> str:
        self._validate()
        odkp_text = [f"Цена Объекта недвижимости составляет {self._fmt(self.total_price)} ({self._num_to_words(self.total_price)}) рублей.\n"]
        for i, block in enumerate(self.blocks_odkp, start=1):
            odkp_text.append(f"{i}. {block}")
        return "\n\n".join(odkp_text)

    def _num_to_words(self, n: int) -> str:
        return num_to_words_ru(n)


# ================= ПРИМЕР ИСПОЛЬЗОВАНИЯ С JSON-СХЕМОЙ =================
if __name__ == "__main__":
    sample_json = {
        "total_price": 8_000_000,
        "blocks": {
            "zadatok": {"enabled": True, "amount": 100_000, "date": "01.11.2026"},
            "cash_before_deal": {"enabled": False, "amount": 0},
            "sbr": {"enabled": True, "amount": 2_000_000},
            "accreditive": {"enabled": False, "amount": 0},
            "mortgage_sber": {
                "enabled": True,
                "amount": 5_000_000,
                "contract_date": "10.11.2026",
                "contract_city": "Москва"
            },
            "msk": {
                "enabled": True,
                "amount": 900_000,
                "is_mortgage_related": True
            },
            "transfer_sber": {"enabled": False, "amount": 0},
            "payoff_other_bank": {"enabled": False, "total_mortgage_amount": 0, "payoff_amount": 0}
        }
    }

    builder = DKP_Payment_Builder.from_json(sample_json)

    print("=== ТЕКСТ ДЛЯ ПДКП (из JSON) ===")
    print(builder.generate_pdkp())
    print("\n=== ТЕКСТ ДЛЯ ОДКП (из JSON) ===")
    print(builder.generate_odkp())
