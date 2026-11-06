"""Готовые строки для экрана: деньги, проценты, часы и итоги списков.

Экран ничего не считает и не форматирует: в JavaScript нет Decimal, любая
арифметика над деньгами там -- float, а `5700000.00` с точкой вместо запятой в
коммерческом предложении не пишут. Поэтому рядом с каждым сырым значением
(его экран отправляет обратно как есть) лежит готовая строка:

    <поле>_text      деньги через money -- «5 985 000,00»; часы и количества --
                     число с запятой без разрядов -- «12,5»
    <поле>_percent   ставка-доля как проценты строкой -- 0,05 -> «5», 0,002 -> «0,2»

Единственная арифметика этого слоя -- вынужденная: ставка * 100 и итог списка
(сумма уже напечатанных строк). Всё остальное берётся из расчёта и контекста.
"""
from decimal import ROUND_HALF_UP, Decimal

from app.render.context import money
from app.web.schema import Entity, plain

_KOPECK = Decimal("0.01")


def percent_text(rate: Decimal | None) -> str | None:
    """0.05 -> «5», 0.055 -> «5,5»; нет ставки -- None."""
    if rate is None:
        return None
    return format((rate * 100).normalize(), "f").replace(".", ",")


def number_text(value: Decimal) -> str:
    """Не деньги: часы и количества. 12.50 -> «12,5», 360 -> «360»."""
    return format(value.normalize(), "f").replace(".", ",")


def money_total(values) -> Decimal:
    """Итог списка денег: сумма строк, КАЖДАЯ из которых округлена до копеек,
    как она будет напечатана (доктрина напечатанной строки: итог сходится с
    видимыми над ним строками). Пустой список -- 0.00."""
    return sum((v.quantize(_KOPECK, rounding=ROUND_HALF_UP) for v in values),
               Decimal("0.00"))


def present(entity: Entity, obj) -> dict:
    """Запись для ответа: сырые значения плюс готовые строки."""
    body = plain(obj)
    for name in entity.money:
        value = getattr(obj, name)
        body[name + "_text"] = None if value is None else money(value)
    for name in entity.numbers:
        body[name + "_text"] = number_text(getattr(obj, name))
    for name in entity.percents:
        body[name + "_percent"] = percent_text(getattr(obj, name))
    if hasattr(obj, "norms"):
        body["norms_text"] = {post: number_text(hours) for post, hours in obj.norms.items()}
    return body


def texts_of(values) -> dict:
    """Статьи расчёта готовыми строками: {«total_wo_vat»: «150 381,00»}."""
    items = values if isinstance(values, dict) else vars(values)
    return {name: money(value) for name, value in items.items()}
