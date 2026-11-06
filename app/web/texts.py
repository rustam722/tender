"""Готовые строки для экрана: деньги, проценты, часы и итоги списков.

Экран ничего не считает и не форматирует: в JavaScript нет Decimal, любая
арифметика над деньгами там -- float, а `5700000.00` с точкой вместо запятой в
коммерческом предложении не пишут. Конвенция та же, что у контекста документов
(задача 8): голое имя -- ГОТОВАЯ строка, <поле>_value -- сырой Decimal строкой
(его экран отправляет обратно):

    amount           деньги через money -- «5 985 000,00»; часы и количества --
                     число с запятой без разрядов -- «12,5»
    amount_value     «5985000.00»
    vat_rate         ставка-доля как есть («0.05»); готовые проценты --
    vat_rate_percent 0,05 -> «5», 0,002 -> «0,2» (ставки -- исключение: `vat_rate`
                     и `vat_percent` -- разные имена, путаницы нет)

При забывчивости безопасно: `row.amount` на экране и `{{ row.amount }}` в шаблоне
дают правильный вид, а не «270.00» с точкой.

Арифметики здесь нет вовсе: форматирование (money, percent, number_ru) живёт в
app.render.context, итог списка (money_total) -- в app.calc, суммы строк и НДС
позиций -- в контексте документов. Этот слой только собирает ответ.
"""
from app.render.context import money, number_ru, percent
from app.web.schema import Entity, plain


def percent_text(rate) -> str | None:
    """Ставка-доля -> «5», «5,5»; нет ставки -- None."""
    return None if rate is None else percent(rate)


def present(entity: Entity, obj) -> dict:
    """Запись для ответа: голое имя -- готовая строка, <поле>_value -- сырое."""
    body = plain(obj)
    for name in entity.money:
        body[name + "_value"] = body[name]
        body[name] = money(getattr(obj, name))
    for name in entity.numbers:
        body[name + "_value"] = body[name]
        body[name] = number_ru(getattr(obj, name))
    for name in entity.percents:
        body[name + "_percent"] = percent_text(getattr(obj, name))
    if hasattr(obj, "norms"):
        body["norms_value"] = body["norms"]
        body["norms"] = {post: number_ru(hours) for post, hours in obj.norms.items()}
    return body


def texts_of(values) -> dict:
    """Статьи расчёта готовыми строками: {«total_wo_vat»: «150 381,00»}; сырые --
    plain(values), под именем <имя>_value."""
    items = values if isinstance(values, dict) else vars(values)
    return {name: money(value) for name, value in items.items()}
