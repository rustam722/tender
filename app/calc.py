"""Калькулятор цены: прямой счёт от состава работ к итогу с НДС и обратный --
подбор часов под заданную цену без НДС.

Чистая арифметика над Decimal, без ввода-вывода. Каждая статья округляется
до копеек (ROUND_HALF_UP), а следующие статьи считаются от уже округлённых.
Итог без НДС -- сумма этих статей, округлённая до total_rounding_digits знаков
(по умолчанию до целых рублей), как ROUND(..., +D88) в книге КП.xlsx.
"""
from dataclasses import dataclass, replace
from decimal import ROUND_HALF_UP, Decimal

from app.models import OrgSettings

KOPECK = Decimal("0.01")


@dataclass
class LaborLine:
    """Строка трудозатрат: должность, число людей, часы на человека, ставка руб/ч."""
    post: str
    headcount: int
    hours: Decimal
    rate: Decimal


@dataclass
class Estimate:
    """Статьи калькуляции, все в рублях с копейками."""
    materials: Decimal
    labor: Decimal          # ФОТ
    insurance: Decimal      # страховые взносы и травматизм (ФОТ + они = расходы на ОТ)
    subcontract: Decimal
    overhead: Decimal
    profit: Decimal
    total_wo_vat: Decimal
    vat: Decimal
    total_with_vat: Decimal


def _money(value: Decimal) -> Decimal:
    """Округление до копеек, ROUND_HALF_UP."""
    return value.quantize(KOPECK, rounding=ROUND_HALF_UP)


def money_total(values) -> Decimal:
    """Итог списка денег: сумма строк, КАЖДАЯ из которых округлена до копеек,
    как она будет напечатана. Доктрина напечатанной строки: итог сходится с
    видимыми над ним строками (три цены по 0,005 печатаются «0,01» и дают 0,03,
    а не 0,02 от суммы сырых). Пустой список -- 0.00.

    Живёт здесь, а не в веб-слое и не на экране: в JavaScript нет Decimal."""
    return sum((_money(Decimal(v)) for v in values), Decimal("0.00"))


def printed_labor(labor: list[LaborLine]) -> list[LaborLine]:
    """Строки ФОТ как они будут НАПЕЧАТАНЫ: сумма строки округлена до копеек
    один раз, на строке, и дальше участвует уже округлённой.

    Доктрина напечатанной строки: ФОТ в документе -- сумма напечатанных сумм
    строк, а не округление суммы сырых произведений. На двух строках
    13 ч x 181,215 руб/ч это 4 711,60 против 4 711,59, и итог расходится на
    две копейки. Строка отдаётся как «1 чел. x сумма строки x 1 руб.»:
    `estimate` округляет ФОТ сам, и от сырых произведений итог разошёлся бы
    с напечатанными строками.

    Те же множители строит `app.render.context._printed` -- оттуда идут цифры
    документов и экранов; здесь они нужны подбору часов, который без них
    минимизировал не ту величину, что напечатают."""
    return [LaborLine(x.post, 1,
                      _money(Decimal(x.headcount) * x.hours * x.rate),
                      Decimal(1))
            for x in labor]


def estimate(labor: list[LaborLine], materials: Decimal, subcontract: Decimal,
             settings: OrgSettings) -> Estimate:
    """Считает все статьи по составу формул книги КП.xlsx, но своей арифметикой.

    Накладные -- от расходов на оплату труда, то есть от ФОТ вместе со
    страховыми (D15 = SUM(D16:D18)); прибыль -- от расходов на ОТ, субподряда
    и накладных, материалы в базу прибыли не входят. Итог без НДС округляется
    до settings.total_rounding_digits знаков.

    Это НЕ воспроизведение книги до копейки: книга складывает неокруглённые
    произведения строк и не округляет страховые взносы, а здесь каждая статья
    округляется до копеек (доктрина напечатанной строки) и следующие считаются
    от уже округлённых. Поэтому итоги расходятся на границе рубля. Книгу
    воспроизводит `workbook_total`, и сборка сравнивает их два итога."""
    fot = _money(sum((Decimal(x.headcount) * x.hours * x.rate for x in labor),
                     Decimal("0")))
    materials = _money(materials)
    subcontract = _money(subcontract)
    # Базы берём из уже округлённых статей (fot, insurance, ...), а не из
    # сырых произведений: каждая строка формы должна пересчитываться вручную
    # от соседних напечатанных строк. Тесты test_rounding_base_is_printed_articles
    # и test_overhead_and_profit_are_taken_from_printed_articles закрепляют это
    # числами, на которых оба подхода расходятся.
    insurance = _money(fot * (settings.insurance_rate + settings.injury_rate))
    labor_cost = fot + insurance                      # расходы на ОТ, D15
    overhead = _money(labor_cost * settings.overhead_rate)
    profit = _money((labor_cost + subcontract + overhead) * settings.profit_rate)
    # Итог -- СУММА напечатанных округлённых статей, округлённая как в книге.
    # Не заменять на округление сырой суммы: итог перестанет сходиться со
    # статьями. Расхождение с суммой статей -- только само округление итога
    # (до 0,50 руб. при целых рублях); в книге оно то же самое.
    articles = materials + labor_cost + subcontract + overhead + profit
    total = _money(articles.quantize(Decimal(1).scaleb(-settings.total_rounding_digits),
                                     rounding=ROUND_HALF_UP))
    vat = _money(total * settings.vat_rate)
    return Estimate(materials=materials, labor=fot, insurance=insurance,
                    subcontract=subcontract, overhead=overhead, profit=profit,
                    total_wo_vat=total, vat=vat, total_with_vat=total + vat)


def workbook_total(labor: list[LaborLine], materials: Decimal, subcontract: Decimal,
                   settings: OrgSettings, digits: int = 0) -> Decimal:
    """Итог без НДС так, как его посчитает книга КП.xlsx: без округления статей,
    с округлением только там, где в книге стоит ROUND. Для сравнения с
    `estimate()`: программа округляет каждую статью до копеек (доктрина
    напечатанной строки), книга -- нет, и на границе рубля итоги расходятся.

    Формулы книги (адреса ячеек листа «Лист1»):

        D16 = F54 = SUM(E*D*C)          ФОТ, сумма НЕокруглённых произведений
        D17 = D16*C17, D18 = D16*C18    страховые и травматизм, без округления
        D15 = SUM(D16:D18)
        D21 = G54                       субподряд, сумма неокруглённых E*F
        D23 = D15 + D21                 (D19, D20, D22 -- литерал 0)
        D24 = ROUND(D15*C24, 2)         накладные
        D25 = ROUND((D23+D24)*C25, 2)   прибыль; материалы в базу не входят
        D26 = ROUND(D11+D23+D24+D25, D88)

    `labor` -- строки как их вводит человек (сырые множители), `materials` и
    `subcontract` -- суммы сырых произведений количества на цену. `digits` --
    D88: пустая клетка значит ROUND(..., 0), то есть целые рубли (форма
    заполняется без D88).

    Ставки берутся из `settings`, а книга считает по своим -- тем, что лежат в
    её клетках C17, C18, C24, C25. Поэтому вызывающий подставляет ставки книги
    в `settings` (`dataclasses.replace`), когда карта ячеек говорит, где они
    лежат (раздел `rates`, `xlsx_render.read_rates`); так делает
    `app.render.context`. Без подстановки итог предсказывает книгу лишь до тех
    пор, пока ставки в настройках не правили."""
    fot = sum((Decimal(x.headcount) * x.hours * x.rate for x in labor),
              Decimal("0"))
    labor_cost = fot + fot * settings.insurance_rate + fot * settings.injury_rate
    direct = labor_cost + subcontract                           # D23
    overhead = _money(labor_cost * settings.overhead_rate)      # D24
    profit = _money((direct + overhead) * settings.profit_rate)  # D25
    total = (materials + direct + overhead + profit).quantize(
        Decimal(1).scaleb(-digits), rounding=ROUND_HALF_UP)
    return _money(total)


class TargetTooLow(ValueError):
    """Целевая цена не оставляет денег на работу (или работы в расчёте нет)."""


def _rub(value: Decimal) -> str:
    """«1 823 686,00»: так же, как context.money (неразрывный пробел между разрядами,
    запятая десятичная). Контекст импортирует calc, поэтому формат повторён здесь;
    тест test_target_too_low_prints_money_the_way_money_does держит их в согласии."""
    return f"{_money(value):,.2f}".translate({ord(","): " ", ord("."): ","})


def price_gap(target: Decimal, result: Estimate, target_includes_vat: bool = False) -> Decimal:
    """Расхождение итога с целевой ценой: итог минус цель, в рублях с копейками.

    Положительное -- итог выше цели, отрицательное -- ниже. Сравнивается тот итог,
    к которому относится цель: с НДС при target_includes_vat, иначе без НДС.
    Итог берётся из расчёта, а не из сложения статей: итог без НДС округлён до
    рублей и может отличаться от суммы напечатанных статей (на данных Арктикгаза --
    на 17 копеек).

    Живёт здесь, а не на экране: в JavaScript нет Decimal, и вычитание денег
    там -- это float."""
    total = result.total_with_vat if target_includes_vat else result.total_wo_vat
    return _money(total - target)


def solve_hours(target: Decimal, labor: list[LaborLine], materials: Decimal,
                subcontract: Decimal, settings: OrgSettings, *,
                target_includes_vat: bool = False) -> list[LaborLine]:
    """Подбирает часы так, чтобы итог был как можно ближе к цели.

    target_includes_vat=False (по умолчанию) -- цель есть цена без НДС;
    True -- цель есть цена с НДС по ставке settings.vat_rate. Приведение цели
    с НДС к цене без НДС делается здесь, а не на экране (там нет Decimal).
    Стартовые часы берутся из цены без НДС = цель / (1 + ставка), а доводка
    меряет расхождение там, где цель задана, то есть по итогу с НДС. При
    целых рублях итога и ставке в целых процентах обе меры дают тот же выбор
    (итог с НДС пропорционален итогу без НДС); различия возможны лишь на
    границе копейки при дробных ставках или округлении итога до копеек.

    Ставки, проценты, материалы и субподряд не трогаются. Один множитель k
    на все строки (замкнутая форма), часы -- целые; затем жадная доводка.
    Исходные строки не меняются, возвращаются новые.

    Доводка меряет итог по НАПЕЧАТАННЫМ строкам (`printed_labor`): именно его
    выведут документ и экран, и минимизировать надо его. Вызывающий обязан
    считать итог подобранных часов так же -- через `printed_labor`."""
    target_wo_vat = target / (1 + settings.vat_rate) if target_includes_vat else target
    fot0 = sum((Decimal(x.headcount) * x.hours * x.rate for x in labor),
               Decimal("0"))
    # a: сколько рублей итога до прибыли дают ФОТ со всеми надбавками.
    # Накладные считаются от расходов на ОТ (ФОТ со страховыми), поэтому
    # наценки перемножаются, а не складываются: ФОТ * 1,302 * 1,05.
    a = (fot0 * (1 + settings.insurance_rate + settings.injury_rate)
         * (1 + settings.overhead_rate))
    if a == 0:
        raise TargetTooLow("В расчёте нет трудозатрат: нет строк с часами и ставкой, "
                           "подбирать часы не из чего. Добавьте строки трудозатрат.")
    k = ((target_wo_vat - materials) / (1 + settings.profit_rate) - subcontract) / a
    if k <= 0:
        fixed = materials + subcontract * (1 + settings.profit_rate)
        price = (f"{_rub(target)} руб. с НДС ({_rub(target_wo_vat)} руб. без НДС)"
                 if target_includes_vat else f"{_rub(target_wo_vat)} руб. без НДС")
        raise TargetTooLow(
            f"Цена {price} не больше неизбежных расходов "
            f"(материалы и субподряд с прибылью — {_rub(fixed)} руб.): на работу "
            "ничего не остаётся. Увеличьте цену или уменьшите материалы и субподряд.")

    hours = [(k * x.hours).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
             for x in labor]

    def gap(current: list[Decimal]) -> Decimal:
        # Итог считается по НАПЕЧАТАННЫМ строкам -- тот самый, который выведут
        # документ и экран калькуляции. По сырым строкам подбор минимизировал
        # не ту величину и показывал расхождение там, где документ попадал
        # в цель ровно.
        lines = [replace(x, hours=h) for x, h in zip(labor, current)]
        return abs(price_gap(target, estimate(printed_labor(lines), materials,
                                              subcontract, settings),
                             target_includes_vat))

    # Жадная доводка: часы округлены, и шаг сетки итогов -- стоимость часа самой
    # дешёвой строки. Каждый проход берёт ровно один сдвиг на час (строка, знак),
    # который сильнее всего уменьшает расхождение. Расхождение в копейках строго
    # убывает, поэтому цикл конечен. Ничья -- первая строка, сначала +1. Часы не
    # уходят ниже нуля; строки без часов не трогаем.
    best = gap(hours)
    while True:
        move = None
        for i, line in enumerate(labor):
            if line.hours <= 0:
                continue
            for step in (1, -1):
                new = hours[i] + step
                if new < 0:
                    continue
                trial = hours[:i] + [new] + hours[i + 1:]
                g = gap(trial)
                if g < best:
                    best, move = g, trial
        if move is None:
            break
        hours = move
    return [replace(x, hours=h) for x, h in zip(labor, hours)]
