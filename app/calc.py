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


def estimate(labor: list[LaborLine], materials: Decimal, subcontract: Decimal,
             settings: OrgSettings) -> Estimate:
    """Считает все статьи по формулам книги КП.xlsx.

    Накладные -- от расходов на оплату труда, то есть от ФОТ вместе со
    страховыми (D15 = SUM(D16:D18)); прибыль -- от расходов на ОТ, субподряда
    и накладных, материалы в базу прибыли не входят. Итог без НДС округляется
    до settings.total_rounding_digits знаков."""
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


class TargetTooLow(ValueError):
    """Целевая цена не оставляет денег на работу (или работы в расчёте нет)."""


def _rub(value: Decimal) -> str:
    return f"{_money(value):,.2f}".replace(",", " ")


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
    Исходные строки не меняются, возвращаются новые."""
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
        lines = [replace(x, hours=h) for x, h in zip(labor, current)]
        return abs(price_gap(target, estimate(lines, materials, subcontract, settings),
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
