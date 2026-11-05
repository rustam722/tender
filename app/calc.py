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


def solve_hours(target_wo_vat: Decimal, labor: list[LaborLine], materials: Decimal,
                subcontract: Decimal, settings: OrgSettings) -> list[LaborLine]:
    """Подбирает часы так, чтобы цена без НДС была как можно ближе к цели.

    Ставки, проценты, материалы и субподряд не трогаются. Один множитель k
    на все строки (замкнутая форма), часы -- целые; затем жадная доводка.
    Исходные строки не меняются, возвращаются новые."""
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
        raise TargetTooLow(
            f"Цена {_rub(target_wo_vat)} руб. без НДС не больше неизбежных расходов "
            f"(материалы и субподряд с прибылью — {_rub(fixed)} руб.): на работу "
            "ничего не остаётся. Увеличьте цену или уменьшите материалы и субподряд.")

    hours = [(k * x.hours).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
             for x in labor]

    def gap(current: list[Decimal]) -> Decimal:
        lines = [replace(x, hours=h) for x, h in zip(labor, current)]
        return abs(target_wo_vat - estimate(lines, materials, subcontract,
                                            settings).total_wo_vat)

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
