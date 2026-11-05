"""Калькулятор цены: прямой счёт от состава работ к итогу с НДС и обратный --
подбор часов под заданную цену без НДС.

Чистая арифметика над Decimal, без ввода-вывода. Каждая статья округляется
до копеек (ROUND_HALF_UP), а следующие статьи считаются от уже округлённых:
итог всегда равен сумме статей, как её видит читатель формы.
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
    insurance: Decimal      # страховые взносы и травматизм
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
    """Считает все статьи. Накладные -- от ФОТ; прибыль -- от ФОТ, страховых,
    субподряда и накладных (материалы в базу прибыли не входят)."""
    fot = _money(sum((Decimal(x.headcount) * x.hours * x.rate for x in labor),
                     Decimal("0")))
    materials = _money(materials)
    subcontract = _money(subcontract)
    # Базы берём из уже округлённых статей (fot, insurance, ...), а не из
    # сырых произведений: каждая строка формы должна пересчитываться вручную
    # от соседних напечатанных строк. Тест test_rounding_base_is_printed_articles
    # закрепляет это числами, на которых оба подхода расходятся.
    insurance = _money(fot * (settings.insurance_rate + settings.injury_rate))
    overhead = _money(fot * settings.overhead_rate)
    profit = _money((fot + insurance + subcontract + overhead)
                    * settings.profit_rate)
    # Итог намеренно СУММА напечатанных округлённых статей. Не заменять на
    # _money(сырой итог): итог перестанет сходиться со своими строками.
    total = materials + fot + insurance + subcontract + overhead + profit
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
    # a: сколько рублей итога до прибыли даёт один рубль ФОТ со всеми надбавками
    a = fot0 * (1 + settings.insurance_rate + settings.injury_rate
                + settings.overhead_rate)
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
