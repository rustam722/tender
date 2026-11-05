"""Калькулятор цены: прямой счёт от состава работ к итогу с НДС.

Чистая арифметика над Decimal, без ввода-вывода. Каждая статья округляется
до копеек (ROUND_HALF_UP), а следующие статьи считаются от уже округлённых:
итог всегда равен сумме статей, как её видит читатель формы.
"""
from dataclasses import dataclass
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
