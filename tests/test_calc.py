from dataclasses import replace
from decimal import Decimal

import pytest

from app.calc import Estimate, LaborLine, estimate
from app.models import OrgSettings

DEFAULTS = OrgSettings(
    insurance_rate=Decimal("0.30"),
    injury_rate=Decimal("0.002"),
    overhead_rate=Decimal("0.05"),
    profit_rate=Decimal("0.10"),
    vat_rate=Decimal("0.05"),
)
# Те же ставки, но итог не округляется до рублей: статьи и их сумма видны до
# копейки. Нужен там, где проверяется арифметика статей, а не округление итога.
KOPECKS = replace(DEFAULTS, total_rounding_digits=2)

# Исходные данные поданного КП для АО «АРКТИКГАЗ».
ARKTIK = [LaborLine("Эколог-проектировщик", 4, Decimal("1200"), Decimal("270")),
          LaborLine("Начальник отдела проектирования", 1, Decimal("360"), Decimal("420"))]


def arktik() -> Estimate:
    return estimate(ARKTIK, Decimal("41686"), Decimal("1620000"), DEFAULTS)


def test_arktikgaz_matches_every_workbook_value():
    # Значения -- кэш формул КП.xlsx (адреса ячеек справа). Модель взята из
    # формул книги, а не восстановлена по виду документа.
    e = arktik()
    assert e.labor == Decimal("1447200.00")        # D16 = F54: 4*1200*270 + 1*360*420
    assert e.insurance == Decimal("437054.40")     # D17 + D18: 434 160 + 2 894,40
    assert e.labor + e.insurance == Decimal("1884254.40")   # D15 = SUM(D16:D18)
    assert e.subcontract == Decimal("1620000.00")  # D21 = G54
    assert e.materials == Decimal("41686.00")      # D11 = F43
    assert e.labor + e.insurance + e.subcontract == Decimal("3504254.40")  # D23, п.7
    assert e.overhead == Decimal("94212.72")       # D24 = ROUND(D15 * 5 %, 2)
    assert e.profit == Decimal("359846.71")        # D25 = ROUND((D23 + D24) * 10 %, 2)
    assert e.total_wo_vat == Decimal("4000000")    # D26 = ROUND(..., 0)


def test_articles_sum_to_3_999_999_83_and_round_to_4_000_000():
    # Ровная цифра получилась снизу вверх: сумма статей 3 999 999,83, а
    # ROUND(..., +D88) при пустой D88 округляет до целых рублей. Подгонки нет.
    e = arktik()
    articles = (e.materials + e.labor + e.insurance
                + e.subcontract + e.overhead + e.profit)
    assert articles == Decimal("3999999.83")
    assert e.total_wo_vat == Decimal("4000000")
    assert e.total_wo_vat - articles == Decimal("0.17")


def test_overhead_base_is_labor_cost_with_insurance_not_fot():
    # «5 % от п. 2» -- это D15 = ФОТ + страховые. От ФОТ было бы 72 360,00.
    e = arktik()
    assert e.overhead != e.labor * DEFAULTS.overhead_rate
    assert e.overhead == (e.labor + e.insurance) * DEFAULTS.overhead_rate
    # база движется вместе со страховыми: без страховых и травматизма накладные
    # считаются от одного ФОТ
    bare = replace(DEFAULTS, insurance_rate=Decimal("0"), injury_rate=Decimal("0"))
    e0 = estimate(ARKTIK, Decimal("41686"), Decimal("1620000"), bare)
    assert e0.insurance == Decimal("0.00")
    assert e0.overhead == Decimal("72360.00")
    # 20 % страховых: (1 447 200 + 289 440) * 5 % = 86 832,00
    e20 = estimate(ARKTIK, Decimal("41686"), Decimal("1620000"),
                   replace(DEFAULTS, insurance_rate=Decimal("0.20"),
                           injury_rate=Decimal("0")))
    assert e20.overhead == Decimal("86832.00")


def test_profit_is_ten_percent_of_base():
    # база: ФОТ + страховые + субподряд + накладные, материалы не входят
    # (п.7 + накладные: 3 504 254,40 + 94 212,72)
    e = arktik()
    base = e.labor + e.insurance + e.subcontract + e.overhead
    assert base == Decimal("3598467.12")
    assert e.profit == Decimal("359846.71")        # 359 846,712


def test_total_sums_all_articles_up_to_rounding():
    e = arktik()
    articles = (e.materials + e.labor + e.insurance
                + e.subcontract + e.overhead + e.profit)
    assert abs(e.total_wo_vat - articles) <= Decimal("0.50")
    assert e.total_wo_vat == Decimal("4000000")
    assert e.vat == Decimal("200000.00")           # 5 % от округлённого итога
    assert e.total_with_vat == Decimal("4200000.00")


def test_total_rounds_to_whole_roubles_half_up():
    # по умолчанию total_rounding_digits == 0, как ROUND(..., +D88) в книге
    assert DEFAULTS.total_rounding_digits == 0
    cases = {"100.49": "100", "100.50": "101", "100.51": "101", "0.49": "0"}
    for raw, whole in cases.items():
        e = estimate([], Decimal(raw), Decimal("0"), DEFAULTS)
        assert e.materials == Decimal(raw)          # статья не округляется до рублей
        assert e.total_wo_vat == Decimal(whole), raw
        assert e.total_with_vat == e.total_wo_vat + e.vat


def test_vat_is_taken_from_the_rounded_total():
    # 100,50 -> 101; 5 % от 101 = 5,05, а от 100,50 было бы 5,03
    e = estimate([], Decimal("100.50"), Decimal("0"), DEFAULTS)
    assert e.total_wo_vat == Decimal("101")
    assert e.vat == Decimal("5.05")
    assert e.total_with_vat == Decimal("106.05")


def test_total_rounding_digits_is_configurable():
    amount = Decimal("123456.78")
    expected = {2: "123456.78", 1: "123456.80", 0: "123457",
                -2: "123500", -3: "123000"}
    for digits, total in expected.items():
        s = replace(DEFAULTS, total_rounding_digits=digits)
        e = estimate([], amount, Decimal("0"), s)
        assert e.total_wo_vat == Decimal(total), digits
    # целые рубли у Арктикгаза -- следствие ширины округления, а не подгонки
    kopecks = estimate(ARKTIK, Decimal("41686"), Decimal("1620000"), KOPECKS)
    assert kopecks.total_wo_vat == Decimal("3999999.83")


def test_total_rounding_digits_has_a_default_but_rates_do_not():
    rates = dict(insurance_rate=Decimal("0.30"), injury_rate=Decimal("0.002"),
                 overhead_rate=Decimal("0.05"), profit_rate=Decimal("0.10"),
                 vat_rate=Decimal("0.05"))
    assert OrgSettings(**rates).total_rounding_digits == 0
    for missing in rates:
        partial = {k: v for k, v in rates.items() if k != missing}
        with pytest.raises(TypeError):
            OrgSettings(**partial)


def test_vat_uses_settings_rate():
    net = Decimal("5700000")
    e5 = estimate([], net, Decimal("0"), DEFAULTS)
    assert e5.total_wo_vat == Decimal("5700000.00")
    assert e5.vat == Decimal("285000.00")
    assert e5.total_with_vat == Decimal("5985000.00")

    s20 = replace(DEFAULTS, vat_rate=Decimal("0.20"))
    e20 = estimate([], net, Decimal("0"), s20)
    assert e20.vat == Decimal("1140000.00")
    assert e20.total_with_vat == Decimal("6840000.00")


def test_empty_labor_gives_zero_fot_but_keeps_materials():
    e = estimate([], Decimal("1000"), Decimal("0"), DEFAULTS)
    assert e.labor == Decimal("0.00")
    assert e.insurance == Decimal("0.00")
    assert e.overhead == Decimal("0.00")
    assert e.profit == Decimal("0.00")
    assert e.materials == Decimal("1000.00")
    assert e.total_wo_vat == Decimal("1000.00")


def test_money_is_rounded_to_kopecks():
    e = estimate([LaborLine("Эколог", 1, Decimal("1"), Decimal("100.005"))],
                 Decimal("0.005"), Decimal("0.015"), DEFAULTS)
    for name, value in vars(e).items():
        assert isinstance(value, Decimal), name
        assert value.as_tuple().exponent == -2, name
    assert e.labor == Decimal("100.01")            # ROUND_HALF_UP: 100,005
    assert e.materials == Decimal("0.01")
    assert e.subcontract == Decimal("0.02")
    assert str(arktik().overhead) == "94212.72"
    # и итог без НДС, округлённый до рублей, остаётся деньгами с копейками
    assert str(arktik().total_wo_vat) == "4000000.00"


# Сырой ФОТ 2,695 округляется до 2,70. Если считать страховые от сырой суммы,
# а не от напечатанного ФОТ, страховые уйдут на копейку (0,81 вместо 0,82), а
# итог статей -- тоже (4,06 вместо 4,07). Выбрана база «напечатанные округлённые
# статьи»: форма заказчика проверяется вручную по строкам, и каждая строка
# должна пересчитываться от соседних.
DIVERGING = [LaborLine("Эколог", 1, Decimal("2.695"), Decimal("1"))]

# Накладные считаются от расходов на ОТ (ФОТ + страховые), поэтому на 2,695
# они от выбора базы не зависят (0,18 в обоих случаях). ФОТ 3,755 расходится
# во всех трёх статьях: от напечатанных -- 1,14 / 0,25 / 0,52, итог 5,67;
# от сырых -- 1,13 / 0,24 / 0,51, итог 5,64.
DIVERGING_OVERHEAD = [LaborLine("Эколог", 1, Decimal("3.755"), Decimal("1"))]


def test_rounding_base_is_printed_articles():
    e = estimate(DIVERGING, Decimal("0"), Decimal("0"), KOPECKS)
    assert e.labor == Decimal("2.70")
    assert e.insurance == Decimal("0.82")      # 2,70 * 0,302 = 0,8154; от сырого 0,81
    # расходы на ОТ 2,70 + 0,82 = 3,52; 3,52 * 0,05 = 0,176
    assert e.overhead == Decimal("0.18")
    assert e.profit == Decimal("0.37")         # (3,52 + 0,18) * 0,10 = 0,370
    assert e.total_wo_vat == Decimal("4.07")   # 2,70 + 0,82 + 0,18 + 0,37; от сырого 4,06
    assert e.vat == Decimal("0.20")            # 4,07 * 0,05 = 0,2035
    assert e.total_with_vat == Decimal("4.27")


def test_overhead_and_profit_are_taken_from_printed_articles():
    e = estimate(DIVERGING_OVERHEAD, Decimal("0"), Decimal("0"), KOPECKS)
    assert e.labor == Decimal("3.76")          # 3,755 -> 3,76 (ROUND_HALF_UP)
    assert e.insurance == Decimal("1.14")      # 3,76 * 0,302 = 1,13552; от сырого 1,13
    assert e.overhead == Decimal("0.25")       # (3,76 + 1,14) * 0,05 = 0,245; от сырого 0,24
    assert e.profit == Decimal("0.52")         # (3,76 + 1,14 + 0,25) * 0,10 = 0,515; от сырого 0,51
    assert e.total_wo_vat == Decimal("5.67")   # от сырого 5,64
    assert e.vat == Decimal("0.28")            # 5,67 * 0,05 = 0,2835
    assert e.total_with_vat == Decimal("5.95")


def test_total_equals_sum_of_rounded_articles_when_not_rounded_to_roubles():
    # при total_rounding_digits=2 итог -- сумма напечатанных статей, а не
    # округление неокруглённой суммы
    e = estimate(DIVERGING, Decimal("0"), Decimal("0"), KOPECKS)
    printed = (e.materials + e.labor + e.insurance
               + e.subcontract + e.overhead + e.profit)
    assert e.total_wo_vat == printed == Decimal("4.07")
    assert e.total_with_vat == e.total_wo_vat + e.vat


def test_default_total_is_the_sum_of_printed_articles_rounded_to_roubles():
    # те же напечатанные статьи 4,07, итог по умолчанию -- 4; НДС от итога
    e = estimate(DIVERGING, Decimal("0"), Decimal("0"), DEFAULTS)
    printed = (e.materials + e.labor + e.insurance
               + e.subcontract + e.overhead + e.profit)
    assert printed == Decimal("4.07")
    assert e.total_wo_vat == Decimal("4")
    assert e.vat == Decimal("0.20")            # 4 * 0,05
    assert e.total_with_vat == Decimal("4.20")


def test_estimate_does_not_mutate_arguments():
    lines = [LaborLine("Эколог", 2, Decimal("10"), Decimal("100"))]
    before = [vars(x).copy() for x in lines]
    estimate(lines, Decimal("1"), Decimal("1"), DEFAULTS)
    assert [vars(x) for x in lines] == before
