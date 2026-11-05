from decimal import Decimal

from app.calc import Estimate, LaborLine, estimate
from app.models import OrgSettings

DEFAULTS = OrgSettings(
    insurance_rate=Decimal("0.30"),
    injury_rate=Decimal("0.002"),
    overhead_rate=Decimal("0.05"),
    profit_rate=Decimal("0.10"),
    vat_rate=Decimal("0.05"),
)

# Исходные данные поданного КП для АО «АРКТИКГАЗ».
ARKTIK = [LaborLine("Эколог-проектировщик", 4, Decimal("1200"), Decimal("270")),
          LaborLine("Начальник отдела проектирования", 1, Decimal("360"), Decimal("420"))]


def arktik() -> Estimate:
    return estimate(ARKTIK, Decimal("41686"), Decimal("1620000"), DEFAULTS)


def test_arktikgaz_labor_matches_submitted():
    e = arktik()
    assert e.labor == Decimal("1447200.00")        # 4*1200*270 + 1*360*420
    assert e.insurance == Decimal("437054.40")     # 30 % + 0,2 %
    assert e.overhead == Decimal("72360.00")       # 5 % от ФОТ


def test_profit_is_ten_percent_of_base():
    # база: ФОТ + страховые + субподряд + накладные, материалы не входят
    e = arktik()
    base = e.labor + e.insurance + e.subcontract + e.overhead
    assert base == Decimal("3576614.40")
    assert e.profit == Decimal("357661.44")


def test_total_sums_all_articles():
    e = arktik()
    assert e.materials == Decimal("41686.00")
    assert e.subcontract == Decimal("1620000.00")
    assert e.total_wo_vat == (e.materials + e.labor + e.insurance
                              + e.subcontract + e.overhead + e.profit)
    assert e.total_wo_vat == Decimal("3975961.84")
    assert e.vat == Decimal("198798.09")           # 5 % от 3 975 961,84 = 198 798,092
    assert e.total_with_vat == Decimal("4174759.93")


def test_vat_uses_settings_rate():
    net = Decimal("5700000")
    e5 = estimate([], net, Decimal("0"), DEFAULTS)
    assert e5.total_wo_vat == Decimal("5700000.00")
    assert e5.vat == Decimal("285000.00")
    assert e5.total_with_vat == Decimal("5985000.00")

    s20 = OrgSettings(**{**DEFAULTS.__dict__, "vat_rate": Decimal("0.20")})
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
    assert str(arktik().overhead) == "72360.00"


def test_total_equals_sum_of_rounded_articles():
    # дробные ставки: каждая статья округлена отдельно, итог — сумма статей
    s = OrgSettings(insurance_rate=Decimal("0.302"), injury_rate=Decimal("0"),
                    overhead_rate=Decimal("0.0333"), profit_rate=Decimal("0.0777"),
                    vat_rate=Decimal("0.05"))
    e = estimate([LaborLine("Эколог", 3, Decimal("7.7"), Decimal("123.45"))],
                 Decimal("10.10"), Decimal("20.20"), s)
    assert e.total_wo_vat == (e.materials + e.labor + e.insurance
                              + e.subcontract + e.overhead + e.profit)
    assert e.total_with_vat == e.total_wo_vat + e.vat


def test_estimate_does_not_mutate_arguments():
    lines = [LaborLine("Эколог", 2, Decimal("10"), Decimal("100"))]
    before = [vars(x).copy() for x in lines]
    estimate(lines, Decimal("1"), Decimal("1"), DEFAULTS)
    assert [vars(x) for x in lines] == before
