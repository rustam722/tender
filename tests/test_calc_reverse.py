import copy
from decimal import Decimal

import pytest

from app.calc import LaborLine, TargetTooLow, estimate, price_gap, solve_hours
from tests.test_calc import ARKTIK, DEFAULTS

MATERIALS = Decimal("41686")
SUBCONTRACT = Decimal("1620000")
TARGET = Decimal("4000000")
# Надбавки выводим из настроек, а не числами: ставки меняются по требованию
# заказчика, и инвариант обязан идти за ними вместе с estimate.
# Накладные считаются от расходов на ОТ (ФОТ со страховыми), поэтому наценки
# перемножаются: A = ФОТ * (1 + страховые + травматизм) * (1 + накладные).
LABOR_MARKUP = ((1 + DEFAULTS.insurance_rate + DEFAULTS.injury_rate)
                * (1 + DEFAULTS.overhead_rate))  # ФОТ -> статьи до прибыли
PROFIT_MARKUP = 1 + DEFAULTS.profit_rate
MARKUP = LABOR_MARKUP * PROFIT_MARKUP           # рубль ФОТ -> рубли итога


def solve(target=TARGET, labor=ARKTIK, materials=MATERIALS, subcontract=SUBCONTRACT):
    return solve_hours(target, labor, materials, subcontract, DEFAULTS)


def total_of(lines, materials=MATERIALS, subcontract=SUBCONTRACT):
    return estimate(lines, materials, subcontract, DEFAULTS).total_wo_vat


def one_hour_step(labor):
    """Стоимость одного часа самой дешёвой строки с наценками: шаг сетки итогов.

    Граница расхождения относится к строкам, которые солвер может двигать:
    с ненулевыми часами и ненулевой стоимостью часа (численность * ставка).
    Строка с нулевыми часами пропускается солвером, строка с нулевой ставкой
    или численностью не меняет итог -- шаг по ним не определён."""
    costs = [x.headcount * x.rate for x in labor
             if x.hours > 0 and x.headcount * x.rate > 0]
    return min(costs) * MARKUP


def exact_multiplier(target=TARGET, labor=ARKTIK, materials=MATERIALS,
                     subcontract=SUBCONTRACT):
    """Замкнутая форма из задачи, посчитанная независимо от реализации."""
    fot0 = sum(x.headcount * x.hours * x.rate for x in labor)
    return ((target - materials) / PROFIT_MARKUP - subcontract) / (fot0 * LABOR_MARKUP)


def test_solved_hours_reproduce_target():
    got = total_of(solve())
    assert abs(got - TARGET) < Decimal("1000")


def test_residual_is_within_one_hour_of_the_cheapest_line():
    # настоящая граница достижимого: расхождение не больше стоимости
    # одного часа самой дешёвой строки с наценками
    got = total_of(solve())
    assert abs(got - TARGET) <= one_hour_step(ARKTIK)


def test_default_markups_multiply_to_the_numbers_in_the_plan():
    # 1,302 * 1,05 = 1,3671 (раньше 1,352); с прибылью 1,3671 * 1,10 = 1,50381
    assert LABOR_MARKUP == Decimal("1.36710")
    assert MARKUP == Decimal("1.503810")
    # шаг сетки: час начальника отдела, 420 * 1,50381 = 631,6002
    assert one_hour_step(ARKTIK) == Decimal("631.600200")


def test_arktikgaz_own_hours_solve_for_four_million():
    # A = 1 447 200 * 1,3671 = 1 978 467,12;
    # k = ((4 000 000 - 41 686) / 1,10 - 1 620 000) / A = 1,0000000772:
    # собственные часы фирмы -- 1200 и 360 -- и есть решение для 4 000 000.
    # Доводке делать нечего: пропорциональное масштабирование уже попадает точно.
    assert Decimal("1.0000000771") < exact_multiplier() < Decimal("1.0000000773")
    solved = solve()
    assert [x.hours for x in solved] == [Decimal("1200"), Decimal("360")]
    assert total_of(solved) == Decimal("4000000")
    assert abs(total_of(solved) - TARGET) == 0


# Цель 3 911 964: k = 0,9595482, масштабирование даёт 1151 и 345 часов и итог
# 3 910 944 (расхождение 1 020). Доводка идёт по шагу в 631,60 руб. (час
# начальника отдела): +1 час -> 3 911 576 (388), ещё +1 -> 3 912 207 (-243);
# следующий шаг в любую сторону расхождение увеличивает (875 / 1 868 / 1 381).
GREEDY_TARGET = Decimal("3911964")


def test_greedy_refinement_beats_plain_scaling():
    k = exact_multiplier(GREEDY_TARGET)
    assert Decimal("0.95954819") < k < Decimal("0.95954820")
    plain = [LaborLine("Эколог-проектировщик", 4, Decimal("1151"), Decimal("270")),
             LaborLine("Начальник отдела проектирования", 1, Decimal("345"), Decimal("420"))]
    assert [round(k * x.hours) for x in ARKTIK] == [1151, 345]
    assert total_of(plain) == Decimal("3910944.00")
    plain_gap = GREEDY_TARGET - total_of(plain)
    assert plain_gap == Decimal("1020.00")

    solved = solve(GREEDY_TARGET)
    assert [x.hours for x in solved] == [Decimal("1151"), Decimal("347")]
    assert total_of(solved) == Decimal("3912207.00")
    refined_gap = abs(total_of(solved) - GREEDY_TARGET)
    assert refined_gap == Decimal("243.00")
    assert refined_gap < plain_gap
    assert refined_gap <= one_hour_step(ARKTIK)


def test_proportions_between_posts_are_preserved():
    # отношение часов 1200:360 сохраняется с точностью до одного часа на строку
    solved = solve()
    k = exact_multiplier()
    for original, new in zip(ARKTIK, solved):
        assert abs(new.hours - k * original.hours) <= 1
    assert abs(solved[0].hours / solved[1].hours - Decimal(1200) / Decimal(360)) < Decimal("0.05")


def test_target_below_fixed_costs_raises():
    with pytest.raises(TargetTooLow, match="неизбежных расходов"):
        solve(Decimal("1000000"))


def test_target_exactly_at_fixed_costs_raises():
    # материалы + субподряд с прибылью: на работу не остаётся ни копейки
    at_fixed = MATERIALS + SUBCONTRACT * PROFIT_MARKUP
    with pytest.raises(TargetTooLow):
        solve(at_fixed)


def test_no_labor_lines_raises_rather_than_dividing_by_zero():
    with pytest.raises(TargetTooLow, match="трудозатрат"):
        solve(labor=[])


def test_lines_without_hours_count_as_no_labor():
    idle = [LaborLine("Эколог", 2, Decimal("0"), Decimal("270"))]
    with pytest.raises(TargetTooLow):
        solve(labor=idle)


def test_hours_are_whole_numbers():
    for target in (TARGET, Decimal("2345678.91"), Decimal("9999999")):
        for line in solve(target):
            assert line.hours == line.hours.to_integral_value()
            assert line.hours.as_tuple().exponent == 0
            assert line.hours >= 0


def test_solve_does_not_mutate_input_lines():
    lines = copy.deepcopy(ARKTIK)
    before = copy.deepcopy(lines)
    solved = solve(labor=lines)
    assert lines == before
    assert len(lines) == len(before)
    assert all(new is not old for new, old in zip(solved, lines))
    # остальные поля строк переносятся как были
    assert [(x.post, x.headcount, x.rate) for x in solved] == \
           [(x.post, x.headcount, x.rate) for x in lines]


def test_solve_is_deterministic():
    assert solve() == solve()
    assert solve(Decimal("3777777.77")) == solve(Decimal("3777777.77"))


def test_far_above_target_terminates_within_bound():
    for target in (Decimal("40000000"), Decimal("400000000")):
        got = total_of(solve(target))
        assert abs(got - target) <= one_hour_step(ARKTIK)


def test_just_above_fixed_costs_terminates_within_bound():
    just_above = MATERIALS + SUBCONTRACT * PROFIT_MARKUP + Decimal("1")
    solved = solve(just_above)
    assert all(x.hours >= 0 for x in solved)
    # на час работы денег не хватает: часы остаются крошечными, а не уходят в минус
    assert sum(x.hours for x in solved) <= 2


def test_proportions_drift_stays_small_across_many_targets():
    # Доводка может увести строку от k * часы больше чем на час: иначе
    # граница остатка недостижима (на двух строках сетка слишком грубая).
    # Закрепляем фактическую границу сдвига -- два часа -- как защиту от
    # расползания; на целевой цене КП сдвиг меньше часа (тест выше).
    for target in range(2_000_000, 9_000_000, 9_973):
        target = Decimal(target)
        k = exact_multiplier(target)
        for original, new in zip(ARKTIK, solve(target)):
            assert abs(new.hours - k * original.hours) <= 2, target


def test_residual_bound_holds_across_many_targets():
    step = one_hour_step(ARKTIK)
    for target in range(2_000_000, 7_000_000, 9_973):
        target = Decimal(target)
        got = total_of(solve(target))
        assert abs(got - target) <= step, target


def test_three_unequal_lines_stay_in_bound_and_in_proportion():
    lines = [LaborLine("Инженер", 3, Decimal("200"), Decimal("350")),
             LaborLine("Техник", 2, Decimal("90"), Decimal("180")),
             LaborLine("Руководитель", 1, Decimal("40"), Decimal("900"))]
    target = Decimal("750000")
    solved = solve(target, lines, Decimal("5000"), Decimal("100000"))
    got = total_of(solved, Decimal("5000"), Decimal("100000"))
    assert abs(got - target) <= one_hour_step(lines)
    k = exact_multiplier(target, lines, Decimal("5000"), Decimal("100000"))
    for original, new in zip(lines, solved):
        assert abs(new.hours - k * original.hours) <= 2


def test_bound_covers_only_lines_the_solver_can_move():
    # Мёртвые строки (нулевые часы, нулевая ставка, нулевая численность) шаг не
    # задают: граница считается по живым строкам, и солвер её выдерживает.
    lines = [LaborLine("Эколог", 4, Decimal("1200"), Decimal("270")),
             LaborLine("Начальник", 1, Decimal("360"), Decimal("420")),
             LaborLine("Без часов", 1, Decimal("0"), Decimal("50")),
             LaborLine("Без ставки", 2, Decimal("100"), Decimal("0")),
             LaborLine("Без людей", 0, Decimal("100"), Decimal("500"))]
    assert one_hour_step(lines) == one_hour_step(ARKTIK)
    solved = solve(labor=lines)
    assert solved[2].hours == 0          # строку без часов солвер не трогает
    got = total_of(solved)
    assert abs(got - TARGET) <= one_hour_step(lines)


# ---- цель «с НДС» и расхождение ------------------------------------------------
# Экран не считает денег: в JavaScript нет Decimal, любая арифметика на экране --
# float. Поэтому приведение цели с НДС к цене без НДС и разница «итог - цель»
# живут здесь, а не на экране.

def solve_vat(target, labor=ARKTIK, materials=MATERIALS, subcontract=SUBCONTRACT,
              settings=DEFAULTS):
    return solve_hours(target, labor, materials, subcontract, settings,
                       target_includes_vat=True)


def with_vat(lines, materials=MATERIALS, subcontract=SUBCONTRACT, settings=DEFAULTS):
    return estimate(lines, materials, subcontract, settings).total_with_vat


def test_target_with_vat_reproduces_the_same_hours_as_its_price_without_vat():
    # 4 000 000 без НДС при 5 % -- ровно 4 200 000 с НДС
    solved = solve_vat(Decimal("4200000"))
    assert [x.hours for x in solved] == [Decimal("1200"), Decimal("360")]
    assert with_vat(solved) == Decimal("4200000.00")


def test_flag_defaults_to_without_vat_and_keeps_the_old_signature():
    assert solve_hours(TARGET, ARKTIK, MATERIALS, SUBCONTRACT, DEFAULTS) == solve()
    assert solve_hours(TARGET, ARKTIK, MATERIALS, SUBCONTRACT, DEFAULTS,
                       target_includes_vat=False) == solve()


def test_with_vat_target_is_never_worse_than_dividing_the_target_by_one_plus_vat():
    # Доводка меряет расхождение там, где его видит заказчик, -- по итогу с НДС.
    # Деление цели на 1,05 и решение «без НДС» -- нижняя планка: на целых рублях
    # и пяти процентах оба выбора совпадают, но не должны быть лучше.
    for target in range(2_100_000, 6_000_000, 9_973):
        target = Decimal(target)
        by_vat = abs(with_vat(solve_vat(target)) - target)
        by_division = abs(with_vat(solve(target / Decimal("1.05"))) - target)
        assert by_vat <= by_division, target


def test_with_vat_residual_stays_within_one_hour_with_vat():
    step = one_hour_step(ARKTIK) * Decimal("1.05")
    for target in range(2_100_000, 6_000_000, 9_973):
        target = Decimal(target)
        assert abs(with_vat(solve_vat(target)) - target) <= step, target


def test_with_vat_uses_the_vat_rate_from_settings():
    from dataclasses import replace
    twenty = replace(DEFAULTS, vat_rate=Decimal("0.20"))
    solved = solve_vat(Decimal("4800000"), settings=twenty)     # 4 000 000 * 1,20
    assert [x.hours for x in solved] == [Decimal("1200"), Decimal("360")]
    zero = replace(DEFAULTS, vat_rate=Decimal("0"))
    assert solve_vat(TARGET, settings=zero) == solve(TARGET)    # без НДС цели совпадают


def test_with_vat_target_below_fixed_costs_says_so_in_terms_of_vat():
    with pytest.raises(TargetTooLow) as caught:
        solve_vat(Decimal("1050000"))
    text = str(caught.value)
    assert "неизбежных расходов" in text and "с НДС" in text and "без НДС" in text


def test_without_vat_message_is_unchanged():
    with pytest.raises(TargetTooLow) as caught:
        solve(Decimal("1000000"))
    assert "с НДС" not in str(caught.value).replace("без НДС", "")


def test_price_gap_is_total_minus_target_signed():
    solved = solve(GREEDY_TARGET)
    result = estimate(solved, MATERIALS, SUBCONTRACT, DEFAULTS)
    assert result.total_wo_vat == Decimal("3912207.00")
    assert price_gap(GREEDY_TARGET, result) == Decimal("243.00")          # выше цели
    assert price_gap(Decimal("3912500"), result) == Decimal("-293.00")    # ниже цели


def test_price_gap_with_vat_compares_with_vat_total():
    result = estimate(solve(), MATERIALS, SUBCONTRACT, DEFAULTS)
    assert result.total_with_vat == Decimal("4200000.00")
    assert price_gap(Decimal("4200000"), result, target_includes_vat=True) == 0
    assert price_gap(Decimal("4199999.50"), result, target_includes_vat=True) \
        == Decimal("0.50")
    # без признака та же цель сравнивается с итогом без НДС
    assert price_gap(Decimal("4200000"), result) == Decimal("-200000.00")


def test_price_gap_shows_kopecks_even_when_total_is_rounded_to_rubles():
    # итог без НДС округлён до рублей, цель -- с копейками: разница видна до копейки
    result = estimate(solve(), MATERIALS, SUBCONTRACT, DEFAULTS)
    assert price_gap(Decimal("3999999.83"), result) == Decimal("0.17")
