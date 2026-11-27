from decimal import Decimal

from ordertool.utils import SKU_RE, money_token, to_decimal


def test_decimal_formats():
    assert to_decimal("29.90") == Decimal("29.90")
    assert to_decimal("1,690.00") == Decimal("1690.00")
    assert to_decimal("19,164.70") == Decimal("19164.70")
    assert to_decimal("2,328") == Decimal("2328")
    assert to_decimal("29,90") == Decimal("29.90")


def test_money_token_can_split_fused_description_suffix():
    assert money_token("EXTRA218.90") == ("218.90", "EXTRA")


def test_sample_style_code_pattern():
    assert SKU_RE.fullmatch("S2620155-0115")
    assert SKU_RE.fullmatch("S2620161-0409")
    assert not SKU_RE.fullmatch("KITCHEN-WARE")
