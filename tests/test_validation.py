from datetime import date
from decimal import Decimal

from ordertool.models import OrderData, PriceRecord, ProductLine
from ordertool.validation import validate_order


STYLE = "S2620155-0115"
ARTICLE = "6814388"
RETAIL = Decimal("49.90")
EXPECTED_BULK = Decimal("19.96")


def make_order(product):
    return OrderData(
        source_pdf="order.pdf",
        source_hash="0" * 64,
        order_number="1001",
        order_date=date(2026, 1, 1),
        document_quantity=product.quantity,
        document_total=product.line_total,
        products=[product],
    )


def catalogue():
    return {
        ARTICLE: PriceRecord(
            sku=ARTICLE,
            article_code=ARTICLE,
            style_code=STYLE,
            description="Steel vase with concrete finish H42cm",
            official_price=RETAIL,
        )
    }


def test_style_code_maps_to_article_and_uses_sixty_percent_bulk_discount():
    product = ProductLine(
        style_code=STYLE,
        description="Steel vase",
        quantity=Decimal("2"),
        unit_price=EXPECTED_BULK,
        line_total=Decimal("39.92"),
        photo_status="MATCHED",
    )
    order = make_order(product)
    validate_order(order, catalogue(), retail_discount_percent=Decimal("60"))

    assert product.article_code == ARTICLE
    assert product.sku == ARTICLE
    assert product.style_code == STYLE
    assert product.retail_price == RETAIL
    assert product.expected_unit_price == EXPECTED_BULK
    assert product.price_difference == Decimal("0.00")
    assert product.validation_status == "VERIFIED"


def test_source_price_is_preserved_and_discounted_price_mismatch_is_reported():
    product = ProductLine(
        style_code=STYLE,
        description="Steel vase",
        quantity=Decimal("2"),
        unit_price=Decimal("20.00"),
        line_total=Decimal("40.00"),
        photo_status="MATCHED",
    )
    order = make_order(product)
    validate_order(order, catalogue(), retail_discount_percent=Decimal("60"))

    assert product.unit_price == Decimal("20.00")
    assert product.line_total == Decimal("40.00")
    assert product.expected_unit_price == EXPECTED_BULK
    assert product.price_difference == Decimal("0.04")
    assert product.validation_status == "DISCREPANCY"
    assert any(issue.startswith("PRICE_HIGHER") for issue in product.issues)


def test_discount_can_be_changed_by_user():
    product = ProductLine(
        style_code=STYLE,
        quantity=Decimal("1"),
        unit_price=RETAIL,
        line_total=RETAIL,
        photo_status="MATCHED",
    )
    order = make_order(product)
    validate_order(order, catalogue(), retail_discount_percent=Decimal("0"))
    assert product.expected_unit_price == RETAIL
    assert product.validation_status == "VERIFIED"
