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


def test_style_code_maps_to_article_and_pdf_retail_price_is_verified():
    # The supplied order PDF shows retail 49.90; bulk buying cost is calculated separately.
    product = ProductLine(
        style_code=STYLE,
        description="Steel vase",
        quantity=Decimal("2"),
        unit_price=RETAIL,
        line_total=Decimal("99.80"),
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
    assert product.expected_line_total == Decimal("39.92")
    assert product.calculated_line_total == Decimal("99.80")
    assert product.validation_status == "VERIFIED"


def test_pdf_price_difference_from_catalogue_retail_is_reported():
    product = ProductLine(
        style_code=STYLE,
        description="Steel vase",
        quantity=Decimal("2"),
        unit_price=Decimal("52.00"),
        line_total=Decimal("104.00"),
        photo_status="MATCHED",
    )
    order = make_order(product)
    validate_order(order, catalogue(), retail_discount_percent=Decimal("60"))

    assert product.unit_price == Decimal("52.00")
    assert product.line_total == Decimal("104.00")
    assert product.expected_unit_price == EXPECTED_BULK
    assert product.price_difference == Decimal("2.10")
    assert product.validation_status == "DISCREPANCY"
    assert any(issue.startswith("PRICE_HIGHER") for issue in product.issues)


def test_discount_can_be_changed_by_user_without_changing_pdf_retail_check():
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
    assert product.price_difference == Decimal("0.00")
    assert product.validation_status == "VERIFIED"
