from datetime import date
from decimal import Decimal

from ordertool.models import OrderData, PriceRecord, ProductLine
from ordertool.validation import validate_order


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


def test_exact_sku_price_and_arithmetic_are_verified():
    product = ProductLine(
        sku="S2620155-0115",
        description="Product",
        quantity=Decimal("2"),
        unit_price=Decimal("10.00"),
        line_total=Decimal("20.00"),
        photo_status="MATCHED",
    )
    order = make_order(product)
    validate_order(order, {
        product.sku: PriceRecord(product.sku, "Product", Decimal("10.00"))
    })
    assert product.validation_status == "VERIFIED"
    assert order.validation_status == "VERIFIED"


def test_price_and_line_total_errors_are_not_silently_corrected():
    product = ProductLine(
        sku="S2620155-0115",
        description="Product",
        quantity=Decimal("2"),
        unit_price=Decimal("12.00"),
        line_total=Decimal("25.00"),
        photo_status="MATCHED",
    )
    order = make_order(product)
    validate_order(order, {
        product.sku: PriceRecord(product.sku, "Product", Decimal("10.00"))
    })
    assert product.unit_price == Decimal("12.00")
    assert product.line_total == Decimal("25.00")
    assert product.validation_status == "DISCREPANCY"
    assert any(issue.startswith("PRICE_HIGHER") for issue in product.issues)
    assert any(issue.startswith("LINE_TOTAL_ERROR") for issue in product.issues)
