from __future__ import annotations
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher

from .models import OrderData, PriceRecord

CENT = Decimal("0.01")


def _round_money(value):
    return value.quantize(CENT, rounding=ROUND_HALF_UP) if value is not None else None


def _description_match(description, price_records):
    source = " ".join((description or "").upper().split())
    if not source:
        return None, 0.0, ""
    scored = []
    for sku, record in price_records.items():
        target = " ".join((record.description or "").upper().split())
        if target:
            scored.append((SequenceMatcher(None, source, target).ratio() * 100, sku))
    scored.sort(reverse=True)
    if not scored:
        return None, 0.0, ""
    score, sku = scored[0]
    margin = score - (scored[1][0] if len(scored) > 1 else 0)
    if score >= 94 and margin >= 5:
        return price_records[sku], score, "DESCRIPTION"
    return None, score, "AMBIGUOUS" if score >= 75 else "UNMATCHED"


def validate_order(order: OrderData, price_records: dict[str, PriceRecord],
                   tolerance: Decimal = Decimal("0.01"), check_kit_multiple: bool = True):
    tolerance = Decimal(tolerance)
    for product in order.products:
        product.issues = [issue for issue in product.issues
                          if not issue.startswith(("PRICE_", "LINE_TOTAL_", "UNMATCHED_"))]
        record = price_records.get((product.sku or "").upper()) if product.sku else None
        if record:
            product.matched_sku = record.sku
            product.match_method = "SKU"
            product.match_confidence = 100.0
        elif not product.sku and product.description:
            record, score, method = _description_match(product.description, price_records)
            product.match_confidence = score
            product.match_method = method
            if record:
                product.matched_sku = record.sku
            elif method == "AMBIGUOUS":
                product.issues.append(f"AMBIGUOUS_PRODUCT_MATCH: best description similarity {score:.1f}%")
            else:
                product.issues.append("UNMATCHED_PRODUCT: product code missing and no reliable description match found")
        else:
            product.issues.append("UNMATCHED_PRODUCT: SKU not found in official price list")

        if not record:
            product.validation_status = "REVIEW"
            continue
        product.official_price = record.official_price
        product.kit_qty = record.kit_qty
        product.catalog_qty_reference = record.catalog_qty
        if record.official_price is None:
            product.issues.append("OFFICIAL_PRICE_MISSING: matched price-list record has no official price")
        elif product.unit_price is None:
            product.issues.append("UNIT_PRICE_MISSING: unit price could not be extracted")
        else:
            product.price_difference = _round_money(product.unit_price - record.official_price)
            if abs(product.price_difference) > tolerance:
                direction = "HIGHER" if product.price_difference > 0 else "LOWER"
                product.issues.append(f"PRICE_{direction}: difference {product.price_difference}")

        if product.quantity is None:
            product.issues.append("QUANTITY_MISSING: ordered quantity could not be extracted")
        else:
            if record.official_price is not None:
                product.expected_line_total = _round_money(product.quantity * record.official_price)
            if check_kit_multiple and record.kit_qty and record.kit_qty > 1:
                if product.quantity % record.kit_qty:
                    product.issues.append(
                        f"KIT_QTY_MISMATCH: quantity {product.quantity} is not a multiple of {record.kit_qty}"
                    )

        if product.quantity is not None and product.unit_price is not None:
            product.calculated_line_total = _round_money(product.quantity * product.unit_price)
        if product.line_total is None:
            product.issues.append("LINE_TOTAL_MISSING: line total could not be extracted")
        elif product.calculated_line_total is not None:
            product.line_total_difference = _round_money(product.line_total - product.calculated_line_total)
            if abs(product.line_total_difference) > tolerance:
                product.issues.append(f"LINE_TOTAL_ERROR: arithmetic difference {product.line_total_difference}")

        if product.photo_status in ("REVIEW", "MISSING"):
            product.notes.append(product.photo_note or "Product photo requires manual review")
        hard_issue = any(issue.startswith(("PRICE_", "LINE_TOTAL_ERROR", "KIT_QTY_MISMATCH"))
                         for issue in product.issues)
        review_issue = bool(product.issues) or product.photo_status in ("REVIEW", "MISSING")
        product.validation_status = "DISCREPANCY" if hard_issue else ("REVIEW" if review_issue else "VERIFIED")

    qtys = [p.quantity for p in order.products if p.quantity is not None]
    totals = [p.line_total for p in order.products if p.line_total is not None]
    sum_qty = sum(qtys, Decimal("0"))
    sum_total = sum(totals, Decimal("0"))
    order.order_issues = []
    if order.document_quantity is not None and sum_qty != order.document_quantity:
        order.order_issues.append(
            f"Document quantity {order.document_quantity} differs from extracted line sum {sum_qty}"
        )
    if order.document_total is not None and abs(sum_total - order.document_total) > tolerance:
        order.order_issues.append(
            f"Document total {order.document_total} differs from extracted line total sum {sum_total}"
        )
    if any(p.validation_status == "DISCREPANCY" for p in order.products) or order.order_issues:
        order.validation_status = "ERRORS"
    elif any(p.validation_status == "REVIEW" for p in order.products) or order.extraction_warnings:
        order.validation_status = "REVIEW"
    else:
        order.validation_status = "VERIFIED"
    return order
