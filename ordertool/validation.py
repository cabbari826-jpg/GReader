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
    for article, record in price_records.items():
        target = " ".join((record.description or "").upper().split())
        if target:
            scored.append((SequenceMatcher(None, source, target).ratio() * 100, article))
    scored.sort(reverse=True)
    if not scored:
        return None, 0.0, ""
    score, article = scored[0]
    margin = score - (scored[1][0] if len(scored) > 1 else 0)
    if score >= 94 and margin >= 5:
        return price_records[article], score, "DESCRIPTION"
    return None, score, "AMBIGUOUS" if score >= 75 else "UNMATCHED"


def validate_order(
    order: OrderData,
    price_records: dict[str, PriceRecord],
    tolerance: Decimal = Decimal("0.01"),
    check_kit_multiple: bool = True,
    retail_discount_percent: Decimal = Decimal("60"),
):
    """Validate PDF prices against Retail Price less the negotiated bulk discount.

    A 60% discount means the expected purchase price is 40% of catalogue Retail Price.
    The source unit price and line total are never overwritten.
    """
    tolerance = Decimal(str(tolerance))
    retail_discount_percent = Decimal(str(retail_discount_percent))
    if not Decimal("0") <= retail_discount_percent <= Decimal("100"):
        raise ValueError("Retail discount must be between 0 and 100 percent.")

    by_style = {}
    for article, record in price_records.items():
        style = (record.style_code or "").strip().upper()
        if style:
            other = by_style.get(style)
            if other and other.article_code != record.article_code:
                raise ValueError(
                    f"Style Code {style} maps to multiple Article Codes "
                    f"({other.article_code}, {record.article_code})."
                )
            by_style[style] = record

    for product in order.products:
        # Remove previous calculated issues so manual photo corrections/revalidation
        # do not duplicate validation findings.
        product.issues = [
            issue for issue in product.issues
            if not issue.startswith((
                "PRICE_", "LINE_TOTAL_", "UNMATCHED_", "AMBIGUOUS_PRODUCT_MATCH",
                "OFFICIAL_PRICE_MISSING", "UNIT_PRICE_MISSING", "QUANTITY_MISSING",
                "KIT_QTY_MISMATCH",
            ))
        ]
        style_code = (product.style_code or "").strip()
        article_code = (product.article_code or product.sku or "").strip()

        record = None
        if article_code and article_code in price_records:
            record = price_records[article_code]
            product.match_method = "ARTICLE_CODE"
            product.match_confidence = 100.0
        if not record and style_code:
            record = by_style.get(style_code.upper())
            if record:
                product.match_method = "STYLE_CODE_TO_ARTICLE"
                product.match_confidence = 100.0
        if not record and not article_code and not style_code and product.description:
            record, score, method = _description_match(product.description, price_records)
            product.match_confidence = score
            product.match_method = method
            if not record and method == "AMBIGUOUS":
                product.issues.append(
                    f"AMBIGUOUS_PRODUCT_MATCH: best description similarity {score:.1f}%"
                )
            elif not record:
                product.issues.append(
                    "UNMATCHED_PRODUCT: no Article Code, Style Code, or reliable description match"
                )

        if not record:
            if not any(issue.startswith(("UNMATCHED_PRODUCT", "AMBIGUOUS_PRODUCT_MATCH")) for issue in product.issues):
                product.issues.append(
                    f"UNMATCHED_PRODUCT: Style Code {style_code or 'not provided'} "
                    "could not be found in the selected catalogue."
                )
            product.validation_status = "REVIEW"
            product.photo_status = "REVIEW" if product.photo_path else product.photo_status
            continue

        canonical_article = (record.article_code or record.sku or "").strip()
        product.article_code = canonical_article
        product.sku = canonical_article  # Canonical product key everywhere in the app.
        product.matched_sku = canonical_article
        if not product.style_code:
            product.style_code = record.style_code
        product.retail_price = record.official_price
        product.discount_percent = retail_discount_percent
        product.kit_qty = record.kit_qty
        product.catalog_qty_reference = record.catalog_qty

        if record.official_price is None:
            product.expected_unit_price = None
            product.official_price = None
            product.issues.append(
                "OFFICIAL_PRICE_MISSING: matched catalogue row has no Retail Price."
            )
        else:
            discount_factor = (Decimal("100") - retail_discount_percent) / Decimal("100")
            product.expected_unit_price = _round_money(record.official_price * discount_factor)
            # The PDF's Price column is the catalogue Retail Price in the supplied
            # order-proposal layout. Compare document prices to retail for source
            # consistency; calculate the negotiated net buying price separately.
            product.official_price = record.official_price
            if product.unit_price is None:
                product.issues.append("UNIT_PRICE_MISSING: unit price could not be extracted.")
            else:
                product.price_difference = _round_money(
                    product.unit_price - record.official_price
                )
                if abs(product.price_difference) > tolerance:
                    direction = "HIGHER" if product.price_difference > 0 else "LOWER"
                    product.issues.append(
                        f"PRICE_{direction}: PDF price {product.unit_price} differs from "
                        f"catalogue Retail Price {record.official_price} by "
                        f"{product.price_difference}. The expected bulk purchase price "
                        f"after {retail_discount_percent}% discount is "
                        f"{product.expected_unit_price}."
                    )

        if product.quantity is None:
            product.issues.append("QUANTITY_MISSING: ordered quantity could not be extracted.")
        else:
            if product.expected_unit_price is not None:
                product.expected_line_total = _round_money(
                    product.quantity * product.expected_unit_price
                )
            if check_kit_multiple and record.kit_qty and record.kit_qty > 1:
                if product.quantity % record.kit_qty:
                    product.issues.append(
                        f"KIT_QTY_MISMATCH: quantity {product.quantity} is not a multiple "
                        f"of catalogue Kit Q.ty {record.kit_qty}."
                    )

        if product.quantity is not None and product.unit_price is not None:
            product.calculated_line_total = _round_money(
                product.quantity * product.unit_price
            )
        if product.line_total is None:
            product.issues.append("LINE_TOTAL_MISSING: line total could not be extracted.")
        elif product.calculated_line_total is not None:
            product.line_total_difference = _round_money(
                product.line_total - product.calculated_line_total
            )
            if abs(product.line_total_difference) > tolerance:
                product.issues.append(
                    f"LINE_TOTAL_ERROR: arithmetic difference {product.line_total_difference}."
                )

        if product.photo_status in ("REVIEW", "MISSING", "NOT_FOUND"):
            product.notes.append(product.photo_note or "Product photo requires manual verification.")
        hard_issue = any(issue.startswith((
            "PRICE_", "LINE_TOTAL_ERROR", "KIT_QTY_MISMATCH"
        )) for issue in product.issues)
        review_issue = bool(product.issues) or product.photo_status in (
            "REVIEW", "MISSING", "NOT_FOUND"
        )
        product.validation_status = (
            "DISCREPANCY" if hard_issue else ("REVIEW" if review_issue else "VERIFIED")
        )

    qtys = [p.quantity for p in order.products if p.quantity is not None]
    totals = [p.line_total for p in order.products if p.line_total is not None]
    sum_qty = sum(qtys, Decimal("0"))
    sum_total = sum(totals, Decimal("0"))
    order.order_issues = []
    if order.document_quantity is not None and sum_qty != order.document_quantity:
        order.order_issues.append(
            f"Document quantity {order.document_quantity} differs from extracted line sum {sum_qty}."
        )
    if order.document_total is not None and abs(sum_total - order.document_total) > tolerance:
        order.order_issues.append(
            f"Document total {order.document_total} differs from extracted line total sum {sum_total}."
        )
    if any(p.validation_status == "DISCREPANCY" for p in order.products) or order.order_issues:
        order.validation_status = "ERRORS"
    elif any(p.validation_status == "REVIEW" for p in order.products) or order.extraction_warnings:
        order.validation_status = "REVIEW"
    else:
        order.validation_status = "VERIFIED"
    return order
