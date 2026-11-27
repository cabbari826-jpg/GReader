from __future__ import annotations
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Any, Optional


@dataclass
class PriceRecord:
    # sku is a backward-compatible alias for the official Article Code.
    sku: str
    description: str = ""
    official_price: Optional[Decimal] = None  # Catalogue Retail Price, before negotiated discount.
    kit_qty: Optional[Decimal] = None
    catalog_qty: Optional[Decimal] = None
    article_code: str = ""
    style_code: str = ""
    source_file: str = ""
    raw: dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        self.article_code = str(self.article_code or self.sku or "").strip()
        self.sku = self.article_code


@dataclass
class ProductLine:
    # sku is the app's canonical product key: Article Code, not PDF Style Code.
    sku: str = ""
    description: str = ""
    quantity: Optional[Decimal] = None
    uom: str = ""
    unit_price: Optional[Decimal] = None
    line_total: Optional[Decimal] = None
    discount: Optional[Decimal] = None
    vat: Optional[Decimal] = None
    colour: str = ""
    page: int = 0
    style_code: str = ""
    article_code: str = ""
    photo_path: Optional[str] = None
    photo_filename: str = ""
    photo_paths: list[str] = field(default_factory=list)
    photo_filenames: list[str] = field(default_factory=list)
    photo_status: str = "NOT_FOUND"
    photo_note: str = ""
    retail_price: Optional[Decimal] = None
    expected_unit_price: Optional[Decimal] = None
    discount_percent: Decimal = Decimal("60")
    official_price: Optional[Decimal] = None  # Backward-compatible alias for expected purchase price.
    price_difference: Optional[Decimal] = None
    expected_line_total: Optional[Decimal] = None
    calculated_line_total: Optional[Decimal] = None
    line_total_difference: Optional[Decimal] = None
    kit_qty: Optional[Decimal] = None
    catalog_qty_reference: Optional[Decimal] = None
    matched_sku: str = ""
    match_method: str = ""
    match_confidence: Optional[float] = None
    validation_status: str = "REVIEW"
    issues: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def __post_init__(self):
        self.article_code = str(self.article_code or self.sku or "").strip()
        self.sku = self.article_code
        self.style_code = str(self.style_code or "").strip()


@dataclass
class OrderData:
    source_pdf: str
    source_hash: str
    order_number: str = ""
    order_date: Optional[date] = None
    customer_name: str = ""
    customer_code: str = ""
    agent: str = ""
    page_count: int = 0
    document_quantity: Optional[Decimal] = None
    document_total: Optional[Decimal] = None
    products: list[ProductLine] = field(default_factory=list)
    extraction_warnings: list[str] = field(default_factory=list)
    order_issues: list[str] = field(default_factory=list)
    conversion_status: str = "PENDING"
    validation_status: str = "PENDING"
    order_workbook: str = ""
    processed_at: str = ""


@dataclass
class BatchResult:
    orders: list[OrderData] = field(default_factory=list)
    validation_report: str = ""
    photo_library: str = ""
    skipped: list[str] = field(default_factory=list)
    errors: list[dict[str, str]] = field(default_factory=list)
    photo_records: list[dict[str, Any]] = field(default_factory=list)
