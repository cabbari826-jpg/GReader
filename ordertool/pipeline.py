from __future__ import annotations
from pathlib import Path
from datetime import datetime
from decimal import Decimal
import json
import traceback

from .models import BatchResult
from .extraction import extract_pdf_order
from .price_list import load_price_list
from .validation import validate_order
from .exports import write_order_workbook, build_validation_report, build_photo_library
from .history import HistoryStore
from .utils import sha256_file


def _json_default(value):
    if isinstance(value, Decimal):
        return str(value)
    if hasattr(value, "isoformat"):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    return str(value)


def write_audit_log(order, log_dir: Path):
    audit = {
        "source_pdf": order.source_pdf,
        "source_hash": order.source_hash,
        "order_number": order.order_number,
        "order_date": order.order_date,
        "customer_name": order.customer_name,
        "page_count": order.page_count,
        "document_quantity": order.document_quantity,
        "document_total": order.document_total,
        "extraction_warnings": order.extraction_warnings,
        "order_issues": order.order_issues,
        "validation_status": order.validation_status,
        "products": [
            {
                "sku": p.sku, "description": p.description, "quantity": p.quantity,
                "unit_price": p.unit_price, "line_total": p.line_total,
                "official_price": p.official_price, "price_difference": p.price_difference,
                "expected_line_total": p.expected_line_total,
                "calculated_line_total": p.calculated_line_total, "page": p.page,
                "photo_filename": p.photo_filename, "photo_status": p.photo_status,
                "matched_sku": p.matched_sku, "match_method": p.match_method,
                "validation_status": p.validation_status, "issues": p.issues, "notes": p.notes,
            } for p in order.products
        ],
    }
    path = log_dir / f"Audit_Order_{order.order_number}_{order.source_hash[:8]}.json"
    path.write_text(json.dumps(audit, ensure_ascii=False, indent=2, default=_json_default), encoding="utf-8")


def process_batch(pdf_paths, price_list_path, output_root, photo_dir=None,
                  tolerance=Decimal("0.01"), check_kit_multiple=True,
                  allow_reprocess=False, history_db=None, progress=None):
    output_root = Path(output_root)
    converted_dir = output_root / "Converted_Orders"
    reports_dir = output_root / "Validation_Reports"
    logs_dir = output_root / "Processing_Logs"
    photo_dir = Path(photo_dir) if photo_dir else output_root / "Product_Photos"
    for directory in (converted_dir, reports_dir, logs_dir, photo_dir):
        directory.mkdir(parents=True, exist_ok=True)
    history = HistoryStore(history_db or (logs_dir / "history.sqlite"))
    price_records = load_price_list(price_list_path)
    result = BatchResult()
    total = len(pdf_paths)
    log_path = logs_dir / "Processing_Log.txt"

    def log(message):
        with open(log_path, "a", encoding="utf-8") as stream:
            stream.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}\n")

    for index, pdf_path in enumerate(pdf_paths, start=1):
        try:
            pdf_path = Path(pdf_path)
            file_hash = sha256_file(pdf_path)
            if history.contains_hash(file_hash) and not allow_reprocess:
                result.skipped.append(f"{pdf_path.name}: already processed (same file hash)")
                log(f"SKIPPED duplicate {pdf_path.name}")
                if progress:
                    progress(index, total, f"Skipped duplicate: {pdf_path.name}")
                continue
            if progress:
                progress(index, total, f"Extracting {pdf_path.name}")
            order = extract_pdf_order(pdf_path, photo_dir, allow_ocr=True)
            validate_order(order, price_records, tolerance=tolerance, check_kit_multiple=check_kit_multiple)
            order.processed_at = datetime.now().isoformat(timespec="seconds")
            if progress:
                progress(index, total, f"Writing Excel for order {order.order_number}")
            order.order_workbook = write_order_workbook(order, converted_dir)
            order.conversion_status = "SUCCESS"
            write_audit_log(order, logs_dir)
            result.orders.append(order)
            for product in order.products:
                paths = product.photo_paths or ([product.photo_path] if product.photo_path else [])
                names = product.photo_filenames or ([product.photo_filename] if product.photo_filename else [])
                if not paths:
                    result.photo_records.append({
                        "sku": product.sku, "description": product.description, "filename": "",
                        "source_pdf": pdf_path.name, "extraction_status": "MISSING",
                        "verification_status": "MISSING", "page": product.page, "image_path": "",
                    })
                else:
                    for photo_index, image_path in enumerate(paths):
                        result.photo_records.append({
                            "sku": product.sku if product.photo_status in ("MATCHED", "MULTIPLE") else "",
                            "description": product.description,
                            "filename": names[photo_index] if photo_index < len(names) else Path(image_path).name,
                            "source_pdf": pdf_path.name, "extraction_status": "EXTRACTED",
                            "verification_status": product.photo_status, "page": product.page,
                            "image_path": image_path,
                        })
            history.record(order)
            log(f"PROCESSED {pdf_path.name}; order={order.order_number}; products={len(order.products)}; validation={order.validation_status}")
        except Exception as exc:
            result.errors.append({
                "filename": Path(pdf_path).name, "error": str(exc), "traceback": traceback.format_exc()
            })
            log(f"ERROR {Path(pdf_path).name}: {exc}")
        if progress:
            progress(index, total, f"Finished {Path(pdf_path).name}")

    if result.orders:
        if progress:
            progress(total, total, "Building validation report and product photo library")
        result.validation_report = build_validation_report(result.orders, reports_dir, Decimal(tolerance))
        result.photo_library = build_photo_library(result.photo_records, output_root)
        for order in result.orders:
            history.record(order, result.validation_report, result.photo_library)
        log(f"BATCH COMPLETE; orders={len(result.orders)}; errors={len(result.errors)}")
    return result
