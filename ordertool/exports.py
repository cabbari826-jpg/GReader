from __future__ import annotations
from pathlib import Path
from datetime import datetime
from decimal import Decimal
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from .models import OrderData
from .utils import unique_path, safe_name, sanitize_excel_text
from .excel_images import embed_native_cell_images

NAVY = "24364B"
BLUE = "2F5597"
PALE_BLUE = "D9EAF7"
GREEN = "E2F0D9"
RED = "FCE4D6"
YELLOW = "FFF2CC"
GREY = "E7E6E6"
WHITE = "FFFFFF"
BORDER = "D7DEE8"


def _base_sheet(ws, freeze=None):
    ws.sheet_view.showGridLines = False
    if freeze:
        ws.freeze_panes = freeze
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _style_header(ws, row, start_col=1, end_col=None):
    end_col = end_col or ws.max_column
    for cells in ws.iter_rows(min_row=row, max_row=row, min_col=start_col, max_col=end_col):
        for cell in cells:
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.font = Font(bold=True, color=WHITE, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="medium", color=BLUE))
    ws.row_dimensions[row].height = 32


def _status_fill(status):
    value = (status or "").upper()
    if value in ("VERIFIED", "MATCHED", "SUCCESS"):
        return GREEN
    if value in ("DISCREPANCY", "ERRORS", "FAILED"):
        return RED
    if value in ("REVIEW", "AMBIGUOUS", "MISSING", "MULTIPLE"):
        return YELLOW
    return GREY


def _format_table(ws, header_row, end_row, name):
    if end_row <= header_row:
        ws.auto_filter.ref = f"A{header_row}:{ws.cell(header_row, ws.max_column).column_letter}{header_row}"
        return
    end_col = ws.cell(header_row, ws.max_column).column_letter
    table = Table(displayName=name, ref=f"A{header_row}:{end_col}{end_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(table)
    ws.auto_filter.ref = table.ref
    _style_header(ws, header_row, 1, ws.max_column)


def _verify_workbook(path, expected_sheets):
    with zipfile.ZipFile(path, "r") as archive:
        bad = archive.testzip()
        if bad:
            raise RuntimeError(f"Generated XLSX failed ZIP integrity check: {bad}")
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        missing = set(expected_sheets) - set(workbook.sheetnames)
        if missing:
            raise RuntimeError(f"Generated workbook is missing worksheets: {sorted(missing)}")
    finally:
        workbook.close()


def _finalize_with_native_images(path, expected_cells):
    count = sum(len(items) for items in expected_cells.values())
    if count:
        embed_native_cell_images(path, expected_cells)
    else:
        _verify_workbook(path, set(expected_cells) or None)


def write_order_workbook(order: OrderData, output_dir: str | Path) -> str:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    target = unique_path(output_dir / f"Order_{safe_name(order.order_number)}.xlsx")
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Order Details"
    _base_sheet(ws, "B8")
    ws.merge_cells("A1:S1")
    ws["A1"] = f"ORDER {order.order_number} - EXTRACTED DETAILS"
    ws["A1"].font = Font(size=16, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[1].height = 30
    metadata = [
        ("Order Number", order.order_number, "Order Date", order.order_date),
        ("Customer", order.customer_name, "Agent", order.agent),
        ("Source PDF", Path(order.source_pdf).name, "Pages", order.page_count),
        ("Document Quantity", float(order.document_quantity) if order.document_quantity is not None else None,
         "Document Total", float(order.document_total) if order.document_total is not None else None),
    ]
    for row_index, values in enumerate(metadata, start=3):
        for col_index, value in zip((1, 2, 4, 5), values):
            cell = ws.cell(row_index, col_index, sanitize_excel_text(value))
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col_index in (1, 4):
                cell.font = Font(bold=True, color=NAVY)
                cell.fill = PatternFill("solid", fgColor=PALE_BLUE)
    headers = [
        "Product Photo", "Article Code", "Style Code", "Product Description", "Ordered Quantity",
        "Unit of Measurement", "Unit Price (PDF)", "Extracted Line Total", "Retail Price",
        "Expected Unit Price (After Discount)", "Price Difference", "Validation Status",
        "Calculated Line Total", "Expected Purchase Total", "Retail Discount (%)",
        "Colour", "Source Page", "Photo Filename", "Validation Notes",
    ]
    header_row, first_data = 7, 8
    for col_index, heading in enumerate(headers, start=1):
        ws.cell(header_row, col_index, heading)
    _style_header(ws, header_row, 1, len(headers))
    image_cells = {"Order Details": []}
    for index, product in enumerate(order.products):
        row = first_data + index
        values = [
            None, product.article_code, product.style_code, product.description,
            float(product.quantity) if product.quantity is not None else None,
            product.uom or "",
            float(product.unit_price) if product.unit_price is not None else None,
            float(product.line_total) if product.line_total is not None else None,
            float(product.retail_price) if product.retail_price is not None else None,
            float(product.expected_unit_price) if product.expected_unit_price is not None else None,
            float(product.price_difference) if product.price_difference is not None else None,
            product.validation_status,
            f'=IF(OR(E{row}="",G{row}=""),"",E{row}*G{row})',
            f'=IF(OR(E{row}="",J{row}=""),"",E{row}*J{row})',
            float(product.discount_percent),
            product.colour, product.page, product.photo_filename,
            "; ".join(product.issues + product.notes),
        ]
        for col_index, value in enumerate(values, start=1):
            cell = ws.cell(row, col_index, sanitize_excel_text(value))
            cell.alignment = Alignment(vertical="center", wrap_text=col_index in (4, 12, 19))
            cell.border = Border(bottom=Side(style="hair", color=BORDER))
        ws.row_dimensions[row].height = 62
        ws.cell(row, 5).number_format = "#,##0.##"
        for col in (7, 8, 9, 10, 11, 13, 14):
            ws.cell(row, col).number_format = '#,##0.00;[Red]-#,##0.00'
        ws.cell(row, 12).fill = PatternFill("solid", fgColor=_status_fill(product.validation_status))
        if product.photo_status in ("REVIEW", "MISSING"):
            ws.cell(row, 1).fill = PatternFill("solid", fgColor=YELLOW)
        if product.photo_status in ("MATCHED", "MULTIPLE") and product.photo_path:
            if Path(product.photo_path).is_file():
                image_cells["Order Details"].append({"cell": f"A{row}", "image": product.photo_path})
    widths = {
        "A": 18, "B": 16, "C": 20, "D": 36, "E": 16, "F": 12, "G": 15, "H": 17,
        "I": 15, "J": 22, "K": 16, "L": 19, "M": 19, "N": 20, "O": 16,
        "P": 18, "Q": 12, "R": 24, "S": 52,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    end_row = max(header_row, first_data + len(order.products) - 1)
    _format_table(ws, header_row, end_row, "OrderProducts")
    ws.freeze_panes = "B8"
    if order.products:
        end = first_data + len(order.products) - 1
        for status, colour in (("VERIFIED", GREEN), ("DISCREPANCY", RED), ("REVIEW", YELLOW)):
            ws.conditional_formatting.add(
                f"L{first_data}:L{end}",
                FormulaRule(formula=[f'L{first_data}="{status}"'],
                            fill=PatternFill("solid", fgColor=colour))
            )

    summary = workbook.create_sheet("Order Summary")
    _base_sheet(summary)
    summary.merge_cells("A1:D1")
    summary["A1"] = f"ORDER {order.order_number} - SUMMARY"
    summary["A1"].font = Font(size=16, bold=True, color=WHITE)
    summary["A1"].fill = PatternFill("solid", fgColor=NAVY)
    start = 3
    summary_values = [
        ("Order Number", order.order_number),
        ("Order Date", order.order_date),
        ("Customer", order.customer_name),
        ("Source PDF", Path(order.source_pdf).name),
        ("Number of Product Lines", f"=COUNTA('Order Details'!B{first_data}:B{first_data + len(order.products)-1})"),
        ("Total Ordered Quantity", f"=SUM('Order Details'!E{first_data}:E{first_data + len(order.products)-1})"),
        ("Extracted Line Value", f"=SUM('Order Details'!H{first_data}:H{first_data + len(order.products)-1})"),
        ("Retail Discount (%)", float(order.products[0].discount_percent) if order.products else 60),
        ("Estimated Savings vs Retail", f"=SUMPRODUCT('Order Details'!I{first_data}:I{first_data + len(order.products)-1}-'Order Details'!J{first_data}:J{first_data + len(order.products)-1},'Order Details'!E{first_data}:E{first_data + len(order.products)-1})"),
        ("Expected Bulk Purchase Value", f"=SUM('Order Details'!N{first_data}:N{first_data + len(order.products)-1})"),
        ("Total VAT", "Not present in source PDF"),
        ("Document Total", float(order.document_total) if order.document_total is not None else None),
        ("Validation Status", order.validation_status),
        ("Lines Requiring Review", sum(p.validation_status == "REVIEW" for p in order.products)),
        ("Lines With Discrepancies", sum(p.validation_status == "DISCREPANCY" for p in order.products)),
    ]
    for row, (label, value) in enumerate(summary_values, start=start):
        summary.cell(row, 1, label).font = Font(bold=True, color=NAVY)
        summary.cell(row, 1).fill = PatternFill("solid", fgColor=PALE_BLUE)
        summary.cell(row, 2, sanitize_excel_text(value))
        if "total" in label.lower() or "value" in label.lower():
            summary.cell(row, 2).number_format = '#,##0.00;[Red]-#,##0.00'
    summary.column_dimensions["A"].width = 32
    summary.column_dimensions["B"].width = 58
    if order.order_issues or order.extraction_warnings:
        audit_row = start + len(summary_values) + 2
        summary.cell(audit_row, 1, "Audit Notes").font = Font(bold=True, color=WHITE)
        summary.cell(audit_row, 1).fill = PatternFill("solid", fgColor=NAVY)
        for row, note in enumerate(order.order_issues + order.extraction_warnings, start=audit_row + 1):
            summary.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
            summary.cell(row, 1, sanitize_excel_text(note)).alignment = Alignment(wrap_text=True, vertical="top")

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.save(target)
    _verify_workbook(target, {"Order Details", "Order Summary"})
    _finalize_with_native_images(target, image_cells)
    order.order_workbook = str(target)
    order.conversion_status = "SUCCESS"
    return str(target)


def build_validation_report(orders: list[OrderData], output_dir: str | Path, tolerance: Decimal):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    target = unique_path(output_dir / f"Validation_Report_{datetime.now():%Y-%m-%d}.xlsx")
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Executive Summary"
    _base_sheet(summary, "A3")
    products = [product for order in orders for product in order.products]
    price_issues = [p for p in products if any(i.startswith("PRICE_") for i in p.issues)]
    quantity_issues = [p for p in products if any(i.startswith(("QUANTITY_", "KIT_QTY_")) for i in p.issues)]
    unmatched = [p for p in products if any(i.startswith(("UNMATCHED_PRODUCT", "AMBIGUOUS_PRODUCT_MATCH", "OFFICIAL_PRICE_MISSING")) for i in p.issues)]
    calculation = [p for p in products if any(i.startswith("LINE_TOTAL_") for i in p.issues)]
    discrepancy_value = sum(
        (abs(p.price_difference * p.quantity) for p in price_issues
         if p.price_difference is not None and p.quantity is not None), Decimal("0")
    ) + sum(
        (abs(p.line_total_difference) for p in calculation if p.line_total_difference is not None),
        Decimal("0")
    )
    summary.merge_cells("A1:D1")
    summary["A1"] = "ORDER VALIDATION - EXECUTIVE SUMMARY"
    summary["A1"].font = Font(size=16, bold=True, color=WHITE)
    summary["A1"].fill = PatternFill("solid", fgColor=NAVY)
    summary.append(["Metric", "Result"])
    _style_header(summary, 2, 1, 2)
    kpis = [
        ("PDF files processed", len(orders)),
        ("Successfully converted orders", sum(o.conversion_status == "SUCCESS" for o in orders)),
        ("Orders with errors", sum(o.validation_status == "ERRORS" for o in orders)),
        ("Orders requiring manual review", sum(o.validation_status == "REVIEW" for o in orders)),
        ("Products checked", len(products)),
        ("Products with price discrepancies", len(price_issues)),
        ("Products with quantity discrepancies", len(quantity_issues)),
        ("Unmatched / ambiguous products", len(unmatched)),
        ("Calculation errors", len(calculation)),
        ("Total value of detected discrepancies", float(discrepancy_value)),
        ("Rounding tolerance", float(tolerance)),
        ("Report generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]
    for row, (label, value) in enumerate(kpis, start=3):
        summary.cell(row, 1, label).font = Font(bold=True, color=NAVY)
        summary.cell(row, 1).fill = PatternFill("solid", fgColor=PALE_BLUE)
        summary.cell(row, 2, value)
    summary.column_dimensions["A"].width = 44
    summary.column_dimensions["B"].width = 26

    detailed = workbook.create_sheet("Detailed Validation")
    headers = [
        "Source PDF", "Order Number", "Order Date", "Page", "Product Code", "Product Description",
        "Extracted Quantity", "Extracted Unit Price", "Official Unit Price", "Price Difference",
        "Extracted Line Total", "Expected Line Total", "Quantity Issue", "Kit Q.ty",
        "Validation Status", "Explanation / Issues", "Photo Status", "Photo Filename",
    ]
    detailed.append(headers)
    for order in orders:
        for product in order.products:
            detailed.append([
                Path(order.source_pdf).name, order.order_number, order.order_date, product.page,
                product.sku, sanitize_excel_text(product.description),
                float(product.quantity) if product.quantity is not None else None,
                float(product.unit_price) if product.unit_price is not None else None,
                float(product.official_price) if product.official_price is not None else None,
                float(product.price_difference) if product.price_difference is not None else None,
                float(product.line_total) if product.line_total is not None else None,
                float(product.expected_line_total) if product.expected_line_total is not None else None,
                "; ".join(i for i in product.issues if i.startswith(("QUANTITY_", "KIT_QTY_"))),
                float(product.kit_qty) if product.kit_qty is not None else None,
                product.validation_status, "; ".join(product.issues + product.notes),
                product.photo_status, product.photo_filename,
            ])
    _format_table(detailed, 1, detailed.max_row, "DetailedValidation")
    detailed.freeze_panes = "A2"

    price_ws = workbook.create_sheet("Price Discrepancies")
    price_ws.append(["Source PDF", "Order Number", "Product Code", "Description", "Quantity",
                     "Extracted Price", "Official Price", "Unit Difference", "Total Exposure",
                     "Status", "Explanation"])
    for order in orders:
        for product in order.products:
            if product.price_difference is not None and abs(product.price_difference) > tolerance:
                exposure = abs(product.price_difference * product.quantity) if product.quantity is not None else None
                price_ws.append([
                    Path(order.source_pdf).name, order.order_number, product.sku, product.description,
                    float(product.quantity) if product.quantity is not None else None,
                    float(product.unit_price) if product.unit_price is not None else None,
                    float(product.official_price) if product.official_price is not None else None,
                    float(product.price_difference), float(exposure) if exposure is not None else None,
                    product.validation_status, "; ".join(i for i in product.issues if i.startswith("PRICE_")),
                ])
    _format_table(price_ws, 1, price_ws.max_row, "PriceDiscrepancies")

    qty_ws = workbook.create_sheet("Quantity Discrepancies")
    qty_ws.append(["Source PDF", "Order Number", "Product Code", "Description",
                   "Extracted Quantity", "Kit Q.ty", "Quantity Issue", "Status"])
    for order in orders:
        for product in order.products:
            issues = [i for i in product.issues if i.startswith(("QUANTITY_", "KIT_QTY_"))]
            if issues:
                qty_ws.append([
                    Path(order.source_pdf).name, order.order_number, product.sku, product.description,
                    float(product.quantity) if product.quantity is not None else None,
                    float(product.kit_qty) if product.kit_qty is not None else None,
                    "; ".join(issues), product.validation_status,
                ])
    _format_table(qty_ws, 1, qty_ws.max_row, "QuantityDiscrepancies")

    unmatched_ws = workbook.create_sheet("Missing and Unmatched Products")
    unmatched_ws.append(["Source PDF", "Order Number", "Page", "Extracted Product Code",
                         "Description", "Match Method", "Confidence", "Issue"])
    for order in orders:
        for product in order.products:
            issues = [i for i in product.issues if i.startswith(
                ("UNMATCHED_PRODUCT", "AMBIGUOUS_PRODUCT_MATCH", "OFFICIAL_PRICE_MISSING")
            )]
            if issues:
                unmatched_ws.append([
                    Path(order.source_pdf).name, order.order_number, product.page, product.sku,
                    product.description, product.match_method, product.match_confidence, "; ".join(issues),
                ])
    _format_table(unmatched_ws, 1, unmatched_ws.max_row, "UnmatchedProducts")

    calc_ws = workbook.create_sheet("Calculation Errors")
    calc_ws.append(["Source PDF", "Order Number", "Page", "Product Code", "Description",
                    "Quantity", "Unit Price", "Extracted Line Total", "Calculated Line Total",
                    "Difference", "Issue"])
    for order in orders:
        for product in order.products:
            issues = [i for i in product.issues if i.startswith("LINE_TOTAL_")]
            if issues:
                calc_ws.append([
                    Path(order.source_pdf).name, order.order_number, product.page, product.sku,
                    product.description, float(product.quantity) if product.quantity is not None else None,
                    float(product.unit_price) if product.unit_price is not None else None,
                    float(product.line_total) if product.line_total is not None else None,
                    float(product.calculated_line_total) if product.calculated_line_total is not None else None,
                    float(product.line_total_difference) if product.line_total_difference is not None else None,
                    "; ".join(issues),
                ])
        for issue in order.order_issues:
            calc_ws.append([Path(order.source_pdf).name, order.order_number, "", "", "Order-level total",
                            "", "", float(order.document_total) if order.document_total is not None else None,
                            "", "", issue])
    _format_table(calc_ws, 1, calc_ws.max_row, "CalculationErrors")

    for ws in workbook.worksheets:
        ws.sheet_view.showGridLines = False
        for column in ws.columns:
            letter = column[0].column_letter
            max_length = max((len(str(cell.value)) for cell in column if cell.value is not None), default=10)
            ws.column_dimensions[letter].width = min(max(max_length + 2, 12), 44)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = Border(bottom=Side(style="hair", color=BORDER))
        ws.row_dimensions[1].height = 34
    for row in range(2, detailed.max_row + 1):
        detailed.cell(row, 15).fill = PatternFill("solid", fgColor=_status_fill(detailed.cell(row, 15).value))
    workbook.save(target)
    _verify_workbook(target, {
        "Executive Summary", "Detailed Validation", "Price Discrepancies",
        "Quantity Discrepancies", "Missing and Unmatched Products", "Calculation Errors",
    })
    return str(target)


def build_photo_library(photo_records: list[dict], output_dir: str | Path) -> str:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    target = unique_path(output_dir / "Product_Photo_Library.xlsx")
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Product Photo Library"
    _base_sheet(ws, "A2")
    headers = [
        "Article Code", "Style Code", "Product Description", "Embedded Photo",
        "Extracted Image Filename", "Original PDF", "Extraction Status",
        "Image Verification Status", "Source Page", "Image Path",
    ]
    ws.append(headers)
    _style_header(ws, 1, 1, len(headers))
    image_cells = {"Product Photo Library": []}
    for row, record in enumerate(photo_records, start=2):
        ws.append([
            record.get("article_code") or record.get("sku"), record.get("style_code"),
            sanitize_excel_text(record.get("description")), None,
            record.get("filename"), record.get("source_pdf"), record.get("extraction_status"),
            record.get("verification_status"), record.get("page"), record.get("image_path"),
        ])
        ws.row_dimensions[row].height = 65
        ws.cell(row, 7).fill = PatternFill("solid", fgColor=_status_fill(record.get("extraction_status")))
        ws.cell(row, 8).fill = PatternFill("solid", fgColor=_status_fill(record.get("verification_status")))
        if record.get("image_path") and Path(record["image_path"]).is_file():
            image_cells["Product Photo Library"].append({"cell": f"D{row}", "image": record["image_path"]})
    widths = {"A": 18, "B": 20, "C": 36, "D": 18, "E": 28, "F": 34, "G": 20, "H": 24, "I": 12, "J": 64}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    _format_table(ws, 1, ws.max_row, "ProductPhotoLibrary")
    workbook.save(target)
    _verify_workbook(target, {"Product Photo Library"})
    _finalize_with_native_images(target, image_cells)
    return str(target)
