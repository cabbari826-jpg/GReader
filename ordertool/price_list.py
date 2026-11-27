from __future__ import annotations
from pathlib import Path
import csv
import re

from .models import PriceRecord
from .utils import to_decimal, SKU_RE, money_token, configure_tesseract

HEADER_SYNONYMS = {
    "sku": {"stylecode", "sku", "productcode", "model", "models", "itemcode", "code", "article"},
    "description": {"styledescription", "description", "productdescription", "productname", "name"},
    "price": {"retailprice", "unitprice", "officialprice", "listprice", "price"},
    "kit_qty": {"kitqty", "kitquantity", "packqty", "packquantity", "minimumorderqty"},
    "catalog_qty": {"quantity", "catalogquantity", "availablequantity"},
    "article_code": {"articlecode", "ean", "barcode"},
}


def _norm_header(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def _find_header_map(values):
    best = None
    for row_index, row in enumerate(values[:30]):
        found = {}
        for col_index, value in enumerate(row):
            key = _norm_header(value)
            for field, aliases in HEADER_SYNONYMS.items():
                if key in aliases and field not in found:
                    found[field] = col_index
        score = sum(k in found for k in ("sku", "price")) + .2 * sum(
            k in found for k in ("description", "kit_qty", "catalog_qty", "article_code")
        )
        if best is None or score > best[0]:
            best = (score, row_index, found)
    if not best or not {"sku", "price"}.issubset(best[2]):
        return None
    return best[1], best[2]


def load_price_list(path):
    paths = list(path) if isinstance(path, (list, tuple)) else [path]
    if not paths:
        raise ValueError("Select at least one official price-list file.")
    combined = {}
    for item in paths:
        source = Path(item)
        suffix = source.suffix.lower()
        if suffix in (".xlsx", ".xlsm"):
            records = _load_excel(source)
        elif suffix == ".csv":
            records = _load_csv(source)
        elif suffix == ".pdf":
            records = _load_pdf(source)
        else:
            raise ValueError(f"Unsupported price-list format: {source.name}")
        for sku, record in records.items():
            prior = combined.get(sku)
            if prior and prior.official_price != record.official_price:
                raise ValueError(
                    f"Conflicting official prices for {sku} in {prior.source_file} and {record.source_file}"
                )
            combined.setdefault(sku, record)
    return combined


def _record(sku, get, source, extra=None):
    return PriceRecord(
        sku=sku,
        description=str(get("description") or "").strip(),
        official_price=to_decimal(get("price")),
        kit_qty=to_decimal(get("kit_qty")),
        catalog_qty=to_decimal(get("catalog_qty")),
        article_code=str(get("article_code") or "").strip(),
        source_file=source.name,
        raw=extra or {},
    )


def _load_excel(path):
    from openpyxl import load_workbook
    workbook = load_workbook(path, read_only=True, data_only=True)
    result = {}
    try:
        for worksheet in workbook.worksheets:
            values = list(worksheet.iter_rows(
                min_row=1, max_row=min(worksheet.max_row, 10000), values_only=True
            ))
            header = _find_header_map(values)
            if not header:
                continue
            header_row, columns = header
            for line_no, row in enumerate(values[header_row + 1:], start=header_row + 2):
                raw_sku = row[columns["sku"]] if len(row) > columns["sku"] else None
                sku = str(raw_sku or "").strip()
                if not sku or not SKU_RE.fullmatch(sku):
                    continue
                def get(field):
                    index = columns.get(field)
                    return row[index] if index is not None and index < len(row) else None
                result[sku.upper()] = _record(
                    sku, get, path, {"sheet": worksheet.title, "row": line_no}
                )
    finally:
        workbook.close()
    if not result:
        raise ValueError(
            "No SKU/price rows found. Expected headings such as Style Code and Retail Price."
        )
    return result


def _load_csv(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.reader(stream))
    header = _find_header_map(rows)
    if not header:
        raise ValueError("CSV must contain recognisable SKU and official price headings.")
    header_row, columns = header
    result = {}
    for row in rows[header_row + 1:]:
        sku = str(row[columns["sku"]] if len(row) > columns["sku"] else "").strip()
        if not sku or not SKU_RE.fullmatch(sku):
            continue
        def get(field):
            index = columns.get(field)
            return row[index] if index is not None and index < len(row) else None
        result[sku.upper()] = _record(sku, get, path)
    if not result:
        raise ValueError("No valid product rows found in CSV price list.")
    return result


def _load_pdf(path):
    import fitz
    document = fitz.open(path)
    result = {}
    try:
        for page in document:
            words = page.get_text("words", sort=True)
            for code in (word for word in words if SKU_RE.fullmatch(word[4])):
                center_y = (code[1] + code[3]) / 2
                line = [word for word in words
                        if abs((word[1] + word[3]) / 2 - center_y) < 7 and word[0] > code[0]]
                prices = []
                for word in line:
                    token, _ = money_token(word[4])
                    if token:
                        prices.append((word[0], token))
                if not prices:
                    continue
                price_token = sorted(prices, key=lambda item: item[0])[-1][1]
                description = " ".join(
                    word[4] for word in line
                    if word[0] < sorted(prices, key=lambda item: item[0])[-1][0]
                )
                result[code[4].upper()] = PriceRecord(
                    code[4], description, to_decimal(price_token), source_file=path.name
                )
        if not result:
            try:
                import io
                import pytesseract
                from PIL import Image
                pytesseract = configure_tesseract(pytesseract)
                for page in document:
                    pixmap = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
                    image = Image.open(io.BytesIO(pixmap.tobytes("png")))
                    data = pytesseract.image_to_data(
                        image, output_type=pytesseract.Output.DICT, config="--psm 6"
                    )
                    lines = {}
                    for index, token in enumerate(data.get("text", [])):
                        token = str(token).strip()
                        try:
                            confidence = float(data["conf"][index])
                        except Exception:
                            confidence = -1
                        if not token or confidence < 35:
                            continue
                        key = (data["block_num"][index], data["par_num"][index], data["line_num"][index])
                        lines.setdefault(key, []).append((data["left"][index], token))
                    for words_in_line in lines.values():
                        words_in_line.sort(key=lambda pair: pair[0])
                        text_line = " ".join(token for _, token in words_in_line)
                        match = re.search(r"[A-Z]\d{6,10}-\d{3,6}", text_line, re.I)
                        if not match:
                            continue
                        sku = match.group(0)
                        price_candidates = []
                        for x, token in words_in_line:
                            value, _ = money_token(token)
                            if value:
                                price_candidates.append((x, value))
                        if not price_candidates:
                            continue
                        price = to_decimal(sorted(price_candidates)[-1][1])
                        desc = " ".join(token for _, token in words_in_line if token != sku)
                        result[sku.upper()] = PriceRecord(sku, desc, price, source_file=path.name)
            except Exception:
                pass
    finally:
        document.close()
    if not result:
        raise ValueError(
            "No SKU/price rows could be read from this PDF. Convert the list to XLSX "
            "or install Tesseract for scanned PDF catalogues."
        )
    return result
