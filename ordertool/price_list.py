from __future__ import annotations
from pathlib import Path
import csv
import re

from .models import PriceRecord
from .utils import to_decimal, money_token, configure_tesseract

HEADER_SYNONYMS = {
    "article_code": {"articlecode", "articlenumber", "articleno", "articlenumbercode", "artcode", "article", "ean", "barcode"},
    "style_code": {"stylecode", "style", "model", "models", "modelcode", "productmodel"},
    "legacy_sku": {"sku", "productcode", "itemcode", "code"},
    "description": {"styledescription", "description", "productdescription", "productname", "name"},
    "price": {"retailprice", "unitprice", "officialprice", "listprice", "price"},
    "kit_qty": {"kitqty", "kitquantity", "packqty", "packquantity", "minimumorderqty"},
    "catalog_qty": {"quantity", "catalogquantity", "availablequantity"},
}


def _norm_header(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def normalize_article_code(value) -> str:
    """Normalise numeric Excel Article Code values without a trailing .0."""
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def _find_header_map(values):
    best = None
    for row_index, row in enumerate(values[:30]):
        found = {}
        for col_index, value in enumerate(row):
            key = _norm_header(value)
            for field, aliases in HEADER_SYNONYMS.items():
                if key in aliases and field not in found:
                    found[field] = col_index
        has_identifier = "article_code" in found or "legacy_sku" in found
        if not has_identifier or "price" not in found:
            continue
        score = 3 + .8 * ("article_code" in found) + .5 * ("style_code" in found)
        score += .2 * sum(k in found for k in ("description", "kit_qty", "catalog_qty"))
        if best is None or score > best[0]:
            best = (score, row_index, found)
    return (best[1], best[2]) if best else None


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
        for article, record in records.items():
            prior = combined.get(article)
            if prior and prior.official_price != record.official_price:
                raise ValueError(
                    f"Conflicting Retail Price values for Article Code {article} in "
                    f"{prior.source_file} and {record.source_file}"
                )
            if prior and prior.style_code and record.style_code and prior.style_code != record.style_code:
                raise ValueError(
                    f"Article Code {article} maps to multiple Style Codes: "
                    f"{prior.style_code} and {record.style_code}"
                )
            combined.setdefault(article, record)
    if not combined:
        raise ValueError("No product rows were found in the selected price lists.")
    return combined


def _record(article_code, get, source, extra=None):
    return PriceRecord(
        sku=article_code,
        article_code=article_code,
        style_code=str(get("style_code") or "").strip(),
        description=str(get("description") or "").strip(),
        official_price=to_decimal(get("price")),
        kit_qty=to_decimal(get("kit_qty")),
        catalog_qty=to_decimal(get("catalog_qty")),
        source_file=source.name,
        raw=extra or {},
    )


def _load_tabular(path, rows):
    header = _find_header_map(rows)
    if not header:
        raise ValueError(
            f"{path.name}: expected Article Code and Retail Price columns "
            "(or a legacy SKU and price column)."
        )
    header_row, columns = header
    result = {}
    for line_no, row in enumerate(rows[header_row + 1:], start=header_row + 2):
        def get(field):
            index = columns.get(field)
            return row[index] if index is not None and index < len(row) else None
        article = normalize_article_code(get("article_code"))
        if not article:
            article = normalize_article_code(get("legacy_sku"))
        if not article:
            continue
        record = _record(article, get, path, {"row": line_no})
        previous = result.get(article)
        if previous and previous.official_price != record.official_price:
            raise ValueError(f"Duplicate Article Code {article} has conflicting Retail Prices in {path.name}.")
        result[article] = record
    if not result:
        raise ValueError(f"{path.name}: no valid Article Code / Retail Price rows found.")
    return result


def _load_excel(path):
    from openpyxl import load_workbook
    workbook = load_workbook(path, read_only=True, data_only=True)
    result = {}
    try:
        for worksheet in workbook.worksheets:
            values = list(worksheet.iter_rows(
                min_row=1, max_row=min(worksheet.max_row, 50000), values_only=True
            ))
            try:
                records = _load_tabular(path, values)
            except ValueError:
                continue
            # Preserve worksheet information in each catalogue row's audit metadata.
            # The row number is already retained by _load_tabular.
            for article, record in records.items():
                record.raw["sheet"] = worksheet.title
                previous = result.get(article)
                if previous and previous.official_price != record.official_price:
                    raise ValueError(
                        f"Conflicting Retail Price for Article Code {article} in worksheet "
                        f"{worksheet.title} of {path.name}"
                    )
                result[article] = record
    finally:
        workbook.close()
    if not result:
        raise ValueError(
            f"{path.name}: no rows found. Expected columns such as Style Code, "
            "Style Description, Retail Price, and Article Code."
        )
    return result


def _load_csv(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.reader(stream))
    return _load_tabular(path, rows)


def _load_pdf(path):
    import fitz
    document = fitz.open(path)
    result = {}
    try:
        # Text-based product catalogues: detect Article Code (numeric) and
        # Style Code on the same visual row where the PDF layout supports it.
        for page in document:
            words = page.get_text("words", sort=True)
            style_words = [word for word in words if re.fullmatch(r"[A-Z]\d{6,10}-\d{3,6}", word[4], re.I)]
            for style_word in style_words:
                center_y = (style_word[1] + style_word[3]) / 2
                line = [word for word in words if abs((word[1] + word[3]) / 2 - center_y) < 7]
                article_tokens = [
                    word[4] for word in line
                    if word[0] > style_word[0] and re.fullmatch(r"\d{6,12}", word[4])
                ]
                price_tokens = []
                for word in line:
                    token, _ = money_token(word[4])
                    if token:
                        price_tokens.append((word[0], token))
                if not article_tokens or not price_tokens:
                    continue
                article = normalize_article_code(article_tokens[-1])
                price = to_decimal(sorted(price_tokens, key=lambda item: item[0])[-1][1])
                description = " ".join(word[4] for word in line if style_word[0] < word[0] < 500)
                result[article] = PriceRecord(
                    sku=article, article_code=article, style_code=style_word[4],
                    description=description, official_price=price, source_file=path.name,
                )
        if not result:
            raise ValueError(
                f"{path.name}: the PDF catalogue did not expose recognizable Article Code, "
                "Style Code, and Retail Price fields. Convert it to XLSX for reliable import."
            )
    finally:
        document.close()
    return result
