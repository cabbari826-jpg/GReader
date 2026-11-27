from __future__ import annotations
import re, io, statistics
from pathlib import Path

import fitz
from PIL import Image

from .models import OrderData, ProductLine
from .utils import SKU_RE, money_token, to_decimal, sha256_file, parse_date, safe_name, configure_tesseract

SKU_TEXT_RE = re.compile(r"^[A-Z]\d{6,10}-\d{3,6}$", re.I)


def _words_for_page(page, allow_ocr=True):
    words = page.get_text("words", sort=True)
    if words and any(SKU_TEXT_RE.fullmatch(w[4].strip()) for w in words):
        return words, False
    if not allow_ocr:
        return words, False
    try:
        import pytesseract
        pytesseract = configure_tesseract(pytesseract)
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        image = Image.open(io.BytesIO(pix.tobytes("png")))
        data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT, config="--psm 6")
        ocr_words = []
        for i, token in enumerate(data.get("text", [])):
            token = str(token).strip()
            try:
                confidence = float(data["conf"][i])
            except Exception:
                confidence = -1
            if not token or confidence < 35:
                continue
            x0, y0 = data["left"][i] / 3, data["top"][i] / 3
            x1 = (data["left"][i] + data["width"][i]) / 3
            y1 = (data["top"][i] + data["height"][i]) / 3
            ocr_words.append((x0, y0, x1, y1, token, 0, i, i))
        return ocr_words, True
    except Exception:
        return words, False


def _chars_for_page(page):
    chars = []
    try:
        raw = page.get_text("rawdict")
        for block in raw.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    for char in span.get("chars", []):
                        text = char.get("c", "")
                        if text and text not in ("\n", "\r"):
                            x0, y0, x1, y1 = char["bbox"]
                            chars.append((x0, y0, x1, y1, text))
    except Exception:
        return []
    return chars


def _parse_row(words, sku_word, page_number, ocr=False, chars=None):
    center_y = (sku_word[1] + sku_word[3]) / 2
    line = [w for w in words if abs((w[1] + w[3]) / 2 - center_y) < 8]
    price = None
    if chars:
        row_chars = [ch for ch in chars if abs((ch[1] + ch[3]) / 2 - center_y) < 8]
        desc_chars = sorted((ch for ch in row_chars if 185 <= ch[0] < 345.5), key=lambda ch: ch[0])
        description = re.sub(r"\s+", " ", "".join(ch[4] for ch in desc_chars)).strip()
        price_chars = sorted(
            (ch for ch in row_chars if 345.5 <= ch[0] < 390 and not ch[4].isspace()),
            key=lambda ch: ch[0],
        )
        price = to_decimal("".join(ch[4] for ch in price_chars))
    else:
        desc_words = sorted((w for w in line if 185 <= w[0] < 345.5), key=lambda w: w[0])
        description = " ".join(w[4] for w in desc_words).strip()
        for word in line:
            if 280 <= word[0] < 395:
                value, prefix = money_token(word[4])
                if value is not None:
                    price = to_decimal(value)
                    if prefix and word[0] < 345.5:
                        description = (description + " " + prefix).strip()
                    break
        description = re.sub(r"\s+", " ", description)

    quantity = None
    for word in line:
        if 380 <= word[0] < 430 and re.fullmatch(r"\d+(?:[.,]\d+)?", word[4]):
            quantity = to_decimal(word[4])
            break
    line_total = None
    for word in line:
        if 425 <= word[0] < 478:
            value, _ = money_token(word[4])
            if value is not None:
                line_total = to_decimal(value)
                break
    colour = re.sub(r"\s+", " ", " ".join(w[4] for w in line if 478 <= w[0] < 575)).strip()
    # The PDF contains Style Code. Article Code is resolved from the official catalogue later.
    product = ProductLine(
        sku="", article_code="", style_code=sku_word[4].strip(),
        description=description, quantity=quantity,
        unit_price=price, line_total=line_total, colour=colour, page=page_number,
        photo_status="REVIEW" if ocr else "NOT_FOUND",
        photo_note="OCR-derived row; image association requires manual confirmation" if ocr else "",
        notes=["OCR text extraction used; please verify this row"] if ocr else [],
    )
    if re.search(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", description):
        product.issues.append(
            "TEXT_ENCODING_REVIEW: PDF text layer contains non-printable characters; cleaned in Excel output"
        )
    return product


def _image_bytes(doc, page, info):
    xref = int(info.get("xref") or 0)
    if xref > 0:
        extracted = doc.extract_image(xref)
        return extracted["image"], extracted.get("ext", "png")
    rect = fitz.Rect(info["bbox"])
    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=rect, alpha=False)
    return pix.tobytes("png"), "png"


def _save_photo(image_bytes, ext, sku, page, order_number, photo_dir):
    import hashlib
    photo_dir.mkdir(parents=True, exist_ok=True)
    image = Image.open(io.BytesIO(image_bytes))
    image.load()
    image_format = (image.format or ext or "PNG").upper()
    if image_format in ("JPG", "JPEG"):
        suffix = ".jpg"
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")
    elif image_format == "PNG":
        suffix = ".png"
        if image.mode not in ("RGB", "RGBA", "L"):
            image = image.convert("RGBA")
    else:
        suffix = ".png"
        if image.mode not in ("RGB", "RGBA", "L"):
            image = image.convert("RGBA")
    base = safe_name(sku) if sku else f"UNASSIGNED_{safe_name(order_number)}_P{page:03d}"
    candidate = photo_dir / f"{base}{suffix}"
    digest = hashlib.sha256(image_bytes).hexdigest()
    if candidate.exists():
        try:
            if hashlib.sha256(candidate.read_bytes()).hexdigest() == digest:
                return str(candidate), candidate.name
        except OSError:
            pass
        number = 2
        while True:
            candidate = photo_dir / f"{base}_{number:02d}{suffix}"
            if not candidate.exists():
                break
            try:
                if hashlib.sha256(candidate.read_bytes()).hexdigest() == digest:
                    return str(candidate), candidate.name
            except OSError:
                pass
            number += 1
    if suffix == ".jpg":
        image.save(candidate, format="JPEG", quality=95, optimize=True)
    else:
        image.save(candidate, format="PNG", optimize=True)
    return str(candidate), candidate.name


def _extract_page_photos(doc, page, products, sku_words, ocr, photo_dir, order_number):
    if not products:
        return
    if ocr:
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        full = Image.open(io.BytesIO(pix.tobytes("png")))
        for product, sku_word in zip(products, sku_words):
            y0 = max(0, int((sku_word[1] - 9) * 3))
            y1 = min(full.height, int((sku_word[1] + 55) * 3))
            x0, x1 = int(25 * 3), min(full.width, int(92 * 3))
            if y1 <= y0 or x1 <= x0:
                product.photo_status = "MISSING"
                continue
            crop = full.crop((x0, y0, x1, y1))
            if crop.width < 30 or crop.height < 30:
                product.photo_status = "MISSING"
                continue
            buffer = io.BytesIO()
            crop.save(buffer, format="PNG")
            product.photo_path, product.photo_filename = _save_photo(
                buffer.getvalue(), "png", "", product.page, order_number, photo_dir
            )
            product.photo_paths = [product.photo_path]
            product.photo_filenames = [product.photo_filename]
            product.photo_status = "REVIEW"
            product.photo_note = "Photo cropped from scanned page; verify SKU association manually"
        return

    infos = page.get_image_info(xrefs=True)
    image_infos = [
        info for info in infos
        if info["bbox"][0] < 112
        and 25 <= info["bbox"][2] - info["bbox"][0] <= 105
        and 25 <= info["bbox"][3] - info["bbox"][1] <= 105
    ]
    if not image_infos:
        for product in products:
            product.photo_status = "MISSING"
            product.photo_note = "No product image found in the row photo column"
        return

    centers = [(word[1] + word[3]) / 2 for word in sku_words]
    offsets = []
    for info in image_infos:
        image_y = (info["bbox"][1] + info["bbox"][3]) / 2
        nearest = min(centers, key=lambda center: abs(image_y - center))
        offsets.append(image_y - nearest)
    expected_offset = statistics.median(offsets) if offsets else 20.0
    matches = {}
    for info in image_infos:
        image_y = (info["bbox"][1] + info["bbox"][3]) / 2
        scored = sorted((abs((image_y - center) - expected_offset), index)
                        for index, center in enumerate(centers))
        residual, product_index = scored[0]
        ambiguous = len(scored) > 1 and abs(scored[1][0] - residual) < 3
        matches.setdefault(product_index, []).append((residual + (100 if ambiguous else 0), info))

    for index, product in enumerate(products):
        candidates = sorted(matches.get(index, []), key=lambda item: item[0])
        if not candidates:
            product.photo_status = "MISSING"
            product.photo_note = "No image could be associated with this product row"
            continue
        score, _ = candidates[0]
        if score > 12:
            product.photo_status = "REVIEW"
            product.photo_note = f"Image-to-row association confidence is low (residual {score:.1f} pt)"
            name_sku = ""
        else:
            product.photo_status = "MATCHED" if len(candidates) == 1 else "MULTIPLE"
            product.photo_note = (
                "Position matched to product row" if len(candidates) == 1
                else f"{len(candidates)} images associated with this SKU"
            )
            name_sku = product.sku
        paths, filenames = [], []
        try:
            for image_index, (_, info) in enumerate(candidates, start=1):
                data, ext = _image_bytes(doc, page, info)
                path, filename = _save_photo(data, ext, name_sku, product.page, order_number, photo_dir)
                if path not in paths:
                    paths.append(path)
                    filenames.append(filename)
            if paths:
                product.photo_path, product.photo_filename = paths[0], filenames[0]
                product.photo_paths, product.photo_filenames = paths, filenames
        except Exception as exc:
            product.photo_status = "REVIEW"
            product.photo_note = f"Image extraction failed: {exc}"


def extract_pdf_order(pdf_path: str | Path, photo_dir: str | Path, allow_ocr=True) -> OrderData:
    pdf_path = Path(pdf_path)
    document = fitz.open(pdf_path)
    try:
        first_text = document[0].get_text("text") if len(document) else ""
        number = re.search(r"Document\s+Number\s*:\s*(\d+)", first_text, re.I)
        date = re.search(r"Document\s+Date\s*:\s*(\d{2}[/-]\d{2}[/-]\d{4})", first_text, re.I)
        totals = re.search(
            r"Document\s+Total\s+Quantity\s+Value\s+([\d.,]+)\s+([\d.,]+)",
            first_text, re.I | re.S
        )
        agent = re.search(r"Agent\s*:\s*([A-Za-zÀ-ÿ .'-]+)", first_text, re.I)
        customer = re.search(r"(?im)^\s*(ITALDIZAYN[^\r\n]*)", first_text)
        order_number = number.group(1) if number else pdf_path.stem
        order = OrderData(
            source_pdf=str(pdf_path),
            source_hash=sha256_file(pdf_path),
            order_number=order_number,
            order_date=parse_date(date.group(1)) if date else None,
            customer_name=customer.group(1).strip() if customer else "",
            agent=agent.group(1).strip() if agent else "",
            page_count=len(document),
            document_quantity=to_decimal(totals.group(1)) if totals else None,
            document_total=to_decimal(totals.group(2)) if totals else None,
        )
        if not number:
            order.extraction_warnings.append("Order number could not be read from the cover page")
        if not date:
            order.extraction_warnings.append("Order date could not be read from the cover page")
        if not totals:
            order.extraction_warnings.append("Document quantity/value could not be read from the cover page")

        for page_index, page in enumerate(document):
            words, used_ocr = _words_for_page(page, allow_ocr=allow_ocr)
            sku_words = sorted(
                (word for word in words if SKU_TEXT_RE.fullmatch(str(word[4]).strip())),
                key=lambda word: (word[1], word[0]),
            )
            if not sku_words:
                continue
            page_chars = [] if used_ocr else _chars_for_page(page)
            page_products = [
                _parse_row(words, sku_word, page_index + 1, ocr=used_ocr, chars=page_chars)
                for sku_word in sku_words
            ]
            if used_ocr:
                order.extraction_warnings.append(
                    f"Page {page_index + 1}: OCR was used; extracted text requires review"
                )
            _extract_page_photos(
                document, page, page_products, sku_words, used_ocr, Path(photo_dir), order_number
            )
            order.products.extend(page_products)

        if not order.products:
            raise ValueError(
                "No product rows were detected. The PDF may be encrypted, scanned with OCR "
                "unavailable, or use an unsupported order layout."
            )
        for product in order.products:
            if not product.description:
                product.issues.append("Missing product description in PDF extraction")
            if product.quantity is None:
                product.issues.append("Quantity could not be extracted")
            if product.unit_price is None:
                product.issues.append("Unit price could not be extracted")
            if product.line_total is None:
                product.issues.append("Line total could not be extracted")
            if product.photo_status in ("REVIEW", "MISSING"):
                product.notes.append(product.photo_note or "Product photo requires review")
        return order
    finally:
        document.close()
