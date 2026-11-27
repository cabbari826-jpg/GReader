from __future__ import annotations
import hashlib, re, os
from decimal import Decimal, InvalidOperation
from pathlib import Path
from datetime import datetime

SKU_RE = re.compile(r"^[A-Z]\d{6,10}-\d{3,6}$", re.I)
MONEY_RE = re.compile(r"^(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})$|^\d+(?:,\d{2})$")
MONEY_SUFFIX_RE = re.compile(r"^(.*?)(\d+(?:[.,]\d{2}))$")


def to_decimal(value):
    if value is None or value == "":
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip().replace(" ", "")
    if not text:
        return None
    # Recognise both 1,234.56 and 1.234,56; a lone comma is treated as decimal.
    if "," in text and "." in text:
        if text.rfind(".") > text.rfind(","):
            text = text.replace(",", "")
        else:
            text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        if re.fullmatch(r"\d{1,3}(?:,\d{3})+", text):
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def money_token(value):
    """Return (decimal_text, prefix_text) for a price-like token."""
    token = str(value).strip().replace(" ", "")
    if MONEY_RE.fullmatch(token):
        return token, ""
    match = MONEY_SUFFIX_RE.fullmatch(token)
    if match and ("." in match.group(2) or "," in match.group(2)):
        return match.group(2), match.group(1)
    return None, ""


def sha256_file(path: str | Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_name(value: str, fallback="item") -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "").strip())
    value = value.strip("._-")
    return value or fallback


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    n = 2
    while True:
        candidate = path.with_name(f"{stem}_{n}{suffix}")
        if not candidate.exists():
            return candidate
        n += 1


def parse_date(text: str):
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            pass
    return None

def configure_tesseract(pytesseract_module=None):
    """Use a bundled portable Tesseract build when present, otherwise PATH."""
    if pytesseract_module is None:
        import pytesseract as pytesseract_module
    import sys
    candidates = []
    env_path = os.environ.get("TESSERACT_CMD")
    if env_path:
        candidates.append(Path(env_path))
    base = Path(__file__).resolve().parent.parent
    candidates.append(base / "resources" / "tesseract" / "tesseract.exe")
    candidates.append(Path(sys.executable).resolve().parent / "resources" / "tesseract" / "tesseract.exe")
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "resources" / "tesseract" / "tesseract.exe")
    for candidate in candidates:
        if candidate.is_file():
            pytesseract_module.pytesseract.tesseract_cmd = str(candidate)
            tessdata = candidate.parent / "tessdata"
            if tessdata.is_dir():
                os.environ.setdefault("TESSDATA_PREFIX", str(tessdata))
            break
    return pytesseract_module


def sanitize_excel_text(value):
    """Remove XML 1.0 forbidden control characters before writing XLSX cells."""
    if not isinstance(value, str):
        return value
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", value)
