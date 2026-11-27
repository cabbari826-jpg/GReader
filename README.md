# PDF Order Converter & Validator

A local-first Windows desktop application for extracting product orders from PDF documents, validating them against official price lists, exporting Excel workbooks, and maintaining a reusable product-photo library.

**Target:** Windows 10/11 x64 · **Runtime:** Python 3.11 · **GUI:** PySide6

## Features

- Batch import of PDF orders, including multi-page documents.
- Coordinate-aware extraction of order details, product rows, quantities, prices, totals, and embedded product images.
- Tesseract OCR fallback for scanned pages (requires a local Tesseract installation).
- SKU-first catalogue matching, with cautious description matching when a SKU is absent.
- Price, quantity, arithmetic, order-total, and optional pack-multiple checks.
- Separate XLSX output per order, a six-sheet validation report, and an embedded-image product-photo library.
- Manual review for uncertain product matches and photo associations.
- Local SQLite processing history and JSON audit files.
- Local processing only: the application does not upload business files to external servers.

## Quick start (Windows)

```powershell
py -3.11 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python app.py
```

Run tests:

```powershell
pip install -r requirements-dev.txt
python -m compileall -q app.py ordertool tests
python -m pytest -q
```

## Build a Windows installer

Install Python 3.11 x64 and Inno Setup 6, then run `BUILD_WINDOWS.bat`. The script creates a standalone application under `dist/PDFOrderValidator`. Compile `installer.iss` with Inno Setup to produce the installer. GitHub Actions can also build the installer when a `v*` tag is pushed.

## Repository structure

```text
app.py
ordertool/              # extraction, catalogue matching, validation, exports, GUI
 tests/                 # automated tests
 docs/                  # architecture notes
 .github/workflows/     # CI and Windows release automation
 BUILD_WINDOWS.bat
 installer.iss
```

## Accuracy and data handling

- SKU is the primary product key; uncertain matches require review.
- Original extracted values are preserved. The application reports discrepancies rather than silently correcting source values.
- Monetary comparisons use Python `Decimal`.
- Missing information such as discounts or VAT is not invented.
- Real orders, price lists, photos, audit logs, and generated reports must stay out of version control.

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the processing design, and [CONTRIBUTING.md](CONTRIBUTING.md) for development checks.

## Known limitations

The first extraction adapter is template-oriented. A substantially different PDF layout may require an adapter update. Product images are embedded XLSX drawing objects anchored to their rows; they are not Microsoft's newer native “Place in Cell” images, so sorting/filtering must be verified in the target Excel version. OCR results and uncertain image links require manual review.

The project has been exercised against private sample documents, which are intentionally not included in this public repository. Review the Windows CI results and perform user acceptance testing with representative files before accounting use.
