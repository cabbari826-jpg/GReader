# PDF Order Converter & Validator

A local-first Windows desktop application for extracting PDF product orders, checking them against official catalogues, exporting Excel files, and maintaining a reusable product-photo library.

**Target:** Windows 10/11 x64 · **Python:** 3.11 · **GUI:** PySide6

## Core capabilities

- Batch import of multi-page PDF orders.
- Coordinate-based table extraction and a local Tesseract OCR fallback.
- Extraction of product codes, descriptions, quantities, prices, totals, and product photos.
- SKU-first catalogue matching, cautious description matching, and manual review for ambiguous associations.
- Decimal-based price, arithmetic, line-total, order-total, and optional pack-size checks.
- Separate order workbooks, six-sheet validation report, product photo library, JSON audit and local SQLite history.
- Local processing only. Order files and catalogues are not uploaded to external services.

## Critical photo requirement: native Excel in-cell pictures

Product photographs must be **native Excel Picture in Cell values**, not pictures floating over or anchored to worksheet cells.

The export pipeline uses Excel's native *Place in Cell* operation on Windows after Python creates the workbook and embeds the local image bytes. It then checks the OOXML rich-value metadata and the expected photo cells. It does not report success or quietly fall back to floating drawings when native in-cell conversion fails.

**Runtime requirement for photo workbooks:** Windows with desktop Microsoft Excel that supports Place in Cell (Microsoft 365 or Excel 2024) and the Python `pywin32` package. Microsoft documents that native cell pictures can be sorted and filtered with worksheet data: [Microsoft Support](https://support.microsoft.com/en-us/excel/insert-picture-in-cell-in-excel).

This dependency is needed only for creating workbooks that contain product photos. The application must explain the compatibility requirement before processing if Excel is unavailable.

See [docs/EXCEL_IN_CELL_IMAGES.md](docs/EXCEL_IN_CELL_IMAGES.md) for the technical approach and acceptance tests.

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

## Build the Windows installer

Install Python 3.11 x64 and Inno Setup 6, then run `BUILD_WINDOWS.bat`. Compile `installer.iss` with Inno Setup. The project also contains a GitHub Actions release workflow to build an installer from a `v*` tag.

## Accuracy and audit rules

- SKU is the primary matching key. Similar descriptions alone do not establish a product match.
- Original extracted values are preserved; the application reports discrepancies rather than correcting prices or quantities.
- Money uses Python `Decimal`.
- Missing discounts, VAT, or other unavailable fields are not invented.
- All uncertain product and image associations are flagged for review.
- Real orders, catalogues, photos, audit logs, and generated business reports must never be committed to this repository.

## Current status

The repository is being assembled from the sample-tested processing implementation. Native in-cell image conversion is now a hard design requirement. The native Excel automation and OOXML verification path still requires Windows acceptance testing on a machine with supported desktop Excel. Do not treat this repository as an accounting-approved production release until the Windows build, in-cell sorting/filtering test, and regression tests pass.

