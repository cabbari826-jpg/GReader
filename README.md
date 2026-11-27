# PDF Order Converter & Validator

A local-first Windows desktop application for extracting product orders from PDF files, validating prices and quantities against official catalogues, exporting Excel workbooks, and maintaining a reusable product-photo library.

**Target:** Windows 10/11 x64 · **Python:** 3.11 · **GUI:** PySide6

## Product identifiers

The PDF contains a **Style Code** (for example, `S2620155-0115`). The official catalogue contains both the Style Code and the numeric **Article Code** (for example, `6814388`). The application links those two fields, but the Article Code is the canonical product identifier after matching.

- Validation reports show both Article Code and Style Code.
- Product-photo filenames use the official Article Code, e.g. `6814388.jpg`.
- The original PDF Style Code remains available for tracing the source row.
- A photo is renamed to an Article Code only after the row and photo association are considered reliable.

## Bulk purchase discount

The catalogue's `Retail Price` is the undiscounted reference. The sample order PDFs also print that Retail Price (for example, 49.90), so PDF price validation compares the extracted PDF price to catalogue Retail Price. This checks the source document's price consistency and avoids falsely flagging every sample line.

Separately, the default negotiated bulk discount is **60% off Retail Price**, so the expected purchase price is 40% of Retail Price:

`Expected purchase price = Retail Price × (1 − discount percentage / 100)`

For example, a Retail Price of 49.90 with a 60% discount produces an estimated bulk purchase price of 19.96.

The discount is editable in the application (0–100%) and remembered for the next run. The source PDF unit price is never changed. The report shows catalogue Retail Price, PDF price difference from retail, discount percentage, expected net unit price, extracted line total, and expected discounted purchase value. The PDF line-total arithmetic is checked against the PDF's own unit price and quantity; the discounted value is shown separately as a cost estimate.

## Features

- Batch import of multi-page PDF orders by file picker or drag-and-drop.
- Coordinate-based extraction using PyMuPDF; local Tesseract OCR fallback for scanned pages.
- Extraction of order fields, Style Codes, descriptions, quantities, prices, totals, colours, and product photographs.
- Catalogue import from XLSX, XLSM, CSV, or supported text-based PDF lists.
- Article Code matching through the catalogue's Style Code mapping, with conservative manual review for unmatched data.
- Decimal-based price, arithmetic, order-total, and configurable pack-multiple checks.
- One order workbook per PDF, a six-sheet validation report, and a reusable product-photo library.
- Native Excel **Picture in Cell** images embedded as rich values, not drawing-layer/floating pictures.
- Manual product-photo review, JSON audit records, and local SQLite processing history.
- In-app update checking and installer download with SHA-256 verification.
- Local document processing. PDFs, price lists, photos, audit records, and reports are not uploaded.

## Native Excel in-cell pictures

The export pipeline writes Excel's native rich-value image representation directly into the XLSX package. Images are stored in the workbook under `xl/media`; worksheet cells refer to them through `vm` metadata and rich-value relationships.

There is no Excel COM automation, no external image URL, and no floating-image fallback. The export verifies that each expected photo cell is bound to a rich-value record and that the image bytes are embedded. Do not open and save the resulting workbook with a library that does not preserve Excel rich values.

Open image-bearing workbooks in a version of Microsoft Excel that supports **Place in Cell**. Sort and filter the table by Article Code during acceptance testing to confirm that images remain with their product records. See [docs/EXCEL_IN_CELL_IMAGES.md](docs/EXCEL_IN_CELL_IMAGES.md).

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

## In-app updates

Use **Check for updates** in the application header. The app checks the latest stable GitHub Release, downloads the installer and its SHA-256 file, verifies the installer, then asks before launching the installer.

To publish an installer that the app can discover:

1. Open the repository's [Windows installer workflow](https://github.com/cabbari826-jpg/GReader/actions/workflows/windows-release.yml).
2. Click **Run workflow** and keep branch `main`.
3. Tick **Publish this installer as a stable GitHub Release for in-app updates**.
4. Start the workflow. When its tests and Windows build succeed, it attaches the installer and SHA-256 checksum to a versioned GitHub Release.
5. Future app launches can find that release through **Check for updates**.

You can also push a version tag such as `v1.0.2`; the same workflow builds and publishes it. A manual run without the publish option creates only an Actions artifact and will not make the updater see the installer.

The updater contacts GitHub only to check release metadata and download the installer. It does not transmit order PDFs, catalogues, customer data, or extracted photos. The installer checksum is verified before the app offers to launch it.

## Build the Windows installer

Install Python 3.11 x64 and Inno Setup 6, then run `BUILD_WINDOWS.bat`. Alternatively, use the GitHub Actions **Windows installer** workflow. See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) and [CONTRIBUTING.md](CONTRIBUTING.md).

## Accuracy and audit rules

- Article Code is canonical; Style Code is retained as the PDF-to-catalogue lookup reference.
- Unmatched or ambiguous catalogue mappings and uncertain photo links require review.
- Original extracted values are preserved. The application reports differences; it does not silently modify source prices or quantities.
- Retail Price is not treated as the purchase price when a negotiated discount is configured.
- Money uses Python `Decimal`.
- Missing discounts, VAT, or fields not present in the source are not invented.
- Real business orders, price lists, photos, audit logs, and generated workbooks must never be committed to this repository.

## Current acceptance status

The GitHub Actions checks cover source compilation, unit tests, and PySide6 window construction. They do not provide an activated Microsoft Excel desktop instance. Before accounting use, open an exported workbook in the target Excel version, sort/filter by Article Code, and verify that each native in-cell image stays with its product row. Also run a batch against current production catalogues and reconcile cover totals.
