# Architecture

## Processing flow

```text
PDF queue
  -> document adapter (PyMuPDF text coordinates; local Tesseract fallback)
  -> OrderData / ProductLine models
  -> official catalogue loader (XLSX / XLSM / CSV / text PDF)
  -> SKU-first matching and Decimal validation
  -> order XLSX + six-sheet batch validation report
  -> extracted photo library + JSON audit + SQLite history
```

## Modules

- `models.py`: explicit order, line, price-list and batch data structures.
- `extraction.py`: template-aware page parsing and product-photo extraction.
- `price_list.py`: catalogue import and SKU/price field detection.
- `validation.py`: matching, tolerance-aware price and arithmetic checks.
- `exports.py`: order workbook, validation report and product-photo library.
- `excel_images.py`: native Excel in-cell image conversion and OOXML verification.
- `pipeline.py`: batch orchestration, duplicate detection, audit and history updates.
- `history.py`: local SQLite persistence.
- `gui.py`: PySide6 desktop interface and manual photo-association workflow.

## Trust and data integrity

- A SKU match is exact and case-normalised. Descriptions are only a cautious secondary key when the source SKU is absent.
- Unmatched or ambiguous products remain in manual review.
- Source values are not overwritten by official prices or calculated values.
- Python `Decimal` is used for monetary arithmetic.
- PDF text and OCR output are treated as extracted evidence, not unquestionable truth.
- No business documents are uploaded to a server.

## Native in-cell image pipeline

openpyxl initially embeds local images as temporary drawings at explicit product row coordinates. On Windows, desktop Excel automation calls `Shape.PlacePictureInCell` to convert each picture into a native cell image. The exporter then checks the workbook's rich-value metadata and verifies the exact photo cells.

There is no floating-picture fallback. If supported Excel automation is unavailable or verification fails, the image-bearing workbook export fails clearly.

See [EXCEL_IN_CELL_IMAGES.md](EXCEL_IN_CELL_IMAGES.md).
