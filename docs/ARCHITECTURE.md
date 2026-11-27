# Architecture

## Processing flow

```text
PDF order queue
  -> PyMuPDF text/coordinate extraction; local Tesseract fallback
  -> PDF Style Code + product row/photo association
  -> official catalogue (Article Code, Style Code, Retail Price, Kit Q.ty)
  -> Style Code -> Article Code mapping
  -> expected net price = Retail Price * (1 - configured bulk discount)
  -> Decimal price, quantity, arithmetic and order-total validation
  -> photos renamed by verified Article Code
  -> order XLSX + validation report + product photo library
  -> JSON audit + SQLite history
```

## Modules

- `models.py`: explicit order, product line, price record and batch models.
- `extraction.py`: template-aware PDF row and product-photo extraction.
- `price_list.py`: Article Code-first catalogue parsing and Style Code mapping.
- `validation.py`: catalogue matching and tolerance-aware discounted price validation.
- `exports.py`: order workbook, validation report and photo-library output.
- `excel_images.py`: native Excel in-cell rich-value OOXML embedding and verification.
- `pipeline.py`: batch orchestration, article-based photo naming, auditing and history.
- `history.py`: local SQLite processing history.
- `gui.py`: PySide6 desktop workflow, review, settings and updater UI.
- `updater.py`: latest release check, installer download and SHA-256 verification.

## Business price rule

The official catalogue value is Retail Price. The default bulk discount is 60%, and the expected purchase price is Retail Price multiplied by 0.40. Users can change the discount in the desktop application. Extracted PDF prices remain unchanged; differences are reported rather than corrected.

## Data integrity and privacy

- Article Code is the canonical identifier; Style Code is retained as the source identifier.
- Unmatched product rows and uncertain image-to-row associations require review.
- Original extracted order values are not overwritten.
- Monetary arithmetic uses `Decimal`.
- All order and catalogue processing is local. The updater communicates with GitHub Releases only; no business files are transmitted.

## Native in-cell photos

openpyxl produces the styled workbook without images. `excel_images.py` then injects the OOXML rich-value records and image media files directly into the XLSX ZIP package. The image is referenced by the worksheet cell's `vm` metadata value; the workbook is not populated with floating drawing shapes. The output is structurally verified before export is marked successful.

See [EXCEL_IN_CELL_IMAGES.md](EXCEL_IN_CELL_IMAGES.md).
