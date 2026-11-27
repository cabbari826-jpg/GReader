# Contributing

1. Use Python 3.11.
2. Create a virtual environment and install `requirements-dev.txt`.
3. Run `python -m compileall -q app.py ordertool tests`.
4. Run `python -m pytest -q`.
5. Never add real order PDFs, catalogues, customer details, extracted photographs, or generated business workbooks.
6. Any workbook-export change must preserve audit values and include a regression test.

## Native photo requirement

A workbook with photos is accepted only when each photo is a native Excel in-cell image. A drawing object positioned over a cell is not equivalent. Validate the workbook with supported desktop Excel, sort/filter the product table by SKU, save, close, and reopen. See [docs/EXCEL_IN_CELL_IMAGES.md](docs/EXCEL_IN_CELL_IMAGES.md).
