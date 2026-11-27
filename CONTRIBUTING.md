# Contributing

1. Use Python 3.11.
2. Create a virtual environment and install `requirements-dev.txt`.
3. Run `python -m compileall -q app.py ordertool tests`.
4. Run `python -m pytest -q`.
5. Never add real order PDFs, catalogues, customer details, extracted photographs, or generated business workbooks.
6. Any change to catalogue parsing must preserve both Article Code and Style Code.
7. Price-check tests must account for the configured retail discount; the default is 60%.
8. In-cell picture tests must verify OOXML rich-value metadata and embedded image bytes. Floating DrawingML images are not accepted.
9. Changes to update/install flows must preserve SHA-256 verification and ask the user before running an installer.

Before deployment, verify the workbook in the target Microsoft Excel version by sorting/filtering the product table by Article Code.
