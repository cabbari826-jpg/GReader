# Native Excel in-cell product photographs

## Requirement

Product photographs must be true Excel **Picture in Cell** values. A DrawingML image anchored over a cell is not sufficient: it is a drawing object, not a cell value, and can become detached from a product when sorting or filtering.

## Implementation

1. PyMuPDF extracts the original embedded photo from the PDF and associates it with a product row using page geometry.
2. The row is linked to the official catalogue using the PDF Style Code.
3. The catalogue's Article Code becomes the canonical product identifier and the photo filename, e.g. `6814388.jpg`.
4. Python writes the normal workbook, tables, formulas, formats, and row heights with openpyxl.
5. The native-image writer patches the XLSX ZIP package with Excel rich-value metadata, local image relationships, embedded image files, and worksheet cell metadata indices.
6. The exporter checks that each expected photo cell has a native rich-value index and that the related image bytes are embedded.

The package structure uses the `xl/metadata.xml` and `xl/richData/*` rich-value parts, with worksheet cell `vm` references and local-image relationships. It does **not** add `xl/drawings/*` pictures or rely on Excel COM automation.

## Why Excel COM was removed

The previous version attempted to open each generated workbook through the Windows Excel automation API. The user's first batch failed with Excel's `Workbooks.Open method failed` error for all three orders. Creating the native rich-value package directly removes that runtime dependency and avoids an Excel startup/session issue during conversion.

## Compatibility and limitations

- The workbook contains actual embedded image bytes; there are no external links or HTTPS image formulas.
- Open it in a Microsoft Excel version that supports **Place in Cell** / native in-cell images.
- Older Excel versions and third-party spreadsheet applications may not render rich-value images correctly.
- Do not re-save the workbook with openpyxl or another library that does not preserve rich-value image parts. Such a save may remove the image metadata.
- Sorting/filtering must be tested in the exact Excel version used by the business before deployment is approved.

## Automated checks

Tests verify the XLSX ZIP integrity, embedded media, rich-value package parts, and each expected worksheet cell's `vm` metadata binding. Those structural tests do not replace the final user acceptance test in desktop Excel.
