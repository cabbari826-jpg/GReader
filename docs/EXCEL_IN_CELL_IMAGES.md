# Native Excel in-cell product photographs

## Requirement

Product images in generated workbooks must be true Excel **Picture in Cell / Place in Cell** values. Ordinary DrawingML pictures with a one-cell or two-cell anchor are not acceptable: they may visually line up with rows but they are drawing-layer objects, not cell values, and can become detached when rows are sorted or filtered.

Microsoft documents native in-cell pictures as sortable/filterable cell values in supported Excel releases:
- Microsoft 365
- Excel 2024

See [Microsoft Support: Insert a picture with the Place in Cell feature](https://support.microsoft.com/en-us/excel/insert-picture-in-cell-in-excel).

## Implementation

1. Python extracts product photos and associates them with SKU rows.
2. openpyxl builds a normal workbook and embeds each local image as a temporary, self-contained picture drawing positioned in the intended Photo cell. It is not linked to the photo folder.
3. On Windows, `ordertool.excel_images.convert_drawings_to_native_cells` opens the workbook through desktop Excel automation and calls Excel's `Shape.PlacePictureInCell` operation on each intended image.
4. Excel saves the workbook using its native rich-value image representation.
5. The application inspects the resulting XLSX package. It requires `xl/metadata.xml`, rich-data image parts, and a cell metadata (`vm`) binding at every expected image cell.
6. If Excel or the necessary in-cell feature is unavailable, export fails with an explicit message. The application must never report a floating-picture workbook as satisfying the in-cell requirement.

The workbook must not be opened and re-saved by openpyxl after native in-cell conversion; libraries that do not understand Excel's rich-value image extension may remove or damage it.

## Runtime requirements and compatibility

- Windows 10/11 x64.
- Desktop Microsoft Excel from Microsoft 365 or Excel 2024 with native Place in Cell support.
- Python dependency: `pywin32`.
- No network URL is used for product photos; local image files are embedded into the workbook.

The converter uses Excel itself to author the native format, avoiding a hand-written rich-data OOXML implementation. It validates the output package before marking export successful.

## Acceptance tests

For an order workbook and photo library:
- Every photo-bearing product row has an in-cell picture value in the Photo column.
- Every target Photo cell has a native rich-value metadata binding.
- The workbook contains rich-data image records and is a valid XLSX ZIP package.
- Open the workbook in supported Excel, sort the product table by SKU, and filter it. Confirm the pictures remain with their corresponding SKU values.
- Save, close, and reopen the workbook. Confirm the pictures still display and remain sortable/filterable.

## Older Excel versions

Older Excel versions may display a compatibility warning or fail to render native in-cell images. The application does not silently substitute floating pictures. On unsupported installations, the user must upgrade Excel or explicitly choose an export without embedded images; that export is marked as not meeting the product-photo acceptance requirement.
