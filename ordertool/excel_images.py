from __future__ import annotations

"""Native Excel 'Place in Cell' image conversion.

openpyxl can embed pictures as drawing objects anchored to cells, but those are
not cell values. This module converts those temporary embedded drawings into
native Excel in-cell images through Excel's own Windows automation interface.
The workbook is deliberately not re-saved by openpyxl after this step.
"""

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


class NativeInCellImagesUnavailable(RuntimeError):
    """Raised when native, sortable Excel cell images cannot be created."""


def _excel_available():
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise NativeInCellImagesUnavailable(
            "Native in-cell photos require Windows, pywin32, and a compatible "
            "desktop version of Microsoft Excel (Microsoft 365 or Excel 2024). "
            "Install or repair Excel, then retry. Floating-picture fallback is disabled."
        ) from exc
    return win32com.client


def _worksheet_part_map(archive: zipfile.ZipFile) -> dict[str, str]:
    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("rel:Relationship", ns)}
    result = {}
    for sheet in workbook.findall("main:sheets/main:sheet", ns):
        rel_id = sheet.attrib[f"{{{ns['r']}}}id"]
        target = targets[rel_id].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        result[sheet.attrib["name"]] = target
    return result


def verify_native_cell_images(workbook_path: str | Path, expected_cells: dict[str, list[str]]) -> None:
    """Verify OOXML rich-value metadata is attached to every expected photo cell."""
    path = Path(workbook_path)
    try:
        with zipfile.ZipFile(path, "r") as archive:
            bad = archive.testzip()
            if bad:
                raise NativeInCellImagesUnavailable(f"Excel produced a damaged XLSX part: {bad}")
            names = set(archive.namelist())
            if "xl/metadata.xml" not in names:
                raise NativeInCellImagesUnavailable(
                    "Excel did not write the native rich-value metadata required for in-cell pictures."
                )
            rich_parts = [name for name in names if name.startswith("xl/richData/")]
            if not rich_parts:
                raise NativeInCellImagesUnavailable(
                    "Excel did not write the native rich-data image records."
                )

            part_map = _worksheet_part_map(archive)
            ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            missing = []
            for sheet_name, cells in expected_cells.items():
                if sheet_name not in part_map:
                    raise NativeInCellImagesUnavailable(
                        f"Expected worksheet '{sheet_name}' was not found after Excel conversion."
                    )
                root = ET.fromstring(archive.read(part_map[sheet_name]))
                cell_map = {
                    cell.attrib.get("r"): cell
                    for cell in root.findall(".//main:sheetData/main:row/main:c", ns)
                }
                for address in cells:
                    cell = cell_map.get(address)
                    if cell is None or "vm" not in cell.attrib:
                        missing.append(f"{sheet_name}!{address}")
            if missing:
                raise NativeInCellImagesUnavailable(
                    "Excel did not store native image values in these expected cells: "
                    + ", ".join(missing[:20])
                )
    except zipfile.BadZipFile as exc:
        raise NativeInCellImagesUnavailable("The workbook is not a valid XLSX file after Excel conversion.") from exc


def convert_drawings_to_native_cells(
    workbook_path: str | Path,
    expected_cells: dict[str, list[str]],
) -> int:
    """Convert all temporary embedded pictures in named worksheets into cell values.

    Expected cells are explicit so we can verify every image landed in the intended
    product row. Images are first inserted as embedded drawings by openpyxl; Excel
    then converts them using Shape.PlacePictureInCell. There is intentionally no
    fallback to floating images.
    """
    path = Path(workbook_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(path)
    expected_count = sum(len(addresses) for addresses in expected_cells.values())
    if expected_count == 0:
        return 0

    win32 = _excel_available()
    excel = None
    workbook = None
    converted = 0
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        # msoAutomationSecurityForceDisable; the generated workbook contains no macros.
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass
        workbook = excel.Workbooks.Open(
            Filename=str(path),
            UpdateLinks=0,
            ReadOnly=False,
            AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )
        for sheet_name, cells in expected_cells.items():
            worksheet = workbook.Worksheets(sheet_name)
            expected = {address.replace("$", "").upper() for address in cells}
            found = set()
            # Iterate backwards because Excel may remove a drawing after conversion.
            for index in range(int(worksheet.Shapes.Count), 0, -1):
                shape = worksheet.Shapes.Item(index)
                # msoPicture = 13; openpyxl embeds raster images as pictures.
                if int(shape.Type) != 13:
                    continue
                address = str(shape.TopLeftCell.Address(False, False)).replace("$", "").upper()
                if address not in expected:
                    continue
                try:
                    shape.PlacePictureInCell()
                except Exception as exc:
                    raise NativeInCellImagesUnavailable(
                        f"Excel could not place the photo in {sheet_name}!{address}. "
                        "This Excel build may not support native in-cell pictures."
                    ) from exc
                found.add(address)
                converted += 1
            absent = expected - found
            if absent:
                raise NativeInCellImagesUnavailable(
                    f"Could not locate the embedded photo drawing(s) for "
                    f"{sheet_name}: {', '.join(sorted(absent))}"
                )
        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None
    except NativeInCellImagesUnavailable:
        raise
    except Exception as exc:
        raise NativeInCellImagesUnavailable(
            "Excel automation could not complete native in-cell image conversion. "
            "Check that desktop Excel is installed, activated, and not displaying "
            f"a modal dialog. Details: {exc}"
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    if converted != expected_count:
        raise NativeInCellImagesUnavailable(
            f"Expected {expected_count} in-cell images, but Excel converted {converted}."
        )
    verify_native_cell_images(path, expected_cells)
    return converted
