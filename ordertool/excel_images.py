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
    """Verify rich-value image parts and a cell metadata binding for every photo cell."""
    path = Path(workbook_path)
    try:
        with zipfile.ZipFile(path, "r") as archive:
            bad = archive.testzip()
            if bad:
                raise NativeInCellImagesUnavailable(f"Damaged XLSX ZIP part: {bad}")
            names = set(archive.namelist())
            if "xl/metadata.xml" not in names:
                raise NativeInCellImagesUnavailable("Excel did not write native rich-value metadata.")
            if not any(name.startswith("xl/richData/") for name in names):
                raise NativeInCellImagesUnavailable("Excel did not write native rich-data image records.")
            part_map = _worksheet_part_map(archive)
            ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            missing = []
            for sheet_name, addresses in expected_cells.items():
                if sheet_name not in part_map:
                    raise NativeInCellImagesUnavailable(
                        f"Worksheet '{sheet_name}' missing after Excel conversion."
                    )
                root = ET.fromstring(archive.read(part_map[sheet_name]))
                cell_map = {
                    cell.attrib.get("r"): cell
                    for cell in root.findall(".//main:sheetData/main:row/main:c", ns)
                }
                for address in addresses:
                    cell = cell_map.get(address.replace("$", "").upper())
                    if cell is None or "vm" not in cell.attrib:
                        missing.append(f"{sheet_name}!{address}")
            if missing:
                raise NativeInCellImagesUnavailable(
                    "Native image value missing from: " + ", ".join(missing[:20])
                )
    except zipfile.BadZipFile as exc:
        raise NativeInCellImagesUnavailable("The generated output is not a valid XLSX file.") from exc


def convert_drawings_to_native_cells(
    workbook_path: str | Path,
    expected_cells: dict[str, list[str]],
) -> int:
    """Use Excel's PlacePictureInCell on each expected embedded picture; never fall back."""
    expected_count = sum(len(items) for items in expected_cells.values())
    if expected_count == 0:
        return 0
    path = Path(workbook_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(path)

    win32 = _excel_available()
    excel = workbook = None
    converted = 0
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass
        workbook = excel.Workbooks.Open(
            Filename=str(path), UpdateLinks=0, ReadOnly=False, AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )
        for sheet_name, addresses in expected_cells.items():
            worksheet = workbook.Worksheets(sheet_name)
            expected = {address.replace("$", "").upper() for address in addresses}
            found = set()
            for index in range(int(worksheet.Shapes.Count), 0, -1):
                shape = worksheet.Shapes.Item(index)
                if int(shape.Type) != 13:  # msoPicture
                    continue
                address = str(shape.TopLeftCell.Address(False, False)).replace("$", "").upper()
                if address not in expected:
                    continue
                try:
                    shape.PlacePictureInCell()
                except Exception as exc:
                    raise NativeInCellImagesUnavailable(
                        f"Excel could not place the photo in {sheet_name}!{address}. "
                        "The installed Excel build may not support Place in Cell."
                    ) from exc
                found.add(address)
                converted += 1
            if expected - found:
                raise NativeInCellImagesUnavailable(
                    f"No embedded picture drawing found for {sheet_name}: "
                    f"{', '.join(sorted(expected - found))}"
                )
        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None
    except NativeInCellImagesUnavailable:
        raise
    except Exception as exc:
        raise NativeInCellImagesUnavailable(
            "Excel automation failed. Check that desktop Excel is installed, activated, "
            f"and not blocked by a modal dialog. Details: {exc}"
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
            f"Expected {expected_count} photos; Excel converted {converted}."
        )
    verify_native_cell_images(path, expected_cells)
    return converted
