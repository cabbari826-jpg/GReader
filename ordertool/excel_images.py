from __future__ import annotations

"""Pure-Python writer for native Excel Picture-in-Cell rich values.

No Excel COM automation and no floating DrawingML pictures are used. The
workbook is first formatted with openpyxl, then this module adds OOXML
rich-value metadata and embeds each local image under xl/media.
"""

from pathlib import Path
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RICH = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
RICH2 = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
RICH_REL = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
ET.register_namespace("", MAIN)
ET.register_namespace("r", OFFICE)


class NativeInCellImagesUnavailable(RuntimeError):
    pass


def _cell_parts(address):
    match = re.fullmatch(r"([A-Za-z]+)([1-9][0-9]*)", address.strip())
    if not match:
        raise ValueError(f"Invalid cell address: {address!r}")
    return match.group(1).upper(), int(match.group(2))


def _col_num(col):
    value = 0
    for char in col:
        value = value * 26 + ord(char) - 64
    return value


def _sheet_part_map(archive):
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {r.attrib["Id"]: r.attrib["Target"]
               for r in rels.findall(f"{{{PKG}}}Relationship")}
    result = {}
    for sheet in workbook.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet"):
        rid = sheet.attrib[f"{{{OFFICE}}}id"]
        target = targets[rid]
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        result[sheet.attrib["name"]] = part
    return result


def _add_rel(root, rel_type, target):
    if any(r.attrib.get("Type") == rel_type for r in root.findall(f"{{{PKG}}}Relationship")):
        return
    used = []
    for rel in root.findall(f"{{{PKG}}}Relationship"):
        match = re.fullmatch(r"rId(\d+)", rel.attrib.get("Id", ""))
        if match:
            used.append(int(match.group(1)))
    rid = 1
    while rid in used:
        rid += 1
    ET.SubElement(root, f"{{{PKG}}}Relationship", {
        "Id": f"rId{rid}", "Type": rel_type, "Target": target,
    })


def _set_image_cell(sheet_root, address, vm):
    col, row_number = _cell_parts(address)
    sheet_data = sheet_root.find(f"{{{MAIN}}}sheetData")
    if sheet_data is None:
        sheet_data = ET.SubElement(sheet_root, f"{{{MAIN}}}sheetData")
    row = next((r for r in sheet_data.findall(f"{{{MAIN}}}row")
                if int(r.attrib.get("r", "0")) == row_number), None)
    if row is None:
        row = ET.Element(f"{{{MAIN}}}row", {"r": str(row_number)})
        rows = list(sheet_data.findall(f"{{{MAIN}}}row"))
        pos = next((i for i, r in enumerate(rows)
                    if int(r.attrib.get("r", "0")) > row_number), len(rows))
        sheet_data.insert(pos, row)
    address = f"{col}{row_number}"
    cell = next((c for c in row.findall(f"{{{MAIN}}}c")
                 if c.attrib.get("r") == address), None)
    if cell is None:
        cell = ET.Element(f"{{{MAIN}}}c", {"r": address})
        cells = list(row.findall(f"{{{MAIN}}}c"))
        pos = next((i for i, c in enumerate(cells)
                    if _col_num(_cell_parts(c.attrib["r"])[0]) > _col_num(col)), len(cells))
        row.insert(pos, cell)
    cell.attrib["t"] = "e"
    cell.attrib["vm"] = str(vm)
    for child in list(cell):
        if child.tag in (f"{{{MAIN}}}v", f"{{{MAIN}}}f", f"{{{MAIN}}}is"):
            cell.remove(child)
    ET.SubElement(cell, f"{{{MAIN}}}v").text = "#VALUE!"


def _xml_bytes(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _metadata(count):
    root = ET.Element(f"{{{MAIN}}}metadata")
    types = ET.SubElement(root, f"{{{MAIN}}}metadataTypes", {"count": "1"})
    ET.SubElement(types, f"{{{MAIN}}}metadataType", {
        "name": "XLRICHVALUE", "minSupportedVersion": "120000", "copy": "1",
        "pasteAll": "1", "pasteValues": "1", "merge": "1", "splitFirst": "1",
        "rowColShift": "1", "clearFormats": "1", "clearComments": "1",
        "assign": "1", "coerce": "1",
    })
    future = ET.SubElement(root, f"{{{MAIN}}}futureMetadata",
                           {"name": "XLRICHVALUE", "count": str(count)})
    for index in range(count):
        bk = ET.SubElement(future, f"{{{MAIN}}}bk")
        ext_list = ET.SubElement(bk, f"{{{MAIN}}}extLst")
        ext = ET.SubElement(ext_list, f"{{{MAIN}}}ext",
                            {"uri": "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"})
        ET.SubElement(ext, f"{{{RICH}}}rvb", {"i": str(index)})
    values = ET.SubElement(root, f"{{{MAIN}}}valueMetadata", {"count": str(count)})
    for index in range(count):
        bk = ET.SubElement(values, f"{{{MAIN}}}bk")
        ET.SubElement(bk, f"{{{MAIN}}}rc", {"t": "1", "v": str(index)})
    return _xml_bytes(root)


def _rich_structure():
    root = ET.Element(f"{{{RICH}}}rvStructures", {"count": "1"})
    struct = ET.SubElement(root, f"{{{RICH}}}s", {"t": "_localImage"})
    ET.SubElement(struct, f"{{{RICH}}}k", {"n": "_rvRel:LocalImageIdentifier", "t": "i"})
    ET.SubElement(struct, f"{{{RICH}}}k", {"n": "CalcOrigin", "t": "i"})
    return _xml_bytes(root)


def _rich_types():
    root = ET.Element(f"{{{RICH2}}}rvTypesInfo")
    flags = ET.SubElement(ET.SubElement(root, f"{{{RICH2}}}global"),
                          f"{{{RICH2}}}keyFlags")
    for name in ("_Self", "_DisplayString", "_Flags", "_Format", "_SubLabel",
                 "_Attribution", "_Icon", "_Display", "_CanonicalPropertyNames",
                 "_ClassificationId"):
        key = ET.SubElement(flags, f"{{{RICH2}}}key", {"name": name})
        if name == "_Self":
            ET.SubElement(key, f"{{{RICH2}}}flag",
                          {"name": "ExcludeFromFile", "value": "1"})
        ET.SubElement(key, f"{{{RICH2}}}flag",
                      {"name": "ExcludeFromCalcComparison", "value": "1"})
    return _xml_bytes(root)


def _rich_values(count):
    root = ET.Element(f"{{{RICH}}}rvData", {"count": str(count)})
    for index in range(count):
        rv = ET.SubElement(root, f"{{{RICH}}}rv", {"s": "0"})
        ET.SubElement(rv, f"{{{RICH}}}v").text = str(index)
        ET.SubElement(rv, f"{{{RICH}}}v").text = "5"
    return _xml_bytes(root)


def _rich_value_rel(count):
    root = ET.Element(f"{{{RICH_REL}}}richValueRels")
    for index in range(count):
        ET.SubElement(root, f"{{{RICH_REL}}}rel",
                      {f"{{{OFFICE}}}id": f"rId{index + 1}"})
    return _xml_bytes(root)


def _rich_value_rel_rels(entries):
    root = ET.Element(f"{{{PKG}}}Relationships")
    for index, entry in enumerate(entries, start=1):
        ext = Path(entry["image"]).suffix.lower().lstrip(".")
        ET.SubElement(root, f"{{{PKG}}}Relationship", {
            "Id": f"rId{index}", "Type": f"{OFFICE}/image",
            "Target": f"../media/image{index}.{ext}",
        })
    return _xml_bytes(root)


def _mime(ext):
    return {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
            "gif": "image/gif", "bmp": "image/bmp", "tif": "image/tiff",
            "tiff": "image/tiff", "webp": "image/webp"}.get(ext, "application/octet-stream")


def _patch_content_types(data, extensions):
    root = ET.fromstring(data)
    defaults = {e.attrib.get("Extension", "").lower()
                for e in root.findall(f"{{{CT}}}Default")}
    for ext in sorted(extensions - defaults):
        ET.SubElement(root, f"{{{CT}}}Default",
                      {"Extension": ext, "ContentType": _mime(ext)})
    overrides = {e.attrib.get("PartName")
                 for e in root.findall(f"{{{CT}}}Override")}
    required = {
        "/xl/metadata.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
        "/xl/richData/rdrichvalue.xml": "application/vnd.ms-excel.rdrichvalue+xml",
        "/xl/richData/rdrichvaluestructure.xml": "application/vnd.ms-excel.rdrichvaluestructure+xml",
        "/xl/richData/rdRichValueTypes.xml": "application/vnd.ms-excel.rdrichvaluetypes+xml",
        "/xl/richData/richValueRel.xml": "application/vnd.ms-excel.richvaluerel+xml",
    }
    for part, content_type in required.items():
        if part not in overrides:
            ET.SubElement(root, f"{{{CT}}}Override",
                          {"PartName": part, "ContentType": content_type})
    return _xml_bytes(root)


def embed_native_cell_images(workbook_path, images_by_sheet):
    """Embed local pictures as native cell values in an existing XLSX file.

    Example: {"Order Details": [{"cell": "A8", "image": "photos/6814388.jpg"}]}
    """
    path = Path(workbook_path)
    entries, seen = [], set()
    for sheet_name, images in images_by_sheet.items():
        for item in images:
            image_path = Path(item["image"])
            if not image_path.is_file():
                raise FileNotFoundError(f"Product image not found: {image_path}")
            ext = image_path.suffix.lower().lstrip(".")
            if ext not in {"png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp"}:
                raise NativeInCellImagesUnavailable(f"Unsupported image type: {image_path.suffix}")
            key = (sheet_name, item["cell"].upper())
            if key in seen:
                raise ValueError(f"More than one image is assigned to {sheet_name}!{item['cell']}")
            seen.add(key)
            _cell_parts(item["cell"])
            entries.append({"sheet": sheet_name, "cell": item["cell"].upper(),
                            "image": str(image_path)})

    if not entries:
        return 0

    temp = path.with_suffix(path.suffix + ".tmp")
    try:
        with zipfile.ZipFile(path, "r") as source:
            bad = source.testzip()
            if bad:
                raise NativeInCellImagesUnavailable(f"Source workbook ZIP is damaged: {bad}")
            sheet_parts = _sheet_part_map(source)
            missing_sheets = {e["sheet"] for e in entries} - set(sheet_parts)
            if missing_sheets:
                raise NativeInCellImagesUnavailable(f"Missing worksheets: {sorted(missing_sheets)}")
            worksheet_roots = {}
            for vm, entry in enumerate(entries, start=1):
                part = sheet_parts[entry["sheet"]]
                worksheet_roots.setdefault(part, ET.fromstring(source.read(part)))
                _set_image_cell(worksheet_roots[part], entry["cell"], vm)

            workbook_rels = ET.fromstring(source.read("xl/_rels/workbook.xml.rels"))
            _add_rel(workbook_rels, f"{OFFICE}/sheetMetadata", "metadata.xml")
            _add_rel(workbook_rels,
                     "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue",
                     "richData/rdrichvalue.xml")
            _add_rel(workbook_rels,
                     "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure",
                     "richData/rdrichvaluestructure.xml")
            _add_rel(workbook_rels,
                     "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes",
                     "richData/rdRichValueTypes.xml")
            _add_rel(workbook_rels,
                     "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel",
                     "richData/richValueRel.xml")

            extensions = {Path(e["image"]).suffix.lower().lstrip(".") for e in entries}
            replacements = {
                "[Content_Types].xml": _patch_content_types(source.read("[Content_Types].xml"), extensions),
                "xl/_rels/workbook.xml.rels": _xml_bytes(workbook_rels),
                "xl/metadata.xml": _metadata(len(entries)),
                "xl/richData/rdrichvaluestructure.xml": _rich_structure(),
                "xl/richData/rdRichValueTypes.xml": _rich_types(),
                "xl/richData/rdrichvalue.xml": _rich_values(len(entries)),
                "xl/richData/richValueRel.xml": _rich_value_rel(len(entries)),
                "xl/richData/_rels/richValueRel.xml.rels": _rich_value_rel_rels(entries),
            }
            replacements.update({part: _xml_bytes(root)
                                 for part, root in worksheet_roots.items()})
            with zipfile.ZipFile(temp, "w", compression=zipfile.ZIP_DEFLATED) as target:
                for name in source.namelist():
                    target.writestr(name, replacements.pop(name) if name in replacements else source.read(name))
                for name, payload in replacements.items():
                    target.writestr(name, payload)
                for index, entry in enumerate(entries, start=1):
                    ext = Path(entry["image"]).suffix.lower().lstrip(".")
                    target.write(entry["image"], f"xl/media/image{index}.{ext}")
        temp.replace(path)
    except Exception:
        temp.unlink(missing_ok=True)
        raise

    verify_native_cell_images(path, images_by_sheet)
    return len(entries)


def verify_native_cell_images(workbook_path, images_by_sheet):
    path = Path(workbook_path)
    try:
        with zipfile.ZipFile(path, "r") as archive:
            if archive.testzip():
                raise NativeInCellImagesUnavailable("Output workbook failed ZIP integrity verification.")
            names = set(archive.namelist())
            required = {
                "xl/metadata.xml", "xl/richData/rdrichvaluestructure.xml",
                "xl/richData/rdRichValueTypes.xml", "xl/richData/rdrichvalue.xml",
                "xl/richData/richValueRel.xml", "xl/richData/_rels/richValueRel.xml.rels",
            }
            missing = required - names
            if missing:
                raise NativeInCellImagesUnavailable(f"Missing rich-value parts: {sorted(missing)}")
            sheet_parts = _sheet_part_map(archive)
            vm = 0
            for sheet_name, images in images_by_sheet.items():
                if sheet_name not in sheet_parts:
                    raise NativeInCellImagesUnavailable(f"Missing sheet {sheet_name}.")
                root = ET.fromstring(archive.read(sheet_parts[sheet_name]))
                for item in images:
                    vm += 1
                    cell = root.find(f".//{{{MAIN}}}c[@r='{item['cell'].upper()}']")
                    if cell is None or cell.attrib.get("t") != "e" or cell.attrib.get("vm") != str(vm):
                        raise NativeInCellImagesUnavailable(
                            f"Photo was not stored as a native image value in {sheet_name}!{item['cell']}."
                        )
            media = [name for name in names if re.fullmatch(r"xl/media/image\d+\.[a-z0-9]+", name)]
            if len(media) < vm:
                raise NativeInCellImagesUnavailable(f"Expected {vm} embedded images, found {len(media)}.")
    except zipfile.BadZipFile as exc:
        raise NativeInCellImagesUnavailable("Output is not a valid XLSX package.") from exc
