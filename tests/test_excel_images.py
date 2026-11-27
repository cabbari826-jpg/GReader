from zipfile import ZipFile
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from PIL import Image

from ordertool.excel_images import (
    NativeInCellImagesUnavailable,
    embed_native_cell_images,
    verify_native_cell_images,
)


def test_invalid_workbook_is_not_accepted(tmp_path):
    workbook = tmp_path / "not-an-xlsx.xlsx"
    workbook.write_bytes(b"not a zip archive")
    try:
        verify_native_cell_images(workbook, {"Order Details": [{"cell": "A8"}]})
    except NativeInCellImagesUnavailable:
        pass
    else:
        raise AssertionError("Invalid workbook must not pass native-image verification")


def test_no_images_is_a_noop(tmp_path):
    assert embed_native_cell_images(tmp_path / "not-created.xlsx", {}) == 0


def test_image_is_embedded_as_native_cell_value_not_drawing(tmp_path):
    workbook_path = tmp_path / "native_cell_image.xlsx"
    image_path = tmp_path / "6814388.jpg"
    Image.new("RGB", (120, 80), (220, 150, 90)).save(image_path, format="JPEG")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Order Details"
    sheet["A7"] = "Product Photo"
    sheet["B7"] = "Article Code"
    sheet["B8"] = "6814388"
    sheet.row_dimensions[8].height = 62
    workbook.save(workbook_path)

    expected = {"Order Details": [{"cell": "A8", "image": str(image_path)}]}
    assert embed_native_cell_images(workbook_path, expected) == 1
    verify_native_cell_images(workbook_path, expected)

    with ZipFile(workbook_path) as archive:
        names = set(archive.namelist())
        assert "xl/media/image1.jpg" in names
        assert "xl/metadata.xml" in names
        assert "xl/richData/rdrichvalue.xml" in names
        assert not any(name.startswith("xl/drawings/") for name in names)
        sheet_xml = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        cell = sheet_xml.find(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c[@r='A8']")
        assert cell is not None
        assert cell.attrib["vm"] == "1"
        assert cell.attrib["t"] == "e"

    # openpyxl can still read the workbook's ordinary data, but it does not render
    # native rich-value images; never resave this workbook with openpyxl.
    reopened = load_workbook(workbook_path, data_only=False)
    assert reopened["Order Details"]["B8"].value == "6814388"
    reopened.close()
