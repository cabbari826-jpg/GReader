from pathlib import Path
import zipfile

import pytest

from ordertool.excel_images import (
    NativeInCellImagesUnavailable,
    convert_drawings_to_native_cells,
    verify_native_cell_images,
)


def test_native_in_cell_conversion_requires_expected_image_cells(tmp_path):
    workbook = tmp_path / "empty.xlsx"
    # An arbitrary non-XLSX file should fail verification, rather than be accepted.
    workbook.write_bytes(b"not a zip archive")
    with pytest.raises(NativeInCellImagesUnavailable):
        verify_native_cell_images(workbook, {"Order Details": ["A8"]})


def test_no_images_is_a_noop(tmp_path):
    workbook = tmp_path / "not-created.xlsx"
    assert convert_drawings_to_native_cells(workbook, {}) == 0


@pytest.mark.windows_excel
def test_excel_native_in_cell_integration():
    """Manual/Windows integration test; run only on a host with desktop Excel."""
    pytest.skip("Requires a dedicated Windows runner with activated Microsoft Excel.")
