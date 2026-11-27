from pathlib import Path

from PIL import Image

from ordertool.models import OrderData, ProductLine
from ordertool.pipeline import _rename_photos_to_article_codes


def test_photo_filename_uses_article_code_not_style_code(tmp_path):
    photo_dir = tmp_path / "Product_Photos"
    photo_dir.mkdir()
    source = photo_dir / "UNASSIGNED_64_P014.png"
    Image.new("RGB", (50, 50), "blue").save(source)

    product = ProductLine(
        article_code="6814388",
        sku="6814388",
        style_code="S2620155-0115",
        photo_path=str(source),
        photo_filename=source.name,
        photo_paths=[str(source)],
        photo_filenames=[source.name],
        photo_status="MATCHED",
    )
    order = OrderData(source_pdf="order.pdf", source_hash="0" * 64, products=[product])

    _rename_photos_to_article_codes(order, photo_dir)

    assert product.photo_filename == "6814388.png"
    assert Path(product.photo_path).is_file()
    assert not source.exists()


def test_uncertain_photo_is_not_assigned_an_article_filename(tmp_path):
    photo_dir = tmp_path / "Product_Photos"
    photo_dir.mkdir()
    source = photo_dir / "UNASSIGNED_66_P014.png"
    Image.new("RGB", (50, 50), "green").save(source)

    product = ProductLine(
        article_code="6814388",
        style_code="S2620155-0115",
        photo_path=str(source),
        photo_paths=[str(source)],
        photo_status="REVIEW",
    )
    order = OrderData(source_pdf="order.pdf", source_hash="0" * 64, products=[product])

    _rename_photos_to_article_codes(order, photo_dir)

    assert source.exists()
    assert product.photo_status == "REVIEW"
    assert product.photo_filename == ""
