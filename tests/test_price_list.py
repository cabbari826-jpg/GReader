from openpyxl import Workbook

from ordertool.price_list import load_price_list


def test_import_uses_article_code_as_canonical_key(tmp_path):
    path = tmp_path / "catalogue.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append([
        "Catalogue", "Catalogue Name", "Page", "Picture", "Style Code",
        "Style Description", "Additional Description", "Favourite", "Option Id",
        "Department", "Function", "Made In", "Composition", "Material", "Event",
        "Segment", "Discount", "Kit Q.ty", "Retail Price", "Article Code",
        "Style Id", "Quantity",
    ])
    sheet.append([
        "3067", "Living room", "Objects", None, "S2620155-0115",
        "VASE CONCRETE FINISH H42", "[6814388] Steel vase", None, 1,
        "OBJECTS", "VASES", "IN", "STEEL", "", "", "PERMANENTE", None,
        1, 49.90, 6814388, "155048125", 0,
    ])
    workbook.save(path)

    products = load_price_list(path)
    assert list(products) == ["6814388"]
    assert products["6814388"].article_code == "6814388"
    assert products["6814388"].style_code == "S2620155-0115"
    assert str(products["6814388"].official_price) == "49.9"
