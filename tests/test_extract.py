from foodmix.extract import parse_extraction


def test_line_category_is_not_taken_from_the_whole_bill():
    text = """
    ```json
    {
      "date": "2026/08/01",
      "category": "蔬菜",
      "lines": [
        {"name": "五花肉", "qty": "1", "unit": "斤", "amount": "20", "category": "肉类"},
        {"name": "可乐", "qty": 1, "unit": "瓶", "amount": 3},
        {"name": "调料包", "qty": 1, "unit": "包", "amount": 4, "category": "调料"}
      ]
    }
    ```
    """
    parsed = parse_extraction(text)
    assert parsed["date"] == "2026-08-01"
    assert [line["category"] for line in parsed["lines"]] == ["肉类", "", ""]
    assert parsed["lines"][0]["raw_name"] == "五花肉"
    assert parsed["lines"][1]["qty"] == 1
