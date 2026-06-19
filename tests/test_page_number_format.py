import pytest

from OfficePDFBinder_Main import AppWorker


@pytest.mark.parametrize(
    ("format_text", "current_number", "total_pages", "expected"),
    [
        ("1", 3, 12, "3"),
        ("1 / 10", 3, 12, "3 / 12"),
        ("Page 1", 3, 12, "Page 3"),
        ("- 1 -", 3, 12, "- 3 -"),
    ],
)
def test_format_page_number(
    format_text, current_number, total_pages, expected
):
    worker = AppWorker("test")

    assert (
        worker._format_page_number(format_text, current_number, total_pages)
        == expected
    )
