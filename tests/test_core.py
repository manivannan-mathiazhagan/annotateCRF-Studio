from types import SimpleNamespace

import fitz

from Scripts import AnnotateCRF_Studio as app


def test_default_font_sizes():
    assert app.DEFAULT_BASE_FONT_SIZE == 10
    assert app.DEFAULT_DOMAIN_FONT_SIZE == 12
    assert app.FONT_SIZE_OPTIONS == list(range(8, 15))


def test_clean_number():
    assert app.clean_number(" 12.5 ") == 12.5
    assert app.clean_number("") is None
    assert app.clean_number("not-a-number", 7) == 7


def test_extract_page_reference():
    assert app.extract_page_reference("Refer to Page 12") == 12
    assert app.extract_page_reference("For annotations refer to page 7") == 7
    assert app.extract_page_reference("See page 3") == 3
    assert app.extract_page_reference("No page reference") is None


def test_wrap_text_by_width_preserves_short_text():
    assert app.wrap_text_by_width("DM.USUBJID", 500) == ["DM.USUBJID"]


def test_wrap_text_by_width_wraps_long_text():
    lines = app.wrap_text_by_width("one two three four five", 45)
    assert len(lines) > 1
    assert " ".join(line.strip() for line in lines) == "one two three four five"


def test_output_pdf_path():
    assert app.sanitize_output_pdf_path("/tmp/crf.pdf") == "/tmp/crf_final.pdf"


def test_get_pdf_base_output_path():
    assert app.get_pdf_base_output_path("/tmp/crf_final.pdf") == "/tmp/crf_final"


def test_refer_page_detection():
    entry = SimpleNamespace(domain="REF", name="REF", annotation="Refer to Page 5")
    assert app.is_refer_page_entry(entry) is True

    normal = SimpleNamespace(domain="DM", name="USUBJID", annotation="DM.USUBJID")
    assert app.is_refer_page_entry(normal) is False


def test_domain_color_map_restarts_by_page_and_ignores_ref_notsub():
    entries = [
        SimpleNamespace(pageno=1, domain="DM", name="", annotation="DM", is_not_submitted=False),
        SimpleNamespace(pageno=1, domain="AE", name="", annotation="AE", is_not_submitted=False),
        SimpleNamespace(pageno=1, domain="REF", name="REF", annotation="Refer to Page 2", is_not_submitted=False),
        SimpleNamespace(pageno=1, domain="NOTSUB", name="", annotation="NOT SUBMITTED", is_not_submitted=True),
        SimpleNamespace(pageno=2, domain="AE", name="", annotation="AE", is_not_submitted=False),
    ]

    page1 = app.get_page_domain_color_map(entries, 1)
    page2 = app.get_page_domain_color_map(entries, 2)

    assert page1["DM"] == app.DOMAIN_COLORS[0]
    assert page1["AE"] == app.DOMAIN_COLORS[1]
    assert "REF" not in page1
    assert "NOTSUB" not in page1
    assert page2["AE"] == app.DOMAIN_COLORS[0]


def test_clamp_rect_to_page():
    rect = fitz.Rect(95, 95, 115, 115)
    clamped = app.clamp_rect_to_page(rect, 100, 100)

    assert clamped.x0 >= 0
    assert clamped.y0 >= 0
    assert clamped.x1 <= 100
    assert clamped.y1 <= 100
    assert clamped.width == rect.width
    assert clamped.height == rect.height
