import fitz
import pytest
from PIL import Image, ImageDraw

from OfficePDFBinder_Main import AppWorker


def pdf_item(path, page_num):
    return {
        "type": "pdf",
        "path": str(path),
        "original_path": str(path),
        "page_num": page_num,
        "rotation": 0,
    }


def find_text_span(page, expected):
    for block in page.get_text("dict").get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                if span.get("text") == expected:
                    return span
    return None


def assert_rendered_pixels_in_visual_rect(page, visual_rect, scale=2):
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    x0 = max(0, int(visual_rect.x0 * scale))
    y0 = max(0, int(visual_rect.y0 * scale))
    x1 = min(pixmap.width, int(visual_rect.x1 * scale) + 1)
    y1 = min(pixmap.height, int(visual_rect.y1 * scale) + 1)

    dark_pixel_found = False
    samples = pixmap.samples
    channels = pixmap.n
    for y in range(y0, y1):
        row_start = y * pixmap.stride
        for x in range(x0, x1):
            offset = row_start + x * channels
            if min(samples[offset : offset + 3]) < 200:
                dark_pixel_found = True
                break
        if dark_pixel_found:
            break

    assert dark_pixel_found


def page_number_settings():
    return {
        "header_enabled": False,
        "footer_enabled": True,
        "font_size": 10,
        "footer": {
            "left": "",
            "center": "",
            "right": "",
            "auto_date": False,
            "auto_page_number": True,
            "page_number_position": "right",
        },
        "page_number_format": "Page 1",
        "page_number_start": 1,
    }


def assert_page_number_at_visual_footer_right(page, expected="Page 1"):
    assert expected in page.get_text()
    span = find_text_span(page, expected)
    assert span is not None
    unrotated_bbox = fitz.Rect(span["bbox"])
    visual_bbox = unrotated_bbox * page.rotation_matrix
    assert page.rect.contains(visual_bbox)
    assert visual_bbox.x0 > page.rect.width / 2
    assert visual_bbox.y0 > page.rect.height * 0.8
    assert_rendered_pixels_in_visual_rect(page, visual_bbox)


def create_special_source(tmp_path, case):
    path = tmp_path / f"special-{case}.pdf"
    document = fitz.open()

    if case == "landscape":
        page = document.new_page(width=500, height=300)
        page.insert_text((36, 72), "LANDSCAPE", fontsize=14)
    elif case == "pre_rotated":
        page = document.new_page(width=300, height=400)
        page.insert_text((36, 72), "PRE-ROTATED", fontsize=14)
        page.set_rotation(90)
    elif case == "cropbox":
        page = document.new_page(width=400, height=500)
        page.insert_text((90, 120), "CROPBOX", fontsize=14)
        page.set_cropbox(fitz.Rect(50, 50, 350, 450))
    elif case == "scanned":
        image_path = tmp_path / "scanned-page.png"
        image = Image.new("RGB", (600, 800), "white")
        draw = ImageDraw.Draw(image)
        draw.rectangle((30, 30, 570, 770), outline="black", width=5)
        draw.text((80, 100), "SCANNED IMAGE", fill="black")
        image.save(image_path)
        page = document.new_page(width=300, height=400)
        page.insert_image(page.rect, filename=str(image_path))
    else:
        raise AssertionError(f"Unknown case: {case}")

    document.save(path)
    document.close()
    return path


@pytest.mark.parametrize(
    ("format_text", "expected", "artifact_name"),
    [
        ("1", "3", "page_number_plain.pdf"),
        ("1 / 10", "3 / 3", "page_number_fraction.pdf"),
        ("Page 1", "Page 3", "page_number_page_label.pdf"),
        ("- 1 -", "- 3 -", "page_number_decorated.pdf"),
    ],
)
def test_page_number_is_saved_and_rendered_at_footer_right(
    pdf_factory,
    tmp_path,
    artifact_saver,
    format_text,
    expected,
    artifact_name,
):
    source = pdf_factory("three-pages.pdf", ["BODY-1", "BODY-2", "BODY-3"])
    output = tmp_path / artifact_name
    worker = AppWorker("merge_save")
    settings = {
        "header_enabled": False,
        "footer_enabled": True,
        "font_size": 10,
        "footer": {
            "left": "",
            "center": "",
            "right": "",
            "auto_date": False,
            "auto_page_number": True,
            "page_number_position": "right",
        },
        "page_number_format": format_text,
        "page_number_start": 1,
    }

    worker._run_merge_save(
        [pdf_item(source, page_num) for page_num in range(3)],
        str(output),
        header_footer_settings=settings,
    )

    artifact_saver(output, f"page_numbers/{artifact_name}")
    with fitz.open(output) as document:
        page = document[2]
        assert expected in page.get_text()

        span = find_text_span(page, expected)
        assert span is not None
        bbox = fitz.Rect(span["bbox"])
        assert page.rect.contains(bbox)
        assert bbox.x0 > page.rect.width / 2
        assert bbox.y0 > page.rect.height * 0.8

        assert_rendered_pixels_in_visual_rect(page, bbox)


@pytest.mark.parametrize("rotation", [0, 90, 180, 270])
def test_page_number_stays_at_visual_footer_right_after_rotation(
    pdf_factory, tmp_path, artifact_saver, rotation
):
    source = pdf_factory(f"rotation-{rotation}.pdf", [f"ROTATION-{rotation}"])
    output = tmp_path / f"page_number_rotation_{rotation}.pdf"
    worker = AppWorker("merge_save")
    item = pdf_item(source, 0)
    item["rotation"] = rotation

    worker._run_merge_save(
        [item],
        str(output),
        header_footer_settings=page_number_settings(),
    )

    artifact_saver(output, f"page_numbers/{output.name}")
    with fitz.open(output) as document:
        page = document[0]
        assert page.rotation == rotation
        assert_page_number_at_visual_footer_right(page)


@pytest.mark.parametrize(
    ("case", "expected_rotation"),
    [
        ("landscape", 0),
        ("pre_rotated", 90),
        ("cropbox", 0),
        ("scanned", 0),
    ],
)
def test_page_number_on_special_pdf_page_types(
    tmp_path, artifact_saver, case, expected_rotation
):
    source = create_special_source(tmp_path, case)
    output = tmp_path / f"page_number_{case}.pdf"
    worker = AppWorker("merge_save")

    worker._run_merge_save(
        [pdf_item(source, 0)],
        str(output),
        header_footer_settings=page_number_settings(),
    )

    artifact_saver(output, f"page_numbers/{output.name}")
    with fitz.open(output) as document:
        page = document[0]
        assert page.rotation == expected_rotation
        assert_page_number_at_visual_footer_right(page)

        if case == "landscape":
            assert page.rect.width > page.rect.height
        elif case == "pre_rotated":
            assert page.rect.width > page.rect.height
        elif case == "cropbox":
            assert page.cropbox.width == pytest.approx(300)
            assert page.cropbox.height == pytest.approx(400)
        elif case == "scanned":
            assert page.get_images(full=True)
