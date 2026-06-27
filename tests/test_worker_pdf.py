import csv
import os
from pathlib import Path

import fitz
import pytest
from PIL import Image

from OfficePDFBinder_Main import AppWorker


def collect_signal(signal):
    calls = []
    signal.connect(lambda *args: calls.append(args))
    return calls


def pdf_item(path, page_num, rotation=0):
    return {
        "type": "pdf",
        "path": str(path),
        "original_path": str(path),
        "page_num": page_num,
        "rotation": rotation,
    }


def blank_item(name="空白ページ 1"):
    return {
        "type": "blank",
        "path": f"__{name}__",
        "original_path": f"__{name}__",
        "page_num": 0,
        "rotation": 0,
        "display_name": name,
        "width": 595,
        "height": 842,
    }


def image_item(path, rotation=0):
    return {
        "type": "image",
        "path": str(path),
        "original_path": str(path),
        "page_num": 0,
        "rotation": rotation,
    }


def test_add_pdf_emits_pages_and_existing_bookmarks(pdf_factory):
    source = pdf_factory(
        "bookmarked.pdf",
        ["first", "second"],
        toc=[[1, "Second page", 2]],
    )
    worker = AppWorker("add_files")
    item_batches = collect_signal(worker.signals.items_ready)
    bookmark_batches = collect_signal(worker.signals.bookmarks_ready)
    finished = collect_signal(worker.signals.finished)

    worker._run_add_files([str(source)])

    assert len(item_batches) == 1
    pages = item_batches[0][0]
    assert [item["page_num"] for item in pages] == [0, 1]
    assert all(item["original_path"] == str(source) for item in pages)
    assert bookmark_batches[0][0] == str(source)
    assert bookmark_batches[0][1][0]["title"] == "Second page"
    assert bookmark_batches[0][1][0]["page_num"] == 1
    assert finished[-1][1:] == ("完了", "ファイルの追加が完了しました。")


@pytest.mark.parametrize(
    ("suffix", "expected_type"),
    [(".docx", "word"), (".xlsx", "excel"), (".pptx", "powerpoint")],
)
def test_add_office_file_classifies_without_opening(
    tmp_path, suffix, expected_type
):
    source = tmp_path / f"sample{suffix}"
    source.write_bytes(b"placeholder")
    worker = AppWorker("add_files")
    items = collect_signal(worker.signals.item_ready)

    worker._run_add_files([str(source)])

    assert items[0][0]["type"] == expected_type
    assert items[0][0]["original_path"] == str(source)


def test_add_image_file_classifies_after_validation(tmp_path):
    source = tmp_path / "photo.png"
    Image.new("RGB", (120, 80), "blue").save(source)
    worker = AppWorker("add_files")
    items = collect_signal(worker.signals.item_ready)

    worker._run_add_files([str(source)])

    assert items[0][0]["type"] == "image"
    assert items[0][0]["original_path"] == str(source)
    assert items[0][0]["page_num"] == 0


def test_add_missing_file_reports_error(tmp_path):
    missing = tmp_path / "missing.pdf"
    worker = AppWorker("add_files")
    errors = collect_signal(worker.signals.error)

    worker._run_add_files([str(missing)])

    assert errors
    assert errors[0][0] == "ファイルが見つかりません"
    assert "missing.pdf" in errors[0][1]


def test_add_broken_pdf_reports_error(tmp_path):
    broken = tmp_path / "broken.pdf"
    broken.write_bytes(b"this is not a PDF")
    worker = AppWorker("add_files")
    errors = collect_signal(worker.signals.error)
    item_batches = collect_signal(worker.signals.items_ready)

    worker._run_add_files([str(broken)])

    assert not item_batches
    assert errors[0][0] == "PDFファイルを読み込めません"
    assert "broken.pdf" in errors[0][1]


def test_cancelled_add_reports_cancelled_without_processing(pdf_factory):
    source = pdf_factory("source.pdf", ["page"])
    worker = AppWorker("add_files")
    worker.is_running = False
    finished = collect_signal(worker.signals.finished)
    item_batches = collect_signal(worker.signals.items_ready)

    worker._run_add_files([str(source)])

    assert not item_batches
    assert finished == [("cancelled", "中止", "処理を中止しました。")]


def test_merge_preserves_requested_page_order_rotation_and_bookmarks(
    pdf_factory, tmp_path
):
    first = pdf_factory(
        "first.pdf",
        ["FIRST-1", "FIRST-2"],
        toc=[[1, "Original first", 1]],
    )
    second = pdf_factory("second.pdf", ["SECOND-1"])
    output = tmp_path / "output" / "merged.pdf"
    worker = AppWorker("merge_save")
    finished = collect_signal(worker.signals.finished)
    errors = collect_signal(worker.signals.error)

    worker._run_merge_save(
        items_data=[
            pdf_item(first, 1, rotation=90),
            pdf_item(second, 0),
            pdf_item(first, 0),
        ],
        output_path=str(output),
        bookmarks=[
            {"title": "Selected second", "path": str(first), "page_num": 1}
        ],
        show_outlines=True,
    )

    assert not errors
    assert output.exists()
    assert finished[-1][1] == "保存完了"
    with fitz.open(output) as document:
        assert document.page_count == 3
        assert "FIRST-2" in document[0].get_text()
        assert "SECOND-1" in document[1].get_text()
        assert "FIRST-1" in document[2].get_text()
        assert document[0].rotation == 90
        toc = document.get_toc()
        assert [entry[1] for entry in toc] == [
            "Selected second",
            "Original first",
        ]
        assert [entry[2] for entry in toc] == [1, 3]


def test_merge_empty_input_reports_error(tmp_path):
    worker = AppWorker("merge_save")
    errors = collect_signal(worker.signals.error)
    output = tmp_path / "empty.pdf"

    worker._run_merge_save([], str(output))

    assert not output.exists()
    assert errors[-1][0] == "保存できるページがありません"


def test_merge_can_insert_blank_page_between_pdf_pages(pdf_factory, tmp_path):
    source = pdf_factory("source.pdf", ["FIRST", "SECOND"])
    output = tmp_path / "with_blank.pdf"
    worker = AppWorker("merge_save")

    worker._run_merge_save(
        [
            pdf_item(source, 0),
            blank_item(),
            pdf_item(source, 1),
        ],
        str(output),
    )

    with fitz.open(output) as document:
        assert document.page_count == 3
        assert "FIRST" in document[0].get_text()
        assert document[1].get_text().strip() == ""
        assert round(document[1].rect.width) == 595
        assert round(document[1].rect.height) == 842
        assert "SECOND" in document[2].get_text()


def test_merge_can_insert_landscape_image_as_a4_landscape_pdf_page(tmp_path):
    image_path = tmp_path / "photo.png"
    Image.new("RGB", (400, 200), "blue").save(image_path)
    output = tmp_path / "image.pdf"
    worker = AppWorker("merge_save")

    worker._run_merge_save([image_item(image_path)], str(output))

    with fitz.open(output) as document:
        assert document.page_count == 1
        page = document[0]
        assert round(page.rect.width) == 842
        assert round(page.rect.height) == 595
        assert page.get_images(full=True)


def test_merge_can_insert_portrait_image_as_a4_portrait_pdf_page(tmp_path):
    image_path = tmp_path / "portrait.png"
    Image.new("RGB", (200, 400), "blue").save(image_path)
    output = tmp_path / "portrait_image.pdf"
    worker = AppWorker("merge_save")

    worker._run_merge_save([image_item(image_path)], str(output))

    with fitz.open(output) as document:
        assert document.page_count == 1
        page = document[0]
        assert round(page.rect.width) == 595
        assert round(page.rect.height) == 842
        assert page.get_images(full=True)


def test_merge_can_disable_image_upscaling(tmp_path):
    image_path = tmp_path / "small.png"
    Image.new("RGB", (100, 50), "blue").save(image_path)
    output = tmp_path / "small_image_no_upscale.pdf"
    worker = AppWorker("merge_save")

    worker._run_merge_save(
        [image_item(image_path)],
        str(output),
        disable_image_upscaling=True,
    )

    with fitz.open(output) as document:
        page = document[0]
        xref = page.get_images(full=True)[0][0]
        rect = page.get_image_rects(xref)[0]
        assert rect.width == pytest.approx(75)
        assert rect.height == pytest.approx(37.5)


def test_merge_expands_small_image_when_upscaling_is_enabled(tmp_path):
    image_path = tmp_path / "small.png"
    Image.new("RGB", (100, 50), "blue").save(image_path)
    output = tmp_path / "small_image_upscaled.pdf"
    worker = AppWorker("merge_save")

    worker._run_merge_save([image_item(image_path)], str(output))

    with fitz.open(output) as document:
        page = document[0]
        xref = page.get_images(full=True)[0][0]
        rect = page.get_image_rects(xref)[0]
        assert rect.width > 100
        assert rect.height > 50


def test_merge_can_rotate_image_page(tmp_path):
    image_path = tmp_path / "photo.png"
    Image.new("RGB", (100, 200), "green").save(image_path)
    output = tmp_path / "rotated_image.pdf"
    worker = AppWorker("merge_save")

    worker._run_merge_save([image_item(image_path, rotation=90)], str(output))

    with fitz.open(output) as document:
        assert document.page_count == 1
        assert document[0].rotation == 90


def test_merge_can_remove_pdf_annotations(tmp_path):
    source = tmp_path / "annotated.pdf"
    output = tmp_path / "without_annotations.pdf"
    document = fitz.open()
    page = document.new_page(width=300, height=400)
    page.insert_text((36, 72), "BODY")
    page.add_text_annot((120, 120), "comment")
    document.save(source)
    document.close()
    worker = AppWorker("merge_save")

    worker._run_merge_save(
        [pdf_item(source, 0)],
        str(output),
        remove_pdf_annotations=True,
    )

    with fitz.open(output) as merged:
        assert list(merged[0].annots() or []) == []
    with fitz.open(source) as original:
        assert len(list(original[0].annots() or [])) == 1


def test_merge_keeps_pdf_annotations_by_default(tmp_path):
    source = tmp_path / "annotated.pdf"
    output = tmp_path / "with_annotations.pdf"
    document = fitz.open()
    page = document.new_page(width=300, height=400)
    page.insert_text((36, 72), "BODY")
    page.add_text_annot((120, 120), "comment")
    document.save(source)
    document.close()
    worker = AppWorker("merge_save")

    worker._run_merge_save([pdf_item(source, 0)], str(output))

    with fitz.open(output) as merged:
        assert len(list(merged[0].annots() or [])) == 1


def test_build_items_data_from_file_paths_expands_pdf_and_classifies_images(
    pdf_factory, tmp_path
):
    pdf = pdf_factory("source.pdf", ["A", "B"])
    image = tmp_path / "photo.png"
    Image.new("RGB", (100, 50), "blue").save(image)
    unsupported = tmp_path / "memo.txt"
    unsupported.write_text("skip", encoding="utf-8")
    worker = AppWorker("batch_merge_subfolders")

    result = worker._build_items_data_from_file_paths(
        [str(pdf), str(image), str(unsupported)]
    )

    assert [item["type"] for item in result["items_data"]] == [
        "pdf",
        "pdf",
        "image",
    ]
    assert [item.get("page_num") for item in result["items_data"]] == [0, 1, 0]
    assert result["skipped_files"] == [str(unsupported)]


def test_batch_merge_subfolders_creates_pdf_per_subfolder_and_utf8_sig_log(
    pdf_factory, tmp_path
):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    first_dir = input_root / "A案件"
    second_dir = input_root / "B案件"
    first_dir.mkdir(parents=True)
    second_dir.mkdir()
    pdf_factory("dummy.pdf", ["unused"])
    source_pdf = pdf_factory("source.pdf", ["PDF"])
    (first_dir / "001.pdf").write_bytes(source_pdf.read_bytes())
    Image.new("RGB", (400, 200), "blue").save(first_dir / "002.jpg")
    (first_dir / "999.txt").write_text("unsupported", encoding="utf-8")
    (second_dir / "memo.txt").write_text("no target", encoding="utf-8")
    worker = AppWorker("batch_merge_subfolders")
    finished = collect_signal(worker.signals.finished)

    worker._run_batch_merge_subfolders(str(input_root), str(output_root))

    assert (output_root / "A案件.pdf").exists()
    assert not (output_root / "B案件.pdf").exists()
    with fitz.open(output_root / "A案件.pdf") as document:
        assert document.page_count == 2
    log_path = next(output_root.glob("batch_log_*.csv"))
    assert log_path.read_bytes().startswith(b"\xef\xbb\xbf")
    with open(log_path, encoding="utf-8-sig", newline="") as log_file:
        rows = list(csv.DictReader(log_file))
    assert [row["結果"] for row in rows] == ["Warning", "Skipped"]
    assert rows[0]["入力サブフォルダ"].endswith("A案件")
    assert rows[0]["スキップファイル数"] == "1"
    assert "未対応ファイル" in rows[0]["メッセージ"]
    assert "成功: 0件" in finished[-1][2]
    assert "警告: 1件" in finished[-1][2]


def test_batch_merge_subfolders_skips_existing_pdf_when_overwrite_is_off(
    pdf_factory, tmp_path
):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    case_dir = input_root / "case"
    case_dir.mkdir(parents=True)
    output_root.mkdir()
    source_pdf = pdf_factory("source.pdf", ["NEW"])
    (case_dir / "001.pdf").write_bytes(source_pdf.read_bytes())
    existing = output_root / "case.pdf"
    existing.write_bytes(b"existing")
    worker = AppWorker("batch_merge_subfolders")

    worker._run_batch_merge_subfolders(
        str(input_root), str(output_root), overwrite=False
    )

    assert existing.read_bytes() == b"existing"
    log_path = next(output_root.glob("batch_log_*.csv"))
    with open(log_path, encoding="utf-8-sig", newline="") as log_file:
        rows = list(csv.DictReader(log_file))
    assert rows[0]["結果"] == "Skipped"
    assert "同名PDF" in rows[0]["メッセージ"]


def test_batch_merge_subfolders_saves_directly_when_output_does_not_exist(
    pdf_factory, tmp_path, monkeypatch
):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    case_dir = input_root / "case"
    case_dir.mkdir(parents=True)
    source_pdf = pdf_factory("source.pdf", ["BODY"])
    (case_dir / "001.pdf").write_bytes(source_pdf.read_bytes())
    worker = AppWorker("batch_merge_subfolders")

    def fail_replace(_src, _dst):
        raise AssertionError("os.replace should not be used for new output PDFs")

    monkeypatch.setattr(os, "replace", fail_replace)

    worker._run_batch_merge_subfolders(str(input_root), str(output_root))

    assert (output_root / "case.pdf").exists()
    assert not list(output_root.glob("*.tmp.pdf"))


def test_batch_merge_subfolders_follows_bookmark_options(pdf_factory, tmp_path):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    case_dir = input_root / "case"
    case_dir.mkdir(parents=True)
    first_pdf = pdf_factory("001.pdf", ["FIRST"])
    second_pdf = pdf_factory("002.pdf", ["SECOND"])
    (case_dir / "001.pdf").write_bytes(first_pdf.read_bytes())
    (case_dir / "002.pdf").write_bytes(second_pdf.read_bytes())
    worker = AppWorker("batch_merge_subfolders")

    worker._run_batch_merge_subfolders(
        str(input_root),
        str(output_root),
        auto_bookmarks_enabled=True,
        show_bookmarks_on_open=True,
    )

    with fitz.open(output_root / "case.pdf") as document:
        assert [entry[1] for entry in document.get_toc()] == ["001", "002"]
        assert document.pagemode == "UseOutlines"


def test_batch_merge_subfolders_can_disable_auto_bookmarks(pdf_factory, tmp_path):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    case_dir = input_root / "case"
    case_dir.mkdir(parents=True)
    source_pdf = pdf_factory("001.pdf", ["FIRST"])
    (case_dir / "001.pdf").write_bytes(source_pdf.read_bytes())
    worker = AppWorker("batch_merge_subfolders")

    worker._run_batch_merge_subfolders(
        str(input_root),
        str(output_root),
        auto_bookmarks_enabled=False,
        show_bookmarks_on_open=False,
    )

    with fitz.open(output_root / "case.pdf") as document:
        assert document.get_toc() == []
        assert document.pagemode in ("UseNone", "")


def test_batch_merge_subfolders_uses_temp_replace_when_overwriting_existing_pdf(
    pdf_factory, tmp_path, monkeypatch
):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    case_dir = input_root / "case"
    case_dir.mkdir(parents=True)
    output_root.mkdir()
    source_pdf = pdf_factory("source.pdf", ["BODY"])
    (case_dir / "001.pdf").write_bytes(source_pdf.read_bytes())
    existing = output_root / "case.pdf"
    existing.write_bytes(b"old")
    calls = []
    real_replace = os.replace
    worker = AppWorker("batch_merge_subfolders")

    def record_replace(src, dst):
        calls.append((Path(src).name, Path(dst).name))
        real_replace(src, dst)

    monkeypatch.setattr(os, "replace", record_replace)

    worker._run_batch_merge_subfolders(
        str(input_root), str(output_root), overwrite=True
    )

    assert len(calls) == 1
    assert calls[0][0].startswith(".case_")
    assert calls[0][0].endswith(".tmp.pdf")
    assert calls[0][1] == "case.pdf"
    assert existing.read_bytes() != b"old"


def test_merge_can_add_header_footer_and_page_number(pdf_factory, tmp_path):
    source = pdf_factory("source.pdf", ["BODY"])
    output = tmp_path / "decorated.pdf"
    worker = AppWorker("merge_save")
    settings = {
        "header_enabled": True,
        "footer_enabled": True,
        "font_size": 10,
        "header": {
            "left": "HEADER",
            "center": "",
            "right": "",
            "auto_date": False,
            "auto_page_number": False,
        },
        "footer": {
            "left": "",
            "center": "FOOTER",
            "right": "",
            "auto_date": False,
            "auto_page_number": True,
            "page_number_position": "right",
        },
        "page_number_format": "Page 1",
        "page_number_start": 1,
    }

    worker._run_merge_save(
        [pdf_item(source, 0)],
        str(output),
        header_footer_settings=settings,
    )

    with fitz.open(output) as document:
        text = document[0].get_text()
        assert "BODY" in text
        assert "HEADER" in text
        assert "FOOTER" in text
        assert "Page 1" in text


def test_merge_uses_mocked_office_conversion_and_removes_temporary_pdf(
    pdf_factory, tmp_path, monkeypatch
):
    office_source = tmp_path / "report.docx"
    office_source.write_bytes(b"placeholder")
    converted = pdf_factory("converted.pdf", ["OFFICE-PAGE"])
    output = tmp_path / "office-merged.pdf"
    worker = AppWorker("merge_save")

    class FakeWord:
        quit_called = False

        def Quit(self):
            self.quit_called = True

    fake_word = FakeWord()
    monkeypatch.setattr(
        worker,
        "_convert_word_to_pdf",
        lambda path, app, suppress_errors, suppress_office_markup=False: (
            str(converted),
            fake_word,
        ),
    )

    worker._run_merge_save(
        [
            {
                "type": "word",
                "path": str(office_source),
                "original_path": str(office_source),
                "rotation": 0,
            }
        ],
        str(output),
    )

    assert fake_word.quit_called is True
    assert not converted.exists()
    with fitz.open(output) as document:
        assert document.page_count == 1
        assert "OFFICE-PAGE" in document[0].get_text()


def test_failed_office_conversion_is_reported_but_pdf_is_saved(
    pdf_factory, tmp_path, monkeypatch
):
    pdf_source = pdf_factory("source.pdf", ["PDF-PAGE"])
    office_source = tmp_path / "failed.docx"
    office_source.write_bytes(b"placeholder")
    output = tmp_path / "partial.pdf"
    worker = AppWorker("merge_save")
    finished = collect_signal(worker.signals.finished)
    monkeypatch.setattr(
        worker,
        "_convert_word_to_pdf",
        lambda path, app, suppress_errors, suppress_office_markup=False: (None, None),
    )

    worker._run_merge_save(
        [
            {
                "type": "word",
                "path": str(office_source),
                "original_path": str(office_source),
                "rotation": 0,
            },
            pdf_item(pdf_source, 0),
        ],
        str(output),
    )

    with fitz.open(output) as document:
        assert document.page_count == 1
        assert "PDF-PAGE" in document[0].get_text()
    assert "failed.docx" in finished[-1][2]
    assert "除外しました" in finished[-1][2]


def test_merge_passes_suppress_office_markup_to_office_converter(
    pdf_factory, tmp_path, monkeypatch
):
    office_source = tmp_path / "report.docx"
    office_source.write_bytes(b"placeholder")
    converted = pdf_factory("converted_markup.pdf", ["OFFICE-PAGE"])
    output = tmp_path / "office-merged.pdf"
    worker = AppWorker("merge_save")
    calls = []

    class FakeWord:
        def Quit(self):
            pass

    def fake_convert(path, app, suppress_errors, suppress_office_markup=False):
        calls.append(suppress_office_markup)
        return str(converted), FakeWord()

    monkeypatch.setattr(worker, "_convert_word_to_pdf", fake_convert)

    worker._run_merge_save(
        [
            {
                "type": "word",
                "path": str(office_source),
                "original_path": str(office_source),
                "rotation": 0,
            }
        ],
        str(output),
        suppress_office_markup=True,
    )

    assert calls == [True]


def test_export_image_creates_jpeg_and_avoids_overwrite(pdf_factory, tmp_path):
    source = pdf_factory("scan.pdf", ["IMAGE"])
    output_dir = tmp_path / "images"
    worker = AppWorker("export_images")
    finished = collect_signal(worker.signals.finished)

    worker._run_export_images(
        [pdf_item(source, 0)], str(output_dir), dpi=72, image_format="JPEG"
    )
    worker._run_export_images(
        [pdf_item(source, 0)], str(output_dir), dpi=72, image_format="JPEG"
    )

    outputs = sorted(path.name for path in output_dir.glob("*.jpg"))
    assert outputs == ["scan_p001.jpg", "scan_p001_001.jpg"]
    assert len(finished) == 2
    assert all(Path(output_dir / name).stat().st_size > 0 for name in outputs)


@pytest.mark.parametrize("dpi", [96, 150, 300, 600])
def test_export_image_dimensions_follow_selected_dpi(pdf_factory, tmp_path, dpi):
    source = pdf_factory("page.pdf", ["IMAGE"])
    output_dir = tmp_path / f"images-{dpi}"
    worker = AppWorker("export_images")

    worker._run_export_images(
        [pdf_item(source, 0)], str(output_dir), dpi=dpi, image_format="JPEG"
    )

    with Image.open(output_dir / "page_p001.jpg") as image:
        expected_width = 300 * dpi / 72
        expected_height = 400 * dpi / 72
        assert abs(image.width - expected_width) <= 1
        assert abs(image.height - expected_height) <= 1
