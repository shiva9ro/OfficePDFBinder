from pathlib import Path

import fitz
import pytest

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
        lambda path, app, suppress_errors: (str(converted), fake_word),
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
        lambda path, app, suppress_errors: (None, None),
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
