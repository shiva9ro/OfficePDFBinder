"""Regression cases found by tracing UI operations through their saved output."""

import configparser
import copy
from pathlib import Path

import fitz
import pytest
from PIL import Image
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialogButtonBox, QFileDialog, QInputDialog, QListWidgetItem, QMessageBox

import OfficePDFBinder_Main as app


def item(path, page=0, kind="pdf", rotation=0):
    return dict(type=kind, path=str(path), original_path=str(path),
                page_num=page, rotation=rotation)


def add_item(window, data):
    widget_item = QListWidgetItem()
    widget_item.setData(Qt.UserRole, data)
    window.page_list_widget.addItem(widget_item)
    return widget_item


def test_rotated_pdf_image_export_matches_pdf_export(tmp_path):
    source = tmp_path / "rotated.pdf"
    with fitz.open() as doc:
        page = doc.new_page(width=200, height=300)
        page.insert_text((20, 50), "TOP")
        page.set_rotation(90)
        doc.save(source)
    worker = app.AppWorker("merge_save")
    worker._run_merge_save([item(source, rotation=90)], str(tmp_path / "merged.pdf"))
    worker._run_export_images([item(source, rotation=90)], str(tmp_path / "images"),
                              dpi=72, image_format="PNG")
    with fitz.open(tmp_path / "merged.pdf") as doc:
        expected = doc[0].get_pixmap()
    with Image.open(tmp_path / "images" / "rotated_p001.png") as image:
        assert image.size == (expected.width, expected.height)
        assert image.convert("RGB").tobytes() == expected.samples


def test_svg_image_export_reaches_worker(main_window, tmp_path, monkeypatch):
    source = tmp_path / "drawing.svg"
    source.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"/>')
    add_item(main_window, item(source, kind="svg")).setSelected(True)
    calls = []
    monkeypatch.setattr(QMessageBox, "information", lambda *_args: None)
    monkeypatch.setattr(QInputDialog, "getItem", lambda *_args: ("150 dpi", True))
    monkeypatch.setattr(QFileDialog, "getExistingDirectory", lambda *_args: str(tmp_path))
    monkeypatch.setattr(main_window, "_run_task", lambda *args, **kwargs: calls.append(kwargs))
    main_window._export_selected_as_images()
    assert calls and calls[0]["items_data"][0]["type"] == "svg"


def test_interleaved_tiff_bookmark_targets_its_own_frame(pdf_factory, tmp_path):
    source = tmp_path / "frames.tiff"
    Image.new("RGB", (30, 30), "red").save(
        source, save_all=True, append_images=[Image.new("RGB", (30, 30), "blue")])
    separator = pdf_factory("separator.pdf", ["separator"])
    output = tmp_path / "merged.pdf"
    app.AppWorker("merge_save")._run_merge_save(
        [item(source, 0, "image"), item(separator), item(source, 1, "image")],
        str(output), bookmarks=[dict(title="Blue", path=str(source), page_num=1)])
    with fitz.open(output) as doc:
        assert doc.get_toc() == [[1, "Blue", 3]]


def test_failed_image_does_not_leave_bookmark_on_next_pdf(pdf_factory, tmp_path):
    broken = tmp_path / "broken.png"
    broken.write_bytes(b"not an image")
    source = pdf_factory("good.pdf", ["good"])
    output = tmp_path / "merged.pdf"
    errors = []
    app.AppWorker("merge_save")._run_merge_save(
        [item(broken, kind="image"), item(source)], str(output),
        bookmarks=[dict(title="Broken image", path=str(broken), page_num=0)],
        collected_errors=errors)
    assert errors
    with fitz.open(output) as doc:
        assert doc.page_count == 1
        assert doc.get_toc() == []


@pytest.mark.parametrize("source_rotation,extra_rotation", [(0, 0), (90, 90), (270, 180)])
def test_pdf_internal_links_survive_reorder_and_interleaving(
    pdf_factory, tmp_path, source_rotation, extra_rotation
):
    source = tmp_path / "links.pdf"
    with fitz.open() as doc:
        for _ in range(3):
            doc.new_page(width=200, height=300)
        doc[0].insert_link({"kind": fitz.LINK_GOTO, "from": fitz.Rect(10, 10, 60, 30),
                            "page": 2, "to": fitz.Point(20, 40)})
        doc[0].set_rotation(source_rotation)
        doc[2].set_rotation(source_rotation)
        doc.save(source)
    separator = pdf_factory("separator.pdf", ["separator"])
    output = tmp_path / "merged.pdf"
    app.AppWorker("merge_save")._run_merge_save(
        [item(source, 2, rotation=extra_rotation), item(separator),
         item(source, 0, rotation=extra_rotation), item(source, 1)], str(output))
    with fitz.open(output) as doc:
        links = doc[2].get_links()
        assert len(links) == 1
        assert links[0]["kind"] == fitz.LINK_GOTO
        assert links[0]["page"] == 0
        assert links[0]["to"] * doc[0].derotation_matrix == fitz.Point(20, 40)
        assert links[0]["from"] * doc[2].derotation_matrix == fitz.Rect(10, 10, 60, 30)


def test_cancelled_add_records_partial_result_for_undo(main_window, pdf_factory):
    source = pdf_factory("partial.pdf", ["1"])
    worker = app.AppWorker("add_files")
    worker.is_running = False
    main_window.current_worker = worker
    add_item(main_window, item(source))
    main_window._on_worker_run_completed(worker)
    assert main_window.merge_action.isEnabled()
    main_window._undo()
    assert main_window.page_list_widget.count() == 0


def test_queued_files_wait_for_worker_exit(main_window, monkeypatch, qtbot):
    worker = app.AppWorker("add_files")
    main_window.current_worker = worker
    main_window._pending_file_paths = ["second.pdf"]
    calls = []
    monkeypatch.setattr(main_window, "_check_duplicate_files", lambda paths: paths)
    monkeypatch.setattr(main_window, "_run_task", lambda *args, **kwargs: calls.append(kwargs))
    main_window.on_worker_finished("add_files", "Done", "Done")
    assert calls == []  # finished is emitted while worker.run() is still active.
    main_window._on_worker_run_completed(worker)
    qtbot.waitUntil(lambda: bool(calls))
    assert calls[0]["file_paths"] == ["second.pdf"]


def test_new_project_undo_restores_header_footer(main_window, pdf_factory, monkeypatch):
    source = pdf_factory("source.pdf", ["1"])
    add_item(main_window, item(source))
    original = {"header_enabled": True, "header": {"left": "Keep this header"}}
    main_window.header_footer_settings = copy.deepcopy(original)
    main_window.page_number_settings["enabled"] = True
    main_window._record_history_change()
    monkeypatch.setattr(app, "_show_standard_question", lambda *_args: QMessageBox.Yes)
    main_window._new_project()
    main_window._undo()
    assert main_window.header_footer_settings == original
    assert main_window.page_number_settings["enabled"]


def test_percent_in_header_does_not_prevent_saving_settings(main_window):
    main_window.header_footer_settings = {"header": {"left": "100% complete"}}
    main_window._save_settings()
    config = configparser.ConfigParser(interpolation=None)
    config.read(main_window.settings_file, encoding="utf-8")
    assert config.get("HeaderFooter", "header_left") == "100% complete"


def test_duplicate_paths_in_same_request_are_not_added_twice(main_window, tmp_path, monkeypatch):
    source = tmp_path / "Case.pdf"
    monkeypatch.setattr(app, "_show_standard_question", lambda *_args: QMessageBox.Yes)
    assert main_window._check_duplicate_files([str(source), str(source).upper()]) == [str(source)]


def test_batch_excludes_its_output_subfolder(pdf_factory, tmp_path):
    root = tmp_path / "input"
    case = root / "case"
    output = root / "output"
    case.mkdir(parents=True)
    source = pdf_factory("source.pdf", ["1"])
    (case / "source.pdf").write_bytes(source.read_bytes())
    worker = app.AppWorker("batch_merge_subfolders")
    worker._run_batch_merge_subfolders(str(root), str(output))
    assert (output / "case.pdf").exists()
    assert not (output / "output.pdf").exists()


def test_save_clear_also_clears_visible_bookmarks(main_window, pdf_factory, monkeypatch):
    source = pdf_factory("source.pdf", ["1"])
    add_item(main_window, item(source))
    main_window.bookmarks = [dict(title="Old", path=str(source), page_num=0)]
    main_window._update_bookmark_tree()
    monkeypatch.setattr(main_window, "_show_copyable_message", lambda *args, **kwargs: QDialogButtonBox.Yes)
    main_window.on_worker_finished("merge_save", "Saved", "No output path")
    assert main_window.page_list_widget.count() == 0
    assert main_window.bookmark_tree.topLevelItemCount() == 0


def test_missing_source_page_does_not_silently_duplicate_last_page(pdf_factory, tmp_path):
    source = pdf_factory("source.pdf", ["1"])
    output = pdf_factory("existing.pdf", ["KEEP"])
    before = output.read_bytes()
    with pytest.raises(ValueError, match="Source page no longer exists"):
        app.AppWorker("merge_save")._run_merge_save_transaction(
            [item(source, 0), item(source, 1)], str(output))
    assert output.read_bytes() == before
    assert not list(tmp_path.glob("*.tmp.pdf"))


def test_link_to_removed_page_is_omitted_but_self_and_web_links_remain(tmp_path):
    source = tmp_path / "links.pdf"
    with fitz.open() as doc:
        doc.new_page()
        doc.new_page()
        for target in (0, 1):
            doc[0].insert_link({"kind": fitz.LINK_GOTO, "from": fitz.Rect(10, 10, 30, 30),
                                "page": target, "to": fitz.Point(20, 40)})
        doc[0].insert_link({"kind": fitz.LINK_URI, "from": fitz.Rect(50, 50, 70, 70),
                            "uri": "https://example.com/"})
        doc.save(source)
    output = tmp_path / "selected.pdf"
    app.AppWorker("merge_save")._run_merge_save([item(source)], str(output))
    with fitz.open(output) as doc:
        links = doc[0].get_links()
        assert len(links) == 2
        assert [link["page"] for link in links if link["kind"] == fitz.LINK_GOTO] == [0]
        assert [link["uri"] for link in links if link["kind"] == fitz.LINK_URI] == ["https://example.com/"]


@pytest.mark.parametrize("partial", [False, True])
def test_overwriting_input_clears_list_and_history_without_asking(
    main_window, pdf_factory, monkeypatch, partial
):
    source = pdf_factory("source.pdf", ["ONE", "TWO"])
    for page in (1, 0):
        add_item(main_window, item(source, page))
    main_window.bookmarks = [dict(title="Second", path=str(source), page_num=1)]
    main_window._update_bookmark_tree()
    main_window._record_history_change()
    main_window.redo_stack.append(main_window._create_state_snapshot())
    calls = []

    def notify(title, message, **kwargs):
        # Clear before displaying the completion notification.
        assert main_window.page_list_widget.count() == 0
        calls.append((message, kwargs))
        return QDialogButtonBox.Ok

    monkeypatch.setattr(main_window, "_show_copyable_message", notify)
    monkeypatch.setattr(app.os, "startfile", lambda *_args: None)
    worker = app.AppWorker("merge_save")
    main_window.current_worker = worker
    worker.signals.finished.connect(main_window.on_worker_finished)
    worker._run_merge_save_transaction(
        [item(source, 1)] if partial else [item(source, 1), item(source, 0)],
        str(source), bookmarks=[])
    assert calls[0][1]["buttons"] == QDialogButtonBox.Ok
    assert "入力PDFに上書きしたため" in calls[0][0]
    assert not main_window.bookmarks
    assert main_window.bookmark_tree.topLevelItemCount() == 0
    assert not main_window.undo_action.isEnabled()
    assert not main_window.redo_action.isEnabled()
    main_window._undo()
    main_window._redo()
    assert main_window.page_list_widget.count() == 0
    with fitz.open(source) as doc:
        assert [p.get_text().strip() for p in doc] == (["TWO"] if partial else ["TWO", "ONE"])
    main_window._on_worker_run_completed(worker)


@pytest.mark.parametrize("existing", [False, True])
def test_other_output_keeps_optional_clear(main_window, pdf_factory, tmp_path, monkeypatch, existing):
    source = pdf_factory("source.pdf", ["ONE"])
    output = pdf_factory("other.pdf", ["OLD"]) if existing else tmp_path / "new.pdf"
    add_item(main_window, item(source))
    main_window._record_history_change()
    calls = []
    monkeypatch.setattr(app.os, "startfile", lambda *_args: None)
    monkeypatch.setattr(main_window, "_show_copyable_message",
                        lambda *args, **kwargs: calls.append(kwargs) or QDialogButtonBox.No)
    worker = app.AppWorker("merge_save")
    main_window.current_worker = worker
    worker.signals.finished.connect(main_window.on_worker_finished)
    worker._run_merge_save_transaction([item(source)], str(output), bookmarks=[])
    assert calls[0]["buttons"] == QDialogButtonBox.Yes | QDialogButtonBox.No
    assert main_window.page_list_widget.count() == 1
    assert main_window.undo_action.isEnabled()
    main_window._on_worker_run_completed(worker)


@pytest.mark.parametrize("failure", ["replace_error", "cancel"])
def test_failed_input_overwrite_keeps_editing_state(main_window, pdf_factory, monkeypatch, failure):
    source = pdf_factory("source.pdf", ["ONE", "TWO"])
    before = source.read_bytes()
    add_item(main_window, item(source, 1))
    main_window._record_history_change()
    snapshot = main_window._create_state_snapshot()
    worker = app.AppWorker("merge_save")
    main_window.current_worker = worker
    if failure == "replace_error":
        def fail_replace(*_args):
            raise PermissionError("locked")
        monkeypatch.setattr(app.os, "replace", fail_replace)
        with pytest.raises(PermissionError):
            worker._run_merge_save_transaction([item(source, 1)], str(source))
    else:
        original_merge = worker._run_merge_save
        def cancel_after_generation(*args, **kwargs):
            result = original_merge(*args, **kwargs)
            worker.is_running = False
            return result
        monkeypatch.setattr(worker, "_run_merge_save", cancel_after_generation)
        worker._run_merge_save_transaction([item(source, 1)], str(source))
    main_window._on_worker_run_completed(worker)
    assert worker.saved_output_path is None
    assert source.read_bytes() == before
    assert main_window._create_state_snapshot() == snapshot


@pytest.mark.parametrize("history", ["undo_stack", "redo_stack"])
def test_overwrite_clears_source_references_in_history(main_window, pdf_factory, history):
    source = pdf_factory("source.pdf", ["ONE"])
    add_item(main_window, item(source))
    getattr(main_window, history).append(main_window._create_state_snapshot())
    main_window.page_list_widget.clear()
    assert main_window._clear_overwritten_input_state(str(source).upper())
    main_window._undo()
    main_window._redo()
    assert main_window.page_list_widget.count() == 0


def test_successful_overwrite_clears_even_when_completion_notification_is_skipped(
    main_window, pdf_factory
):
    source = pdf_factory("source.pdf", ["ONE", "TWO"])
    add_item(main_window, item(source, 1))
    main_window._record_history_change()
    worker = app.AppWorker("merge_save")
    main_window.current_worker = worker
    worker._run_merge_save_transaction([item(source, 1)], str(source), emit_completion=False)
    worker.is_running = False  # Cancellation arrived after the successful replacement.
    main_window._on_worker_run_completed(worker)
    assert main_window.page_list_widget.count() == 0
    assert not main_window.undo_action.isEnabled()
