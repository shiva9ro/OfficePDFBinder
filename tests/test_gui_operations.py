import copy

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QListWidgetItem, QMessageBox

import OfficePDFBinder_Main as app_module


def item_data(path, page_num=0, item_type="pdf", rotation=0):
    data = {
        "type": item_type,
        "path": str(path),
        "original_path": str(path),
        "rotation": rotation,
    }
    if item_type == "pdf":
        data["page_num"] = page_num
    return data


def add_raw_item(window, data):
    item = QListWidgetItem()
    item.setData(Qt.ItemDataRole.UserRole, copy.deepcopy(data))
    window.page_list_widget.addItem(item)
    return item


def current_page_order(window):
    return [
        window.page_list_widget.item(row)
        .data(Qt.ItemDataRole.UserRole)
        .get("page_num")
        for row in range(window.page_list_widget.count())
    ]


def test_move_multiple_selected_pages_preserves_relative_order(
    main_window, pdf_factory
):
    source = pdf_factory("pages.pdf", ["1", "2", "3", "4"])
    items = [add_raw_item(main_window, item_data(source, page)) for page in range(4)]
    main_window._record_history_change()
    items[1].setSelected(True)
    items[2].setSelected(True)

    main_window._move_down()

    assert current_page_order(main_window) == [0, 3, 1, 2]
    assert sorted(
        main_window.page_list_widget.row(item)
        for item in main_window.page_list_widget.selectedItems()
    ) == [2, 3]


def test_move_to_top_and_bottom(main_window, pdf_factory):
    source = pdf_factory("pages.pdf", ["1", "2", "3", "4"])
    items = [add_raw_item(main_window, item_data(source, page)) for page in range(4)]
    items[2].setSelected(True)

    main_window._move_to_top()
    assert current_page_order(main_window) == [2, 0, 1, 3]

    main_window.page_list_widget.clearSelection()
    main_window.page_list_widget.item(0).setSelected(True)
    main_window._move_to_bottom()
    assert current_page_order(main_window) == [0, 1, 3, 2]


def test_rotate_affects_only_pdf_and_wraps_at_360(main_window, pdf_factory, tmp_path):
    source = pdf_factory("page.pdf", ["1"])
    pdf = add_raw_item(main_window, item_data(source, rotation=270))
    office = add_raw_item(
        main_window,
        item_data(tmp_path / "document.docx", item_type="word"),
    )
    pdf.setSelected(True)
    office.setSelected(True)

    main_window._rotate_selected(90)

    assert pdf.data(Qt.ItemDataRole.UserRole)["rotation"] == 0
    assert office.data(Qt.ItemDataRole.UserRole)["rotation"] == 0


def test_delete_selected_removes_orphan_bookmark(
    main_window, pdf_factory, monkeypatch
):
    source = pdf_factory("pages.pdf", ["1", "2"])
    first = add_raw_item(main_window, item_data(source, 0))
    add_raw_item(main_window, item_data(source, 1))
    main_window.bookmarks = [
        {
            "title": "First",
            "path": str(source),
            "page_num": 0,
            "auto": False,
        }
    ]
    first.setSelected(True)
    monkeypatch.setattr(
        QMessageBox,
        "question",
        lambda *args, **kwargs: QMessageBox.StandardButton.Yes,
    )

    main_window._delete_selected()

    assert current_page_order(main_window) == [1]
    assert main_window.bookmarks == [
        {
            "title": "pages",
            "path": str(source),
            "page_num": 1,
            "auto": True,
        }
    ]


def test_undo_and_redo_restore_rotation(main_window, pdf_factory):
    source = pdf_factory("page.pdf", ["1"])
    item = add_raw_item(main_window, item_data(source, 0))
    main_window._record_history_change()
    item.setSelected(True)

    main_window._rotate_selected(90)
    assert main_window.page_list_widget.item(0).data(Qt.UserRole)["rotation"] == 90
    assert main_window.undo_action.isEnabled()

    main_window._undo()
    assert main_window.page_list_widget.item(0).data(Qt.UserRole)["rotation"] == 0
    assert main_window.redo_action.isEnabled()

    main_window._redo()
    assert main_window.page_list_widget.item(0).data(Qt.UserRole)["rotation"] == 90


def test_auto_bookmarks_are_created_once_per_file(main_window, pdf_factory):
    first = pdf_factory("first.pdf", ["1", "2"])
    second = pdf_factory("second.pdf", ["1"])
    add_raw_item(main_window, item_data(first, 0))
    add_raw_item(main_window, item_data(first, 1))
    add_raw_item(main_window, item_data(second, 0))

    main_window._generate_bookmarks_from_list()

    assert [(b["title"], b["page_num"]) for b in main_window.bookmarks] == [
        ("first", 0),
        ("second", 0),
    ]
    assert all(bookmark["auto"] for bookmark in main_window.bookmarks)


def test_manual_bookmark_prevents_duplicate_auto_bookmark(main_window, pdf_factory):
    source = pdf_factory("first.pdf", ["1"])
    add_raw_item(main_window, item_data(source, 0))
    main_window.bookmarks = [
        {
            "title": "Custom",
            "path": str(source),
            "page_num": 0,
            "auto": False,
        }
    ]

    main_window._generate_bookmarks_from_list()

    assert main_window.bookmarks == [
        {
            "title": "Custom",
            "path": str(source),
            "page_num": 0,
            "auto": False,
        }
    ]


def test_duplicate_check_returns_only_new_files(
    main_window, pdf_factory, tmp_path, monkeypatch
):
    existing = pdf_factory("existing.pdf", ["1"])
    new_file = tmp_path / "new.pdf"
    add_raw_item(main_window, item_data(existing, 0))
    monkeypatch.setattr(
        QMessageBox,
        "question",
        lambda *args, **kwargs: QMessageBox.StandardButton.Yes,
    )

    result = main_window._check_duplicate_files([str(existing), str(new_file)])

    assert result == [str(new_file)]


def test_ipc_queue_uses_natural_filename_order(
    main_window, tmp_path, monkeypatch
):
    paths = [
        str(tmp_path / "document10.pdf"),
        str(tmp_path / "document2.pdf"),
        str(tmp_path / "document1.pdf"),
    ]
    main_window._ipc_pending_file_paths = paths.copy()
    calls = []
    monkeypatch.setattr(main_window, "_check_duplicate_files", lambda value: value)
    monkeypatch.setattr(
        main_window,
        "_run_task",
        lambda task, label, **kwargs: calls.append((task, label, kwargs)),
    )

    main_window._flush_ipc_pending_files()

    assert main_window._ipc_pending_file_paths == []
    assert calls[0][2]["file_paths"] == [
        str(tmp_path / "document1.pdf"),
        str(tmp_path / "document2.pdf"),
        str(tmp_path / "document10.pdf"),
    ]


def test_status_bar_counts_pdf_pages_and_unique_office_files(
    main_window, pdf_factory, tmp_path
):
    pdf = pdf_factory("pages.pdf", ["1", "2"])
    word = tmp_path / "report.docx"
    add_raw_item(main_window, item_data(pdf, 0))
    add_raw_item(main_window, item_data(pdf, 1))
    add_raw_item(main_window, item_data(word, item_type="word"))
    add_raw_item(main_window, item_data(word, item_type="word"))
    main_window.page_list_widget.item(0).setSelected(True)

    main_window.update_status_bar()

    assert main_window.file_status_label.text() == "PDFページ: 2 | Word: 1"
    assert main_window.selection_status_label.text() == "選択中: 1件"


def test_history_is_limited_to_configured_maximum(main_window):
    for index in range(app_module.MAX_HISTORY + 5):
        main_window.auto_bookmarks_enabled = bool(index % 2)
        main_window._record_history_change()

    assert len(main_window.undo_stack) == app_module.MAX_HISTORY


def test_cancel_worker_marks_current_worker_and_updates_dialog(main_window):
    class DummyWorker:
        is_running = True

    main_window.current_worker = DummyWorker()
    main_window.setup_progress_dialog("テスト")

    main_window.cancel_worker()

    assert main_window.current_worker.is_running is False
    assert main_window.progress_dialog.labelText() == "処理を中止しています..."


def test_action_states_follow_selection_boundaries(main_window, pdf_factory):
    source = pdf_factory("pages.pdf", ["1", "2"])
    first = add_raw_item(main_window, item_data(source, 0))
    add_raw_item(main_window, item_data(source, 1))

    main_window._update_page_mode_actions_state()
    assert main_window.merge_action.isEnabled()
    assert not main_window.delete_action.isEnabled()

    first.setSelected(True)
    main_window._update_page_mode_actions_state()
    assert main_window.delete_action.isEnabled()
    assert not main_window.move_up_action.isEnabled()
    assert main_window.move_down_action.isEnabled()

    main_window.page_list_widget.clearSelection()
    main_window.page_list_widget.item(1).setSelected(True)
    main_window._update_page_mode_actions_state()
    assert main_window.move_up_action.isEnabled()
    assert not main_window.move_down_action.isEnabled()
