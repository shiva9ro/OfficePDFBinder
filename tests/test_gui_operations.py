import copy

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialogButtonBox,
    QFileDialog,
    QInputDialog,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
)

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


def test_batch_dialog_output_tracks_input_until_user_edits(qtbot, tmp_path):
    initial = str(tmp_path / "initial")
    dialog = app_module.BatchMergeSubfoldersDialog(initial_path=initial)
    qtbot.addWidget(dialog)

    first_input = str(tmp_path / "input1")
    dialog.input_root_edit.setText(first_input)
    assert dialog.output_root_edit.text() == first_input

    custom_output = str(tmp_path / "custom-output")
    dialog.output_root_edit.setText(custom_output)
    dialog.output_root_edit.textEdited.emit(custom_output)

    second_input = str(tmp_path / "input2")
    dialog.input_root_edit.setText(second_input)
    assert dialog.output_root_edit.text() == custom_output


def test_page_list_selection_style_has_no_filename_border(main_window):
    style = main_window.page_list_widget.styleSheet()

    assert "QListWidget::item:selected" in style
    assert "border: 2px" not in style
    assert "outline: none" in style


def test_office_converter_settings_are_three_choices(main_window):
    action_texts = {
        action.text() for action in main_window.office_converter_action_group.actions()
    }

    assert action_texts == {
        "自動",
        "Microsoft Officeを優先",
        "LibreOfficeを優先",
    }
    assert not hasattr(main_window, "select_libreoffice_path_action")
    assert not hasattr(main_window, "clear_libreoffice_path_action")


def test_settings_menu_groups_image_and_comment_options(main_window):
    actions = [
        action for action in main_window.settings_menu.actions() if not action.isSeparator()
    ]
    office_converter_action = next(
        action
        for action in actions
        if action.menu() and action.menu().title() == "Office変換エンジン"
    )

    assert actions.index(main_window.disable_image_upscaling_action) < actions.index(
        main_window.remove_pdf_annotations_action
    )
    assert actions.index(main_window.remove_pdf_annotations_action) < actions.index(
        main_window.suppress_office_markup_action
    )
    assert actions.index(main_window.suppress_office_markup_action) < actions.index(
        office_converter_action
    )


def test_copyable_message_dialog_keeps_copy_button_auxiliary(main_window, qtbot):
    dialog, button_box, _, default_widget = main_window._build_copyable_message_dialog(
        "保存完了",
        "PDFを保存しました:\nC:/tmp/sample.pdf\n\n一覧をクリアしますか？",
        icon=QMessageBox.Icon.Information,
        buttons=(
            QDialogButtonBox.StandardButton.Yes
            | QDialogButtonBox.StandardButton.No
        ),
        default_button=QDialogButtonBox.StandardButton.No,
    )
    qtbot.addWidget(dialog)

    copy_buttons = [
        button
        for button in dialog.findChildren(QPushButton)
        if button.text() == "コピー"
    ]
    assert len(copy_buttons) == 1
    assert copy_buttons[0].focusPolicy() == Qt.FocusPolicy.NoFocus
    assert not copy_buttons[0].autoDefault()
    assert not copy_buttons[0].isDefault()

    yes_button = button_box.button(QDialogButtonBox.StandardButton.Yes)
    no_button = button_box.button(QDialogButtonBox.StandardButton.No)
    assert yes_button.text() == "はい"
    assert no_button.text() == "いいえ"
    assert default_widget == no_button
    assert no_button.isDefault()
    assert no_button.autoDefault()
    assert not yes_button.isDefault()
    assert not yes_button.autoDefault()

    dialog.show()
    qtbot.wait(50)
    assert dialog.focusWidget() == no_button


def test_toolbar_prioritizes_add_and_merge_actions(main_window):
    actions = [action for action in main_window.main_toolbar.actions() if not action.isSeparator()]

    assert actions[:2] == [main_window.add_action, main_window.merge_action]
    assert "min-width: 68px;" in app_module.QSS
    assert main_window.merge_action.text() == "結合\n保存"


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
        app_module,
        "_show_standard_question",
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
        app_module,
        "_show_standard_question",
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
    image = tmp_path / "photo.png"
    add_raw_item(main_window, item_data(pdf, 0))
    add_raw_item(main_window, item_data(pdf, 1))
    add_raw_item(main_window, item_data(word, item_type="word"))
    add_raw_item(main_window, item_data(word, item_type="word"))
    add_raw_item(main_window, item_data(image, item_type="image"))
    main_window.page_list_widget.item(0).setSelected(True)

    main_window.update_status_bar()

    assert main_window.file_status_label.text() == "PDFページ: 2 | Word: 1 | 画像: 1"
    assert main_window.selection_status_label.text() == "選択中: 1件"


def test_status_bar_counts_svg_as_image(main_window, tmp_path):
    svg = tmp_path / "diagram.svg"
    add_raw_item(main_window, item_data(svg, item_type="svg"))

    main_window.update_status_bar()

    assert main_window.file_status_label.text() == "画像: 1"


def test_add_blank_page_inserts_after_selection_and_counts_as_pdf_page(
    main_window, pdf_factory
):
    source = pdf_factory("pages.pdf", ["1", "2"])
    add_raw_item(main_window, item_data(source, 0))
    second = add_raw_item(main_window, item_data(source, 1))
    second.setSelected(True)

    main_window._add_blank_page()

    assert main_window.page_list_widget.count() == 3
    blank_data = main_window.page_list_widget.item(2).data(Qt.ItemDataRole.UserRole)
    assert blank_data["type"] == "blank"
    assert blank_data["display_name"] == "空白ページ 1"
    assert main_window.file_status_label.text() == "PDFページ: 3"
    assert main_window.page_list_widget.item(2).isSelected()


def test_blank_page_does_not_create_automatic_bookmark(main_window):
    main_window._add_blank_page()

    assert main_window.bookmarks == []


def test_blank_page_selection_disables_rotation_and_image_export(main_window):
    main_window._add_blank_page()

    assert main_window.delete_action.isEnabled()
    assert main_window.export_selected_pdf_action.isEnabled()
    assert not main_window.export_selected_images_action.isEnabled()
    assert not main_window.rot_left_action.isEnabled()
    assert not main_window.rot_right_action.isEnabled()


def test_image_selection_can_rotate_but_not_export_as_image(main_window, tmp_path):
    image = tmp_path / "photo.png"
    item = add_raw_item(main_window, item_data(image, item_type="image"))
    item.setSelected(True)

    main_window._update_page_mode_actions_state()

    assert main_window.rot_left_action.isEnabled()
    assert main_window.rot_right_action.isEnabled()
    assert not main_window.export_selected_images_action.isEnabled()


def test_image_export_uses_selected_dpi_and_defaults_to_300(
    main_window, pdf_factory, tmp_path, monkeypatch
):
    source = pdf_factory("page.pdf", ["1"])
    item = add_raw_item(main_window, item_data(source, 0))
    item.setSelected(True)
    calls = []

    def select_dpi(_parent, _title, _label, items, current, editable):
        assert items == ["96 dpi", "150 dpi", "300 dpi", "600 dpi"]
        assert items[current] == "300 dpi"
        assert editable is False
        return "150 dpi", True

    monkeypatch.setattr(QInputDialog, "getItem", select_dpi)
    monkeypatch.setattr(
        QFileDialog,
        "getExistingDirectory",
        lambda *_args, **_kwargs: str(tmp_path / "images"),
    )
    monkeypatch.setattr(
        main_window,
        "_run_task",
        lambda task, label, **kwargs: calls.append((task, label, kwargs)),
    )

    main_window._export_selected_as_images()

    assert calls[0][0] == "export_images"
    assert calls[0][2]["dpi"] == 150
    assert calls[0][2]["image_format"] == "JPEG"


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
