import OfficePDFBinder_Main as app_module


def test_header_footer_dialog_preserves_internal_settings(qtbot):
    dialog = app_module.HeaderFooterSettingsDialog()
    qtbot.addWidget(dialog)

    assert dialog.windowTitle() == "ヘッダー・フッターの設定"
    assert dialog.header_page_number_position_combo.currentText() == "なし"

    dialog.header_enable_checkbox.setChecked(True)
    dialog.header_page_number_position_combo.setCurrentText("中央")
    dialog.footer_page_number_position_combo.setCurrentText("右")
    dialog.page_number_format_combo.setCurrentText("Page 1")

    settings = dialog.get_settings()

    assert settings["header_enabled"] is True
    assert settings["header"]["page_number_position"] == "center"
    assert settings["footer"]["page_number_position"] == "right"
    assert settings["page_number_format"] == "Page 1"


def test_main_window_default_ui_remains_japanese(qtbot, tmp_path, monkeypatch):
    settings_file = tmp_path / "OfficePDFBinder" / "settings.ini"
    monkeypatch.setattr(
        app_module,
        "_get_settings_file_path",
        lambda: str(settings_file),
    )

    window = app_module.OfficePDFBinderApp()
    qtbot.addWidget(window)

    assert window.windowTitle() == "Office PDF Binder"
    assert window.add_action.text() == "ファイル\n追加"
    assert window.delete_action.text() == "選択項目\n削除"
    assert window.save_action.text() == "名前を付けて保存(&S)..."
    assert window.auto_bookmark_action.isChecked() is True
    assert settings_file.parent.is_dir()
