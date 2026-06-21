from pathlib import Path

import OfficePDFBinder_Main as app_module


def test_portable_mode_is_enabled_only_by_marker(tmp_path, monkeypatch):
    monkeypatch.setattr(app_module, "_get_runtime_dir", lambda: str(tmp_path))

    assert app_module._is_portable_mode() is False

    (tmp_path / app_module.PORTABLE_MARKER_FILENAME).touch()

    assert app_module._is_portable_mode() is True


def test_normal_mode_office_temp_pdf_uses_default_temp(tmp_path, monkeypatch):
    office_path = tmp_path / "source" / "sample.docx"
    office_path.parent.mkdir()
    office_path.touch()
    expected_temp_dir = tmp_path / "system-temp"
    expected_temp_dir.mkdir()
    monkeypatch.setattr(app_module, "_is_portable_mode", lambda: False)
    monkeypatch.setattr(app_module.tempfile, "tempdir", str(expected_temp_dir))

    temp_pdf = Path(app_module._create_office_temp_pdf_path(str(office_path)))
    try:
        assert temp_pdf.parent == expected_temp_dir
        assert temp_pdf.name.startswith("sample_")
        assert temp_pdf.suffix == ".pdf"
    finally:
        temp_pdf.unlink(missing_ok=True)


def test_portable_mode_office_temp_pdf_uses_source_folder(tmp_path, monkeypatch):
    office_path = tmp_path / "shared-folder" / "sample.docx"
    office_path.parent.mkdir()
    office_path.touch()
    monkeypatch.setattr(app_module, "_is_portable_mode", lambda: True)

    temp_pdf = Path(app_module._create_office_temp_pdf_path(str(office_path)))
    try:
        assert temp_pdf.parent == office_path.parent
        assert temp_pdf.name.startswith("sample_")
        assert temp_pdf.suffix == ".pdf"
    finally:
        temp_pdf.unlink(missing_ok=True)


def test_portable_mode_window_does_not_create_or_save_settings(
    qtbot, tmp_path, monkeypatch
):
    settings_file = tmp_path / "AppData" / "OfficePDFBinder" / "settings.ini"
    monkeypatch.setattr(app_module, "_is_portable_mode", lambda: True)
    monkeypatch.setattr(
        app_module,
        "_get_settings_file_path",
        lambda: str(settings_file),
    )

    window = app_module.OfficePDFBinderApp()
    qtbot.addWidget(window)
    window._save_settings()

    assert window.settings_file is None
    assert window.settings_dir is None
    assert not settings_file.parent.exists()
    assert window.windowTitle() == "Office PDF Binder（ポータブル版）"


def test_portable_mode_debug_log_does_not_write_file(tmp_path, monkeypatch):
    log_path = tmp_path / "debug.log"
    monkeypatch.setattr(app_module, "_ENABLE_DEBUG_LOG", True)
    monkeypatch.setattr(app_module, "_DEBUG_LOG_PATH", str(log_path))
    monkeypatch.setattr(app_module, "_is_portable_mode", lambda: True)

    app_module._debug_log("test")

    assert not log_path.exists()
