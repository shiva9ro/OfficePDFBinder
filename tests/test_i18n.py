import re
from pathlib import Path
from xml.etree import ElementTree

import i18n
import OfficePDFBinder_Main as app_module
from i18n import normalize_language, resolve_language


def test_normalize_language_accepts_qt_and_inno_names():
    assert normalize_language("ja_JP") == "ja"
    assert normalize_language("japanese") == "ja"
    assert normalize_language("en-US") == "en"
    assert normalize_language("english") == "en"


def test_language_environment_override_has_highest_priority(tmp_path):
    (tmp_path / "OfficePDFBinder.language").write_text("japanese", encoding="utf-8")

    result = resolve_language(
        tmp_path,
        environ={"OFFICEPDFBINDER_LANGUAGE": "en"},
        system_locale="ja_JP",
    )

    assert result == "en"


def test_installer_language_file_is_used(tmp_path):
    (tmp_path / "OfficePDFBinder.language").write_text("english", encoding="utf-8")

    result = resolve_language(tmp_path, environ={}, system_locale="ja_JP")

    assert result == "en"


def test_os_language_is_used_without_explicit_setting(tmp_path):
    assert resolve_language(tmp_path, environ={}, system_locale="ja_JP") == "ja"
    assert resolve_language(tmp_path, environ={}, system_locale="en_US") == "en"


def test_non_japanese_os_uses_english(tmp_path):
    assert resolve_language(tmp_path, environ={}, system_locale="fr_FR") == "en"


def test_english_translation_catalog_is_complete():
    catalog = Path(__file__).parents[1] / "translations" / "OfficePDFBinder_en.ts"
    root = ElementTree.parse(catalog).getroot()
    messages = root.findall("./context/message")

    assert len(messages) >= 80
    assert all(message.findtext("translation") for message in messages)
    assert all(
        message.find("translation").get("type") != "unfinished" for message in messages
    )
    for message in messages:
        source_fields = set(
            re.findall(r"\{[A-Za-z_][A-Za-z0-9_]*\}", message.findtext("source"))
        )
        translated_fields = set(
            re.findall(r"\{[A-Za-z_][A-Za-z0-9_]*\}", message.findtext("translation"))
        )
        assert translated_fields == source_fields


def test_compiled_english_translation_is_loaded(qapp, monkeypatch):
    project_root = Path(__file__).parents[1]
    monkeypatch.setenv("OFFICEPDFBINDER_LANGUAGE", "en")

    language, loaded = i18n.install_app_translator(qapp, project_root)
    try:
        assert language == "en"
        assert loaded is True
        assert i18n.translate("MainWindow", "ページを整理(&P)") == "Organize Pages (&P)"
    finally:
        if i18n._active_translator is not None:
            qapp.removeTranslator(i18n._active_translator)
            i18n._active_translator = None
            i18n._active_language = i18n.DEFAULT_LANGUAGE


def test_header_footer_internal_values_do_not_depend_on_translation(
    qapp, qtbot, monkeypatch
):
    project_root = Path(__file__).parents[1]
    monkeypatch.setenv("OFFICEPDFBINDER_LANGUAGE", "en")
    language, loaded = i18n.install_app_translator(qapp, project_root)
    assert (language, loaded) == ("en", True)

    try:
        dialog = app_module.HeaderFooterSettingsDialog()
        qtbot.addWidget(dialog)
        dialog.header_page_number_position_combo.setCurrentIndex(
            dialog.header_page_number_position_combo.findData("center")
        )

        assert dialog.windowTitle() == "Header and Footer Settings"
        assert dialog.header_page_number_position_combo.currentText() == "Center"
        assert dialog.get_settings()["header"]["page_number_position"] == "center"
    finally:
        if i18n._active_translator is not None:
            qapp.removeTranslator(i18n._active_translator)
            i18n._active_translator = None
            i18n._active_language = i18n.DEFAULT_LANGUAGE


def test_main_window_uses_english_catalog(qapp, qtbot, tmp_path, monkeypatch):
    project_root = Path(__file__).parents[1]
    settings_file = tmp_path / "OfficePDFBinder" / "settings.ini"
    monkeypatch.setenv("OFFICEPDFBINDER_LANGUAGE", "en")
    monkeypatch.setattr(
        app_module, "_get_settings_file_path", lambda: str(settings_file)
    )
    language, loaded = i18n.install_app_translator(qapp, project_root)
    assert (language, loaded) == ("en", True)

    try:
        window = app_module.OfficePDFBinderApp()
        qtbot.addWidget(window)

        assert window.add_action.text() == "Add\nFiles"
        assert window.delete_action.text() == "Delete\nSelected"
        assert window.merge_action.text() == "Combine and\nSave PDF"
        assert window.bookmark_dock.windowTitle() == "Bookmarks"
        assert Path(window.user_manual_path).name == "README.html"
    finally:
        if i18n._active_translator is not None:
            qapp.removeTranslator(i18n._active_translator)
            i18n._active_translator = None
            i18n._active_language = i18n.DEFAULT_LANGUAGE


def test_english_manual_and_bilingual_license_are_packaged_sources():
    project_root = Path(__file__).parents[1]
    japanese_markdown = (project_root / "README.ja.md").read_text(encoding="utf-8")
    english_markdown = (project_root / "README.md").read_text(encoding="utf-8")
    japanese_html = (project_root / "README.ja.html").read_text(encoding="utf-8")
    english_html = (project_root / "README.html").read_text(encoding="utf-8")
    license_text = (project_root / "LICENSE.txt").read_text(encoding="utf-8")

    assert "[English](README.md)" in japanese_markdown
    assert "[日本語](README.ja.md)" in english_markdown
    assert 'href="README.html"' in japanese_html
    assert 'href="README.ja.html"' in english_html
    assert '<html lang="ja">' in japanese_html
    assert '<html lang="en">' in english_html
    assert "Office PDF Binder - User Manual" in english_html
    assert "Restricted Portable Mode" in english_html
    assert "English\n-------" in license_text
    assert "日本語\n------" in license_text
    assert (
        "GNU AFFERO GENERAL PUBLIC LICENSE\nVersion 3, 19 November 2007" in license_text
    )


def test_portable_build_regenerates_and_packages_source_archive():
    project_root = Path(__file__).parents[1]
    build_script = (project_root / "build_portable.ps1").read_text(encoding="utf-8")

    assert '".\\create_source_archive.ps1"' in build_script
    assert '"source.zip"' in build_script
