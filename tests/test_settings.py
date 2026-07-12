import configparser

import OfficePDFBinder_Main as app_module


def test_save_settings_writes_current_user_preferences(main_window):
    main_window.last_used_path = r"C:\Documents"
    main_window.auto_bookmarks_enabled = False
    main_window.show_bookmarks_on_open = False
    main_window.remove_pdf_annotations = True
    main_window.disable_image_upscaling = True
    main_window.suppress_office_markup = True
    main_window.office_converter = "libreoffice"
    main_window.libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.com"
    main_window.page_number_settings.update(
        {
            "enabled": True,
            "is_header": True,
            "alignment": "right",
            "format": "Page 1",
            "start_number": 5,
            "font_size": 12,
        }
    )
    main_window.header_footer_settings = {
        "header_enabled": True,
        "footer_enabled": False,
        "font_size": 11,
        "header": {
            "left": "Department",
            "center": "Report",
            "right": "2026",
            "auto_date": True,
        },
        "footer": {
            "left": "Internal",
            "center": "",
            "right": "",
            "auto_date": False,
        },
    }

    main_window._save_settings()

    config = configparser.ConfigParser()
    config.read(main_window.settings_file, encoding="utf-8")
    assert config.get("Paths", "last_used") == r"C:\Documents"
    assert config.getboolean("Bookmarks", "auto_generation") is False
    assert config.getboolean("Bookmarks", "show_on_open") is False
    assert config.getboolean("PdfImport", "remove_annotations") is True
    assert config.getboolean("ImageConversion", "disable_upscaling") is True
    assert config.getboolean("OfficeConversion", "suppress_markup") is True
    assert config.get("OfficeConversion", "converter") == "libreoffice"
    assert not config.has_option("OfficeConversion", "libreoffice_path")
    assert config.getboolean("PageNumbers", "enabled") is True
    assert config.get("PageNumbers", "alignment") == "right"
    assert config.getint("PageNumbers", "start_number") == 5
    assert config.get("HeaderFooter", "header_center") == "Report"


def test_load_settings_restores_supported_preferences(
    qtbot, tmp_path, monkeypatch
):
    settings_file = tmp_path / "OfficePDFBinder" / "settings.ini"
    settings_file.parent.mkdir(parents=True)
    settings_file.write_text(
        "[Paths]\nlast_used = C:\\Work\n"
        "[Bookmarks]\nauto_generation = false\nshow_on_open = false\n"
        "[PdfImport]\nremove_annotations = true\n"
        "[ImageConversion]\ndisable_upscaling = true\n"
        "[OfficeConversion]\nsuppress_markup = true\n"
        "converter = libreoffice\n"
        "libreoffice_path = C:\\Program Files\\LibreOffice\\program\\soffice.com\n",
        encoding="utf-8",
    )
    monkeypatch.setattr(
        app_module,
        "_get_settings_file_path",
        lambda: str(settings_file),
    )

    window = app_module.OfficePDFBinderApp()
    qtbot.addWidget(window)

    assert window.last_used_path == r"C:\Work"
    assert window.auto_bookmarks_enabled is False
    assert window.show_bookmarks_on_open is False
    assert window.remove_pdf_annotations is True
    assert window.disable_image_upscaling is True
    assert window.suppress_office_markup is True
    assert window.office_converter == "libreoffice"
    assert window.libreoffice_path == ""


def test_load_broken_settings_continues_with_defaults(
    qtbot, tmp_path, monkeypatch
):
    settings_file = tmp_path / "OfficePDFBinder" / "settings.ini"
    settings_file.parent.mkdir(parents=True)
    settings_file.write_text("not valid ini content", encoding="utf-8")
    monkeypatch.setattr(
        app_module,
        "_get_settings_file_path",
        lambda: str(settings_file),
    )

    window = app_module.OfficePDFBinderApp()
    qtbot.addWidget(window)

    assert window.last_used_path == ""
    assert window.auto_bookmarks_enabled is True
    assert window.show_bookmarks_on_open is True
    assert window.remove_pdf_annotations is False
    assert window.disable_image_upscaling is False
    assert window.suppress_office_markup is False
