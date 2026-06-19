import os
import shutil
from pathlib import Path

import fitz
import pytest
from PySide6.QtGui import QIcon


# Keep GUI tests independent from the interactive Windows desktop.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture
def pdf_factory(tmp_path):
    def create(name, page_texts, toc=None):
        path = tmp_path / name
        document = fitz.open()
        for text in page_texts:
            page = document.new_page(width=300, height=400)
            page.insert_text((36, 72), text, fontsize=14)
        if toc:
            document.set_toc(toc)
        document.save(path)
        document.close()
        return path

    return create


@pytest.fixture
def main_window(qtbot, tmp_path, monkeypatch):
    import OfficePDFBinder_Main as app_module

    settings_file = tmp_path / "AppData" / "OfficePDFBinder" / "settings.ini"
    monkeypatch.setattr(
        app_module,
        "_get_settings_file_path",
        lambda: str(settings_file),
    )

    window = app_module.OfficePDFBinderApp()
    qtbot.addWidget(window)
    monkeypatch.setattr(window, "_create_thumbnail", lambda _data, _size=None: QIcon())
    return window


@pytest.fixture
def artifact_saver():
    def save(source, relative_path):
        if os.environ.get("OFFICEPDFBINDER_KEEP_TEST_ARTIFACTS") != "1":
            return None

        destination = Path(__file__).parent / "artifacts" / relative_path
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
        return destination

    return save
