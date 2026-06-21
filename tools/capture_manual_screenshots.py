"""Capture repeatable English screenshots for the English README."""

import os
from pathlib import Path
import shutil
import sys
import tempfile
import time

from PySide6.QtCore import Qt
from PySide6.QtGui import QPainter, QPixmap
from PySide6.QtWidgets import QApplication


PROJECT_ROOT = Path(__file__).parents[1]
SAMPLES_DIR = PROJECT_ROOT / "samples"
OUTPUT_DIR = PROJECT_ROOT / "docs" / "images"
WINDOW_WIDTH = 911
WINDOW_HEIGHT = 1014

SAMPLE_FILES = [
    "一般会計概算要求・要望額.pdf",
    "excel_sample.xlsx",
    "概算要求に当たっての基本的な方針について.pdf",
    "powerpoint_sample.pptx",
    "国土交通省・公共事業関係予算のポイント.pdf",
    "総務省予算のポイント.pdf",
    "word_sample.docx",
]


def process_events_until(app, condition, timeout=60):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        app.processEvents()
        if condition():
            return
        time.sleep(0.05)
    raise TimeoutError("Timed out while preparing the screenshot")


def resize_window(window, width, height):
    window.show()
    window.resize(width, height)
    window.move(20, 20)
    QApplication.processEvents()


def capture_widget(widget, destination):
    QApplication.processEvents()
    pixmap = widget.grab()
    if not pixmap.save(str(destination), "PNG"):
        raise RuntimeError(f"Could not save screenshot: {destination}")


def capture_window_with_dialog(window, dialog, destination):
    """Render two Qt top-level widgets without depending on desktop stacking."""
    QApplication.processEvents()
    window_pixmap = window.grab()
    dialog_pixmap = dialog.grab()
    combined = QPixmap(window_pixmap.size())
    combined.fill(Qt.GlobalColor.transparent)
    painter = QPainter(combined)
    painter.drawPixmap(0, 0, window_pixmap)
    x = (window_pixmap.width() - dialog_pixmap.width()) // 2
    y = (window_pixmap.height() - dialog_pixmap.height()) // 2
    painter.drawPixmap(x, y, dialog_pixmap)
    painter.end()
    if not combined.save(str(destination), "PNG"):
        raise RuntimeError(f"Could not save screenshot: {destination}")


def main():
    os.environ["OFFICEPDFBINDER_LANGUAGE"] = "en"
    sys.path.insert(0, str(PROJECT_ROOT))
    import OfficePDFBinder_Main as app_module
    import i18n

    missing = [name for name in SAMPLE_FILES if not (SAMPLES_DIR / name).is_file()]
    if missing:
        raise FileNotFoundError(f"Missing sample files: {missing}")

    app = QApplication.instance() or QApplication([])
    language, loaded = i18n.install_app_translator(app, PROJECT_ROOT)
    if (language, loaded) != ("en", True):
        raise RuntimeError("English translation could not be loaded")

    settings_root = Path(tempfile.mkdtemp(prefix="OfficePDFBinder_screenshot_"))
    app_module._get_settings_file_path = lambda: str(settings_root / "settings.ini")

    try:
        window = app_module.OfficePDFBinderApp()
        resize_window(window, WINDOW_WIDTH, WINDOW_HEIGHT)
        window._add_files_from_paths([str(SAMPLES_DIR / name) for name in SAMPLE_FILES])
        process_events_until(
            app,
            lambda: window.page_list_widget.count() == 11
            and window.threadpool.activeThreadCount() == 0,
        )
        # Flush queued signals and final thumbnail updates.
        for _ in range(10):
            app.processEvents()
            time.sleep(0.05)

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        capture_widget(window, OUTPUT_DIR / "screenshot-main-en.png")

        first_pdf_item = window.page_list_widget.item(0)
        first_pdf_item.setSelected(True)
        dialog = app_module.HeaderFooterSettingsDialog(window)
        dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
        dialog.show()
        QApplication.processEvents()
        QApplication.processEvents()
        capture_window_with_dialog(
            window, dialog, OUTPUT_DIR / "screenshot-header-footer-en.png"
        )

        dialog.close()
        window.close()
    finally:
        if i18n._active_translator is not None:
            app.removeTranslator(i18n._active_translator)
        shutil.rmtree(settings_root, ignore_errors=True)


if __name__ == "__main__":
    main()
