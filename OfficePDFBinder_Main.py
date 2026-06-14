import configparser
import copy
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from multiprocessing import Pool, cpu_count

import fitz
from PIL import Image, ImageDraw, ImageFont
from PySide6.QtCore import (
    QEvent,
    QMimeData,
    QObject,
    QPoint,
    QRunnable,
    QRect,
    QSize,
    Qt,
    QThreadPool,
    QTimer,
    QUrl,
    Signal,
    Slot,
)
from PySide6.QtGui import (
    QAction,
    QColor,
    QDesktopServices,
    QDrag,
    QDragEnterEvent,
    QDragMoveEvent,
    QDropEvent,
    QIcon,
    QImage,
    QKeySequence,
    QMouseEvent,
    QPainter,
    QPen,
    QPixmap,
    QWheelEvent,
)
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDockWidget,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMenuBar,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QSpinBox,
    QStatusBar,
    QToolBar,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

_STARTUP_TIME = time.perf_counter()
INITIAL_WINDOW_X = 100
INITIAL_WINDOW_Y = 100
INITIAL_WINDOW_WIDTH = 900
INITIAL_WINDOW_HEIGHT = 800
MIN_WINDOW_WIDTH = 760
MIN_WINDOW_HEIGHT = 560
WINDOW_SCREEN_MARGIN = 40
APP_ICON_FILENAME = "app.ico"

from version import APP_NAME, APP_VERSION

# --- Optional libraries ---
try:
    import qtawesome as qta
except ImportError:
    qta = None
    print("警告: qtawesomeが未インストールです。アイコンは表示されません。")

# --- Windows-specific for Office conversion ---
if sys.platform == "win32":
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        win32com = None
        print("警告: pywin32が未インストールです。Officeファイルの処理はできません。")
else:
    win32com = None

# --- 定数定義 ---
# ファイル拡張子
SUPPORTED_WORD = (".docx", ".doc", ".docm")
SUPPORTED_EXCEL = (".xlsx", ".xls", ".xlsm")
SUPPORTED_POWERPOINT = (".pptx", ".ppt", ".pptm")
SUPPORTED_PDF = (".pdf",)
ALL_SUPPORTED_EXTENSIONS = (
    SUPPORTED_PDF + SUPPORTED_WORD + SUPPORTED_EXCEL + SUPPORTED_POWERPOINT
)

# Officeファイル変換用の定数（Microsoft Office COM定数）
WORD_SAVE_AS_PDF_FORMAT = 17  # wdFormatPDF
EXCEL_EXPORT_PDF_TYPE = 0  # xlTypePDF
POWERPOINT_SAVE_AS_PDF_FORMAT = 32  # ppSaveAsPDF

# ディスク容量チェック
MIN_FREE_SPACE_MB = 100

# UI定数
THUMBNAIL_WIDTH = 150
THUMBNAIL_HEIGHT = 212
GRID_ITEM_PADDING_X = 20
GRID_ITEM_PADDING_Y = 40
MAX_HISTORY = 15


class DropListWidget(QListWidget):
    """ドラッグ&ドロップ対応のQListWidget"""

    files_dropped = Signal(list)  # ドロップされたファイルパスのリストを送信
    zoom_requested = Signal(int)  # ズーム要求（正の値で拡大、負の値で縮小）

    def event(self, e):
        """ドラッグイベントを処理"""
        event_type = e.type()

        if event_type == QEvent.Type.DragEnter:
            if isinstance(e, QDragEnterEvent):
                self.dragEnterEvent(e)
                return True
        elif event_type == QEvent.Type.DragMove:
            if isinstance(e, QDragMoveEvent):
                self.dragMoveEvent(e)
                return True
        elif event_type == QEvent.Type.Drop:
            if isinstance(e, QDropEvent):
                self.dropEvent(e)
                return True

        return super().event(e)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setDropIndicatorShown(True)  # ドロップインジケーターを表示
        self._dragged_items = []  # ドラッグ中のアイテムを保持
        self._drag_start_position = QPoint()  # ドラッグ開始位置
        self._drop_row = -1  # ドロップ位置（描画用）
        self._pressed_item = None  # マウスボタンを押したときのアイテム
        self._is_dragging = False  # ドラッグ中かどうか
        self._pressed_item_was_selected = False
        self._drag_candidate_items = []
        self._drag_timer = QTimer()  # 長押し検出用タイマー
        self._drag_timer.setSingleShot(True)
        self._drag_timer.timeout.connect(self._enable_drag_mode)
        self._drag_mode_enabled = False  # ドラッグモードが有効かどうか

    def dragEnterEvent(self, event: QDragEnterEvent):
        mime_data = event.mimeData()
        # 内部ドラッグ（リスト内のアイテム移動）かチェック
        if mime_data.hasFormat("application/x-qabstractitemmodeldatalist"):
            # ドロップ位置を初期化
            self._drop_row = -1
            event.acceptProposedAction()
            event.setDropAction(Qt.DropAction.MoveAction)
        elif mime_data.hasUrls():
            # サポートされているファイル形式かチェック
            urls = mime_data.urls()
            has_supported_file = False
            for url in urls:
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    if os.path.isfile(file_path):
                        ext = os.path.splitext(file_path)[1].lower()
                        if ext in ALL_SUPPORTED_EXTENSIONS:
                            has_supported_file = True
                            break
            if has_supported_file:
                event.acceptProposedAction()
                event.setDropAction(Qt.DropAction.CopyAction)
            else:
                event.ignore()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        mime_data = event.mimeData()
        if mime_data.hasFormat("application/x-qabstractitemmodeldatalist"):
            # ドロップ位置を記録（描画用）
            drop_row = self.indexAt(event.position().toPoint()).row()
            if drop_row < 0:
                drop_row = self.count()
            self._drop_row = drop_row
            # 再描画をトリガー
            self.viewport().update()
            event.acceptProposedAction()
            event.setDropAction(Qt.DropAction.MoveAction)
        elif mime_data.hasUrls():
            # dragEnterEventで受け入れた場合、dragMoveEventでも受け入れる
            event.acceptProposedAction()
            event.setDropAction(Qt.DropAction.CopyAction)
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        mime_data = event.mimeData()
        # 内部ドラッグ（リスト内のアイテム移動）の場合
        if mime_data.hasFormat("application/x-qabstractitemmodeldatalist"):
            # ドロップ位置を取得
            drop_row = self.indexAt(event.position().toPoint()).row()
            if drop_row < 0:
                drop_row = self.count()

            # ドラッグ中のアイテムを使用（なければ選択されたアイテムを使用）
            if hasattr(self, "_dragged_items") and self._dragged_items:
                selected_items = self._dragged_items
            else:
                # フォールバック: 選択されたアイテムを取得
                selected_items = []
                for i in range(self.count()):
                    item = self.item(i)
                    if item and item.isSelected():
                        selected_items.append(item)

            if not selected_items:
                event.ignore()
                return

            # 選択されたアイテムの行番号を取得
            selected_rows = [self.row(item) for item in selected_items]

            # ドロップ位置が選択範囲内の場合は無視
            if drop_row >= min(selected_rows) and drop_row <= max(selected_rows) + 1:
                event.ignore()
                return

            # 選択されたアイテムを行番号の降順でソート（安全に削除するため）
            items_to_move = sorted(
                [(self.row(item), item) for item in selected_items],
                key=lambda x: x[0],
                reverse=True,
            )

            # アイテムを削除
            for row, item in items_to_move:
                self.takeItem(row)
                # ドロップ位置が削除位置より前の場合、ドロップ位置を調整
                if drop_row > row:
                    drop_row -= 1

            # アイテムを新しい位置に挿入（元の順序を保持）
            insert_pos = drop_row
            for row, item in reversed(items_to_move):
                self.insertItem(insert_pos, item)
                item.setSelected(True)
                insert_pos += 1

            # ドロップ位置にスクロール
            if drop_row < self.count():
                self.scrollToItem(self.item(drop_row))

            event.acceptProposedAction()
            event.setDropAction(Qt.DropAction.MoveAction)

            # ドロップ位置をリセット
            self._drop_row = -1
            self.viewport().update()

            # 親ウィジェットに移動が完了したことを通知（しおりの再生成など）
            if hasattr(self, "_on_items_moved_callback"):
                self._on_items_moved_callback()
        elif mime_data.hasUrls():
            # 外部ファイルのドロップ
            file_paths = []
            urls = mime_data.urls()
            for url in urls:
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    if os.path.isfile(file_path):
                        ext = os.path.splitext(file_path)[1].lower()
                        if ext in ALL_SUPPORTED_EXTENSIONS:
                            file_paths.append(file_path)
            if file_paths:
                self.files_dropped.emit(file_paths)
                event.acceptProposedAction()
                event.setDropAction(Qt.DropAction.CopyAction)
            else:
                event.ignore()
        else:
            event.ignore()

    def mousePressEvent(self, event: QMouseEvent):
        """マウスボタンが押されたときの処理"""
        if event.button() == Qt.MouseButton.LeftButton:
            # ドラッグ開始位置を記録
            self._drag_start_position = event.position().toPoint()
            # 押されたアイテムを記録
            item = self.itemAt(event.position().toPoint())
            self._pressed_item = item
            self._is_dragging = False
            self._drag_mode_enabled = False
            self._pressed_item_was_selected = bool(item and item.isSelected())
            self._drag_candidate_items = (
                self.selectedItems().copy() if self._pressed_item_was_selected else []
            )

            modifiers = event.modifiers()
            is_multi_select_operation = bool(
                modifiers
                & (
                    Qt.KeyboardModifier.ControlModifier
                    | Qt.KeyboardModifier.ShiftModifier
                )
            )
            if item and not is_multi_select_operation:
                # アイテム上の通常左クリックだけを長押しDD候補にする。
                self._drag_timer.start(QApplication.startDragTime())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent):
        """マウス移動時の処理（ドラッグ開始判定）"""
        if not (event.buttons() & Qt.MouseButton.LeftButton):
            super().mouseMoveEvent(event)
            return

        # 既にドラッグ中の場合は何もしない
        if self._is_dragging:
            super().mouseMoveEvent(event)
            return

        # 長押し成立後の移動は、範囲選択ではなくアイテム移動DDとして扱う。
        if self._drag_mode_enabled:
            if self._pressed_item:
                if self._pressed_item_was_selected:
                    selected_items = [
                        item
                        for item in self._drag_candidate_items
                        if self.row(item) >= 0
                    ]
                    if selected_items:
                        self.clearSelection()
                        for item in selected_items:
                            item.setSelected(True)
                else:
                    selected_items = [self._pressed_item]
                    self.clearSelection()
                    self._pressed_item.setSelected(True)

                self._drag_timer.stop()
                self._is_dragging = True
                self._start_drag_for_selected_items(
                    selected_items, event.position().toPoint()
                )
                return

        # 長押し成立前に一定距離以上動いた場合は、通常の範囲選択として扱う。
        if hasattr(self, "_drag_start_position"):
            move_distance = (
                event.position().toPoint() - self._drag_start_position
            ).manhattanLength()
            # Qt標準のドラッグ開始距離を超えた場合は、選択操作とみなす。
            if move_distance >= QApplication.startDragDistance():
                self._drag_timer.stop()
                self._drag_mode_enabled = False
                # 通常のマウス移動処理（範囲選択など）
                super().mouseMoveEvent(event)
                return

        # 通常のマウス移動処理（範囲選択など）
        super().mouseMoveEvent(event)

    def _enable_drag_mode(self):
        """長押しでドラッグモードを有効化"""
        # マウスボタンがまだ押されている場合のみ有効化
        if QApplication.mouseButtons() & Qt.MouseButton.LeftButton:
            self._drag_mode_enabled = True

    def mouseReleaseEvent(self, event: QMouseEvent):
        """マウスボタンが離されたときの処理"""
        if event.button() == Qt.MouseButton.LeftButton:
            # タイマーを停止
            self._drag_timer.stop()
            self._drag_mode_enabled = False
            self._pressed_item = None
            self._pressed_item_was_selected = False
            self._drag_candidate_items = []
        super().mouseReleaseEvent(event)

    def startDrag(self, supportedActions):
        """複数選択されたアイテムのドラッグを開始（オーバーライド）"""
        # 選択されたアイテムを取得
        selected_items = self.selectedItems()
        if not selected_items:
            # 標準の動作にフォールバック
            super().startDrag(supportedActions)
            return

        # ドラッグ中のアイテムを保持（ドロップ時に使用）
        self._dragged_items = selected_items.copy()

        # MIMEデータを作成
        mime_data = QMimeData()
        # 標準のMIMEタイプを使用（Qtの内部ドラッグ&ドロップ用）
        mime_data.setData("application/x-qabstractitemmodeldatalist", b"")

        # ドラッグオブジェクトを作成
        drag = QDrag(self)
        drag.setMimeData(mime_data)

        # ドラッグを開始
        result = drag.exec(supportedActions)
        # ドラッグが終了したら位置をリセット
        self._is_dragging = False
        self._drag_mode_enabled = False
        self._pressed_item = None
        self._pressed_item_was_selected = False
        self._drag_candidate_items = []
        self._drag_timer.stop()
        if result == Qt.DropAction.IgnoreAction:
            self._drop_row = -1
            self.viewport().update()
        # ドラッグ中のアイテムをクリア
        self._dragged_items = []

    def _start_drag_for_selected_items(self, selected_items, hot_spot):
        """選択されたアイテムのドラッグを開始（非推奨：startDragを使用）"""
        # startDragを呼び出す
        self.startDrag(Qt.DropAction.MoveAction)

    def paintEvent(self, event):
        """カスタムドロップインジケーターを描画"""
        super().paintEvent(event)

        # ドロップ位置が有効な場合、線を描画
        if (
            self._drop_row >= 0
            and hasattr(self, "_dragged_items")
            and self._dragged_items
        ):
            painter = QPainter(self.viewport())
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            indicator_color = QColor("#00a8ff")
            indicator_pen = QPen(indicator_color, 4, Qt.PenStyle.SolidLine)
            indicator_pen.setCapStyle(Qt.PenCapStyle.RoundCap)
            painter.setPen(indicator_pen)

            # ドロップ位置のX座標を計算（アイテムの左端）
            if self._drop_row < self.count():
                # アイテムの左端に縦線を描画
                item_rect = self.visualItemRect(self.item(self._drop_row))
                x = item_rect.left()
                top_y = item_rect.top()
                bottom_y = item_rect.bottom()
            else:
                # 最後のアイテムの右端に縦線を描画
                if self.count() > 0:
                    last_item_rect = self.visualItemRect(self.item(self.count() - 1))
                    x = last_item_rect.right()
                    top_y = last_item_rect.top()
                    bottom_y = last_item_rect.bottom()
                else:
                    x = 0
                    top_y = 0
                    bottom_y = self.viewport().height()

            # 縦線を描画（上下に少し余白を持たせる）
            margin = 5
            start_y = top_y + margin
            end_y = bottom_y - margin
            highlight_rect = QRect(x - 7, start_y, 14, max(1, end_y - start_y))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(0, 168, 255, 70))
            painter.drawRoundedRect(highlight_rect, 5, 5)
            painter.setPen(indicator_pen)
            painter.drawLine(x, start_y, x, end_y)
            painter.setBrush(indicator_color)
            painter.drawEllipse(QPoint(x, start_y), 6, 6)
            painter.drawEllipse(QPoint(x, end_y), 6, 6)

    def wheelEvent(self, event: QWheelEvent):
        """マウスホイールイベントを処理（Ctrl+ホイールでズーム）"""
        # Ctrlキーが押されている場合
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            # ホイールの回転量を取得（正の値で上方向、負の値で下方向）
            delta = event.angleDelta().y()
            if delta > 0:
                # 上方向に回転 = 拡大
                self.zoom_requested.emit(1)
            elif delta < 0:
                # 下方向に回転 = 縮小
                self.zoom_requested.emit(-1)
            event.accept()
        else:
            # Ctrlキーが押されていない場合は通常のスクロール動作
            super().wheelEvent(event)


QSS = """
QMainWindow { background-color: #2c313c; }
QWidget#CustomMenuBarContainer {
    background-color: #353b48;
}
/* コンテナの中にあるQMenuBarは、背景を透明にして親の色に溶け込ませる */
QWidget#CustomMenuBarContainer QMenuBar {
    background-color: transparent;
    color: #f5f6fa;
    font-size: 10pt;
    font-family: 'Meiryo UI', 'Segoe UI';
}
QWidget#CustomMenuBarContainer QMenuBar::item {
    background-color: transparent;
    padding: 4px 10px;
}
QWidget#CustomMenuBarContainer QMenuBar::item:selected { /* マウスオーバー時 */
    background-color: #5d6776;
    border-radius: 4px;
}
QMenu {
    background-color: #4a5260;
    color: #f5f6fa;
    border: 1px solid #5d6776;
    padding: 5px;
}
QMenu::item { padding: 5px 25px 5px 25px; }
QMenu::item:selected { background-color: #0984e3; border-radius: 4px; }
QMenu::separator { height: 1px; background-color: #5d6776; margin: 5px 0px; }

QToolBar { background-color: #353b48; border: none; padding: 5px; spacing: 5px; }
QToolBar::handle { image: none; }
QToolButton {
    background-color: #4a5260;
    color: #f5f6fa;
    border-radius: 4px;
    padding: 4px 6px;
    font-size: 10.5pt;
    font-family: 'Meiryo UI', 'Segoe UI';
    min-width: 75px;
}
QToolButton:hover { background-color: #5d6776; }
QToolButton:pressed { background-color: #0984e3; }
QToolButton:disabled { background-color: #414752; color: #7a828e; }
QListWidget { background-color: #2c313c; border: none; color: #f5f6fa; font-size: 10.5pt; padding: 5px; font-family: 'Meiryo UI', 'Segoe UI'; }
QStatusBar {
    background-color: #353b48;
    color: #f5f6fa;
    font-size: 10.5pt;
}
QStatusBar QLabel {
    color: #f5f6fa;
    font-size: 10.5pt;
}
QStatusBar::size-grip {
    background: transparent;
    border: none;
}
QProgressDialog { background-color: #353b48; color: white; }
QProgressDialog QLabel { color: white; }
QProgressBar {
    color: #f5f6fa; /* パーセント表示の文字色を白に設定 */
    background-color: #4a5260; /* バーの背景色をボタンと同じに設定 */
    border: 1px solid #5d6776; /* 枠線を設定 */
    border-radius: 5px;
    text-align: center; /* テキストを中央揃えに */
}

QProgressBar::chunk {
    background-color: #0984e3; /* 進捗部分の色を選択色と同じ青に設定 */
    border-radius: 4px;
    margin: 1px; /* 枠線の内側に表示するためのマージン */
}
"""


class WorkerSignals(QObject):
    item_ready = Signal(dict)  # 単一アイテム用（Word/Excel/PowerPoint用）
    items_ready = Signal(list)  # バッチ用（PDFの複数ページ用）
    bookmarks_ready = Signal(
        str, list
    )  # PDFファイルから読み込んだしおり（ファイルパス、しおりリスト）
    progress = Signal(int, str)
    non_cancellable_started = Signal(str)
    non_cancellable_finished = Signal()
    finished = Signal(str, str, str)
    error = Signal(str, str)


class AppWorker(QRunnable):
    def __init__(self, task_name, **kwargs):
        super().__init__()
        self.signals = WorkerSignals()
        self.task_name = task_name
        self.kwargs = kwargs
        self.is_running = True

    @Slot()
    def run(self):
        if sys.platform == "win32" and win32com:
            pythoncom.CoInitialize()
        try:
            if self.task_name == "add_files":
                self._run_add_files(**self.kwargs)
            elif self.task_name == "merge_save":
                self._run_merge_save(**self.kwargs)
            elif self.task_name == "export_images":
                self._run_export_images(**self.kwargs)
        except Exception as e:
            self.signals.error.emit(
                "予期せぬエラー", f"{self.task_name} 実行中にエラーが発生しました:\n{e}"
            )
        finally:
            if sys.platform == "win32" and win32com:
                pythoncom.CoUninitialize()

    def _run_add_files(self, file_paths):
        total_files = len(file_paths)
        for i, path in enumerate(file_paths):
            if not self.is_running:
                break
            progress_percent = int((i / total_files) * 100)
            progress_text = (
                f"読み込み中 ({i+1}/{total_files}): {os.path.basename(path)}"
            )
            self.signals.progress.emit(progress_percent, progress_text)

            # ファイルアクセス権限チェック
            if not os.path.exists(path):
                self.signals.error.emit(
                    "ファイルが見つかりません",
                    f"ファイル '{os.path.basename(path)}' が見つかりません。\n\n"
                    "考えられる原因:\n"
                    "・ファイルが移動または削除された\n"
                    "・ファイルパスが変更された\n"
                    "・ネットワークドライブが切断された\n\n"
                    "解決方法:\n"
                    "・ファイルの場所を確認してください\n"
                    "・ファイルが存在することを確認してください",
                )
                continue

            if not os.access(path, os.R_OK):
                self.signals.error.emit(
                    "ファイルを読み取れません",
                    f"ファイル '{os.path.basename(path)}' を読み取る権限がありません。\n\n"
                    "考えられる原因:\n"
                    "・ファイルが他のアプリケーションで開かれている\n"
                    "・ファイルのアクセス権限が不足している\n"
                    "・ファイルが読み取り専用になっている\n\n"
                    "解決方法:\n"
                    "・ファイルを開いている他のアプリケーションを閉じてください\n"
                    "・ファイルのプロパティでアクセス権限を確認してください\n"
                    "・管理者権限でアプリケーションを実行してみてください",
                )
                continue

            file_ext = os.path.splitext(path)[1].lower()
            if file_ext == ".pdf":
                try:
                    with fitz.open(path) as doc:
                        total_pages = len(doc)

                        # PDFファイルから既存のしおり（TOC）を読み込む
                        pdf_bookmarks = []
                        try:
                            toc = doc.get_toc()
                            if toc:
                                # 階層構造のしおりをフラットなリストに変換
                                for item in toc:
                                    if len(item) >= 3:
                                        title = item[1]
                                        page_num = (
                                            item[2] - 1
                                        )  # PyMuPDFは1ベース、アプリは0ベース
                                        if 0 <= page_num < total_pages:
                                            pdf_bookmarks.append(
                                                {
                                                    "title": title,
                                                    "path": path,
                                                    "page_num": page_num,
                                                    "auto": False,  # PDFから読み込んだしおりは手動しおりとして扱う
                                                }
                                            )
                        except Exception as e:
                            # しおりの読み込みに失敗しても処理は続行
                            print(
                                f"警告: PDFファイル '{os.path.basename(path)}' のしおりを読み込めませんでした: {e}"
                            )

                        # すべてのページ情報をリストにまとめる（バッチ処理）
                        pages_data = []
                        for page_num in range(total_pages):
                            if not self.is_running:
                                break
                            pages_data.append(
                                {
                                    "type": "pdf",
                                    "path": path,
                                    "page_num": page_num,
                                    "rotation": 0,
                                    "original_path": path,
                                }
                            )

                        # バッチで一度にシグナル発行（100回→1回に削減）
                        if pages_data and self.is_running:
                            self.signals.items_ready.emit(pages_data)

                        # PDFから読み込んだしおりを送信
                        if pdf_bookmarks and self.is_running:
                            self.signals.bookmarks_ready.emit(path, pdf_bookmarks)
                except Exception as e:
                    self.signals.error.emit(
                        "PDFファイルを読み込めませんでした",
                        f"PDFファイル '{os.path.basename(path)}' を読み込むことができませんでした。\n\n"
                        f"エラー詳細: {e}\n\n"
                        "考えられる原因:\n"
                        "・PDFファイルが破損している\n"
                        "・PDFファイルが暗号化されている\n"
                        "・PDFファイルの形式がサポートされていない\n\n"
                        "解決方法:\n"
                        "・PDFファイルが正常に開けるか確認してください\n"
                        "・別のPDFビューアーでファイルを開いてみてください\n"
                        "・ファイルが破損していないか確認してください",
                    )
            elif file_ext in SUPPORTED_WORD:
                self.signals.item_ready.emit(
                    {"type": "word", "path": path, "rotation": 0, "original_path": path}
                )
            elif file_ext in SUPPORTED_EXCEL:
                self.signals.item_ready.emit(
                    {
                        "type": "excel",
                        "path": path,
                        "rotation": 0,
                        "original_path": path,
                    }
                )
            elif file_ext in SUPPORTED_POWERPOINT:
                self.signals.item_ready.emit(
                    {
                        "type": "powerpoint",
                        "path": path,
                        "rotation": 0,
                        "original_path": path,
                    }
                )
        if self.is_running:
            self.signals.progress.emit(100, "ファイルの読み込みが完了しました。")
            self.signals.finished.emit(
                self.task_name, "完了", "ファイルの追加が完了しました。"
            )
        else:
            self.signals.finished.emit("中止", "中止", "処理が中止されました。")

    def _run_merge_save(
        self,
        items_data,
        output_path,
        bookmarks=None,
        show_outlines=True,
        page_number_settings=None,
        header_footer_settings=None,
    ):
        # 出力ディレクトリのアクセス権限チェック
        output_dir = os.path.dirname(output_path)
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
            except Exception as e:
                self.signals.error.emit(
                    "出力ディレクトリを作成できませんでした",
                    f"出力ディレクトリ '{output_dir}' を作成できませんでした。\n\n"
                    f"エラー詳細: {e}\n\n"
                    "考えられる原因:\n"
                    "・ディレクトリへのアクセス権限が不足している\n"
                    "・ディスクが満杯になっている\n"
                    "・パスが無効である\n\n"
                    "解決方法:\n"
                    "・別の保存先を選択してください\n"
                    "・管理者権限でアプリケーションを実行してみてください\n"
                    "・ディスクの空き容量を確認してください",
                )
                return

        if not os.access(output_dir, os.W_OK):
            self.signals.error.emit(
                "保存先に書き込めません",
                f"出力ディレクトリ '{output_dir}' に書き込む権限がありません。\n\n"
                "考えられる原因:\n"
                "・ディレクトリへの書き込み権限が不足している\n"
                "・ディレクトリが読み取り専用になっている\n"
                "・ネットワークドライブが切断された\n\n"
                "解決方法:\n"
                "・別の保存先を選択してください\n"
                "・ディレクトリのプロパティでアクセス権限を確認してください\n"
                "・管理者権限でアプリケーションを実行してみてください",
            )
            return

        # ディスク容量の簡易チェック（最低100MB必要と仮定）
        try:
            stat = shutil.disk_usage(output_dir)
            free_space_mb = stat.free / (1024 * 1024)
            if free_space_mb < MIN_FREE_SPACE_MB:
                self.signals.error.emit(
                    "ディスク容量が不足しています",
                    f"出力ディレクトリ '{output_dir}' の空き容量が不足しています。\n\n"
                    f"現在の空き容量: {free_space_mb:.1f}MB\n"
                    f"必要な空き容量: {MIN_FREE_SPACE_MB}MB以上\n\n"
                    "解決方法:\n"
                    "・不要なファイルを削除してディスク容量を確保してください\n"
                    "・別のドライブに保存してください\n"
                    "・ディスククリーンアップを実行してください",
                )
                return
        except Exception as e:
            # ディスク容量チェックに失敗しても処理は続行（警告のみ）
            print(f"警告: ディスク容量チェックに失敗しました: {e}")

        final_doc = fitz.open()
        temp_files_to_clean = []
        failed_office_conversions = []
        try:
            # --- ステップ1: しおり情報を準備 ---
            pending_toc_entries = []
            bookmark_map = {}
            if bookmarks:
                for b in bookmarks:
                    path = b.get("path")
                    if not path:
                        continue
                    bookmark_map.setdefault(path, []).append(
                        {
                            "title": b.get("title", "無題"),
                            "page_num": b.get("page_num", 0),
                        }
                    )
                for entries in bookmark_map.values():
                    entries.sort(key=lambda entry: entry.get("page_num", 0))

            # --- ステップ2: items_dataの順序を保持しながら、同じファイルの連続するページをまとめる ---
            ordered_tasks = []
            current_task = None
            current_path = None

            for item in items_data:
                path = item["original_path"]
                item_type = item["type"]

                # 新しいファイルまたは前のファイルと異なる場合は新しいタスクを作成
                if path != current_path:
                    # 前のタスクを保存
                    if current_task:
                        ordered_tasks.append((current_path, current_task))

                    # 新しいタスクを作成
                    current_path = path
                    current_task = {
                        "type": item_type,
                        "pages_to_add": [],
                        "rotations": {},
                    }

                # PDFの場合はページ番号を追加
                if item_type == "pdf":
                    page_num = item["page_num"]
                    current_task["pages_to_add"].append(page_num)
                    current_task["rotations"][page_num] = item.get("rotation", 0)

            # 最後のタスクを保存
            if current_task:
                ordered_tasks.append((current_path, current_task))

            # Officeファイルの変換（アプリケーションを再利用、順序を保持）
            office_apps = {}  # タイプごとのアプリケーションインスタンス
            office_conversion_started = False
            try:
                for path, task in ordered_tasks:
                    if task["type"] not in ("word", "excel", "powerpoint"):
                        continue

                    office_type = task["type"]
                    app_name = office_type.capitalize()
                    if office_type == "excel":
                        app_name = "Excel"
                    elif office_type == "powerpoint":
                        app_name = "PowerPoint"

                    # アプリケーションインスタンスを取得または作成
                    if office_type not in office_apps:
                        self.signals.non_cancellable_started.emit(
                            f"{app_name}ファイルを変換中..."
                        )
                        office_conversion_started = True
                        office_apps[office_type] = None

                    app_instance = office_apps[office_type]

                    converter = None
                    if office_type == "word":
                        converter = self._convert_word_to_pdf
                    elif office_type == "excel":
                        converter = self._convert_excel_to_pdf
                    elif office_type == "powerpoint":
                        converter = self._convert_powerpoint_to_pdf

                    if converter:
                        temp_pdf, app_instance = converter(
                            path, app_instance, suppress_errors=True
                        )
                        office_apps[office_type] = (
                            app_instance  # 更新されたインスタンスを保存
                        )
                        if not temp_pdf:
                            failed_office_conversions.append(
                                (app_name, os.path.basename(path))
                            )
                            continue
                        temp_files_to_clean.append(temp_pdf)
                        # 変換後のPDFパスをタスクに保存
                        task["converted_pdf_path"] = temp_pdf

                # すべてのOfficeアプリケーションを終了
                for office_type, app_instance in office_apps.items():
                    if app_instance:
                        try:
                            app_instance.Quit()
                        except Exception as e:
                            app_name = office_type.capitalize()
                            if office_type == "excel":
                                app_name = "Excel"
                            elif office_type == "powerpoint":
                                app_name = "PowerPoint"
                            print(
                                f"警告: {app_name}アプリケーションの終了に失敗しました: {e}"
                            )

                # 変換が行われた場合のみ終了シグナルを送信
                if office_conversion_started:
                    self.signals.non_cancellable_finished.emit()
            except Exception:
                # エラー時もアプリケーションを終了
                for app_instance in office_apps.values():
                    if app_instance:
                        try:
                            app_instance.Quit()
                        except Exception:
                            pass
                raise

            # 順序を保持したタスクリストを使用
            all_tasks = ordered_tasks

            total_files = len(all_tasks)
            for i, (path, task) in enumerate(all_tasks):
                if not self.is_running:
                    break
                self.signals.progress.emit(
                    int(((i + 1) / total_files) * 100),
                    f"処理中: {os.path.basename(path)}",
                )

                # このファイルのページが挿入される前の総ページ数を、しおりの開始位置として使う。
                page_offset_for_bookmark = final_doc.page_count

                # Officeファイルの場合は変換済みPDFを使用
                if task["type"] in ("word", "excel", "powerpoint"):
                    source_path = task.get("converted_pdf_path")
                    if not source_path:
                        # 変換に失敗したOfficeファイルはスキップ
                        continue
                else:
                    source_path = path

                source_doc = fitz.open(source_path)

                pages_to_insert = task["pages_to_add"] or list(
                    range(source_doc.page_count)
                )

                # アプリ内で管理しているしおりを追加（優先度: 高）
                file_bookmarks = bookmark_map.get(path, [])
                app_bookmark_pages = set()  # アプリ内のしおりが使用するページ番号を記録
                if file_bookmarks:
                    for bm in file_bookmarks:
                        page_num = bm.get("page_num", 0)
                        if task["pages_to_add"]:
                            try:
                                relative_index = pages_to_insert.index(page_num)
                            except ValueError:
                                continue
                        else:
                            if not pages_to_insert:
                                continue
                            relative_index = max(
                                0, min(page_num, len(pages_to_insert) - 1)
                            )
                        target_page_index = page_offset_for_bookmark + relative_index
                        app_bookmark_pages.add(target_page_index)
                        pending_toc_entries.append(
                            {
                                "title": bm.get("title", "無題"),
                                "page_index": target_page_index,
                            }
                        )

                # 既存PDFのしおりを保持（アプリ内のしおりと重複しない場合のみ）
                existing_toc = source_doc.get_toc()
                if existing_toc:
                    for toc_entry in existing_toc:
                        # TOCエントリの形式: [level, title, page_num, ...]
                        if len(toc_entry) < 3:
                            continue
                        original_page_num = (
                            toc_entry[2] - 1
                        )  # TOCのページ番号は1ベースなので0ベースに変換

                        # ページ選択がある場合、選択範囲内のしおりのみを処理
                        if task["pages_to_add"]:
                            if original_page_num not in pages_to_insert:
                                continue
                            try:
                                relative_index = pages_to_insert.index(
                                    original_page_num
                                )
                            except ValueError:
                                continue
                        else:
                            if original_page_num < 0 or original_page_num >= len(
                                pages_to_insert
                            ):
                                continue
                            relative_index = original_page_num

                        target_page_index = page_offset_for_bookmark + relative_index

                        # アプリ内のしおりと重複しない場合のみ追加
                        if target_page_index not in app_bookmark_pages:
                            pending_toc_entries.append(
                                {
                                    "title": (
                                        toc_entry[1] if len(toc_entry) > 1 else "無題"
                                    ),
                                    "page_index": target_page_index,
                                }
                            )

                # ページ挿入と回転の処理
                for page_num in pages_to_insert:
                    page_offset = final_doc.page_count
                    final_doc.insert_pdf(
                        source_doc, from_page=page_num, to_page=page_num
                    )
                    rotation = task["rotations"].get(page_num, 0)
                    if rotation != 0:
                        page_to_rotate = final_doc[page_offset]
                        # ページ挿入直後の回転状態を確認
                        before_rotation = page_to_rotate.rotation
                        _debug_log(
                            f"[ROTATION DEBUG] ページ挿入後、回転適用前: "
                            f"ファイル={os.path.basename(path)}, "
                            f"元のページ番号={page_num}, "
                            f"final_docのページインデックス={page_offset}, "
                            f"適用前の回転={before_rotation}, "
                            f"適用する回転={rotation}"
                        )

                        # 画像が含まれているPDFかどうかを判定
                        # 画像が含まれている場合は、テキストの有無に関わらず画像コンテンツを回転させる
                        # （画像とテキストが混在している場合でも、画像だけが回転されない問題を防ぐ）
                        is_likely_scanned = False
                        try:
                            image_list = page_to_rotate.get_images()
                            text_dict = page_to_rotate.get_text("dict")
                            text_length = len(text_dict.get("blocks", []))
                            # 画像が含まれている場合は、画像コンテンツを回転させる
                            is_likely_scanned = len(image_list) > 0
                            _debug_log(
                                f"[ROTATION DEBUG] ページタイプ判定: "
                                f"ファイル={os.path.basename(path)}, "
                                f"final_docのページインデックス={page_offset}, "
                                f"画像数={len(image_list)}, "
                                f"テキストブロック数={text_length}, "
                                f"スキャン画像PDFの可能性={is_likely_scanned}"
                            )
                        except Exception as e:
                            _debug_log(f"[ROTATION DEBUG] ページタイプ判定エラー: {e}")

                        # 元の回転を考慮して、最終的な回転角度を計算
                        # 最終的な回転角度 = (元の回転 + 新しい回転) % 360
                        final_rotation = (before_rotation + rotation) % 360

                        _debug_log(
                            f"[ROTATION DEBUG] 回転計算: "
                            f"元の回転={before_rotation}, "
                            f"新しい回転={rotation}, "
                            f"最終的な回転={final_rotation}"
                        )

                        # set_rotation()でページ全体（画像も含む）を回転させる
                        page_to_rotate.set_rotation(final_rotation)

                        # 回転適用後の状態を確認
                        after_rotation = page_to_rotate.rotation
                        _debug_log(
                            f"[ROTATION DEBUG] 回転適用後: "
                            f"ファイル={os.path.basename(path)}, "
                            f"final_docのページインデックス={page_offset}, "
                            f"適用後の回転={after_rotation}"
                        )

                source_doc.close()

            if self.is_running:
                # ループ中に生成したTOCリストをPDFへ設定する。
                toc_list = []
                if pending_toc_entries:
                    for entry in pending_toc_entries:
                        page_index = entry.get("page_index")
                        if page_index is None or page_index >= final_doc.page_count:
                            continue
                        page = final_doc[page_index]
                        top_left = self._convert_visual_top_left_to_internal_coords_for_bookmark(
                            page
                        )
                        dest = {
                            "kind": fitz.LINK_GOTO,
                            "page": page_index,
                            "to": fitz.Point(top_left.x, top_left.y),
                            "zoom": 0,
                        }
                        toc_list.append(
                            [1, entry.get("title", "無題"), page_index + 1, dest]
                        )
                if toc_list:
                    final_doc.set_toc(toc_list)

                # ヘッダー・フッターを追加（ページ番号も統合済み、日本語は japan-s、数字・英字は helv を使用）
                if header_footer_settings:
                    # ヘッダー・フッター追加前の回転状態を確認
                    _debug_log(
                        f"[ROTATION DEBUG] ヘッダー・フッター追加前: "
                        f"総ページ数={final_doc.page_count}"
                    )
                    for page_idx in range(final_doc.page_count):
                        page = final_doc[page_idx]
                        if page.rotation != 0:
                            _debug_log(
                                f"[ROTATION DEBUG] ヘッダー・フッター追加前の回転: "
                                f"ページインデックス={page_idx}, "
                                f"回転={page.rotation}"
                            )
                    self._add_header_footer(final_doc, header_footer_settings)
                    # ヘッダー・フッター追加後の回転状態を確認
                    _debug_log(
                        f"[ROTATION DEBUG] ヘッダー・フッター追加後: "
                        f"総ページ数={final_doc.page_count}"
                    )
                    for page_idx in range(final_doc.page_count):
                        page = final_doc[page_idx]
                        if page.rotation != 0:
                            _debug_log(
                                f"[ROTATION DEBUG] ヘッダー・フッター追加後の回転: "
                                f"ページインデックス={page_idx}, "
                                f"回転={page.rotation}"
                            )

                final_doc.set_pagemode("UseOutlines" if show_outlines else "UseNone")

                # 保存前の回転状態を確認
                _debug_log(
                    f"[ROTATION DEBUG] 保存前: 総ページ数={final_doc.page_count}"
                )
                for page_idx in range(final_doc.page_count):
                    page = final_doc[page_idx]
                    if page.rotation != 0:
                        _debug_log(
                            f"[ROTATION DEBUG] 保存前の回転: "
                            f"ページインデックス={page_idx}, "
                            f"回転={page.rotation}"
                        )

                self.signals.progress.emit(95, "ファイルを最適化して保存中...")
                if final_doc.page_count == 0:
                    self.signals.error.emit(
                        "保存できるページがありません",
                        "PDFに保存できるページがありませんでした。\n\n"
                        "Officeファイルの変換に失敗した場合は、Microsoft Officeが正しくインストールされているか確認してください。",
                    )
                    return
                final_doc.save(output_path, garbage=4, deflate=True, clean=True)

                # 保存後の回転状態を確認（保存したファイルを再度開いて確認）
                try:
                    saved_doc = fitz.open(output_path)
                    _debug_log(
                        f"[ROTATION DEBUG] 保存後（ファイル再読み込み）: 総ページ数={saved_doc.page_count}"
                    )
                    for page_idx in range(saved_doc.page_count):
                        page = saved_doc[page_idx]
                        if page.rotation != 0:
                            _debug_log(
                                f"[ROTATION DEBUG] 保存後の回転: "
                                f"ページインデックス={page_idx}, "
                                f"回転={page.rotation}"
                            )
                    saved_doc.close()
                except Exception as e:
                    _debug_log(f"[ROTATION DEBUG] 保存後の確認エラー: {e}")
                self.signals.progress.emit(100, "保存完了")
                message = f"PDFを正常に保存しました:\n{output_path}"
                if failed_office_conversions:
                    skipped_files = "\n".join(
                        f"・{file_name}（{app_name}）"
                        for app_name, file_name in failed_office_conversions
                    )
                    message += (
                        "\n\n"
                        "ただし、以下のOfficeファイルは変換に失敗したためスキップしました:\n"
                        f"{skipped_files}\n\n"
                        "Microsoft Officeがインストールされているか、対象ファイルをOfficeで開けるか確認してください。"
                    )
                self.signals.finished.emit(
                    self.task_name,
                    "保存完了",
                    message,
                )

        finally:
            final_doc.close()
            for temp_f in temp_files_to_clean:
                try:
                    if os.path.exists(temp_f):
                        os.remove(temp_f)
                except Exception as e:
                    print(f"一時ファイルの削除に失敗: {temp_f}, エラー: {e}")

    def _convert_office_to_pdf(
        self,
        office_path,
        app_name,
        save_as_format,
        export_format_type=None,
        app_instance=None,
        suppress_errors=False,
    ):
        """OfficeファイルをPDFに変換（アプリケーションインスタンスを再利用可能）"""
        if not win32com:
            if not suppress_errors:
                self.signals.error.emit(
                    "Officeファイル処理に必要なコンポーネントが見つかりません",
                    "Officeファイルを処理するために必要なpywin32パッケージが見つかりません。\n\n"
                    "解決方法:\n"
                    "・以下のコマンドでpywin32をインストールしてください:\n"
                    "  pip install pywin32\n"
                    "・インストール後、アプリケーションを再起動してください",
                )
            return None, None

        # アプリケーションインスタンスが提供されていない場合は新規作成
        app = app_instance
        if app is None:
            try:
                app = win32com.client.Dispatch(f"{app_name}.Application")
                app.DisplayAlerts = False
                if app_name == "PowerPoint":
                    try:
                        app.WindowState = 2
                        app.Top = -4000
                        app.Left = -4000
                    except Exception as e:
                        print(
                            f"警告: PowerPointウィンドウの最小化/移動に失敗しました。: {e}"
                        )
                elif hasattr(app, "Visible"):
                    app.Visible = False
            except Exception as e:
                if not suppress_errors:
                    self.signals.error.emit(
                        f"{app_name}アプリケーションの起動に失敗しました",
                        f"'{os.path.basename(office_path)}' の変換に失敗しました。\n\n"
                        f"エラー詳細: {e}\n\n"
                        "考えられる原因:\n"
                        f"・{app_name}がインストールされていない\n"
                        f"・{app_name}アプリケーションが他のプロセスで使用されている\n\n"
                        "解決方法:\n"
                        f"・Microsoft {app_name}が正しくインストールされているか確認してください\n"
                        f"・{app_name}アプリケーションをすべて閉じてから再試行してください",
                    )
                return None, None

        doc = None
        try:
            temp_pdf_path = tempfile.NamedTemporaryFile(
                suffix=".pdf",
                prefix=f"{os.path.splitext(os.path.basename(office_path))[0]}_",
                delete=False,
            ).name

            open_method = getattr(
                app,
                (
                    "Presentations"
                    if app_name == "PowerPoint"
                    else ("Workbooks" if app_name == "Excel" else "Documents")
                ),
            )

            # ファイルを読み取り専用で開く
            doc = open_method.Open(os.path.abspath(office_path), ReadOnly=True)

            if export_format_type is not None:
                doc.ExportAsFixedFormat(
                    Type=export_format_type, Filename=os.path.abspath(temp_pdf_path)
                )
            else:
                doc.SaveAs(os.path.abspath(temp_pdf_path), FileFormat=save_as_format)

            # アプリケーションインスタンスを返す（再利用のため）
            return temp_pdf_path, app
        except Exception as e:
            if not suppress_errors:
                self.signals.error.emit(
                    f"{app_name}ファイルの変換に失敗しました",
                    f"'{os.path.basename(office_path)}' の変換に失敗しました。\n\n"
                    f"エラー詳細: {e}\n\n"
                    "考えられる原因:\n"
                    f"・{app_name}がインストールされていない\n"
                    f"・{app_name}アプリケーションが他のプロセスで使用されている\n"
                    "・ファイルが破損している\n"
                    "・ファイルがパスワードで保護されている\n\n"
                    "解決方法:\n"
                    f"・Microsoft {app_name}が正しくインストールされているか確認してください\n"
                    f"・{app_name}アプリケーションをすべて閉じてから再試行してください\n"
                    "・ファイルが正常に開けるか確認してください\n"
                    "・ファイルがパスワード保護されていないか確認してください",
                )
            return None, app
        finally:
            if doc:
                # ファイルを閉じる際に、いかなる変更も保存しないことを明示
                # アプリケーションごとにClose()メソッドの引数が異なる
                try:
                    if app_name == "Word":
                        # WordのDocument.Close()はSaveChangesを位置引数またはキーワード引数で受け取る
                        doc.Close(0)  # 0 = wdDoNotSaveChanges
                    elif app_name == "Excel":
                        # ExcelのWorkbook.Close()はSaveChangesを位置引数で受け取る
                        doc.Close(False)  # False = 変更を保存しない
                    elif app_name == "PowerPoint":
                        # PowerPointのPresentation.Close()は引数なし
                        doc.Close()
                except Exception:
                    # Close()が失敗した場合は無視（既に閉じられている可能性がある）
                    pass
            # app_instanceが提供されている場合は終了しない（再利用のため）
            # 提供されていない場合は終了する（後でQuitを呼ぶ必要がある）

    def _convert_word_to_pdf(self, path, app_instance=None, suppress_errors=False):
        result, app = self._convert_office_to_pdf(
            path,
            "Word",
            save_as_format=WORD_SAVE_AS_PDF_FORMAT,
            app_instance=app_instance,
            suppress_errors=suppress_errors,
        )
        return result, app

    def _convert_excel_to_pdf(self, path, app_instance=None, suppress_errors=False):
        result, app = self._convert_office_to_pdf(
            path,
            "Excel",
            save_as_format=None,
            export_format_type=EXCEL_EXPORT_PDF_TYPE,
            app_instance=app_instance,
            suppress_errors=suppress_errors,
        )
        return result, app

    def _convert_powerpoint_to_pdf(self, path, app_instance=None, suppress_errors=False):
        result, app = self._convert_office_to_pdf(
            path,
            "PowerPoint",
            save_as_format=POWERPOINT_SAVE_AS_PDF_FORMAT,
            app_instance=app_instance,
            suppress_errors=suppress_errors,
        )
        return result, app

    def _convert_visual_top_left_to_internal_coords_for_bookmark(self, page):
        """回転を考慮し、視覚的左上を内部座標系の左上に戻す（しおり用）

        ポイント:
        - `page.rect` は回転後の視覚的な矩形
        - 内部座標系のサイズは `page.mediabox`（回転前の幅・高さ）を採用する
        - 90/270度は縦横が入れ替わるため、内部サイズは mediabox をそのまま使い、逆変換のみで吸収する
        """
        rect = page.rect
        rotation = page.rotation
        mediabox = page.mediabox  # 回転前の内部座標系サイズ

        # 視覚的な左上を矩形基準の相対座標に
        visual_x = rect.tl.x - rect.x0
        visual_y = rect.tl.y - rect.y0

        internal_width = mediabox.width
        internal_height = mediabox.height

        # 視覚→内部の逆変換
        if rotation == 0:
            internal_x = visual_x
            internal_y = visual_y
        elif rotation == 90:
            # (vx, vy) → (H - vy, vx)
            internal_x = internal_height - visual_y
            internal_y = visual_x
        elif rotation == 180:
            # (vx, vy) → (W - vx, H - vy)
            internal_x = internal_width - visual_x
            internal_y = -internal_height - visual_y
        elif rotation == 270:
            # (vx, vy) → (vy, W - vx)
            internal_x = visual_y
            internal_y = -internal_width - visual_x
        else:
            internal_x = visual_x
            internal_y = visual_y

        # mediabox の原点を考慮して絶対座標に
        return fitz.Point(mediabox.x0 + internal_x, mediabox.y0 + internal_y)

    def _calculate_position_with_rotation(self, page, is_header, alignment, margin=20):
        """回転を考慮した位置を計算する

        PyMuPDFの座標系:
        - 原点は左下 (rect.x0, rect.y0)
        - x軸は右方向、y軸は上方向
        - page.set_rotation()で回転を設定すると、ページ全体が回転する
        - insert_text()でテキストを挿入する際、元の座標系で位置を指定する
        - 挿入されたテキストは、ページの回転と一緒に回転される

        回転後の視覚的な位置にテキストを配置するには、元の座標系での位置を逆変換する必要があります。

        座標変換（回転行列の逆変換）:
        - 90度時計回り: (vx, vy) → (height-vy, vx)
        - 180度: (vx, vy) → (width-vx, height-vy)
        - 270度: (vx, vy) → (vy, width-vx)

        Args:
            page: PDFページオブジェクト
            is_header: True=ヘッダー、False=フッター
            alignment: "left", "center", "right"
            margin: 端からのマージン（ポイント）

        Returns:
            fitz.Point: 描画位置（元の座標系での位置）
        """
        rect = page.rect
        rotation = page.rotation
        width = rect.width
        height = rect.height

        # page.rectが回転後の視覚的なrectを返している場合、
        # widthとheightは既に回転後の視覚的な値なので、そのまま使う
        # 回転のない横長PDFと同じロジックで計算

        # 回転角度に応じて座標を計算
        if rotation == 0:
            # 無回転の場合: 直接座標を計算
            # ヘッダー: 上端からmargin離れた位置 (y = margin)
            # フッター: 下端からmargin離れた位置 (y = height - margin)
            if is_header:  # ヘッダーの場合
                y_pos = margin  # 上端からmargin離れた位置
            else:  # フッターの場合
                y_pos = height - margin  # 下端からmargin離れた位置

            # 配置に応じたx座標
            if alignment == "center":
                x_pos = rect.x0 + width / 2
            elif alignment == "right":
                x_pos = rect.x1 - margin
            else:  # left
                x_pos = rect.x0 + margin

        else:
            # 回転ありの場合: 視覚的な座標を計算してから、元の座標系に変換
            # ステップ1: 視覚的な座標系で位置を計算（回転後の状態として扱う）
            # フッター: 視覚的な下端からmargin離れた位置
            # ヘッダー: 視覚的な上端からmargin離れた位置
            if not is_header:  # フッターの場合
                visual_y = (
                    margin  # 視覚的な下端からmargin離れた位置（元の座標系では下端）
                )
            else:  # ヘッダーの場合
                visual_y = (
                    height - margin
                )  # 視覚的な上端からmargin離れた位置（元の座標系では上端）

            # 配置に応じた視覚的なx座標
            if alignment == "center":
                visual_x = width / 2
            elif alignment == "right":
                visual_x = width - margin
            else:  # left
                visual_x = margin

            # ステップ2: 回転角度に応じて、視覚的な座標を元の座標系に変換
            if rotation == 90:
                # 90度回転: 視覚的な座標 (vx, vy) → 元の座標
                # 90度回転では、視覚的な上下が元の座標系の左右に対応
                # フッター（視覚的下端）→ 元の座標系の右端
                # ヘッダー（視覚的上端）→ 元の座標系の左端
                # 視覚的な左右は元の座標系の上下に対応
                # alignmentの調整は_insert_text_with_mixed_fontsでy座標に対して行うため、
                # y_posはvisual_xをそのまま使う（270度回転と同じ）
                # ただし、90度回転時のみrightとleftを入れ替える
                if not is_header:  # フッターの場合
                    x_pos = rect.x0 + (height - margin)
                else:  # ヘッダーの場合
                    x_pos = rect.x0 + margin
                # 90度回転時のみ、rightとleftを入れ替える
                if alignment == "center":
                    visual_x_for_y = visual_x
                elif alignment == "right":
                    if is_header:
                        # ヘッダー: 視覚的right → 元の座標系ではleft側（margin）
                        visual_x_for_y = margin
                    else:
                        # フッター: 視覚的right → 元の座標系ではleft側（margin）（leftとrightを入れ替え）
                        visual_x_for_y = margin
                else:  # left
                    if is_header:
                        # ヘッダー: 視覚的left → 元の座標系ではright側（width - margin）
                        visual_x_for_y = width - margin
                    else:
                        # フッター: 視覚的left → 元の座標系ではright側（width - margin）（leftとrightを入れ替え）
                        visual_x_for_y = width - margin
                y_pos = rect.y0 + visual_x_for_y
            elif rotation == 180:
                # 180度回転: y座標を無回転と逆にする
                if is_header:  # ヘッダーの場合
                    y_pos = height - margin  # 無回転のフッターと同じ位置
                else:  # フッターの場合
                    y_pos = margin  # 無回転のヘッダーと同じ位置

                # x座標は左右が逆になる
                if alignment == "center":
                    x_pos = rect.x0 + width / 2
                elif alignment == "right":
                    x_pos = rect.x0 + margin
                else:  # left
                    x_pos = rect.x1 - margin

            elif rotation == 270:
                # 270度回転: 視覚的な座標 (vx, vy) → 元の座標
                # 270度回転では、視覚的な上下が元の座標系の左右に対応（90度と逆）
                # フッター（視覚的下端）→ 元の座標系の左端
                # ヘッダー（視覚的上端）→ 元の座標系の右端
                if not is_header:  # フッターの場合
                    x_pos = rect.x0 + margin
                else:  # ヘッダーの場合
                    x_pos = rect.x0 + (height - margin)
                y_pos = rect.y0 + visual_x
            else:
                # 回転なしとして扱う（念のため）
                x_pos = rect.x0 + visual_x
                y_pos = rect.y0 + (height - visual_y)

        result = fitz.Point(x_pos, y_pos)
        return result

    def _format_page_number(self, format_text, current_number, total_pages):
        """ページ番号をフォーマットする

        Args:
            format_text: フォーマット文字列（"1", "1 / 10", "Page 1", "- 1 -"など）
            current_number: 現在のページ番号
            total_pages: 総ページ数

        Returns:
            str: フォーマットされたページ番号文字列
        """
        if " / " in format_text:
            # "1 / 10" 形式: 分子は現在のページ番号、分母は総ページ数
            return f"{current_number} / {total_pages}"
        elif "Page " in format_text:
            # "Page 1" 形式: 現在のページ番号を入れる
            return f"Page {current_number}"
        elif "-" in format_text and format_text.count("-") >= 2:
            # "- 1 -" 形式: 現在のページ番号を入れる
            return f"- {current_number} -"
        else:
            # "1" 形式: 現在のページ番号のみ
            return str(current_number)

    def _add_page_numbers(self, doc, settings):
        """PDFドキュメントの各ページにページ番号を追加"""
        total_pages = len(doc)
        start_number = settings.get("start_number", 1)
        format_text = settings.get("format", "1")
        font_size = settings.get("font_size", 10)
        is_header = settings.get("is_header", False)
        alignment = settings.get("alignment", "center")

        for page_num in range(total_pages):
            page = doc[page_num]

            # 元の回転角度を保存
            original_rotation = page.rotation

            # ページの回転を一時的に0に戻す
            if original_rotation != 0:
                page.set_rotation(0)

            # ページ番号を計算
            current_number = start_number + page_num
            page_text = self._format_page_number(
                format_text, current_number, total_pages
            )

            # 元の回転を先に復元
            if original_rotation != 0:
                page.set_rotation(original_rotation)

            # 回転後の座標系で位置を再計算
            point = self._calculate_position_with_rotation(
                page, is_header, alignment, margin=20
            )

            # テキストを追加（回転後の座標系で位置を計算済み）
            self._insert_text_with_mixed_fonts(
                page,
                point,
                page_text,
                font_size,
                alignment,
                text_rotation=original_rotation,  # テキスト自体は回転させない（ページの回転に従う）
            )

    def _insert_text_with_mixed_fonts(
        self,
        page,
        point,
        text,
        font_size,
        alignment="left",
        text_rotation=0,
    ):
        """テキストを日本語部分と非日本語部分に分けて、それぞれ異なるフォントで描画する。

        日本語文字には japan-s（明朝体）、数字・英字には helv（Helvetica）を使用。

        Args:
            page: PDFページオブジェクト
            point: 描画開始位置（fitz.Point）
            text: 描画するテキスト
            font_size: フォントサイズ
            alignment: 配置（"left", "center", "right"）
            text_rotation: テキストの回転角度（0, 90, 180, 270）
        """
        if not text:
            return

        # 日本語文字（全角文字）かどうかを判定する関数
        def is_japanese(char):
            code = ord(char)
            # ひらがな、カタカナ、漢字、全角記号など
            return (
                0x3040 <= code <= 0x309F  # ひらがな
                or 0x30A0 <= code <= 0x30FF  # カタカナ
                or 0x4E00 <= code <= 0x9FAF  # CJK統合漢字
                or 0x3400 <= code <= 0x4DBF  # CJK統合漢字拡張A
                or 0x20000 <= code <= 0x2A6DF  # CJK統合漢字拡張B
                or 0xFF00 <= code <= 0xFFEF  # 全角記号
            )

        # テキストを日本語部分と非日本語部分に分割
        segments = []
        current_segment = ""
        current_is_japanese = None

        for char in text:
            char_is_japanese = is_japanese(char)

            if current_is_japanese is None:
                # 最初の文字
                current_is_japanese = char_is_japanese
                current_segment = char
            elif current_is_japanese == char_is_japanese:
                # 同じ種類の文字が続く
                current_segment += char
            else:
                # 種類が変わった
                segments.append((current_segment, current_is_japanese))
                current_segment = char
                current_is_japanese = char_is_japanese

        # 最後のセグメントを追加
        if current_segment:
            segments.append((current_segment, current_is_japanese))

        # 各セグメントの幅を計算して、全体の幅を求める
        total_width = 0
        segment_widths = []
        for segment_text, is_jp in segments:
            if is_jp:
                # 日本語: 全角文字として扱う（フォントサイズ × 文字数）
                width = len(segment_text) * font_size
            else:
                # 非日本語: 半角文字として扱う（フォントサイズ × 0.6 × 文字数）
                width = len(segment_text) * font_size * 0.6
            segment_widths.append(width)
            total_width += width

        # 配置に応じて開始位置を調整
        # 90度/270度回転時は、視覚的な左右が元の座標系の上下に対応するため、
        # alignmentの調整をy座標に対して行う必要がある
        if text_rotation == 90:
            # 90度回転時: alignmentの調整をy座標に対して行う
            # _calculate_position_with_rotationで計算された位置は基本位置なので、
            # テキストの幅を考慮した調整が必要
            if alignment == "center":
                # 中央配置: 全体の幅の半分だけ下にずらす（y座標を減らす）
                y_start = point.y + total_width / 2
            elif alignment == "right":
                y_start = point.y + total_width
            else:
                # 左配置: そのまま
                y_start = point.y
            x_start = point.x
        elif text_rotation == 270:
            # 270度回転時: alignmentの調整をy座標に対して行う
            if alignment == "center":
                # 中央配置: 全体の幅の半分だけ下にずらす（y座標を減らす）
                y_start = point.y - total_width / 2
            elif alignment == "right":
                # 右配置: 全体の幅だけ下にずらす（y座標を減らす）
                y_start = point.y - total_width
            else:
                # 左配置: そのまま
                y_start = point.y
            x_start = point.x
        elif text_rotation == 180:
            # 180度回転時: 左右が逆なので、調整も逆にする
            if alignment == "center":
                # 中央配置: 全体の幅の半分だけ右にずらす（x座標を増やす）
                x_start = point.x + total_width / 2
            elif alignment == "right":
                # 右配置: 全体の幅だけ右にずらす（x座標を増やす）
                x_start = point.x + total_width
            else:
                # 左配置: そのまま
                x_start = point.x
            y_start = point.y
        else:
            # 0度回転時: alignmentの調整をx座標に対して行う（従来通り）
            if alignment == "center":
                # 中央配置: 全体の幅の半分だけ左にずらす
                x_start = point.x - total_width / 2
            elif alignment == "right":
                # 右配置: 全体の幅だけ左にずらす
                x_start = point.x - total_width
            else:
                # 左配置: そのまま
                x_start = point.x
            y_start = point.y

        # 各セグメントを描画
        if text_rotation == 90 or text_rotation == 270:
            # 90度/270度回転時: y座標を調整
            current_y = y_start
            for (segment_text, is_jp), segment_width in zip(segments, segment_widths):
                if segment_text:
                    font_name = "japan" if is_jp else "helv"
                    segment_point = fitz.Point(x_start, current_y)
                    text_rotate = text_rotation
                    page.insert_text(
                        segment_point,
                        segment_text,
                        fontsize=font_size,
                        color=(0, 0, 0),
                        fontname=font_name,
                        rotate=text_rotate,
                    )
                    current_y += segment_width
        else:
            # 0度/180度回転時: x座標を調整（従来通り）
            current_x = x_start
            for (segment_text, is_jp), segment_width in zip(segments, segment_widths):
                if segment_text:
                    font_name = "japan" if is_jp else "helv"
                    segment_point = fitz.Point(current_x, y_start)
                    # ページの回転を元に戻すと、テキストも一緒に回転する
                    # テキストの向きを正すために、テキスト自体を同じ方向に回転させる
                    # 90度回転している場合、テキストも90度回転させれば、正しい向きになる
                    text_rotate = text_rotation
                    page.insert_text(
                        segment_point,
                        segment_text,
                        fontsize=font_size,
                        color=(0, 0, 0),
                        fontname=font_name,
                        rotate=text_rotate,
                    )
                    current_x += segment_width

    def _add_header_footer(self, doc, settings):
        """PDFドキュメントの各ページにヘッダー・フッターを追加"""
        font_size = settings.get("font_size", 10)
        # 日付形式: 2025/11/08（半角数字・ゼロ埋め）
        now = datetime.now()
        current_date = f"{now.year}/{now.month:02d}/{now.day:02d}"

        # ページ番号の設定を取得（統合された設定から）
        total_pages = len(doc)
        page_number_format = settings.get("page_number_format", "1")
        page_number_start = settings.get("page_number_start", 1)

        for page_num in range(len(doc)):
            page = doc[page_num]

            # 元の回転角度を保存（ヘッダー・フッター追加時に使用）
            original_rotation = page.rotation

            # ページ番号を計算（必要に応じて使用）
            current_number = page_number_start + page_num
            page_number_text = self._format_page_number(
                page_number_format, current_number, total_pages
            )

            # ヘッダーを追加
            if settings.get("header_enabled") and "header" in settings:
                header = settings["header"]

                # 左
                left_text = header.get("left", "")
                if (
                    header.get("auto_page_number")
                    and header.get("page_number_position") == "left"
                ):
                    left_text = page_number_text
                if left_text:
                    point = self._calculate_position_with_rotation(
                        page,
                        is_header=True,
                        alignment="left",
                        margin=20,
                    )
                    self._insert_text_with_mixed_fonts(
                        page,
                        point,
                        left_text,
                        font_size,
                        "left",
                        text_rotation=original_rotation,
                    )

                # 中央
                center_text = header.get("center", "")
                if (
                    header.get("auto_page_number")
                    and header.get("page_number_position") == "center"
                ):
                    center_text = page_number_text
                if center_text:
                    point = self._calculate_position_with_rotation(
                        page, is_header=True, alignment="center", margin=20
                    )
                    self._insert_text_with_mixed_fonts(
                        page,
                        point,
                        center_text,
                        font_size,
                        "center",
                        text_rotation=original_rotation,
                    )

                # 右（主に日付を想定）
                right_text = header.get("right", "")
                if header.get("auto_date"):
                    right_text = current_date
                elif (
                    header.get("auto_page_number")
                    and header.get("page_number_position") == "right"
                ):
                    right_text = page_number_text

                if right_text:
                    point = self._calculate_position_with_rotation(
                        page, is_header=True, alignment="right", margin=20
                    )
                    self._insert_text_with_mixed_fonts(
                        page,
                        point,
                        right_text,
                        font_size,
                        "right",
                        text_rotation=original_rotation,
                    )

            # フッターを追加
            if settings.get("footer_enabled") and "footer" in settings:
                footer = settings["footer"]

                # 左
                left_text = footer.get("left", "")
                if (
                    footer.get("auto_page_number")
                    and footer.get("page_number_position") == "left"
                ):
                    left_text = page_number_text
                if left_text:
                    point = self._calculate_position_with_rotation(
                        page,
                        is_header=False,
                        alignment="left",
                        margin=20,
                    )
                    self._insert_text_with_mixed_fonts(
                        page,
                        point,
                        left_text,
                        font_size,
                        "left",
                        text_rotation=original_rotation,
                    )

                # 中央
                center_text = footer.get("center", "")
                if (
                    footer.get("auto_page_number")
                    and footer.get("page_number_position") == "center"
                ):
                    center_text = page_number_text
                if center_text:
                    point = self._calculate_position_with_rotation(
                        page, is_header=False, alignment="center", margin=20
                    )
                    self._insert_text_with_mixed_fonts(
                        page,
                        point,
                        center_text,
                        font_size,
                        "center",
                        text_rotation=original_rotation,
                    )

                # 右（主に日付を想定）
                right_text = footer.get("right", "")
                if footer.get("auto_date"):
                    right_text = current_date
                elif (
                    footer.get("auto_page_number")
                    and footer.get("page_number_position") == "right"
                ):
                    right_text = page_number_text

                if right_text:
                    point = self._calculate_position_with_rotation(
                        page, is_header=False, alignment="right", margin=20
                    )
                    self._insert_text_with_mixed_fonts(
                        page,
                        point,
                        right_text,
                        font_size,
                        "right",
                        text_rotation=original_rotation,
                    )

    def _run_export_images(self, items_data, output_dir, dpi=300, image_format="JPEG"):
        """選択されたページを画像として書き出す"""
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
            except Exception as e:
                self.signals.error.emit(
                    "出力ディレクトリを作成できませんでした",
                    f"出力ディレクトリ '{output_dir}' を作成できませんでした。\n\n"
                    f"エラー詳細: {e}",
                )
                return

        if not os.access(output_dir, os.W_OK):
            self.signals.error.emit(
                "保存先に書き込めません",
                f"出力ディレクトリ '{output_dir}' に書き込む権限がありません。",
            )
            return

        temp_files_to_clean = []

        # Officeファイルをタイプごとにグループ化してアプリケーションを再利用
        office_items_by_type = {"word": [], "excel": [], "powerpoint": []}
        pdf_items = []

        for item in items_data:
            item_type = item.get("type")
            if item_type in ("word", "excel", "powerpoint"):
                office_items_by_type[item_type].append(item)
            else:
                pdf_items.append(item)

        # Officeファイルの変換（アプリケーションを再利用）
        try:
            for office_type in ("word", "excel", "powerpoint"):
                if not office_items_by_type[office_type]:
                    continue

                app_name = office_type.capitalize()
                if office_type == "excel":
                    app_name = "Excel"
                elif office_type == "powerpoint":
                    app_name = "PowerPoint"

                app_instance = None
                for item in office_items_by_type[office_type]:
                    if not self.is_running:
                        break

                    original_path = item.get("original_path")
                    if not original_path:
                        continue

                    converter = None
                    if office_type == "word":
                        converter = self._convert_word_to_pdf
                    elif office_type == "excel":
                        converter = self._convert_excel_to_pdf
                    elif office_type == "powerpoint":
                        converter = self._convert_powerpoint_to_pdf

                    if converter:
                        temp_pdf, app_instance = converter(original_path, app_instance)
                        if not temp_pdf:
                            continue
                        temp_files_to_clean.append(temp_pdf)
                        # 変換後のPDFパスをアイテムに保存
                        item["converted_pdf_path"] = temp_pdf

                # アプリケーションを終了
                if app_instance:
                    try:
                        app_instance.Quit()
                    except Exception as e:
                        print(
                            f"警告: {app_name}アプリケーションの終了に失敗しました: {e}"
                        )
        except Exception as e:
            self.signals.error.emit(
                "Officeファイルの変換に失敗しました",
                f"Officeファイルの変換中にエラーが発生しました。\n\nエラー詳細: {e}",
            )
            return

        try:
            # PDFファイルと変換済みOfficeファイルを処理
            all_items = pdf_items + [
                item
                for items in office_items_by_type.values()
                for item in items
                if item.get("converted_pdf_path")
            ]

            for idx, item in enumerate(all_items):
                if not self.is_running:
                    break

                self.signals.progress.emit(
                    int((idx / len(all_items)) * 100),
                    f"画像変換中: {os.path.basename(item.get('original_path', ''))}",
                )

                item_type = item.get("type")
                original_path = item.get("original_path")

                if not original_path:
                    continue

                # Officeファイルの場合は変換済みPDFを使用
                if item_type in ("word", "excel", "powerpoint"):
                    source_path = item.get("converted_pdf_path", original_path)
                else:
                    source_path = original_path

                # PDFから画像に変換
                try:
                    with fitz.open(source_path) as doc:
                        if item_type == "pdf":
                            page_num = item.get("page_num", 0)
                            if page_num >= len(doc):
                                continue
                            page = doc[page_num]
                            rotation = item.get("rotation", 0)
                            if rotation != 0:
                                page.set_rotation(rotation)
                        else:
                            # Officeファイルの場合は最初のページ
                            page = doc[0]

                        # 解像度を設定（DPIからスケールを計算）
                        zoom = dpi / 72.0
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat)

                        # ファイル名を生成
                        base_name = os.path.splitext(os.path.basename(original_path))[0]
                        if item_type == "pdf":
                            page_suffix = f"_p{item.get('page_num', 0) + 1:03d}"
                        else:
                            page_suffix = "_p001"

                        ext = "jpg" if image_format == "JPEG" else image_format.lower()
                        output_filename = f"{base_name}{page_suffix}.{ext}"
                        output_path = os.path.join(output_dir, output_filename)

                        # ファイル名の重複を避ける
                        counter = 1
                        while os.path.exists(output_path):
                            output_filename = (
                                f"{base_name}{page_suffix}_{counter:03d}.{ext}"
                            )
                            output_path = os.path.join(output_dir, output_filename)
                            counter += 1

                        # 画像を保存
                        if image_format == "JPEG":
                            pix.save(output_path, output="jpeg", jpg_quality=95)
                        else:
                            pix.save(output_path, output=image_format.lower())

                except Exception as e:
                    self.signals.error.emit(
                        "画像変換エラー",
                        f"'{os.path.basename(original_path)}' の画像変換に失敗しました。\n\n"
                        f"エラー詳細: {e}",
                    )
                    continue

        finally:
            # 一時ファイルをクリーンアップ
            for temp_f in temp_files_to_clean:
                try:
                    if os.path.exists(temp_f):
                        os.remove(temp_f)
                except Exception as e:
                    print(f"一時ファイルの削除に失敗: {temp_f}, エラー: {e}")

        if self.is_running:
            self.signals.progress.emit(100, "画像書き出し完了")
            self.signals.finished.emit(
                self.task_name,
                "完了",
                f"画像を正常に書き出しました:\n{output_dir}",
            )


class HeaderFooterSettingsDialog(QDialog):
    """ヘッダー・フッター設定ダイアログ"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ヘッダー・フッターの設定")
        self.setMinimumWidth(500)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # ヘッダーを追加するかどうか
        self.header_enable_checkbox = QCheckBox("ヘッダーを追加する")
        self.header_enable_checkbox.setChecked(False)
        layout.addWidget(self.header_enable_checkbox)

        # ヘッダー設定グループ
        header_group = QWidget()
        header_layout = QVBoxLayout(header_group)

        # ヘッダー左
        header_left_layout = QHBoxLayout()
        header_left_layout.addWidget(QLabel("左:"))
        self.header_left_text = QLineEdit()
        self.header_left_text.setPlaceholderText("例: ○○市役所")
        header_left_layout.addWidget(self.header_left_text)
        header_layout.addLayout(header_left_layout)

        # ヘッダー中央
        header_center_layout = QHBoxLayout()
        header_center_layout.addWidget(QLabel("中央:"))
        self.header_center_text = QLineEdit()
        self.header_center_text.setPlaceholderText("例: 文書名")
        header_center_layout.addWidget(self.header_center_text)
        header_layout.addLayout(header_center_layout)

        # ヘッダー右
        header_right_layout = QHBoxLayout()
        header_right_layout.addWidget(QLabel("右:"))
        self.header_right_text = QLineEdit()
        self.header_right_text.setPlaceholderText("例: 令和○年○月○日")
        header_right_layout.addWidget(self.header_right_text)
        header_layout.addLayout(header_right_layout)

        # 日付を自動挿入（ヘッダー右）
        self.header_date_checkbox = QCheckBox("右側に現在の日付を自動挿入")
        self.header_date_checkbox.setChecked(False)
        header_layout.addWidget(self.header_date_checkbox)

        # ページ番号を自動挿入（ヘッダー）
        page_number_header_layout = QHBoxLayout()
        page_number_header_layout.addWidget(QLabel("ページ番号を自動挿入:"))
        self.header_page_number_position_combo = QComboBox()
        self.header_page_number_position_combo.addItems(["なし", "左", "中央", "右"])
        self.header_page_number_position_combo.setCurrentText("なし")
        page_number_header_layout.addWidget(self.header_page_number_position_combo)
        page_number_header_layout.addStretch()
        header_layout.addLayout(page_number_header_layout)

        layout.addWidget(header_group)

        # チェックボックスが外されていても、設定内容は表示・編集可能にする
        # （チェックボックスは「保存時に適用するかどうか」のフラグとして機能）

        layout.addWidget(QFrame())  # 区切り線の代わり

        # フッターを追加するかどうか
        self.footer_enable_checkbox = QCheckBox("フッターを追加する")
        self.footer_enable_checkbox.setChecked(False)
        layout.addWidget(self.footer_enable_checkbox)

        # フッター設定グループ
        footer_group = QWidget()
        footer_layout = QVBoxLayout(footer_group)

        # フッター左
        footer_left_layout = QHBoxLayout()
        footer_left_layout.addWidget(QLabel("左:"))
        self.footer_left_text = QLineEdit()
        self.footer_left_text.setPlaceholderText("例: ○○市役所")
        footer_left_layout.addWidget(self.footer_left_text)
        footer_layout.addLayout(footer_left_layout)

        # フッター中央
        footer_center_layout = QHBoxLayout()
        footer_center_layout.addWidget(QLabel("中央:"))
        self.footer_center_text = QLineEdit()
        self.footer_center_text.setPlaceholderText("例: 文書名")
        footer_center_layout.addWidget(self.footer_center_text)
        footer_layout.addLayout(footer_center_layout)

        # フッター右
        footer_right_layout = QHBoxLayout()
        footer_right_layout.addWidget(QLabel("右:"))
        self.footer_right_text = QLineEdit()
        self.footer_right_text.setPlaceholderText("例: 令和○年○月○日")
        footer_right_layout.addWidget(self.footer_right_text)
        footer_layout.addLayout(footer_right_layout)

        # 日付を自動挿入（フッター右）
        self.footer_date_checkbox = QCheckBox("右側に現在の日付を自動挿入")
        self.footer_date_checkbox.setChecked(False)
        footer_layout.addWidget(self.footer_date_checkbox)

        # ページ番号を自動挿入（フッター）
        page_number_footer_layout = QHBoxLayout()
        page_number_footer_layout.addWidget(QLabel("ページ番号を自動挿入:"))
        self.footer_page_number_position_combo = QComboBox()
        self.footer_page_number_position_combo.addItems(["なし", "左", "中央", "右"])
        self.footer_page_number_position_combo.setCurrentText("なし")
        page_number_footer_layout.addWidget(self.footer_page_number_position_combo)
        page_number_footer_layout.addStretch()
        footer_layout.addLayout(page_number_footer_layout)

        layout.addWidget(footer_group)

        # チェックボックスが外されていても、設定内容は表示・編集可能にする
        # （チェックボックスは「保存時に適用するかどうか」のフラグとして機能）

        # フォントサイズ
        font_size_layout = QHBoxLayout()
        font_size_layout.addWidget(QLabel("フォントサイズ:"))
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setMinimum(6)
        self.font_size_spin.setMaximum(72)
        self.font_size_spin.setValue(10)
        font_size_layout.addWidget(self.font_size_spin)
        font_size_layout.addStretch()
        layout.addLayout(font_size_layout)

        # ページ番号設定
        page_number_group = QGroupBox("ページ番号設定")
        page_number_layout = QVBoxLayout(page_number_group)

        # 表示形式
        format_layout = QHBoxLayout()
        format_layout.addWidget(QLabel("表示形式:"))
        self.page_number_format_combo = QComboBox()
        self.page_number_format_combo.addItems(["1", "1 / 10", "Page 1", "- 1 -"])
        format_layout.addWidget(self.page_number_format_combo)
        format_layout.addStretch()
        page_number_layout.addLayout(format_layout)

        # 開始番号
        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("開始番号:"))
        self.page_number_start_spin = QSpinBox()
        self.page_number_start_spin.setMinimum(1)
        self.page_number_start_spin.setMaximum(9999)
        self.page_number_start_spin.setValue(1)
        start_layout.addWidget(self.page_number_start_spin)
        start_layout.addStretch()
        page_number_layout.addLayout(start_layout)

        layout.addWidget(page_number_group)

        # ボタン
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def get_settings(self):
        """設定を取得（チェックボックスの状態を使用）"""
        header_enabled = self.header_enable_checkbox.isChecked()
        footer_enabled = self.footer_enable_checkbox.isChecked()

        # ヘッダーの設定
        header_left = self.header_left_text.text()
        header_center = self.header_center_text.text()
        header_right = self.header_right_text.text()
        header_auto_date = self.header_date_checkbox.isChecked()

        header_page_number_position = None
        header_page_number_text = self.header_page_number_position_combo.currentText()
        if header_page_number_text != "なし":
            if header_page_number_text == "中央":
                header_page_number_position = "center"
            elif header_page_number_text == "右":
                header_page_number_position = "right"
            else:
                header_page_number_position = "left"

        # フッターの設定
        footer_left = self.footer_left_text.text()
        footer_center = self.footer_center_text.text()
        footer_right = self.footer_right_text.text()
        footer_auto_date = self.footer_date_checkbox.isChecked()

        footer_page_number_position = None
        footer_page_number_text = self.footer_page_number_position_combo.currentText()
        if footer_page_number_text != "なし":
            if footer_page_number_text == "中央":
                footer_page_number_position = "center"
            elif footer_page_number_text == "右":
                footer_page_number_position = "right"
            else:
                footer_page_number_position = "left"

        settings = {
            "header_enabled": header_enabled,
            "footer_enabled": footer_enabled,
            "font_size": self.font_size_spin.value(),
        }

        settings["header"] = {
            "left": header_left,
            "center": header_center,
            "right": header_right,
            "auto_date": header_auto_date,
            "auto_page_number": header_page_number_position is not None,
            "page_number_position": header_page_number_position,
        }

        settings["footer"] = {
            "left": footer_left,
            "center": footer_center,
            "right": footer_right,
            "auto_date": footer_auto_date,
            "auto_page_number": footer_page_number_position is not None,
            "page_number_position": footer_page_number_position,
        }

        # ページ番号のフォーマットと開始番号
        settings["page_number_format"] = self.page_number_format_combo.currentText()
        settings["page_number_start"] = self.page_number_start_spin.value()

        return settings

    def set_settings(self, settings):
        """設定を反映"""
        if not settings:
            return

        # ヘッダー・フッターのON/OFF
        self.header_enable_checkbox.setChecked(settings.get("header_enabled", False))
        self.footer_enable_checkbox.setChecked(settings.get("footer_enabled", False))

        # ヘッダー設定
        if "header" in settings:
            header = settings["header"]
            self.header_left_text.setText(header.get("left", ""))
            self.header_center_text.setText(header.get("center", ""))
            self.header_right_text.setText(header.get("right", ""))
            self.header_date_checkbox.setChecked(header.get("auto_date", False))
            position = header.get("page_number_position")
            if position == "center":
                self.header_page_number_position_combo.setCurrentText("中央")
            elif position == "right":
                self.header_page_number_position_combo.setCurrentText("右")
            elif position == "left":
                self.header_page_number_position_combo.setCurrentText("左")
            else:
                self.header_page_number_position_combo.setCurrentText("なし")

        # フッター設定
        if "footer" in settings:
            footer = settings["footer"]
            self.footer_left_text.setText(footer.get("left", ""))
            self.footer_center_text.setText(footer.get("center", ""))
            self.footer_right_text.setText(footer.get("right", ""))
            self.footer_date_checkbox.setChecked(footer.get("auto_date", False))
            position = footer.get("page_number_position")
            if position == "center":
                self.footer_page_number_position_combo.setCurrentText("中央")
            elif position == "right":
                self.footer_page_number_position_combo.setCurrentText("右")
            elif position == "left":
                self.footer_page_number_position_combo.setCurrentText("左")
            else:
                self.footer_page_number_position_combo.setCurrentText("なし")

        # フォントサイズ
        self.font_size_spin.setValue(settings.get("font_size", 10))

        # ページ番号設定
        self.page_number_format_combo.setCurrentText(
            settings.get("page_number_format", "1")
        )
        self.page_number_start_spin.setValue(settings.get("page_number_start", 1))


class OfficePDFBinderApp(QMainWindow):
    # --- レイアウト定数 ---
    # グローバル定数を参照
    THUMBNAIL_WIDTH = THUMBNAIL_WIDTH
    THUMBNAIL_HEIGHT = THUMBNAIL_HEIGHT

    GRID_ITEM_WIDTH = THUMBNAIL_WIDTH + GRID_ITEM_PADDING_X  # サムネイル幅 + 左右の余白
    GRID_ITEM_HEIGHT = (
        THUMBNAIL_HEIGHT + GRID_ITEM_PADDING_Y
    )  # サムネイル高 + テキスト領域の高さ

    THUMBNAIL_SIZE = QSize(THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT)
    GRID_ITEM_SIZE = QSize(GRID_ITEM_WIDTH, GRID_ITEM_HEIGHT)

    # IPC で受信したファイルパスをメインスレッド側に伝えるシグナル
    ipc_files_received = Signal(list)

    def __init__(self, initial_geometry=None, initially_maximized=False):
        super().__init__()
        self.setWindowTitle("Office PDF Binder")
        self.setWindowIcon(_get_app_icon())
        self.setMinimumSize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)
        if initial_geometry is None:
            initial_geometry = _default_window_geometry()
        self.setGeometry(initial_geometry)
        self._restore_maximized = initially_maximized
        self.threadpool = QThreadPool()
        self.current_worker = None
        self.bookmarks = []
        self.auto_bookmarks_enabled = True
        self.show_bookmarks_on_open = True
        # ページ番号設定のデフォルト値
        self.page_number_settings = {
            "enabled": False,
            "is_header": False,
            "alignment": "center",
            "format": "1",
            "start_number": 1,
            "font_size": 10,
        }
        # ヘッダー・フッター設定のデフォルト値
        self.header_footer_settings = None
        if getattr(sys, "frozen", False):
            self.install_dir = os.path.dirname(sys.executable)
        else:
            self.install_dir = os.path.dirname(os.path.abspath(__file__))
        self.user_manual_path = os.path.join(self.install_dir, "README.html")
        # IPC シグナルとハンドラを接続
        self.ipc_files_received.connect(self._add_files_from_paths)
        # ズームは約1.75倍刻みの5段階（最小約0.33〜最大約3.06）
        self.zoom_levels = [1 / (1.75**2), 1 / 1.75, 1.0, 1.75, 1.75**2]
        self.zoom_index = 2  # 1.0（デフォルト）
        self.zoom_level = self.zoom_levels[self.zoom_index]
        self.zoom_wheel_accumulator = 0  # ホイールイベントの累積量
        self.zoom_timer = QTimer()
        self.zoom_timer.setSingleShot(True)
        self.zoom_timer.timeout.connect(self._process_accumulated_zoom)
        self.thumbnail_regen_timer = QTimer()
        self.thumbnail_regen_timer.setSingleShot(True)
        self.thumbnail_regen_timer.timeout.connect(self._regenerate_thumbnails)
        self.pending_thumbnail_size = None  # 再生成待ちのサムネイルサイズ
        max_zoom = max(self.zoom_levels)
        self.max_thumbnail_size = QSize(
            int(THUMBNAIL_WIDTH * max_zoom), int(THUMBNAIL_HEIGHT * max_zoom)
        )
        self.thumbnail_cache = {}
        self.undo_stack = []
        self.redo_stack = []
        self._pending_file_paths = []  # IPC経由で複数ファイルが送られてきた場合のキュー
        # IPC経由のファイル追加を短時間バッファし、自然順でまとめて処理するためのキューとタイマー
        self._ipc_pending_file_paths = []
        self._ipc_batch_timer = QTimer(self)
        self._ipc_batch_timer.setSingleShot(True)
        self._ipc_batch_timer.timeout.connect(self._flush_ipc_pending_files)

        # ユーザーのAppDataフォルダ内に設定ファイルを保存します
        # これにより、ユーザーごとの設定が保持され、アクセス権の問題も回避できます
        self.settings_dir = os.path.dirname(_get_settings_file_path())
        os.makedirs(self.settings_dir, exist_ok=True)
        self.settings_file = _get_settings_file_path()

        self.config = configparser.ConfigParser()
        self.last_used_path = ""  # 最後に使用したパスを保持する変数
        self._load_settings()  # 起動時に設定を読み込む
        # アプリ起動時に設定をリセット（常に初期状態から始める）
        self._reset_page_settings()

        self.setup_ui()
        self.create_actions()
        self.create_toolbar()
        self.create_menu_bar()
        self.progress_dialog = None
        self.apply_stylesheet()
        self.update_status_bar()
        self._update_page_mode_actions_state()
        self._setup_shortcuts()
        # 初期ズームレベルの状態を設定
        self._apply_zoom()
        self._record_history_change(initial=True)

        if self._restore_maximized:
            self.setWindowState(self.windowState() | Qt.WindowState.WindowMaximized)

    def _load_settings(self):
        """設定ファイルから設定を読み込みます。"""
        try:
            if os.path.exists(self.settings_file):
                self.config.read(self.settings_file, encoding="utf-8")
                # 'Paths'セクションの'last_used'というキーから値を取得
                # fallback=""は、もしキーが存在しなかった場合に空文字を返す安全装置
                self.last_used_path = self.config.get("Paths", "last_used", fallback="")
                self.auto_bookmarks_enabled = self.config.getboolean(
                    "Bookmarks", "auto_generation", fallback=True
                )
                self.show_bookmarks_on_open = self.config.getboolean(
                    "Bookmarks", "show_on_open", fallback=True
                )
                # ページ番号とヘッダー・フッターの設定は読み込まない
                # （アプリ起動時に常にリセットされるため）
        except Exception as e:
            print(f"設定の読み込み中にエラーが発生しました: {e}")
            # 読み込みに失敗した場合も、空のパスで続行する
            self.last_used_path = ""
            # ページ番号とヘッダー・フッターの設定は、アプリ起動時にリセットされるため、ここでは処理しない

    def _save_settings(self):
        """現在の設定を設定ファイルに保存します。"""
        try:
            # 'Paths'というセクションがなければ作成
            if not self.config.has_section("Paths"):
                self.config.add_section("Paths")
            if not self.config.has_section("Bookmarks"):
                self.config.add_section("Bookmarks")
            if not self.config.has_section("PageNumbers"):
                self.config.add_section("PageNumbers")
            if not self.config.has_section("HeaderFooter"):
                self.config.add_section("HeaderFooter")
            if not self.config.has_section("Window"):
                self.config.add_section("Window")

            # 'last_used'というキーに、現在のパスを設定
            self.config.set("Paths", "last_used", self.last_used_path)
            self.config.set(
                "Bookmarks", "auto_generation", str(self.auto_bookmarks_enabled)
            )
            self.config.set(
                "Bookmarks", "show_on_open", str(self.show_bookmarks_on_open)
            )
            window_rect = self.normalGeometry() if self.isMaximized() else self.geometry()
            window_rect = _fit_rect_to_available_geometry(window_rect)
            self.config.set("Window", "x", str(window_rect.x()))
            self.config.set("Window", "y", str(window_rect.y()))
            self.config.set("Window", "width", str(window_rect.width()))
            self.config.set("Window", "height", str(window_rect.height()))
            self.config.set("Window", "maximized", str(self.isMaximized()))
            # ページ番号設定を保存
            self.config.set(
                "PageNumbers", "enabled", str(self.page_number_settings["enabled"])
            )
            self.config.set(
                "PageNumbers", "is_header", str(self.page_number_settings["is_header"])
            )
            self.config.set(
                "PageNumbers", "alignment", self.page_number_settings["alignment"]
            )
            self.config.set(
                "PageNumbers", "format", self.page_number_settings["format"]
            )
            self.config.set(
                "PageNumbers",
                "start_number",
                str(self.page_number_settings["start_number"]),
            )
            self.config.set(
                "PageNumbers", "font_size", str(self.page_number_settings["font_size"])
            )
            # ヘッダー・フッター設定を保存（チェックボックスが外されていても設定内容は保持）
            if self.header_footer_settings:
                self.config.set(
                    "HeaderFooter",
                    "header_enabled",
                    str(self.header_footer_settings.get("header_enabled", False)),
                )
                self.config.set(
                    "HeaderFooter",
                    "footer_enabled",
                    str(self.header_footer_settings.get("footer_enabled", False)),
                )
                self.config.set(
                    "HeaderFooter",
                    "font_size",
                    str(self.header_footer_settings.get("font_size", 10)),
                )
                # チェックボックスが外されていても、設定内容は保存する
                if "header" in self.header_footer_settings:
                    header = self.header_footer_settings["header"]
                    self.config.set(
                        "HeaderFooter", "header_left", header.get("left", "")
                    )
                    self.config.set(
                        "HeaderFooter", "header_center", header.get("center", "")
                    )
                    self.config.set(
                        "HeaderFooter", "header_right", header.get("right", "")
                    )
                    self.config.set(
                        "HeaderFooter",
                        "header_auto_date",
                        str(header.get("auto_date", False)),
                    )
                if "footer" in self.header_footer_settings:
                    footer = self.header_footer_settings["footer"]
                    self.config.set(
                        "HeaderFooter", "footer_left", footer.get("left", "")
                    )
                    self.config.set(
                        "HeaderFooter", "footer_center", footer.get("center", "")
                    )
                    self.config.set(
                        "HeaderFooter", "footer_right", footer.get("right", "")
                    )
                    self.config.set(
                        "HeaderFooter",
                        "footer_auto_date",
                        str(footer.get("auto_date", False)),
                    )

            # 設定ファイルに書き出す
            with open(self.settings_file, "w", encoding="utf-8") as configfile:
                self.config.write(configfile)
        except Exception as e:
            print(f"設定の保存中にエラーが発生しました: {e}")

    def get_icon(self, name, color="white"):
        if qta:
            return qta.icon(name, color=color)
        return QIcon()

    def _setup_shortcuts(self):
        """キーボードショートカットを設定する"""
        # アクションにショートカットが既に設定されているので、ここでは追加の設定のみ
        # 将来の拡張用にメソッドを用意
        pass

    def setup_ui(self):
        self.central_widget = QWidget()
        self.central_widget.setAcceptDrops(
            True
        )  # 中央ウィジェットにもドロップを有効にする
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout(self.central_widget)
        layout.setContentsMargins(0, 0, 0, 0)

        # --- 1. ページモード用ウィジェット ---
        self.page_list_widget = DropListWidget()
        self.page_list_widget.files_dropped.connect(self._handle_dropped_files)
        self.page_list_widget.zoom_requested.connect(self._on_zoom_requested)
        self.page_list_widget.setViewMode(QListWidget.IconMode)
        self.page_list_widget.setResizeMode(QListWidget.Adjust)
        self.page_list_widget.setGridSize(self.GRID_ITEM_SIZE)
        self.page_list_widget.setUniformItemSizes(True)
        self.page_list_widget.setMovement(QListWidget.Static)
        self.page_list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        self.page_list_widget.setIconSize(self.THUMBNAIL_SIZE)
        page_item_style = """
            QListWidget::item { background-color: #4a5260; border: none; border-radius: 5px; padding: 10px; margin: 0px; }
            QListWidget::item:selected { background-color: #0984e3; border: 2px solid #74b9ff; color: white; }
            QListWidget::item:hover:!selected { background-color: #5d6776; }
        """
        self.page_list_widget.setStyleSheet(page_item_style)
        self.page_list_widget.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.page_list_widget.customContextMenuRequested.connect(
            self._show_page_context_menu
        )
        self.page_list_widget.itemSelectionChanged.connect(self.update_status_bar)
        self.page_list_widget.itemSelectionChanged.connect(
            self._update_page_mode_actions_state
        )
        self.page_list_widget.itemSelectionChanged.connect(
            self._update_bookmark_add_state
        )
        self.page_list_widget.itemDoubleClicked.connect(self._open_source_file)
        layout.addWidget(self.page_list_widget)

        # ドラッグ&ドロップ完了時のコールバックを設定
        self.page_list_widget._on_items_moved_callback = self._on_items_moved
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.file_status_label = QLabel()
        self.selection_status_label = QLabel()
        self.status_bar.addWidget(self.file_status_label)
        self.status_bar.addWidget(self.selection_status_label)

        # --- しおりパネル ---
        self.bookmark_dock = QDockWidget("しおり", self)
        self.bookmark_dock.setObjectName("BookmarkDock")
        bookmark_widget = QWidget()
        bookmark_layout = QVBoxLayout(bookmark_widget)
        bookmark_layout.setContentsMargins(5, 5, 5, 5)
        self.BOOKMARK_INDEX_ROLE = Qt.UserRole + 1
        self.bookmark_tree = QTreeWidget()
        self.bookmark_tree.setHeaderHidden(True)
        self.bookmark_tree.itemDoubleClicked.connect(self._navigate_to_bookmark)
        self.bookmark_tree.itemSelectionChanged.connect(
            self._update_bookmark_buttons_state
        )
        bookmark_layout.addWidget(self.bookmark_tree)

        button_row = QHBoxLayout()
        self.bookmark_add_button = QPushButton("追加")
        self.bookmark_add_button.setToolTip("選択ページにしおりを追加")
        self.bookmark_add_button.clicked.connect(self._add_bookmark_from_selection)
        button_row.addWidget(self.bookmark_add_button)

        self.bookmark_rename_button = QPushButton("名前変更")
        self.bookmark_rename_button.setToolTip("選択したしおりの名前を変更")
        self.bookmark_rename_button.clicked.connect(self._rename_selected_bookmark)
        button_row.addWidget(self.bookmark_rename_button)

        self.bookmark_delete_button = QPushButton("削除")
        self.bookmark_delete_button.setToolTip("選択したしおりを削除")
        self.bookmark_delete_button.clicked.connect(self._delete_selected_bookmark)
        button_row.addWidget(self.bookmark_delete_button)

        bookmark_layout.addLayout(button_row)

        info_label = QLabel(
            "※自動しおりはメニューからON/OFFできます。自動しおりも削除・編集可能です（編集すると手動しおりに変換されます）。"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #bdc3c7; font-size: 9pt;")
        bookmark_layout.addWidget(info_label)
        self.bookmark_dock.setWidget(bookmark_widget)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.bookmark_dock)
        self.bookmark_dock.setVisible(False)
        self.bookmark_dock.visibilityChanged.connect(
            self._on_bookmark_dock_visibility_changed
        )
        self._update_bookmark_add_state()
        self._update_bookmark_buttons_state()
        self._on_bookmark_dock_visibility_changed(self.bookmark_dock.isVisible())

    def create_actions(self):
        # --- ファイル操作グループ ---
        self.add_action = QAction(
            self.get_icon("fa5s.plus-circle", color="#2ecc71"), "ファイル\n追加", self
        )
        self.add_action.setToolTip("ファイル\n追加")
        self.add_action.triggered.connect(self._add_files)

        self.delete_action = QAction(
            self.get_icon("fa5s.trash-alt", color="#e74c3c"), "選択項目\n削除", self
        )
        self.delete_action.setToolTip("選択項目\n削除 (Delete)")
        self.delete_action.setShortcut("Delete")
        self.delete_action.triggered.connect(self._delete_selected)

        # --- ページ編集グループ ---
        self.rot_left_action = QAction(
            self.get_icon("fa5s.undo", color="#3498db"), "左へ90°\n回転", self
        )
        self.rot_left_action.setToolTip(
            "選択したPDFページを左に90度回転させます。\n複数のページを選択して一括で回転することもできます。"
        )
        self.rot_left_action.triggered.connect(lambda: self._rotate_selected(-90))

        self.rot_right_action = QAction(
            self.get_icon("fa5s.redo", color="#3498db"), "右へ90°\n回転", self
        )
        self.rot_right_action.setToolTip(
            "選択したPDFページを右に90度回転させます。\n複数のページを選択して一括で回転することもできます。"
        )
        self.rot_right_action.triggered.connect(lambda: self._rotate_selected(90))

        self.move_to_top_action = QAction(
            self.get_icon("fa5s.angle-double-up", color="#f1c40f"),
            "一番上へ\n移動",
            self,
        )
        self.move_to_top_action.setToolTip(
            "選択したページをリストの一番上に移動します。\n複数のページを選択して一括で移動することもできます。"
        )
        self.move_to_top_action.triggered.connect(self._move_to_top)

        self.move_up_action = QAction(
            self.get_icon("fa5s.arrow-up", color="#f1c40f"), "上へ\n移動", self
        )
        self.move_up_action.setToolTip(
            "選択したページを1つ上に移動します。\n複数のページを選択して一括で移動することもできます。"
        )
        self.move_up_action.triggered.connect(self._move_up)

        self.move_down_action = QAction(
            self.get_icon("fa5s.arrow-down", color="#f1c40f"), "下へ\n移動", self
        )
        self.move_down_action.setToolTip(
            "選択したページを1つ下に移動します。\n複数のページを選択して一括で移動することもできます。"
        )
        self.move_down_action.triggered.connect(self._move_down)

        self.move_to_bottom_action = QAction(
            self.get_icon("fa5s.angle-double-down", color="#f1c40f"),
            "一番下へ\n移動",
            self,
        )
        self.move_to_bottom_action.setToolTip(
            "選択したページをリストの一番下に移動します。\n複数のページを選択して一括で移動することもできます。"
        )
        self.move_to_bottom_action.triggered.connect(self._move_to_bottom)

        # --- 保存とモード変更グループ ---
        self.merge_action = QAction(
            self.get_icon("fa5s.save", color="#3498db"), "名前を付けて\n保存", self
        )
        self.merge_action.setToolTip("名前を付けて保存")
        self.merge_action.triggered.connect(self._merge_and_save)

        self.about_action = QAction("アプリ情報(&I)", self)
        self.about_action.setToolTip("このアプリケーションについての情報を表示します")
        self.about_action.triggered.connect(self._show_about_dialog)

        self.open_manual_action = QAction("マニュアル(&M)", self)
        self.open_manual_action.setToolTip(
            "ユーザーマニュアル (HTML) を既定のブラウザで開きます"
        )
        self.open_manual_action.triggered.connect(self._open_user_manual)

        # --- しおりパネルアクション ---
        self.bookmark_panel_action = QAction("しおり(&B)", self)
        self.bookmark_panel_action.setCheckable(True)
        self.bookmark_panel_action.setChecked(False)
        self.bookmark_panel_action.triggered.connect(self._toggle_bookmark_panel)

        # --- メニュー用アクション（キーボードショートカット付き） ---
        # ファイルメニュー
        self.new_action = QAction("新規(&N)", self)
        self.new_action.setShortcut(QKeySequence.New)
        self.new_action.setToolTip("新しいプロジェクトを開始")
        self.new_action.triggered.connect(self._new_project)

        self.open_action = QAction("ファイルを追加(&O)...", self)
        self.open_action.setShortcut(QKeySequence.Open)
        self.open_action.setIcon(self.get_icon("fa5s.folder-open", color="#2ecc71"))
        self.open_action.setToolTip("ファイルを追加")
        self.open_action.triggered.connect(self._add_files)

        self.save_action = QAction("名前を付けて保存(&S)...", self)
        self.save_action.setShortcut(QKeySequence.Save)
        self.save_action.setIcon(self.get_icon("fa5s.save", color="#3498db"))
        self.save_action.setToolTip("名前を付けて保存")
        self.save_action.triggered.connect(self._merge_and_save)

        self.export_selected_pdf_action = QAction(
            "選択ページをPDFとして書き出し(&E)...", self
        )
        self.export_selected_pdf_action.setToolTip(
            "選択したページをPDFファイルとして書き出します"
        )
        self.export_selected_pdf_action.triggered.connect(self._export_selected_as_pdf)

        self.export_selected_images_action = QAction(
            "選択ページを画像として書き出し(&I)...", self
        )
        self.export_selected_images_action.setToolTip(
            "選択したページをJPEG画像として書き出します"
        )
        self.export_selected_images_action.triggered.connect(
            self._export_selected_as_images
        )

        self.exit_action = QAction("終了(&X)", self)
        self.exit_action.setShortcut(QKeySequence.Quit)
        self.exit_action.setToolTip("アプリケーションを終了")
        self.exit_action.triggered.connect(self.close)

        # 編集メニュー
        self.undo_action = QAction("元に戻す(&U)", self)
        self.undo_action.setShortcut(QKeySequence.Undo)
        self.undo_action.setToolTip("操作を元に戻す")
        self.undo_action.setEnabled(False)
        self.undo_action.triggered.connect(self._undo)

        self.redo_action = QAction("やり直す(&R)", self)
        self.redo_action.setShortcut(QKeySequence.Redo)
        self.redo_action.setToolTip("操作をやり直す")
        self.redo_action.setEnabled(False)
        self.redo_action.triggered.connect(self._redo)

        self.select_all_action = QAction("すべて選択(&A)", self)
        self.select_all_action.setShortcut(QKeySequence.SelectAll)
        self.select_all_action.setToolTip("すべてのアイテムを選択")
        self.select_all_action.triggered.connect(self._select_all)

        # 表示メニュー
        self.zoom_in_action = QAction("拡大(&I)", self)
        self.zoom_in_action.setShortcut(QKeySequence.ZoomIn)
        self.zoom_in_action.setToolTip("表示を拡大 (Ctrl++)")
        self.zoom_in_action.triggered.connect(self._zoom_in)

        self.zoom_out_action = QAction("縮小(&O)", self)
        self.zoom_out_action.setShortcut(QKeySequence.ZoomOut)
        self.zoom_out_action.setToolTip("表示を縮小 (Ctrl+-)")
        self.zoom_out_action.triggered.connect(self._zoom_out)

        self.zoom_fit_action = QAction("表示サイズに合わせる(&F)", self)
        self.zoom_fit_action.setShortcut("Ctrl+0")
        self.zoom_fit_action.setToolTip("表示サイズに合わせる (Ctrl+0)")
        self.zoom_fit_action.triggered.connect(self._zoom_fit)

        # ブックマーク設定
        self.auto_bookmark_action = QAction("ファイルごとに自動しおり", self)
        self.auto_bookmark_action.setCheckable(True)
        self.auto_bookmark_action.setChecked(self.auto_bookmarks_enabled)
        self.auto_bookmark_action.setToolTip(
            "ファイルの先頭ページに自動でしおりを追加します"
        )
        self.auto_bookmark_action.triggered.connect(self._toggle_auto_bookmarks)

        self.show_outline_on_open_action = QAction("PDF閲覧時にしおりを自動表示", self)
        self.show_outline_on_open_action.setCheckable(True)
        self.show_outline_on_open_action.setChecked(self.show_bookmarks_on_open)
        self.show_outline_on_open_action.setToolTip(
            "結合後のPDFをビューアで開く際に、しおりペインを自動で表示します"
        )
        self.show_outline_on_open_action.triggered.connect(
            self._toggle_show_bookmarks_on_open
        )

        # ヘッダー・フッター設定
        self.header_footer_settings_action = QAction(
            "ヘッダー・フッター設定(&H)...", self
        )
        self.header_footer_settings_action.setShortcut("Ctrl+H")
        self.header_footer_settings_action.setToolTip(
            "ヘッダー・フッターの設定を行います (Ctrl+H)"
        )
        self.header_footer_settings_action.triggered.connect(
            self._show_header_footer_settings
        )

    def create_toolbar(self):
        toolbar = QToolBar("メインツールバー")
        toolbar.setIconSize(QSize(22, 22))
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        toolbar.setAllowedAreas(Qt.ToolBarArea.TopToolBarArea)

        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)

        self.addToolBar(toolbar)

        # --- グループ1: ファイル操作 ---
        toolbar.addAction(self.add_action)
        toolbar.addAction(self.delete_action)
        toolbar.addSeparator()

        # --- グループ2: ページ・ファイル編集 ---
        # 2-1: 回転
        toolbar.addAction(self.rot_left_action)
        toolbar.addAction(self.rot_right_action)

        # 2-2: 順序編集 (移動とモード変更)
        toolbar.addAction(self.move_to_top_action)
        toolbar.addAction(self.move_up_action)
        toolbar.addAction(self.move_down_action)
        toolbar.addAction(self.move_to_bottom_action)
        toolbar.addSeparator()

        # --- グループ3: 保存 ---
        toolbar.addAction(self.merge_action)

    def create_menu_bar(self):
        """
        QHBoxLayoutを使い、メニューバーを作成します。
        ヘルプメニューは右端に固定されます。
        """
        # 1. メニューバー全体を乗せるためのコンテナWidgetを作成
        custom_menubar_widget = QWidget()
        custom_menubar_widget.setObjectName("CustomMenuBarContainer")

        # 2. 水平ボックスレイアウトを作成し、コンテナにセット
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        custom_menubar_widget.setLayout(layout)

        # 3. メニューバーを作成
        menubar = QMenuBar()

        # ファイルメニュー
        file_menu = menubar.addMenu("ファイル(&F)")
        file_menu.addAction(self.new_action)
        file_menu.addAction(self.open_action)
        file_menu.addSeparator()
        file_menu.addAction(self.save_action)
        file_menu.addSeparator()
        file_menu.addAction(self.export_selected_pdf_action)
        file_menu.addAction(self.export_selected_images_action)
        file_menu.addSeparator()
        file_menu.addAction(self.exit_action)

        # 編集メニュー
        edit_menu = menubar.addMenu("編集(&E)")
        edit_menu.addAction(self.undo_action)
        edit_menu.addAction(self.redo_action)
        edit_menu.addSeparator()
        edit_menu.addAction(self.select_all_action)
        edit_menu.addSeparator()
        edit_menu.addAction(self.delete_action)
        edit_menu.addSeparator()
        # ページ操作サブメニュー
        page_ops_menu = edit_menu.addMenu("ページ操作(&P)")
        page_ops_menu.addAction(self.move_to_top_action)
        page_ops_menu.addAction(self.move_up_action)
        page_ops_menu.addAction(self.move_down_action)
        page_ops_menu.addAction(self.move_to_bottom_action)
        page_ops_menu.addSeparator()
        page_ops_menu.addAction(self.rot_left_action)
        page_ops_menu.addAction(self.rot_right_action)

        # 表示メニュー
        view_menu = menubar.addMenu("表示(&V)")
        view_menu.addAction(self.zoom_in_action)
        view_menu.addAction(self.zoom_out_action)
        view_menu.addAction(self.zoom_fit_action)
        view_menu.addSeparator()
        view_menu.addAction(self.bookmark_panel_action)

        # 設定メニュー
        settings_menu = menubar.addMenu("設定(&S)")
        settings_menu.addAction(self.auto_bookmark_action)
        settings_menu.addAction(self.show_outline_on_open_action)
        settings_menu.addSeparator()
        settings_menu.addAction(self.header_footer_settings_action)

        # ヘルプ関連アクション（メニュー直下に表示）
        menubar.addAction(self.open_manual_action)
        menubar.addAction(self.about_action)

        # 4. レイアウトにメニューバーと伸縮する空白を追加
        layout.addWidget(menubar)
        layout.addStretch(1)  # 右端に空白を追加

        # 5. この自作ウィジェットを、QMainWindowのメニュー領域に設定
        self.setMenuWidget(custom_menubar_widget)

    def apply_stylesheet(self):
        self.setStyleSheet(QSS)

    def _run_task(self, task_name, user_facing_name, **kwargs):
        # デバッグ用: タスク開始前のスレッド数を記録
        _debug_log(
            f"[DEBUG] _run_task: task_name={task_name}, "
            f"activeThreadCount(before)={self.threadpool.activeThreadCount()}"
        )
        if self.threadpool.activeThreadCount() > 0:
            _debug_log(
                f"[DEBUG] _run_task: task_name={task_name} を開始しようとしたが、"
                "既に別のタスクが実行中のためキャンセルしました。"
            )
            QMessageBox.warning(self, "処理中", "現在別の処理を実行中です。")
            return
        self.current_worker = AppWorker(task_name, **kwargs)
        self.setup_progress_dialog(user_facing_name)
        self.current_worker.signals.progress.connect(self.update_progress)
        self.current_worker.signals.finished.connect(self.on_worker_finished)
        self.current_worker.signals.error.connect(self.on_worker_error)
        self.current_worker.signals.non_cancellable_started.connect(
            self._on_non_cancellable_started
        )
        self.current_worker.signals.non_cancellable_finished.connect(
            self._on_non_cancellable_finished
        )
        if task_name == "add_files":
            self.current_worker.signals.item_ready.connect(self.add_item_to_view)
            self.current_worker.signals.items_ready.connect(self.add_items_to_view)
            self.current_worker.signals.bookmarks_ready.connect(
                self._load_bookmarks_from_pdf
            )
        self.threadpool.start(self.current_worker)

    def _check_duplicate_files(self, file_paths):
        """既に追加されているファイルをチェックし、重複を除外する"""
        existing_paths = set()
        for i in range(self.page_list_widget.count()):
            item = self.page_list_widget.item(i)
            item_data = item.data(Qt.UserRole)
            if item_data and "original_path" in item_data:
                existing_paths.add(os.path.abspath(item_data["original_path"]))

        new_paths = []
        duplicate_paths = []
        for path in file_paths:
            abs_path = os.path.abspath(path)
            if abs_path in existing_paths:
                duplicate_paths.append(os.path.basename(path))
            else:
                new_paths.append(path)

        if duplicate_paths:
            msg = "以下のファイルは既に追加されています:\n\n" + "\n".join(
                duplicate_paths
            )
            if new_paths:
                msg += f"\n\n{len(new_paths)}個の新しいファイルを追加します。"
                reply = QMessageBox.question(
                    self,
                    "重複ファイル",
                    msg,
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                )
                if reply == QMessageBox.StandardButton.No:
                    return []
            else:
                QMessageBox.information(self, "重複ファイル", msg)
                return []

        return new_paths

    def _add_files(self):
        file_dialog_filter = "対応ファイル (*.pdf *.docx *.doc *.docm *.xlsx *.xls *.xlsm *.pptx *.ppt *.pptm);;All Files (*)"
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "ファイルを選択", self.last_used_path, file_dialog_filter
        )

        if file_paths:
            # 重複チェック
            file_paths = self._check_duplicate_files(file_paths)
        if file_paths:
            self.last_used_path = os.path.dirname(file_paths[0])
            self._run_task("add_files", "ファイルを追加", file_paths=file_paths)

    def _handle_dropped_files(self, file_paths):
        """ドロップされたファイルを処理する"""
        if file_paths:
            # 重複チェック
            file_paths = self._check_duplicate_files(file_paths)
            if file_paths:
                # 最初のファイルのディレクトリをlast_used_pathに設定
                self.last_used_path = os.path.dirname(file_paths[0])
                self._run_task("add_files", "ファイルを追加", file_paths=file_paths)

    def _flush_ipc_pending_files(self):
        """IPC経由で一時バッファしたファイルを自然順でまとめて追加する"""
        if not self._ipc_pending_file_paths:
            return

        # バッファを取り出してクリア
        pending = self._ipc_pending_file_paths
        self._ipc_pending_file_paths = []

        now_str = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        _debug_log(
            f"[DEBUG] _flush_ipc_pending_files: start {now_str}, buffered={pending}"
        )

        # Windows 標準の自然順（ファイル名順）でソート
        def _natural_key(p):
            name = os.path.basename(p)
            return [
                int(t) if t.isdigit() else t.lower()
                for t in re.split(r"([0-9]+)", name)
            ]

        pending.sort(key=_natural_key)

        _debug_log(f"[DEBUG] _flush_ipc_pending_files: buffered={pending}")

        # 重複チェック
        pending = self._check_duplicate_files(pending)
        if not pending:
            return

        _debug_log(f"[DEBUG] _flush_ipc_pending_files: deduped={pending}")

        # 既存の add_files ワーカーが実行中かチェック
        active_count = self.threadpool.activeThreadCount()
        is_add_files_running = (
            self.current_worker is not None
            and getattr(self.current_worker, "task_name", None) == "add_files"
            and getattr(self.current_worker, "is_running", False)
        )
        _debug_log(
            "[DEBUG] _flush_ipc_pending_files: "
            f"activeThreadCount={active_count}, "
            f"is_add_files_running={is_add_files_running}"
        )

        if is_add_files_running:
            _debug_log(
                f"[DEBUG] _flush_ipc_pending_files: add_files 実行中のためキューに追加: {pending}"
            )
            self._pending_file_paths.extend(pending)
            return

        # 最初のファイルのディレクトリをlast_used_pathに設定
        self.last_used_path = os.path.dirname(pending[0])
        self._run_task("add_files", "ファイルを追加", file_paths=pending)

    def _add_files_from_paths(self, file_paths):
        """コマンドライン引数からファイルパスを受け取って追加する"""
        _debug_log(f"[DEBUG] _add_files_from_paths: 呼び出し file_paths={file_paths}")
        if file_paths:
            # ファイルパスをフィルタリング（存在するファイルのみ）
            valid_paths = []
            for path in file_paths:
                if os.path.isfile(path):
                    ext = os.path.splitext(path)[1].lower()
                    if ext in ALL_SUPPORTED_EXTENSIONS:
                        valid_paths.append(path)

            if valid_paths:
                # IPC経由は短時間バッファに入れてから自然順でまとめて処理する
                now_str = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                _debug_log(
                    f"[DEBUG] _add_files_from_paths: buffer extend at {now_str}, paths={valid_paths}"
                )
                self._ipc_pending_file_paths.extend(valid_paths)
                # 最後の到着から500ms後にまとめて処理（到着ごとにリスタート）
                self._ipc_batch_timer.stop()
                self._ipc_batch_timer.start(500)

    def add_item_to_view(self, item_data):
        item = QListWidgetItem()
        item.setData(Qt.UserRole, item_data)
        base_filename = os.path.basename(item_data["original_path"])

        # テキストを設定
        self._update_item_text(item, item_data, base_filename)

        item.setIcon(self._create_thumbnail(item_data))
        self.page_list_widget.addItem(item)

    def _add_pdf_items_sequential(self, pdf_items):
        """通常処理（比較用）"""
        for item_data in pdf_items:
            self.add_item_to_view(item_data)

    def add_items_to_view(self, items_data):
        """複数のアイテムをバッチで追加（PDFの複数ページ用）"""
        # PDFのサムネイル生成を並列化
        pdf_items = [item for item in items_data if item.get("type") == "pdf"]
        non_pdf_items = [item for item in items_data if item.get("type") != "pdf"]

        # 非PDFアイテムは通常通り処理
        for item_data in non_pdf_items:
            self.add_item_to_view(item_data)

        # PDFアイテムは通常処理（ファイルごとにグループ化して最適化）
        if pdf_items:

            # ファイルごとにグループ化（ファイルを1回だけ開くため）
            files_dict = {}
            for item_data in pdf_items:
                file_path = item_data["original_path"]
                if file_path not in files_dict:
                    files_dict[file_path] = []
                files_dict[file_path].append(item_data)

            # ファイルごとに処理（ファイルを1回だけ開く）
            for file_path, file_items in files_dict.items():
                try:
                    with fitz.open(file_path) as doc:
                        for item_data in file_items:
                            # サムネイルを事前に生成してキャッシュに保存
                            key = self._thumbnail_cache_key(item_data)
                            if key not in self.thumbnail_cache:
                                try:
                                    page = doc.load_page(item_data["page_num"])
                                    rotation = item_data.get("rotation", 0)
                                    matrix = fitz.Matrix(1, 1).prerotate(rotation)

                                    # スケーリング計算
                                    rect = page.rect
                                    if rotation in (90, 270):
                                        page_width, page_height = (
                                            rect.height,
                                            rect.width,
                                        )
                                    else:
                                        page_width, page_height = (
                                            rect.width,
                                            rect.height,
                                        )

                                    max_width = self.max_thumbnail_size.width()
                                    max_height = self.max_thumbnail_size.height()
                                    scale_x = max_width / page_width
                                    scale_y = max_height / page_height
                                    scale = min(scale_x, scale_y)

                                    matrix = fitz.Matrix(scale, scale).prerotate(
                                        rotation
                                    )
                                    pix = page.get_pixmap(matrix=matrix, alpha=False)

                                    img = QImage(
                                        pix.samples,
                                        pix.width,
                                        pix.height,
                                        pix.stride,
                                        QImage.Format_RGB888,
                                    )
                                    base_pixmap = QPixmap.fromImage(img)
                                    self.thumbnail_cache[key] = base_pixmap
                                except Exception as e:
                                    print(f"サムネイル生成エラー: {e}")

                            # UIに追加
                            self.add_item_to_view(item_data)
                except Exception as e:
                    print(f"ファイルオープンエラー: {e}")
                    # エラー時は通常通り処理
                    for item_data in file_items:
                        self.add_item_to_view(item_data)

    def _add_pdf_items_parallel(self, pdf_items):
        """PDFアイテムのサムネイル生成を並列化"""
        if not pdf_items:
            return

        # 並列処理の準備
        max_size_tuple = (
            self.max_thumbnail_size.width(),
            self.max_thumbnail_size.height(),
        )
        num_processes = min(cpu_count(), len(pdf_items), 8)  # 最大8プロセス（並列処理）
        # ワーカー関数に渡す引数を準備
        args_list = [(item_data, max_size_tuple) for item_data in pdf_items]

        # 並列処理でサムネイルを生成
        results = []
        try:
            with Pool(processes=num_processes) as pool:
                results = pool.map(self._build_base_thumbnail_worker, args_list)
        except Exception:
            import traceback

            traceback.print_exc()
            # エラー時は通常処理にフォールバック
            for item_data in pdf_items:
                self.add_item_to_view(item_data)
            return

        # 結果を処理してUIに追加（メインスレッドでQPixmapを作成）
        for result in results:
            if result is None:
                continue

            item_data = result["item_data"]
            item = QListWidgetItem()
            item.setData(Qt.UserRole, item_data)
            base_filename = os.path.basename(item_data["original_path"])

            # テキストを設定
            self._update_item_text(item, item_data, base_filename)

            # サムネイルを生成
            if result.get("success"):
                try:
                    # 画像データからQPixmapを作成（メインスレッドで実行）
                    img = QImage(
                        result["image_data"],
                        result["width"],
                        result["height"],
                        result["stride"],
                        QImage.Format_RGB888,
                    )
                    max_size = QSize(result["max_size"][0], result["max_size"][1])
                    base_pixmap = QPixmap.fromImage(
                        img.scaled(
                            max_size, Qt.KeepAspectRatio, Qt.SmoothTransformation
                        )
                    )

                    # キャッシュに保存
                    key = self._thumbnail_cache_key(item_data)
                    self.thumbnail_cache[key] = base_pixmap

                    # 最終的なサムネイルを作成
                    thumbnail = self._create_thumbnail(item_data)
                    item.setIcon(thumbnail)
                except Exception as e:
                    print(f"サムネイル作成エラー: {e}")
                    # エラー時はダミーサムネイル
                    item.setIcon(
                        self._create_dummy_thumbnail(
                            f"Error\nP{item_data.get('page_num', 0)+1}",
                            "#e74c3c",
                            QSize(THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT),
                        )
                    )
            else:
                # エラー時はダミーサムネイル
                item.setIcon(
                    self._create_dummy_thumbnail(
                        f"PDF Error\nP{item_data.get('page_num', 0)+1}",
                        "#e74c3c",
                        QSize(THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT),
                    )
                )

        self.page_list_widget.addItem(item)

    def _add_pdf_items_parallel_by_file(self, pdf_items):
        """PDFアイテムのサムネイル生成をファイルごとに並列化"""
        if not pdf_items:
            return

        # ファイルごとにグループ化
        files_dict = {}
        for item_data in pdf_items:
            file_path = item_data["original_path"]
            if file_path not in files_dict:
                files_dict[file_path] = []
            files_dict[file_path].append(item_data)

        # 並列処理の準備
        max_size_tuple = (
            self.max_thumbnail_size.width(),
            self.max_thumbnail_size.height(),
        )
        # プロセス数の決定:
        # - CPUコア数以下（コア数と同じが最適）
        # - ファイル数以下（ファイル数より多くしても意味がない）
        # - 最大8プロセス（メモリ使用量とオーバーヘッドを考慮）
        num_processes = min(cpu_count(), len(files_dict), 8)  # 最大8プロセス
        # ワーカー関数に渡す引数を準備（ファイルごと）
        args_list = [
            (file_path, file_items, max_size_tuple)
            for file_path, file_items in files_dict.items()
        ]

        # 並列処理でサムネイルを生成（ファイルごと）
        results = []
        try:
            with Pool(processes=num_processes) as pool:
                results = pool.map(
                    self._build_base_thumbnails_for_file_worker, args_list
                )
        except Exception:
            import traceback

            traceback.print_exc()
            # エラー時は通常処理にフォールバック
            for item_data in pdf_items:
                self.add_item_to_view(item_data)
            return

        # 結果を処理してUIに追加（メインスレッドでQPixmapを作成）
        for file_results in results:
            if file_results is None:
                continue

            for result in file_results:
                if result is None:
                    continue

                # タプル形式を展開
                # 成功時: (True, original_path, page_num, rotation, type, image_data, width, height, stride)
                # 失敗時: (False, original_path, page_num, rotation, type, error_message)
                success = result[0]
                original_path = result[1]
                page_num = result[2]
                rotation = result[3]
                item_type = result[4]

                # item_dataを再構築（QListWidgetItemに保存するため）
                item_data = {
                    "type": item_type,
                    "path": original_path,  # 一時ファイルのパス（存在しない場合はoriginal_pathを使用）
                    "original_path": original_path,
                    "page_num": page_num,
                    "rotation": rotation,
                }

                item = QListWidgetItem()
                item.setData(Qt.UserRole, item_data)
                base_filename = os.path.basename(original_path)

                # テキストを設定
                self._update_item_text(item, item_data, base_filename)

                # サムネイルを生成
                if success:
                    try:
                        # タプルから画像データを取得
                        image_data = result[5]
                        width = result[6]
                        height = result[7]
                        stride = result[8]

                        # 画像データからQPixmapを作成（メインスレッドで実行）
                        img = QImage(
                            image_data,
                            width,
                            height,
                            stride,
                            QImage.Format_RGB888,
                        )
                        # ワーカー側で既にスケーリング済みなので、そのままQPixmapに変換
                        base_pixmap = QPixmap.fromImage(img)

                        # キャッシュに保存
                        key = self._thumbnail_cache_key(item_data)
                        self.thumbnail_cache[key] = base_pixmap

                        # 最終的なサムネイルを作成
                        thumbnail = self._create_thumbnail(item_data)
                        item.setIcon(thumbnail)
                    except Exception as e:
                        print(f"サムネイル作成エラー: {e}")
                        # エラー時はダミーサムネイル
                        item.setIcon(
                            self._create_dummy_thumbnail(
                                f"Error\nP{page_num+1}",
                                "#e74c3c",
                                QSize(THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT),
                            )
                        )
                else:
                    # エラー時はダミーサムネイル
                    item.setIcon(
                        self._create_dummy_thumbnail(
                            f"PDF Error\nP{page_num+1}",
                            "#e74c3c",
                            QSize(THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT),
                        )
                    )

                self.page_list_widget.addItem(item)

    @staticmethod
    def _build_base_thumbnails_for_file_worker(args):
        """マルチプロセッシング用のワーカー関数（ファイルごとに全ページを処理）"""
        file_path, file_items, max_size_tuple = args

        results = []
        try:
            # ファイルを1回だけ開く
            with fitz.open(file_path) as doc:
                for item_data in file_items:
                    page_num = item_data.get("page_num", 0)
                    rotation = item_data.get("rotation", 0)
                    item_type = item_data.get("type", "pdf")
                    original_path = item_data.get("original_path", file_path)

                    try:
                        page = doc.load_page(page_num)
                        # 回転を考慮したマトリックス
                        matrix = fitz.Matrix(1, 1).prerotate(rotation)

                        # ページの元のサイズを取得
                        rect = page.rect
                        # 回転後のサイズを計算
                        if rotation in (90, 270):
                            page_width, page_height = rect.height, rect.width
                        else:
                            page_width, page_height = rect.width, rect.height

                        # max_sizeに合わせてスケールを計算
                        max_width, max_height = max_size_tuple
                        scale_x = max_width / page_width
                        scale_y = max_height / page_height
                        scale = min(scale_x, scale_y)  # アスペクト比を保つ

                        # スケーリングをマトリックスに適用
                        matrix = fitz.Matrix(scale, scale).prerotate(rotation)

                        pix = page.get_pixmap(
                            matrix=matrix,
                            alpha=False,
                        )
                        # タプル形式で返す（pickleのオーバーヘッド削減）
                        # (success, original_path, page_num, rotation, type, image_data, width, height, stride)
                        results.append(
                            (
                                True,
                                original_path,
                                page_num,
                                rotation,
                                item_type,
                                bytes(pix.samples),
                                pix.width,
                                pix.height,
                                pix.stride,
                            )
                        )
                    except Exception as e:
                        # エラー時もタプル形式
                        # (success, original_path, page_num, rotation, type, error_message)
                        results.append(
                            (
                                False,
                                original_path,
                                page_num,
                                rotation,
                                item_type,
                                str(e),
                            )
                        )
        except Exception as e:
            # ファイルを開けない場合、すべてのページをエラーとして返す
            for item_data in file_items:
                results.append(
                    (
                        False,
                        item_data.get("original_path", file_path),
                        item_data.get("page_num", 0),
                        item_data.get("rotation", 0),
                        item_data.get("type", "pdf"),
                        str(e),
                    )
                )
        return results

    def _create_thumbnail(self, item_data, size=None):
        if size is None:
            # 現在のズームレベルに基づいたサイズを計算
            new_width = int(THUMBNAIL_WIDTH * self.zoom_level)
            new_height = int(THUMBNAIL_HEIGHT * self.zoom_level)
            size = QSize(new_width, new_height)

        base_pixmap = self._get_or_create_base_thumbnail(item_data)
        if base_pixmap is None:
            # フォールバックとしてダミーサムネイルを返す
            return self._create_dummy_thumbnail("[No Image]", "#7f8c8d", size)

        if base_pixmap.size() == size:
            return base_pixmap

        # アスペクト比を保持して縮小
        scaled_pixmap = base_pixmap.scaled(
            size, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )

        # 縦長を基準に固定サイズのキャンバスを作成（回転したページでも同じサイズを保証）
        result_pixmap = QPixmap(size.width(), size.height())
        result_pixmap.fill(Qt.GlobalColor.transparent)

        # 縮小したサムネイルを中央配置
        painter = QPainter(result_pixmap)
        x_offset = (size.width() - scaled_pixmap.width()) // 2
        y_offset = (size.height() - scaled_pixmap.height()) // 2
        painter.drawPixmap(x_offset, y_offset, scaled_pixmap)
        painter.end()

        return result_pixmap

    def _get_or_create_base_thumbnail(self, item_data):
        """最大サイズのサムネイルをキャッシュし、そこから縮小して使い回す"""
        key = self._thumbnail_cache_key(item_data)
        cached = self.thumbnail_cache.get(key)
        if cached:
            return cached

        pixmap = self._build_base_thumbnail(item_data)
        if pixmap:
            self.thumbnail_cache[key] = pixmap
        return pixmap

    @staticmethod
    def _build_base_thumbnail_worker(args):
        """マルチプロセッシング用のワーカー関数（PyMuPDFの処理のみ）"""
        item_data, max_size_tuple = args

        if item_data["type"] == "pdf":
            try:
                with fitz.open(item_data["path"]) as doc:
                    page = doc.load_page(item_data["page_num"])
                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(1, 1).prerotate(item_data["rotation"]),
                        alpha=False,
                    )
                    # 画像データをバイト配列として返す（QPixmapはメインスレッドで作成）
                    return {
                        "success": True,
                        "item_data": item_data,
                        "image_data": bytes(pix.samples),
                        "width": pix.width,
                        "height": pix.height,
                        "stride": pix.stride,
                        "max_size": max_size_tuple,
                    }
            except Exception as e:
                return {
                    "success": False,
                    "item_data": item_data,
                    "error": str(e),
                    "max_size": max_size_tuple,
                }
        return None

    def _build_base_thumbnail(self, item_data):
        size = self.max_thumbnail_size
        if item_data["type"] == "pdf":
            try:
                with fitz.open(item_data["path"]) as doc:
                    page = doc.load_page(item_data["page_num"])
                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(1, 1).prerotate(item_data["rotation"]),
                        alpha=False,
                    )
                    img = QImage(
                        pix.samples,
                        pix.width,
                        pix.height,
                        pix.stride,
                        QImage.Format_RGB888,
                    )
                    return QPixmap.fromImage(
                        img.scaled(size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    )
            except Exception:
                return self._create_dummy_thumbnail(
                    f"PDF Error\nP{item_data.get('page_num', 0)+1}", "#e74c3c", size
                )
        elif item_data["type"] == "word":
            return self._create_file_type_thumbnail(
                "W", "Word", "#185ABD", size, "document"
            )
        elif item_data["type"] == "excel":
            return self._create_file_type_thumbnail(
                "X", "Excel", "#107C41", size, "spreadsheet"
            )
        elif item_data["type"] == "powerpoint":
            return self._create_file_type_thumbnail(
                "P", "PowerPoint", "#C43E1C", size, "presentation"
            )
        else:
            return self._create_file_type_thumbnail("F", "File", "#6B7280", size)

    def _thumbnail_cache_key(self, item_data):
        return (
            item_data.get("type"),
            item_data.get("path"),
            item_data.get("original_path"),
            item_data.get("page_num"),
            item_data.get("rotation", 0),
        )

    def _invalidate_thumbnail_cache(self, item_data):
        key = self._thumbnail_cache_key(item_data)
        if key in self.thumbnail_cache:
            del self.thumbnail_cache[key]

    def _create_dummy_thumbnail(self, text, color, size):
        img = Image.new("RGBA", (size.width(), size.height()), (74, 82, 96, 255))
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("arial.ttf", 30)
        except IOError:
            font = ImageFont.load_default()
        text_bbox = draw.textbbox((0, 0), text, font=font)
        draw.text(
            (
                (size.width() - (text_bbox[2] - text_bbox[0])) / 2,
                (size.height() - (text_bbox[3] - text_bbox[1])) / 2,
            ),
            text,
            fill=color,
            font=font,
        )
        return QPixmap.fromImage(
            QImage(img.tobytes(), img.width, img.height, QImage.Format_RGBA8888)
        )

    def _create_file_type_thumbnail(
        self, badge_text, label, accent_color, size, content_type="document"
    ):
        """Office系ファイル用の自前サムネイルを描画する。"""
        width = size.width()
        height = size.height()
        img = Image.new("RGBA", (width, height), (45, 52, 64, 255))
        draw = ImageDraw.Draw(img)

        try:
            badge_font = ImageFont.truetype("arialbd.ttf", max(34, width // 3))
        except IOError:
            badge_font = ImageFont.load_default()

        margin = max(4, width // 32)
        card_left = margin
        card_top = margin
        card_right = width - margin
        card_bottom = height - margin
        accent = accent_color
        try:
            accent_rgb = tuple(
                int(accent_color.lstrip("#")[i : i + 2], 16) for i in (0, 2, 4)
            )
        except Exception:
            accent_rgb = (107, 114, 128)
        accent_soft = tuple(int(248 * 0.82 + value * 0.18) for value in accent_rgb)
        accent_line = accent_rgb

        shadow_offset = max(3, width // 36)
        draw.rounded_rectangle(
            (
                card_left + shadow_offset,
                card_top + shadow_offset,
                card_right + shadow_offset // 2,
                card_bottom + shadow_offset,
            ),
            radius=max(5, width // 24),
            fill=(15, 23, 42, 60),
        )
        draw.rounded_rectangle(
            (card_left, card_top, card_right, card_bottom),
            radius=max(5, width // 24),
            fill=(*accent_soft, 255),
            outline=(*accent_line, 255),
            width=1,
        )
        content_left = card_left + max(12, width // 12)
        content_top = card_top + max(18, height // 9)
        content_right = card_right - max(10, width // 14)
        content_bottom = card_bottom - max(14, height // 10)
        if content_type == "spreadsheet":
            self._draw_spreadsheet_thumbnail_content(
                draw,
                content_left,
                content_top,
                content_right,
                content_bottom,
                accent,
            )
        elif content_type == "presentation":
            self._draw_presentation_thumbnail_content(
                draw,
                content_left,
                content_top,
                content_right,
                content_bottom,
                accent,
            )
        else:
            self._draw_document_thumbnail_content(
                draw,
                content_left,
                content_top,
                content_right,
                content_bottom,
                accent_line,
            )

        icon_size = min(width, height) // 3 + min(width, height) // 10
        icon_left = card_left + max(6, width // 26)
        icon_top = card_top + max(6, height // 30)
        draw.rounded_rectangle(
            (
                icon_left + shadow_offset,
                icon_top + shadow_offset,
                icon_left + icon_size + shadow_offset,
                icon_top + icon_size + shadow_offset,
            ),
            radius=max(5, icon_size // 10),
            fill=(15, 23, 42, 70),
        )
        draw.rounded_rectangle(
            (
                icon_left,
                icon_top,
                icon_left + icon_size,
                icon_top + icon_size,
            ),
            radius=max(5, icon_size // 10),
            fill=accent,
        )
        badge_bbox = draw.textbbox((0, 0), badge_text, font=badge_font)
        badge_width = badge_bbox[2] - badge_bbox[0]
        badge_height = badge_bbox[3] - badge_bbox[1]
        draw.text(
            (
                icon_left + (icon_size - badge_width) / 2 - badge_bbox[0],
                icon_top + (icon_size - badge_height) / 2 - badge_bbox[1] - 1,
            ),
            badge_text,
            fill=(255, 255, 255, 255),
            font=badge_font,
        )

        return QPixmap.fromImage(
            QImage(img.tobytes(), img.width, img.height, QImage.Format_RGBA8888)
        )

    def _draw_document_thumbnail_content(self, draw, left, top, right, bottom, accent):
        line_color = (*accent, 255)
        line_height = max(3, (bottom - top) // 18)
        line_gap = max(8, (bottom - top) // 7)
        y = top
        while y + line_height <= bottom:
            draw.rounded_rectangle(
                (left, y, right, y + line_height),
                radius=2,
                fill=line_color,
            )
            y += line_gap

    def _draw_spreadsheet_thumbnail_content(self, draw, left, top, right, bottom, accent):
        try:
            accent_rgb = tuple(
                int(accent.lstrip("#")[i : i + 2], 16) for i in (0, 2, 4)
            )
        except Exception:
            accent_rgb = (107, 114, 128)
        accent_line = tuple(int(203 * 0.25 + value * 0.75) for value in accent_rgb)
        accent_fill = tuple(int(248 * 0.9 + value * 0.1) for value in accent_rgb)
        rows = 5
        cols = 4
        cell_width = max(8, (right - left) // cols)
        cell_height = max(7, (bottom - top) // rows)
        table_right = left + cell_width * cols
        table_bottom = top + cell_height * rows

        draw.rounded_rectangle(
            (left, top, table_right, table_bottom),
            radius=4,
            fill=(*accent_fill, 255),
            outline=(*accent_line, 255),
            width=1,
        )
        for col in range(1, cols):
            x = left + col * cell_width
            draw.line((x, top, x, table_bottom), fill=(*accent_line, 255), width=2)
        for row in range(1, rows):
            y = top + row * cell_height
            draw.line((left, y, table_right, y), fill=(*accent_line, 255), width=2)

    def _draw_presentation_thumbnail_content(
        self, draw, left, top, right, bottom, accent
    ):
        diameter = min(right - left, bottom - top)
        if diameter <= 0:
            return
        cx = left + (right - left - diameter) // 2
        cy = top + (bottom - top - diameter) // 2
        chart_box = (cx, cy, cx + diameter, cy + diameter)
        draw.ellipse(chart_box, fill=(226, 91, 55, 255))
        draw.pieslice(chart_box, start=270, end=45, fill=accent)
        draw.pieslice(chart_box, start=45, end=120, fill=(190, 57, 31, 255))
        draw.ellipse(chart_box, outline=accent, width=2)

    def _update_page_mode_actions_state(self):
        """
        ページモードの全てのアクションの状態を、リストの状態に応じて一元管理します。
        """
        total_count = self.page_list_widget.count()
        selected_count = len(self.page_list_widget.selectedItems())

        # --- Step 1: アイテムが1つも無い場合 ---
        if total_count == 0:
            self.delete_action.setEnabled(False)
            self.rot_left_action.setEnabled(False)
            if hasattr(self, "export_selected_pdf_action"):
                self.export_selected_pdf_action.setEnabled(False)
            if hasattr(self, "export_selected_images_action"):
                self.export_selected_images_action.setEnabled(False)
            self.rot_right_action.setEnabled(False)
            self.move_to_top_action.setEnabled(False)
            self.move_up_action.setEnabled(False)
            self.move_down_action.setEnabled(False)
            self.move_to_bottom_action.setEnabled(False)
            self.merge_action.setEnabled(False)
            return  # これで処理終了

        # --- Step 2: アイテムは有るが、何も選択されていない場合 ---
        if selected_count == 0:
            # 保存は可能
            self.merge_action.setEnabled(True)
            # 選択が必要なアクションはすべて無効
            self.delete_action.setEnabled(False)
            self.rot_left_action.setEnabled(False)
            self.rot_right_action.setEnabled(False)
            self.move_to_top_action.setEnabled(False)
            self.move_up_action.setEnabled(False)
            self.move_down_action.setEnabled(False)
            self.move_to_bottom_action.setEnabled(False)
            if hasattr(self, "export_selected_pdf_action"):
                self.export_selected_pdf_action.setEnabled(False)
            if hasattr(self, "export_selected_images_action"):
                self.export_selected_images_action.setEnabled(False)
            return  # これで処理終了

        # --- Step 3: アイテムが選択されている場合 ---
        # 選択されていれば常に有効なアクション
        self.delete_action.setEnabled(True)
        self.rot_left_action.setEnabled(True)
        self.rot_right_action.setEnabled(True)
        self.merge_action.setEnabled(True)
        if hasattr(self, "export_selected_pdf_action"):
            self.export_selected_pdf_action.setEnabled(True)
        if hasattr(self, "export_selected_images_action"):
            self.export_selected_images_action.setEnabled(True)

        # 移動アクションの状態判定
        selected_rows = [
            self.page_list_widget.row(item)
            for item in self.page_list_widget.selectedItems()
        ]
        min_row = min(selected_rows)
        max_row = max(selected_rows)

        # 上方向への移動
        can_move_up = min_row > 0
        self.move_to_top_action.setEnabled(can_move_up)
        self.move_up_action.setEnabled(can_move_up)

        # 下方向への移動
        can_move_down = max_row < total_count - 1
        self.move_to_bottom_action.setEnabled(can_move_down)
        self.move_down_action.setEnabled(can_move_down)

    def _reset_page_settings(self):
        """ページ番号とヘッダー・フッターの設定をリセット"""
        # ページ番号設定をリセット
        self.page_number_settings = {
            "enabled": False,
            "is_header": False,
            "alignment": "center",
            "format": "1",
            "start_number": 1,
            "font_size": 10,
        }
        # ヘッダー・フッター設定をリセット
        self.header_footer_settings = None

        # チェックボックスの状態を更新

    def _new_project(self):
        """新しいプロジェクトを開始（リストをクリア）"""
        if self.page_list_widget.count() == 0:
            return
        reply = QMessageBox.question(
            self,
            "新規プロジェクト",
            "現在のプロジェクトをクリアして新規プロジェクトを開始しますか？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.page_list_widget.clear()
            self.thumbnail_cache.clear()
            self.bookmarks.clear()
            # 設定をリセット（新しいプロジェクトを開始するため）
            self._reset_page_settings()
            if hasattr(self, "bookmark_dock") and self.bookmark_dock.isVisible():
                self._update_bookmark_tree()
            self.update_status_bar()
            self._update_page_mode_actions_state()
            self._record_history_change()

    def _select_all(self):
        """現在のリストのすべてのアイテムを選択"""
        self.page_list_widget.selectAll()

    def _delete_selected(self):
        if not self.page_list_widget.selectedItems():
            return
        if (
            QMessageBox.question(
                self,
                "確認",
                f"{len(self.page_list_widget.selectedItems())}個のアイテムを削除しますか？",
            )
            == QMessageBox.StandardButton.Yes
        ):
            for item in sorted(
                self.page_list_widget.selectedItems(),
                key=lambda i: self.page_list_widget.row(i),
                reverse=True,
            ):
                item_data = item.data(Qt.UserRole)
                if item_data:
                    self._invalidate_thumbnail_cache(item_data)
                self.page_list_widget.takeItem(self.page_list_widget.row(item))
            self._generate_bookmarks_from_list()
            self.update_status_bar()
            self._update_page_mode_actions_state()
            self._record_history_change()

    def _rotate_selected(self, angle):
        """
        選択されたPDFページのサムネイルを指定された角度(angle)で回転させる
        """
        changed = False
        for item in self.page_list_widget.selectedItems():
            data = item.data(Qt.UserRole)
            # PDFページアイテムのみを対象とする
            if data["type"] == "pdf":
                # 現在の回転角度に新しい角度を加算し、360で割った余りを新しい角度とする
                # (例: 270 + 90 = 360 -> 0)
                data["rotation"] = (data["rotation"] + angle) % 360
                self._invalidate_thumbnail_cache(data)
                item.setData(Qt.UserRole, data)
                # 新しい回転角度でサムネイルを再生成
                item.setIcon(self._create_thumbnail(data))
                changed = True
        if changed:
            self._record_history_change()

    def _open_source_file(self, item: QListWidgetItem):
        try:
            item_data = item.data(Qt.UserRole)
            if not item_data or "original_path" not in item_data:
                return
            original_path = item_data["original_path"]
            if os.path.exists(original_path):
                os.startfile(original_path)
            else:
                QMessageBox.warning(
                    self,
                    "ファイルエラー",
                    f"ファイルが見つかりません:\n{original_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"ファイルを開けませんでした:\n{e}")

    def _open_user_manual(self):
        """配布物に含まれるユーザーマニュアルHTMLを既定ブラウザで開く"""
        if os.path.exists(self.user_manual_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.user_manual_path))
        else:
            QMessageBox.warning(
                self,
                "ユーザーマニュアル",
                f"ユーザーマニュアルが見つかりませんでした。\n{self.user_manual_path}",
            )

    def _move_items(self, direction):
        # --- ステップ1: 選択されたアイテムを行番号順で取得 ---
        selected_rows = sorted(
            [
                self.page_list_widget.row(item)
                for item in self.page_list_widget.selectedItems()
            ]
        )
        if not selected_rows:
            return

        # --- ステップ2: 境界チェック ---
        if direction == -1 and selected_rows[0] == 0:
            return  # 先頭にいるのに「上へ」は押せない
        if direction == 1 and selected_rows[-1] == self.page_list_widget.count() - 1:
            return  # 末尾にいるのに「下へ」は押せない

        # --- ステップ3: 逆順スワップ法の実行 ---
        if direction == -1:  # 「上へ」移動の場合
            # ループを前から (昇順で) 処理する
            for row in selected_rows:
                # アイテムを一つ上のアイテムと入れ替える
                # takeItemで取り出し、insertItemで一つ上の位置に挿入する
                item = self.page_list_widget.takeItem(row)
                self.page_list_widget.insertItem(row - 1, item)
        else:  # 「下へ」移動の場合
            # ループを後ろから (降順で) 処理する
            for row in reversed(selected_rows):
                # アイテムを一つ下のアイテムと入れ替える
                item = self.page_list_widget.takeItem(row)
                self.page_list_widget.insertItem(row + 1, item)

        # --- ステップ4: 移動後のアイテムを選択状態に戻す ---
        # QListWidgetの仕様上、takeItem/insertItemを使うと選択が解除されるため
        for row in selected_rows:
            # 移動後の新しい行番号 (row + direction) にあるアイテムを選択する
            new_row = row + direction
            if 0 <= new_row < self.page_list_widget.count():
                self.page_list_widget.item(new_row).setSelected(True)

        # --- ステップ5: 選択範囲の先頭が見えるようにスクロール ---
        first_item_new_row = selected_rows[0] + direction
        if 0 <= first_item_new_row < self.page_list_widget.count():
            first_item = self.page_list_widget.item(first_item_new_row)
            if first_item:
                self.page_list_widget.scrollToItem(
                    first_item, QAbstractItemView.ScrollHint.EnsureVisible
                )

        # --- ステップ6: グリッドサイズを再設定（回転したページがある場合の描画領域の縮みを防ぐ） ---
        self._update_grid_size()
        # --- ステップ7: しおりの対象行がズレるので、再生成 ---
        self._generate_bookmarks_from_list()
        self._update_page_mode_actions_state()
        self._record_history_change()

    def _on_items_moved(self):
        """ドラッグ&ドロップでアイテムが移動されたときに呼ばれる"""
        # グリッドサイズを再設定
        self._update_grid_size()
        # しおりの対象行がズレるので、再生成
        self._generate_bookmarks_from_list()
        # アクションの状態を更新
        self._update_page_mode_actions_state()
        # 履歴に記録
        self._record_history_change()

    def _move_up(self):
        self._move_items(-1)

    def _move_down(self):
        self._move_items(1)

    def _move_to_top(self):
        """選択されているアイテムをリストの一番上に移動させます。"""
        selected_items = self.page_list_widget.selectedItems()
        if not selected_items:
            return

        # 選択されたアイテムを行番号が降順になるようにソート
        # (リストから安全に削除するため)
        selected_rows = sorted(
            [self.page_list_widget.row(item) for item in selected_items], reverse=True
        )

        # 既に先頭にある場合は何もしない
        if selected_rows[-1] == 0:  # 降順なので一番小さい行番号は末尾
            return

        # 降順でアイテムをリストから取り出す
        items_to_move = [self.page_list_widget.takeItem(row) for row in selected_rows]

        # 取り出した順とは逆（元のリストでの昇順）で、リストの先頭に挿入していく
        insert_pos = 0
        for item in reversed(items_to_move):
            self.page_list_widget.insertItem(insert_pos, item)
            item.setSelected(True)
            insert_pos += 1

        # 先頭にスクロール
        self.page_list_widget.scrollToItem(
            self.page_list_widget.item(0), QAbstractItemView.ScrollHint.EnsureVisible
        )
        # グリッドサイズを再設定（回転したページがある場合の描画領域の縮みを防ぐ）
        self._update_grid_size()
        self._generate_bookmarks_from_list()
        self._update_page_mode_actions_state()
        self._record_history_change()

    def _move_to_bottom(self):
        """選択されているアイテムをリストの一番下に移動させます。"""
        selected_items = self.page_list_widget.selectedItems()
        if not selected_items:
            return
        selected_rows = sorted(
            [self.page_list_widget.row(item) for item in selected_items], reverse=True
        )
        if selected_rows[0] == self.page_list_widget.count() - 1:
            return
        items_to_move = [self.page_list_widget.takeItem(row) for row in selected_rows]
        new_last_index = self.page_list_widget.count()
        for item in reversed(items_to_move):
            self.page_list_widget.addItem(item)
            item.setSelected(True)
        self.page_list_widget.scrollToItem(
            self.page_list_widget.item(new_last_index),
            QAbstractItemView.ScrollHint.EnsureVisible,
        )
        # グリッドサイズを再設定（回転したページがある場合の描画領域の縮みを防ぐ）
        self._update_grid_size()
        self._generate_bookmarks_from_list()
        self._update_page_mode_actions_state()
        self._record_history_change()

    def _create_state_snapshot(self):
        """現在のページリスト状態をスナップショットとして取得"""
        page_items = []
        for i in range(self.page_list_widget.count()):
            item_data = self.page_list_widget.item(i).data(Qt.UserRole)
            if item_data is not None:
                page_items.append(copy.deepcopy(item_data))
        selection_rows = [
            self.page_list_widget.row(item)
            for item in self.page_list_widget.selectedItems()
        ]
        return {
            "page_items": page_items,
            "bookmarks": copy.deepcopy(self.bookmarks),
            "selection_rows": selection_rows,
            "last_used_path": self.last_used_path,
            "auto_bookmarks_enabled": self.auto_bookmarks_enabled,
            "show_bookmarks_on_open": self.show_bookmarks_on_open,
        }

    def _apply_state_snapshot(self, snapshot):
        """スナップショットからページリスト状態を復元"""
        self.page_list_widget.clear()
        for item_data in snapshot.get("page_items", []):
            self.add_item_to_view(copy.deepcopy(item_data))

        self.bookmarks = copy.deepcopy(snapshot.get("bookmarks", []))
        self.last_used_path = snapshot.get("last_used_path", self.last_used_path)
        self.auto_bookmarks_enabled = snapshot.get(
            "auto_bookmarks_enabled", self.auto_bookmarks_enabled
        )
        self.show_bookmarks_on_open = snapshot.get(
            "show_bookmarks_on_open", self.show_bookmarks_on_open
        )
        if hasattr(self, "auto_bookmark_action"):
            self.auto_bookmark_action.setChecked(self.auto_bookmarks_enabled)
        if hasattr(self, "show_outline_on_open_action"):
            self.show_outline_on_open_action.setChecked(self.show_bookmarks_on_open)
        # ページ番号とヘッダー・フッターのチェックボックスの状態を更新

        if hasattr(self, "bookmark_dock") and self.bookmark_dock.isVisible():
            self._update_bookmark_tree()

        self._apply_zoom()
        self.update_status_bar()

        # 選択状態を復元
        self.page_list_widget.clearSelection()
        for row in snapshot.get("selection_rows", []):
            if 0 <= row < self.page_list_widget.count():
                self.page_list_widget.item(row).setSelected(True)
        self._update_page_mode_actions_state()

    def _remove_auto_bookmarks(self):
        """自動生成されたしおりのみを取り除く"""
        before = len(self.bookmarks)
        self.bookmarks = [b for b in self.bookmarks if not b.get("auto", False)]
        return before != len(self.bookmarks)

    def _prune_orphan_bookmarks(self):
        """存在しないページを指しているしおりを削除（パスは正規化して比較）"""
        valid_keys = {
            (
                os.path.abspath(
                    self.page_list_widget.item(row).data(Qt.UserRole)["original_path"]
                ),
                self.page_list_widget.item(row).data(Qt.UserRole).get("page_num", 0),
            )
            for row in range(self.page_list_widget.count())
            if self.page_list_widget.item(row)
            and self.page_list_widget.item(row).data(Qt.UserRole)
        }
        before = len(self.bookmarks)
        self.bookmarks = [
            b
            for b in self.bookmarks
            if (os.path.abspath(b.get("path", "")), b.get("page_num", 0)) in valid_keys
        ]
        return before != len(self.bookmarks)

    def _sort_bookmarks_by_page(self):
        """現在のページ順に合わせてしおりを並び替える"""
        lookup = self._build_bookmark_lookup()
        self.bookmarks.sort(
            key=lambda b: lookup.get(
                (os.path.abspath(b.get("path", "")), b.get("page_num", 0)), float("inf")
            )
        )

    def _build_bookmark_lookup(self):
        """(path, page_num) -> 現在のページインデックス（パスは正規化して比較）"""
        lookup = {}
        for row in range(self.page_list_widget.count()):
            item = self.page_list_widget.item(row)
            if not item:
                continue
            data = item.data(Qt.UserRole)
            if not data:
                continue
            key = (os.path.abspath(data["original_path"]), data.get("page_num", 0))
            if key not in lookup:
                lookup[key] = row
        return lookup

    def _states_equal(self, state_a, state_b):
        return (
            state_a.get("page_items", []) == state_b.get("page_items", [])
            and state_a.get("bookmarks", []) == state_b.get("bookmarks", [])
            and state_a.get("auto_bookmarks_enabled")
            == state_b.get("auto_bookmarks_enabled")
            and state_a.get("show_bookmarks_on_open")
            == state_b.get("show_bookmarks_on_open")
        )

    def _record_history_change(self, initial=False):
        """履歴に現在の状態を記録"""
        snapshot = self._create_state_snapshot()
        if self.undo_stack and self._states_equal(self.undo_stack[-1], snapshot):
            if not initial:
                self.redo_stack.clear()
                self._update_history_actions_state()
            return
        self.undo_stack.append(snapshot)
        if len(self.undo_stack) > MAX_HISTORY:
            self.undo_stack.pop(0)
        if not initial:
            self.redo_stack.clear()
        self._update_history_actions_state()

    def _update_history_actions_state(self):
        """Undo/Redoアクションの有効/無効を更新"""
        can_undo = len(self.undo_stack) > 1
        can_redo = len(self.redo_stack) > 0
        self.undo_action.setEnabled(can_undo)
        self.redo_action.setEnabled(can_redo)

    def _format_bookmark_title(self, item_data, include_page=True):
        """設定に応じたしおりタイトルを生成"""
        if not item_data:
            return "無題"
        base = os.path.splitext(os.path.basename(item_data["original_path"]))[0]
        if include_page and item_data.get("type") == "pdf":
            page_num = item_data.get("page_num", 0)
            base += f" P.{page_num + 1}"
        return base

    def _undo(self):
        """直前の状態に戻す"""
        if len(self.undo_stack) <= 1:
            return
        current_state = self.undo_stack.pop()
        self.redo_stack.append(current_state)
        previous_state = copy.deepcopy(self.undo_stack[-1])
        self._apply_state_snapshot(previous_state)
        self._update_history_actions_state()

    def _redo(self):
        """一度戻した操作をやり直す"""
        if not self.redo_stack:
            return
        next_state = copy.deepcopy(self.redo_stack.pop())
        self._apply_state_snapshot(next_state)
        self.undo_stack.append(next_state)
        if len(self.undo_stack) > MAX_HISTORY:
            self.undo_stack.pop(0)
        self._update_history_actions_state()

    def _load_bookmarks_from_pdf(self, file_path, pdf_bookmarks):
        """PDFファイルから読み込んだしおりを追加（既存のしおりと重複しないように）"""
        normalized_path = os.path.abspath(file_path)
        existing_keys = {
            (os.path.abspath(b.get("path", "")), b.get("page_num", 0))
            for b in self.bookmarks
        }

        for bookmark in pdf_bookmarks:
            bookmark_key = (normalized_path, bookmark.get("page_num", 0))
            if bookmark_key not in existing_keys:
                bookmark["path"] = file_path  # 元のパスを保持（正規化前）
                self.bookmarks.append(bookmark)
                existing_keys.add(bookmark_key)

        # しおりツリーを更新
        if hasattr(self, "bookmark_dock") and self.bookmark_dock.isVisible():
            self._update_bookmark_tree()

    def _generate_bookmarks_from_list(self):
        """
        ページリストに基づいた自動しおりを再生成します。
        手動で作成したしおりは保持し、自動しおりのみを更新します。
        既存のしおり（手動・自動問わず）がある場合は、新しい自動しおりを追加しません。
        """
        removed_auto = self._remove_auto_bookmarks()
        orphan_removed = self._prune_orphan_bookmarks()

        if not self.auto_bookmarks_enabled:
            if (removed_auto or orphan_removed) and getattr(
                self.bookmark_dock, "isVisible", lambda: False
            )():
                self._update_bookmark_tree()
            return

        # 既存のしおりのキー（path, page_num）をセットとして保持（パスは正規化して比較）
        existing_bookmark_keys = {
            (os.path.abspath(b.get("path", "")), b.get("page_num", 0))
            for b in self.bookmarks
        }

        last_original_path = None
        auto_created = False
        for row in range(self.page_list_widget.count()):
            item = self.page_list_widget.item(row)
            if not item:
                continue
            item_data = item.data(Qt.UserRole)
            if not item_data:
                continue
            current_original_path = item_data["original_path"]

            if current_original_path != last_original_path:
                bookmark_key = (
                    os.path.abspath(current_original_path),
                    item_data.get("page_num", 0),
                )
                # 既存のしおりがない場合のみ自動しおりを追加
                if bookmark_key not in existing_bookmark_keys:
                    title = self._format_bookmark_title(item_data, include_page=False)
                    self.bookmarks.append(
                        {
                            "title": title,
                            "path": current_original_path,
                            "page_num": item_data.get("page_num", 0),
                            "auto": True,
                        }
                    )
                    auto_created = True
                    # 追加したしおりのキーを既存キーセットに追加（同じファイルの複数ページで重複追加を防ぐ）
                    existing_bookmark_keys.add(bookmark_key)
                last_original_path = current_original_path

        self._sort_bookmarks_by_page()
        if (
            (auto_created or removed_auto or orphan_removed)
            and hasattr(self, "bookmark_dock")
            and self.bookmark_dock.isVisible()
        ):
            self._update_bookmark_tree()

    def _update_item_text(self, item, item_data, base_filename=None):
        """アイテムのテキストを設定（ズームレベルに応じて省略処理）"""
        if base_filename is None:
            base_filename = os.path.basename(item_data["original_path"])

        if item_data["type"] == "pdf":
            # PDFの場合は、ファイル名とページ番号を2行で表示（ツールチップ用）
            page_num_text = f".P{item_data['page_num'] + 1}"
            icon_text = f"{base_filename}\n{page_num_text.lstrip('.')}"
        else:
            # PDF以外（Word, Excel, PPT）の場合は、ファイル名の下に空の2行目を作る
            icon_text = f"{base_filename}\n"

        # ファイル名とページ番号を1行で表示（幅がある場合のみページ番号も表示）
        if item_data["type"] == "pdf":
            # PDFの場合は、ファイル名とページ番号を1行で表示
            # 形式: "ファイル名.P1"（半角ドット、P、数字、スペースなし）
            # ファイル名とページ番号を結合
            combined_text = f"{base_filename}{page_num_text}"
            # テキスト表示領域の幅を計算
            font_metrics = self.page_list_widget.fontMetrics()
            current_width = int(THUMBNAIL_WIDTH * self.zoom_level)
            max_width = max(20, current_width - 12)

            # 幅が十分にある場合はファイル名とページ番号の両方を表示
            if font_metrics.horizontalAdvance(combined_text) <= max_width:
                final_text = combined_text
            else:
                # 幅が足りない場合も、ページ番号 suffix は必ず残す。
                final_text = self._elide_text_keep_suffix(
                    base_filename, page_num_text, max_width
                )
        else:
            # PDF以外の場合は、ファイル名のみ
            final_text = self._elide_text(base_filename)

        item.setText(final_text)
        item.setToolTip(icon_text)  # ツールチップには改行ありの完全なテキストを設定
        item.setTextAlignment(Qt.AlignCenter)

        # QListWidgetのIconModeでは、改行が表示されない可能性があるため、
        # フォントサイズを小さくするか、テキスト表示領域を調整する必要がある可能性がある

    def _elide_text_keep_suffix(self, text, suffix, max_width):
        """末尾の suffix を残したまま、先頭側のテキストだけ省略する。"""
        font_metrics = self.page_list_widget.fontMetrics()
        suffix_width = font_metrics.horizontalAdvance(suffix)
        filename_width = max(0, max_width - suffix_width)

        if filename_width <= 0:
            return suffix

        filename_text = font_metrics.elidedText(text, Qt.ElideRight, filename_width)
        final_text = f"{filename_text}{suffix}"

        while (
            filename_width > 0
            and font_metrics.horizontalAdvance(final_text) > max_width
        ):
            filename_width -= 4
            filename_text = font_metrics.elidedText(text, Qt.ElideRight, filename_width)
            final_text = f"{filename_text}{suffix}"

        return final_text

    def _elide_text(self, text):
        """
        指定されたテキストを、サムネイルの幅に合わせてピクセル単位で省略する。
        """
        # 空文字列の場合はそのまま返す
        if not text:
            return text

        # 現在のウィジェットのフォントから、寸法計算機（QFontMetrics）を取得
        font_metrics = self.page_list_widget.fontMetrics()

        # テキスト表示領域の幅は、グリッドサイズの幅を使用
        # グリッドサイズ = アイコンサイズ + パディング
        current_width = int(THUMBNAIL_WIDTH * self.zoom_level)
        grid_width = current_width + GRID_ITEM_PADDING_X

        # テキスト表示領域はグリッドサイズの幅だが、左右に少し余白がある
        max_width = grid_width - 10

        # テキスト全体のピクセル幅を計算
        if font_metrics.horizontalAdvance(text) <= max_width:
            return text  # 幅に収まる場合はそのまま返す

        # 幅を超える場合は、末尾に "..." を付けて省略する
        # Qtには便利な省略機能が組み込まれている
        return font_metrics.elidedText(text, Qt.ElideRight, max_width)

    def _merge_and_save(self):
        if self.page_list_widget.count() == 0:
            QMessageBox.warning(self, "警告", "結合するアイテムがありません。")
            return

        # 保存済みのヘッダー・フッター設定を使用
        header_footer_settings = (
            copy.deepcopy(self.header_footer_settings)
            if self.header_footer_settings
            else {}
        )

        # ページ番号の設定をヘッダー/フッターの設定に統合
        if self.page_number_settings.get("enabled"):
            page_number_format = self.page_number_settings.get("format", "1")
            page_number_start = self.page_number_settings.get("start_number", 1)
            is_header = self.page_number_settings.get("is_header", False)
            alignment = self.page_number_settings.get("alignment", "center")

            # ページ番号の位置を決定（left, center, right）
            page_number_position = alignment

            # ヘッダーまたはフッターの設定に統合
            if is_header:
                if "header" not in header_footer_settings:
                    header_footer_settings["header"] = {}
                header_footer_settings["header"]["auto_page_number"] = True
                header_footer_settings["header"][
                    "page_number_position"
                ] = page_number_position
            else:
                if "footer" not in header_footer_settings:
                    header_footer_settings["footer"] = {}
                header_footer_settings["footer"]["auto_page_number"] = True
                header_footer_settings["footer"][
                    "page_number_position"
                ] = page_number_position

            # ページ番号のフォーマットと開始番号を設定に追加
            header_footer_settings["page_number_format"] = page_number_format
            header_footer_settings["page_number_start"] = page_number_start

        output_path, _ = QFileDialog.getSaveFileName(
            self, "名前を付けて保存", self.last_used_path, "PDF Files (*.pdf)"
        )

        if not output_path:
            return

        self.last_used_path = os.path.dirname(output_path)

        items_data = [
            self.page_list_widget.item(i).data(Qt.UserRole)
            for i in range(self.page_list_widget.count())
        ]
        bookmarks_export = self._prepare_bookmarks_for_export(items_data)
        self._run_task(
            "merge_save",
            "名前を付けて保存",
            items_data=items_data,
            output_path=output_path,
            bookmarks=bookmarks_export,
            show_outlines=self.show_bookmarks_on_open,
            header_footer_settings=header_footer_settings,
        )

    def _prepare_bookmarks_for_export(self, items_data):
        """現在のしおりをPDF出力用に整形"""
        self._prune_orphan_bookmarks()
        export = []
        for bookmark in self.bookmarks:
            path = bookmark.get("path")
            if not path:
                continue
            export.append(
                {
                    "title": bookmark.get("title", "無題"),
                    "path": path,
                    "page_num": bookmark.get("page_num", 0),
                }
            )
        return export

    def _export_selected_as_pdf(self):
        """選択されたページをPDFとして書き出す"""
        selected_items = self.page_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.information(
                self, "書き出し", "書き出すページを選択してください。"
            )
            return

        # 選択されたアイテムのデータを取得（選択順を保持）
        selected_items_sorted = sorted(
            selected_items, key=lambda item: self.page_list_widget.row(item)
        )
        items_data = [item.data(Qt.UserRole) for item in selected_items_sorted]

        # 保存先を選択
        default_filename = f"exported_{len(items_data)}pages.pdf"
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "選択ページをPDFとして保存",
            os.path.join(self.last_used_path, default_filename),
            "PDF Files (*.pdf)",
        )

        if not output_path:
            return

        self.last_used_path = os.path.dirname(output_path)

        # しおりは選択範囲に含まれるものだけを抽出
        bookmarks_export = []
        for bookmark in self.bookmarks:
            bookmark_path = bookmark.get("path")
            bookmark_page = bookmark.get("page_num", 0)
            # 選択されたアイテムの中に該当するものがあるかチェック
            for item_data in items_data:
                if (
                    item_data.get("original_path") == bookmark_path
                    and item_data.get("type") == "pdf"
                    and item_data.get("page_num") == bookmark_page
                ):
                    bookmarks_export.append(
                        {
                            "title": bookmark.get("title", "無題"),
                            "path": bookmark_path,
                            "page_num": bookmark_page,
                        }
                    )
                    break

        # ヘッダー・フッター設定を使用
        header_footer_settings = (
            copy.deepcopy(self.header_footer_settings)
            if self.header_footer_settings
            else {}
        )

        # ページ番号の設定をヘッダー/フッターの設定に統合
        if self.page_number_settings.get("enabled"):
            page_number_format = self.page_number_settings.get("format", "1")
            page_number_start = self.page_number_settings.get("start_number", 1)
            is_header = self.page_number_settings.get("is_header", False)
            alignment = self.page_number_settings.get("alignment", "center")

            # ページ番号の位置を決定（left, center, right）
            page_number_position = alignment

            # ヘッダーまたはフッターの設定に統合
            if is_header:
                if "header" not in header_footer_settings:
                    header_footer_settings["header"] = {}
                header_footer_settings["header"]["auto_page_number"] = True
                header_footer_settings["header"][
                    "page_number_position"
                ] = page_number_position
            else:
                if "footer" not in header_footer_settings:
                    header_footer_settings["footer"] = {}
                header_footer_settings["footer"]["auto_page_number"] = True
                header_footer_settings["footer"][
                    "page_number_position"
                ] = page_number_position

            # ページ番号のフォーマットと開始番号を設定に追加
            header_footer_settings["page_number_format"] = page_number_format
            header_footer_settings["page_number_start"] = page_number_start

        self._run_task(
            "merge_save",
            "選択ページをPDFとして書き出し",
            items_data=items_data,
            output_path=output_path,
            bookmarks=bookmarks_export if bookmarks_export else None,
            show_outlines=False,  # 部分書き出しではしおりを自動表示しない
            header_footer_settings=header_footer_settings,
        )

    def _show_header_footer_settings(self):
        """ヘッダー・フッター設定ダイアログを表示"""
        dialog = HeaderFooterSettingsDialog(self)

        # 設定が None の場合は、デフォルト値を設定
        if not self.header_footer_settings:
            self.header_footer_settings = {
                "header_enabled": False,
                "footer_enabled": False,
                "font_size": 10,
                "header": {
                    "left": "",
                    "center": "",
                    "right": "",
                    "auto_date": False,
                },
                "footer": {
                    "left": "",
                    "center": "",
                    "right": "",
                    "auto_date": False,
                },
            }

        # 現在の設定をダイアログに反映
        dialog.set_settings(self.header_footer_settings)

        # ページ番号設定も統合されている場合、それも反映
        if self.page_number_settings.get("enabled"):
            is_header = self.page_number_settings.get("is_header", False)
            alignment = self.page_number_settings.get("alignment", "center")
            if is_header:
                if alignment == "center":
                    dialog.header_page_number_position_combo.setCurrentText("中央")
                elif alignment == "right":
                    dialog.header_page_number_position_combo.setCurrentText("右")
                else:
                    dialog.header_page_number_position_combo.setCurrentText("左")
            else:
                if alignment == "center":
                    dialog.footer_page_number_position_combo.setCurrentText("中央")
                elif alignment == "right":
                    dialog.footer_page_number_position_combo.setCurrentText("右")
                else:
                    dialog.footer_page_number_position_combo.setCurrentText("左")

            dialog.page_number_format_combo.setCurrentText(
                self.page_number_settings.get("format", "1")
            )
            dialog.page_number_start_spin.setValue(
                self.page_number_settings.get("start_number", 1)
            )

        if dialog.exec() == QDialog.DialogCode.Accepted:
            settings = dialog.get_settings()
            # チェックボックスが外されていても、設定内容は保持する
            self.header_footer_settings = settings

            # ページ番号設定を統合された設定から取得して、後方互換性のためにpage_number_settingsにも反映
            # （既存のコードがpage_number_settingsを参照している可能性があるため）
            header_has_page_number = settings.get("header", {}).get(
                "auto_page_number", False
            )
            footer_has_page_number = settings.get("footer", {}).get(
                "auto_page_number", False
            )

            if header_has_page_number or footer_has_page_number:
                self.page_number_settings["enabled"] = True
                if header_has_page_number:
                    self.page_number_settings["is_header"] = True
                    self.page_number_settings["alignment"] = settings.get(
                        "header", {}
                    ).get("page_number_position", "center")
                else:
                    self.page_number_settings["is_header"] = False
                    self.page_number_settings["alignment"] = settings.get(
                        "footer", {}
                    ).get("page_number_position", "center")
                self.page_number_settings["format"] = settings.get(
                    "page_number_format", "1"
                )
                self.page_number_settings["start_number"] = settings.get(
                    "page_number_start", 1
                )
            else:
                self.page_number_settings["enabled"] = False

            self._save_settings()
            # チェックボックスの状態を更新

    def _export_selected_as_images(self):
        """選択されたページを画像として書き出す"""
        selected_items = self.page_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.information(
                self, "書き出し", "書き出すページを選択してください。"
            )
            return

        # 保存先フォルダを選択
        output_dir = QFileDialog.getExistingDirectory(
            self,
            "画像の保存先フォルダを選択",
            self.last_used_path,
        )

        if not output_dir:
            return

        self.last_used_path = output_dir

        # 選択されたアイテムのデータを取得（選択順を保持）
        selected_items_sorted = sorted(
            selected_items, key=lambda item: self.page_list_widget.row(item)
        )
        items_data = [item.data(Qt.UserRole) for item in selected_items_sorted]

        # 画像書き出しは PDF ページのみを対象とし、Office ファイルは対象外とする
        pdf_items = [d for d in items_data if d and d.get("type") == "pdf"]
        non_pdf_items = [d for d in items_data if d and d.get("type") != "pdf"]

        if not pdf_items:
            # すべて Office など PDF 以外の場合は何もしない
            QMessageBox.information(
                self,
                "書き出し",
                "画像として書き出せるのは PDF ページのみです。\n"
                "Word / Excel / PowerPoint などは対象外です。",
            )
            return

        if non_pdf_items:
            # 混在している場合は PDF だけを書き出すことを通知
            QMessageBox.information(
                self,
                "書き出し",
                "PDF 以外のファイルが含まれていますが、\n"
                "画像として書き出されるのは PDF ページのみです。",
            )

        # 実際にワーカーへ渡すのは PDF ページのみ
        items_data = pdf_items

        self._run_task(
            "export_images",
            "選択ページを画像として書き出し",
            items_data=items_data,
            output_dir=output_dir,
            dpi=300,
            image_format="JPEG",
        )

    def setup_progress_dialog(self, user_facing_name):
        title = f"{user_facing_name}しています..."
        label = "準備しています..."
        _debug_log(
            f"[DEBUG] setup_progress_dialog: title='{title}', "
            f"label='{label}', activeThreadCount={self.threadpool.activeThreadCount()}"
        )
        self.progress_dialog = QProgressDialog(label, "キャンセル", 0, 100, self)
        self.progress_dialog.setWindowTitle(title)
        self.progress_dialog.setMinimumWidth(400)  # 単位はピクセルです
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.setAutoReset(True)
        self.progress_dialog.canceled.connect(self.cancel_worker)
        self.progress_dialog.show()

    def cancel_worker(self):
        if self.current_worker:
            self.current_worker.is_running = False
        if self.progress_dialog:
            self.progress_dialog.setLabelText("処理を中止しています...")
            self.progress_dialog.setCancelButton(None)

    def update_progress(self, value, message):
        try:
            if self.progress_dialog:
                self.progress_dialog.setValue(value)
                self.progress_dialog.setLabelText(message)
        except AttributeError:
            # progress_dialog がこのメソッドの実行中に閉じられた場合に発生する
            # 安全に無視できるため、何もしない
            pass

    def on_worker_finished(self, task_name, title, message):
        _debug_log(
            f"[DEBUG] on_worker_finished: task_name={task_name}, "
            f"title={title}, message={message}, "
            f"activeThreadCount(before)={self.threadpool.activeThreadCount()}"
        )
        if self.progress_dialog:
            _debug_log(
                "[DEBUG] on_worker_finished: progress_dialog が存在するため close() を呼びます。"
            )
            self.progress_dialog.close()
        self.progress_dialog = None
        self.current_worker = None
        if task_name == "add_files":
            self._generate_bookmarks_from_list()
            self._update_page_mode_actions_state()
            self._record_history_change()
            # キューに溜まったファイルパスがあれば処理
            if self._pending_file_paths:
                pending_files = self._pending_file_paths[:]
                self._pending_file_paths.clear()
                _debug_log(
                    f"[DEBUG] on_worker_finished: "
                    f"_pending_file_paths に溜まったファイルを検出。"
                    f"count={len(pending_files)}, paths={pending_files}"
                )
                # 重複チェック
                pending_files = self._check_duplicate_files(pending_files)
                if pending_files:
                    self.last_used_path = os.path.dirname(pending_files[0])
                    _debug_log(
                        "[DEBUG] on_worker_finished: "
                        "キュー分の add_files タスクを再実行します。"
                    )
                    self._run_task(
                        "add_files", "ファイルを追加", file_paths=pending_files
                    )
        self.update_status_bar()
        if task_name == "merge_save" and title == "保存完了":
            # メッセージから保存先パスを抽出
            # メッセージ形式: "PDFを正常に保存しました:\n{output_path}"
            output_path = None
            if "\n" in message:
                lines = message.split("\n")
                if len(lines) >= 2:
                    output_path = lines[1].strip()

            # PDFを標準アプリで開く
            if output_path and os.path.exists(output_path):
                try:
                    # Windowsの場合、os.startfile() を使用
                    if sys.platform == "win32":
                        os.startfile(output_path)
                    else:
                        # その他のOSの場合
                        QDesktopServices.openUrl(QUrl.fromLocalFile(output_path))
                except Exception as e:
                    print(f"PDFファイルを開けませんでした: {e}")

            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setWindowTitle(title)
            msg_box.setText(f"{message}\n\nリストをクリアしますか？")
            msg_box.setStandardButtons(
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            msg_box.setDefaultButton(QMessageBox.StandardButton.No)
            if msg_box.exec() == QMessageBox.StandardButton.Yes:
                self.page_list_widget.clear()
                self.thumbnail_cache.clear()
                self.bookmarks.clear()
                # 設定をリセット（新しいプロジェクトを開始するため）
                self._reset_page_settings()
                self.update_status_bar()
                self._update_page_mode_actions_state()
                self._record_history_change()

    def on_worker_error(self, title, message):
        if self.progress_dialog:
            self.progress_dialog.close()
        self.progress_dialog = None
        self.current_worker = None
        QMessageBox.critical(self, title, message)
        self.update_status_bar()

    def update_status_bar(self, force_page_mode=False):
        if self.page_list_widget.count() == 0:
            self.file_status_label.setText("アイテムを追加してください")
            self.selection_status_label.setText("")
        else:
            counts = {"pdf": 0, "word": 0, "excel": 0, "powerpoint": 0}
            unique_files = set()
            for i in range(self.page_list_widget.count()):
                data = self.page_list_widget.item(i).data(Qt.UserRole)
                if data["type"] == "pdf":
                    counts["pdf"] += 1
                elif data["original_path"] not in unique_files:
                    counts[data["type"]] += 1
                    unique_files.add(data["original_path"])

            status_parts = [
                f"PDFページ: {counts['pdf']}" if counts["pdf"] > 0 else None,
                f"Word: {counts['word']}" if counts["word"] > 0 else None,
                f"Excel: {counts['excel']}" if counts["excel"] > 0 else None,
                (f"PPT: {counts['powerpoint']}" if counts["powerpoint"] > 0 else None),
            ]

            file_info_text = " | ".join(filter(None, status_parts))
            selection_info_text = (
                f"選択中: {len(self.page_list_widget.selectedItems())}件"
            )

            self.file_status_label.setText(file_info_text)
            self.selection_status_label.setText(selection_info_text)

    def closeEvent(self, event):
        self._save_settings()  # ← [追加] アプリ終了前に設定を保存

        if self.threadpool.activeThreadCount() > 0:
            if (
                QMessageBox.question(
                    self, "確認", "処理が実行中です。本当に終了しますか？"
                )
                == QMessageBox.StandardButton.Yes
            ):
                if self.current_worker:
                    self.current_worker.is_running = False
                self.threadpool.waitForDone(2000)
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def _show_about_dialog(self):
        """アプリ情報とライセンス関連ファイルへのアクセスを提供するカスタムダイアログ"""

        dialog = QDialog(self)
        dialog.setWindowTitle(f"アプリ情報 - {APP_NAME}")
        dialog.setMinimumWidth(450)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(5)

        title_label = QLabel(APP_NAME)
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)

        layout.addWidget(title_label)

        version_label = QLabel(f"Version: {APP_VERSION}")
        version_label.setAlignment(Qt.AlignRight)

        layout.addWidget(version_label)

        desc_label = QLabel(
            "PDFおよびOfficeドキュメントを結合・編集するためのツールです。<br>"
            "Copyright (C) 2026 Takeshi Kashiwagi"
        )
        desc_label.setAlignment(Qt.AlignLeft)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("padding-left: 15px;")

        layout.addWidget(desc_label)

        # 区切り線
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)

        license_info = QLabel(
            "本ソフトウェアは <b>GNU AFFERO GENERAL PUBLIC LICENSE v3.0 (AGPL-3.0)</b> の下で提供されます。<br><br>"
            "使用しているオープンソースライブラリおよび本ソフトウェア自体の"
            "ライセンス条項、ソースコードは以下のボタンから確認できます。"
        )
        license_info.setAlignment(Qt.AlignLeft)
        license_info.setWordWrap(True)
        license_info.setStyleSheet("padding-left: 15px;")

        layout.addWidget(license_info)

        layout.addSpacing(10)

        btn_layout = QVBoxLayout()
        btn_layout.setSpacing(5)

        btn_license = QPushButton("使用許諾契約書 (AGPL-3.0) を表示")
        btn_license.setIcon(self.get_icon("fa5s.file-contract", "#f1c40f"))
        btn_license.clicked.connect(lambda: self._open_local_file("LICENSE.txt"))

        btn_notice = QPushButton("サードパーティライセンス (NOTICE) を表示")
        btn_notice.setIcon(self.get_icon("fa5s.file-alt", "#f1c40f"))
        btn_notice.clicked.connect(lambda: self._open_local_file("NOTICE.txt"))

        btn_source = QPushButton("ソースコード格納フォルダを開く")
        btn_source.setIcon(self.get_icon("fa5s.code", "#2ecc71"))
        btn_source.clicked.connect(lambda: self._open_file_folder("source.zip"))

        btn_layout.addWidget(btn_license)
        btn_layout.addWidget(btn_notice)
        btn_layout.addWidget(btn_source)

        layout.addLayout(btn_layout)

        layout.addSpacing(5)
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)

        dialog.exec()

    def _open_local_file(self, filename):
        """アプリと同じフォルダにあるテキストファイルを開く"""
        try:
            # 実行ファイル(exe)のある場所を取得
            base_path = (
                os.path.dirname(sys.executable)
                if getattr(sys, "frozen", False)
                else os.path.dirname(__file__)
            )
            file_path = os.path.join(base_path, filename)

            if os.path.exists(file_path):
                os.startfile(file_path)
            else:
                QMessageBox.warning(
                    self,
                    "ファイル不明",
                    f"'{filename}' が見つかりません。\nインストールフォルダを確認してください。",
                )
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"ファイルを開けませんでした:\n{e}")

    def _open_file_folder(self, filename):
        """指定したファイルが選択された状態でエクスプローラーを開く"""
        try:
            base_path = (
                os.path.dirname(sys.executable)
                if getattr(sys, "frozen", False)
                else os.path.dirname(__file__)
            )
            file_path = os.path.join(base_path, filename)

            if os.path.exists(file_path):
                # Windowsのエクスプローラーでファイルを選択状態で開くコマンド
                subprocess.run(["explorer", "/select,", file_path])
            else:
                QMessageBox.warning(
                    self,
                    "ファイル不明",
                    f"'{filename}' が見つかりません。\nインストールフォルダを確認してください。",
                )
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"フォルダを開けませんでした:\n{e}")

    @Slot(str)
    def _on_non_cancellable_started(self, message):
        """中断不可処理が始まったときに呼び出されるスロット。"""
        if self.progress_dialog:
            self.progress_dialog.setLabelText(message)
            self.progress_dialog.setCancelButtonText(None)  # これでボタンが非表示になる

    def _toggle_bookmark_panel(self, checked):
        """ブックマークパネルの表示/非表示を切り替えます。"""
        self.bookmark_dock.setVisible(checked)

    def _on_bookmark_dock_visibility_changed(self, visible):
        """ドックのクローズボタンなどで状態が変わった際にUIを同期"""
        if hasattr(self, "bookmark_panel_action"):
            self.bookmark_panel_action.blockSignals(True)
            self.bookmark_panel_action.setChecked(visible)
            self.bookmark_panel_action.blockSignals(False)
        if visible:
            self._update_bookmark_tree()
        self._update_bookmark_add_state()
        self._update_bookmark_buttons_state()

    def _update_bookmark_tree(self):
        """ブックマークツリーを更新します。"""
        if not hasattr(self, "bookmark_tree"):
            return
        self._prune_orphan_bookmarks()
        self._sort_bookmarks_by_page()
        self.bookmark_tree.clear()
        for idx, bookmark in enumerate(self.bookmarks):
            display_title = bookmark.get("title", "無題")
            if bookmark.get("auto"):
                display_title += " (自動)"
            item = QTreeWidgetItem([display_title])
            tooltip = f"{bookmark.get('title','無題')}\n{os.path.basename(bookmark.get('path',''))}"
            if "page_num" in bookmark:
                tooltip += f" / P.{bookmark.get('page_num', 0) + 1}"
            item.setToolTip(0, tooltip)
            item.setData(0, Qt.UserRole, bookmark)
            item.setData(0, self.BOOKMARK_INDEX_ROLE, idx)
            self.bookmark_tree.addTopLevelItem(item)
        self.bookmark_tree.expandAll()
        self._update_bookmark_buttons_state()

    def _update_bookmark_add_state(self):
        """ページ選択状況に応じて追加ボタンの状態を更新"""
        if not hasattr(self, "bookmark_add_button"):
            return
        has_selection = bool(self.page_list_widget.selectedItems())
        self.bookmark_add_button.setEnabled(has_selection)

    def _update_bookmark_buttons_state(self):
        """ブックマーク選択に応じて編集ボタンを更新（自動しおりも編集可能）"""
        if not hasattr(self, "bookmark_tree"):
            return
        current_item = self.bookmark_tree.currentItem()
        if not current_item:
            enable_edit = False
        else:
            bookmark = current_item.data(0, Qt.UserRole)
            # 自動しおりも編集可能（編集時に手動しおりに変換される）
            enable_edit = bool(bookmark)
        if hasattr(self, "bookmark_rename_button"):
            self.bookmark_rename_button.setEnabled(enable_edit)
        if hasattr(self, "bookmark_delete_button"):
            self.bookmark_delete_button.setEnabled(enable_edit)

    def _show_page_context_menu(self, pos):
        """ページリストの右クリックメニュー"""
        if not hasattr(self, "page_list_widget"):
            return

        item = self.page_list_widget.itemAt(pos)
        if item and not item.isSelected():
            self.page_list_widget.clearSelection()
            item.setSelected(True)

        menu = QMenu(self)
        menu.addAction(self.add_action)

        has_selection = bool(self.page_list_widget.selectedItems())
        if has_selection:
            menu.addSeparator()
            menu.addAction(self.rot_left_action)
            menu.addAction(self.rot_right_action)
            menu.addSeparator()
            menu.addAction(self.move_to_top_action)
            menu.addAction(self.move_up_action)
            menu.addAction(self.move_down_action)
            menu.addAction(self.move_to_bottom_action)

            bookmark_action = QAction("選択ページにしおり", self)
            bookmark_action.triggered.connect(self._add_bookmark_from_selection)
            menu.addSeparator()
            menu.addAction(bookmark_action)
            menu.addSeparator()
            menu.addAction(self.export_selected_pdf_action)
            menu.addAction(self.export_selected_images_action)
            menu.addSeparator()
            menu.addAction(self.delete_action)

        global_pos = self.page_list_widget.viewport().mapToGlobal(pos)
        menu.exec(global_pos)

    def _add_bookmark_from_selection(self):
        """選択中のページから手動しおりを追加"""
        selected_items = self.page_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "しおり", "ページを選択してください。")
            return
        item = selected_items[0]
        item_data = item.data(Qt.UserRole)
        if not item_data:
            return
        default_title = self._format_bookmark_title(item_data)
        title, ok = QInputDialog.getText(
            self,
            "しおりを追加",
            "しおり名:",
            QLineEdit.EchoMode.Normal,
            default_title,
        )
        if not ok:
            return
        title = title.strip()
        if not title:
            QMessageBox.warning(self, "しおり", "しおり名を入力してください。")
            return
        self.bookmarks.append(
            {
                "title": title,
                "path": item_data["original_path"],
                "page_num": item_data.get("page_num", 0),
                "auto": False,
            }
        )
        self._sort_bookmarks_by_page()
        self._update_bookmark_tree()
        self._record_history_change()

    def _rename_selected_bookmark(self):
        """選択したしおりの名前を変更（自動しおりの場合は手動しおりに変換）"""
        item = self.bookmark_tree.currentItem()
        if not item:
            return
        bookmark = item.data(0, Qt.UserRole)
        if not bookmark:
            return
        index = item.data(0, self.BOOKMARK_INDEX_ROLE)
        if index is None or not (0 <= int(index) < len(self.bookmarks)):
            QMessageBox.warning(self, "しおり", "しおり情報を取得できませんでした。")
            return
        index = int(index)
        is_auto = bookmark.get("auto", False)
        new_title, ok = QInputDialog.getText(
            self,
            "しおりの名前変更",
            "新しい名前:",
            QLineEdit.EchoMode.Normal,
            bookmark.get("title", ""),
        )
        if not ok:
            return
        new_title = new_title.strip()
        if not new_title:
            QMessageBox.warning(self, "しおり", "名前を入力してください。")
            return
        self.bookmarks[index]["title"] = new_title
        # 自動しおりを編集した場合は手動しおりに変換（次回自動生成時に再生成されない）
        if is_auto:
            self.bookmarks[index]["auto"] = False
        self._update_bookmark_tree()
        self._record_history_change()

    def _delete_selected_bookmark(self):
        """選択したしおりを削除（自動しおりも削除可能、次回自動生成時に再生成される）"""
        item = self.bookmark_tree.currentItem()
        if not item:
            return
        bookmark = item.data(0, Qt.UserRole)
        if not bookmark:
            return
        index = item.data(0, self.BOOKMARK_INDEX_ROLE)
        if index is None or not (0 <= int(index) < len(self.bookmarks)):
            QMessageBox.warning(self, "しおり", "しおり情報を取得できませんでした。")
            return
        index = int(index)
        self.bookmarks.pop(index)
        self._update_bookmark_tree()
        self._record_history_change()

    def _toggle_auto_bookmarks(self, checked):
        """自動しおりのON/OFF切り替え"""
        self.auto_bookmarks_enabled = checked
        self.auto_bookmark_action.setChecked(checked)
        self._generate_bookmarks_from_list()
        self._save_settings()
        self._record_history_change()

    def _toggle_show_bookmarks_on_open(self, checked):
        """PDF保存時のしおり表示挙動を切り替え"""
        self.show_bookmarks_on_open = checked
        self.show_outline_on_open_action.setChecked(checked)
        self._save_settings()
        self._record_history_change()

    def _navigate_to_bookmark(self, item, column):
        """ブックマークをクリックしたときに該当ページに移動します。"""
        if not item:
            return
        bookmark = item.data(0, Qt.UserRole)
        if not bookmark:
            return
        target_row = self._find_row_for_bookmark(bookmark)
        if target_row is None:
            QMessageBox.warning(
                self, "しおり", "該当するページが見つかりませんでした。"
            )
            return
        target_item = self.page_list_widget.item(target_row)
        if target_item:
            self.page_list_widget.setCurrentItem(target_item)
            self.page_list_widget.scrollToItem(
                target_item, QAbstractItemView.ScrollHint.PositionAtCenter
            )

    def _find_row_for_bookmark(self, bookmark):
        """しおりが指すページの現在の行番号を取得（パスは正規化して比較）"""
        target_key = (
            os.path.abspath(bookmark.get("path", "")),
            bookmark.get("page_num", 0),
        )
        for row in range(self.page_list_widget.count()):
            item = self.page_list_widget.item(row)
            if not item:
                continue
            data = item.data(Qt.UserRole)
            if not data:
                continue
            key = (os.path.abspath(data["original_path"]), data.get("page_num", 0))
            if key == target_key:
                return row
        return None

    def _zoom_in(self):
        """サムネイルを拡大"""
        if self.zoom_index < len(self.zoom_levels) - 1:
            self.zoom_index += 1
            self.zoom_level = self.zoom_levels[self.zoom_index]
            self._apply_zoom()

    def _zoom_out(self):
        """サムネイルを縮小"""
        if self.zoom_index > 0:
            self.zoom_index -= 1
            self.zoom_level = self.zoom_levels[self.zoom_index]
            self._apply_zoom()

    def _zoom_fit(self):
        """サムネイルサイズをデフォルトに戻す"""
        default_index = self.zoom_levels.index(1.0)
        if self.zoom_index != default_index:
            self.zoom_index = default_index
            self.zoom_level = self.zoom_levels[self.zoom_index]
            self._apply_zoom()

    def _on_zoom_requested(self, direction):
        """マウスホイールからのズーム要求を処理（デバウンス処理付き）"""
        # ホイールイベントの累積量を更新
        self.zoom_wheel_accumulator += direction

        # タイマーをリセット（150ms後に処理を実行）
        self.zoom_timer.stop()
        self.zoom_timer.start(150)  # 150ms待機

    def _process_accumulated_zoom(self):
        """累積されたホイールイベントを処理"""
        if self.zoom_wheel_accumulator == 0:
            return

        # 累積量に応じてズーム段階を移動（閾値: 3回分）
        zoom_steps = self.zoom_wheel_accumulator // 3
        if zoom_steps == 0:
            return

        new_index = self.zoom_index + zoom_steps
        new_index = max(0, min(len(self.zoom_levels) - 1, new_index))
        if new_index == self.zoom_index:
            self.zoom_wheel_accumulator = 0
            return

        self.zoom_index = new_index
        self.zoom_level = self.zoom_levels[self.zoom_index]

        # 累積量をリセット
        self.zoom_wheel_accumulator = 0

        # ズームを適用（一度だけ）
        self._apply_zoom()

    def _apply_zoom(self):
        """ズームレベルを適用してサムネイルサイズとグリッドサイズを更新"""
        # 念のためズームレベルを現在の段階に合わせる
        self.zoom_level = self.zoom_levels[self.zoom_index]

        # 新しいサムネイルサイズを計算
        new_width = int(THUMBNAIL_WIDTH * self.zoom_level)
        new_height = int(THUMBNAIL_HEIGHT * self.zoom_level)
        new_thumbnail_size = QSize(new_width, new_height)

        # 新しいグリッドサイズを計算
        new_grid_width = new_width + GRID_ITEM_PADDING_X
        new_grid_height = new_height + GRID_ITEM_PADDING_Y
        new_grid_size = QSize(new_grid_width, new_grid_height)

        # ページリストウィジェットに適用（アイコンサイズとグリッドサイズ）
        self.page_list_widget.setIconSize(new_thumbnail_size)
        self.page_list_widget.setGridSize(new_grid_size)

        # ズームレベルの状態に応じてボタンの有効/無効を更新
        self.zoom_in_action.setEnabled(self.zoom_index < len(self.zoom_levels) - 1)
        self.zoom_out_action.setEnabled(self.zoom_index > 0)

        # サムネイルを即座にスケール（キャッシュから取得してスケール）
        # ページリストのアイテムのサムネイルを更新
        for i in range(self.page_list_widget.count()):
            item = self.page_list_widget.item(i)
            if item:
                item_data = item.data(Qt.UserRole)
                if item_data:
                    # キャッシュから取得してスケール（再生成ではない）
                    item.setIcon(self._create_thumbnail(item_data, new_thumbnail_size))
                    # テキストも再計算（ズームレベルに応じて省略が変わる）
                    self._update_item_text(item, item_data)

    def _update_grid_size(self):
        """グリッドサイズを現在のズームレベルに基づいて再設定する"""
        # 現在のズームレベルに基づいたサイズを計算
        new_width = int(THUMBNAIL_WIDTH * self.zoom_level)
        new_height = int(THUMBNAIL_HEIGHT * self.zoom_level)
        new_thumbnail_size = QSize(new_width, new_height)

        # 新しいグリッドサイズを計算
        new_grid_width = new_width + GRID_ITEM_PADDING_X
        new_grid_height = new_height + GRID_ITEM_PADDING_Y
        new_grid_size = QSize(new_grid_width, new_grid_height)

        # ページリストウィジェットに適用
        self.page_list_widget.setIconSize(new_thumbnail_size)
        self.page_list_widget.setGridSize(new_grid_size)

    def _regenerate_thumbnails(self):
        """サムネイルを再生成（遅延実行）"""
        if self.pending_thumbnail_size is None:
            return

        # 現在のズームレベルに基づいたサイズを計算（確実に最新のサイズを使用）
        new_width = int(THUMBNAIL_WIDTH * self.zoom_level)
        new_height = int(THUMBNAIL_HEIGHT * self.zoom_level)
        new_thumbnail_size = QSize(new_width, new_height)

        # ページリストのアイテムのサムネイルとテキストを再生成
        for i in range(self.page_list_widget.count()):
            item = self.page_list_widget.item(i)
            if item:
                item_data = item.data(Qt.UserRole)
                if item_data:
                    # 新しいサイズでサムネイルを再生成
                    item.setIcon(self._create_thumbnail(item_data, new_thumbnail_size))
                    # テキストも再計算（ズームレベルに応じて省略が変わる）
                    self._update_item_text(item, item_data)

        self.pending_thumbnail_size = None

    @Slot()
    def _on_non_cancellable_finished(self):
        """中断不可処理が終わったときに呼び出されるスロット。"""
        if self.progress_dialog:
            self.progress_dialog.setCancelButtonText(
                "キャンセル"
            )  # これでボタンが再表示される


# --- 単一インスタンス制御用の簡易サーバ ---
_SINGLE_INSTANCE_SERVER_NAME = "OfficePDFBinder_SingleInstance"

# デバッグログ用（ファイルにも出力）
_DEBUG_LOG_PATH = os.path.join(os.path.expanduser("~"), "OfficePDFBinder_debug.log")

# デフォルトでは配布ビルドでログを書かない。必要な場合は環境変数で有効化する。
_ENABLE_DEBUG_LOG = os.environ.get("OFFICEPDFBINDER_DEBUG_LOG", "").lower() in (
    "1",
    "true",
    "yes",
    "on",
)


def _debug_log(msg):
    """デバッグログをコンソールとファイルに出力（_ENABLE_DEBUG_LOG が True のときのみ）"""
    if not _ENABLE_DEBUG_LOG:
        return
    log_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(log_msg)
    try:
        with open(_DEBUG_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(log_msg + "\n")
    except Exception:
        pass  # ログファイル書き込み失敗は無視


def _startup_log(msg):
    """起動時の処理時間を測るためのデバッグログ。"""
    elapsed = time.perf_counter() - _STARTUP_TIME
    _debug_log(f"[STARTUP +{elapsed:.3f}s] {msg}")


def _get_app_icon():
    """アプリのウィンドウアイコンを取得する。"""
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    icon_path = os.path.join(base_dir, APP_ICON_FILENAME)
    if os.path.exists(icon_path):
        return QIcon(icon_path)
    return QIcon()


def _get_settings_file_path():
    """ユーザー設定ファイルのパスを返す。"""
    return os.path.join(os.environ["APPDATA"], "OfficePDFBinder", "settings.ini")


def _available_geometry_for_rect(rect=None):
    """指定位置に対応する画面の使用可能領域を取得する。"""
    screen = None
    app = QApplication.instance()
    if app is not None and rect is not None:
        center = rect.center()
        screen = app.screenAt(center)
    if screen is None:
        screen = QApplication.primaryScreen()
    if screen is not None:
        return screen.availableGeometry()
    return QRect(0, 0, INITIAL_WINDOW_WIDTH, INITIAL_WINDOW_HEIGHT)


def _fit_rect_to_available_geometry(rect):
    """ウィンドウ矩形を現在利用できる画面領域内に収める。"""
    available = _available_geometry_for_rect(rect)
    max_width = max(MIN_WINDOW_WIDTH, available.width() - WINDOW_SCREEN_MARGIN)
    max_height = max(MIN_WINDOW_HEIGHT, available.height() - WINDOW_SCREEN_MARGIN)
    width = min(max(rect.width(), MIN_WINDOW_WIDTH), max_width)
    height = min(max(rect.height(), MIN_WINDOW_HEIGHT), max_height)

    left_limit = available.left()
    top_limit = available.top()
    right_limit = available.right() - width + 1
    bottom_limit = available.bottom() - height + 1

    x = min(max(rect.x(), left_limit), max(left_limit, right_limit))
    y = min(max(rect.y(), top_limit), max(top_limit, bottom_limit))
    return QRect(x, y, width, height)


def _default_window_geometry_for_available(available):
    """指定された画面領域に対する既定のウィンドウ位置とサイズを返す。"""
    width = min(INITIAL_WINDOW_WIDTH, max(MIN_WINDOW_WIDTH, available.width() - 80))
    height = min(INITIAL_WINDOW_HEIGHT, max(MIN_WINDOW_HEIGHT, available.height() - 80))
    x = available.left() + min(INITIAL_WINDOW_X, max(0, (available.width() - width) // 2))
    y = available.top() + min(INITIAL_WINDOW_Y, max(0, (available.height() - height) // 3))
    return _fit_rect_to_available_geometry(QRect(x, y, width, height))


def _default_window_geometry():
    """初回起動時のウィンドウ位置とサイズを画面に合わせて決める。"""
    available = _available_geometry_for_rect()
    return _default_window_geometry_for_available(available)


def _restore_window_geometry(rect):
    """保存済み位置を復元する。収まらない場合は既定位置寄りに戻す。"""
    available = _available_geometry_for_rect(rect)
    fitted = _fit_rect_to_available_geometry(rect)
    if fitted == rect:
        return fitted
    default_rect = _default_window_geometry_for_available(available)
    return QRect(default_rect.x(), default_rect.y(), fitted.width(), fitted.height())


def _load_startup_window_geometry():
    """起動直後に使うウィンドウ位置を設定ファイルから取得する。"""
    settings_file = _get_settings_file_path()
    if not os.path.exists(settings_file):
        return _default_window_geometry(), False

    config = configparser.ConfigParser()
    try:
        config.read(settings_file, encoding="utf-8")
        if not config.has_section("Window"):
            return _default_window_geometry(), False
        rect = QRect(
            config.getint("Window", "x", fallback=INITIAL_WINDOW_X),
            config.getint("Window", "y", fallback=INITIAL_WINDOW_Y),
            config.getint("Window", "width", fallback=INITIAL_WINDOW_WIDTH),
            config.getint("Window", "height", fallback=INITIAL_WINDOW_HEIGHT),
        )
        maximized = config.getboolean("Window", "maximized", fallback=False)
        return _restore_window_geometry(rect), maximized
    except Exception:
        return _default_window_geometry(), False


def _send_files_to_running_instance(file_paths):
    """既に起動中のインスタンスにファイルパスを送信する。

    成功した場合は True（＝自分はすぐ終了してよい）、失敗した場合は False を返す。
    """
    # QApplicationが必要なので、ここで作成（既に存在する場合は再利用）
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
        _startup_log("child QApplication created")

    local_socket = QLocalSocket()

    if not file_paths:
        # 空の場合は接続確認のみ（引数なし起動時の多重起動防止用）
        local_socket.connectToServer(_SINGLE_INSTANCE_SERVER_NAME)
        if local_socket.waitForConnected(300):  # 300ms待機
            local_socket.disconnectFromServer()
            _startup_log("existing instance detected without file arguments")
            return True
        _startup_log("no existing instance detected without file arguments")
        return False

    _debug_log(
        f"[DEBUG] _send_files_to_running_instance: {len(file_paths)}個のファイルを送信しようとしています: {file_paths}"
    )

    local_socket.connectToServer(_SINGLE_INSTANCE_SERVER_NAME)
    if not local_socket.waitForConnected(500):  # 500ms待機
        _debug_log(
            f"[DEBUG] _send_files_to_running_instance: 接続失敗 ({local_socket.errorString()})"
        )
        _startup_log("child failed to connect to existing instance")
        return False

    try:
        # 改行区切りでパスを送信（UTF-8）
        payload = "\n".join(file_paths).encode("utf-8", errors="ignore")
        local_socket.write(payload)
        local_socket.flush()
        local_socket.waitForBytesWritten(1000)  # 1秒待機
        _debug_log("[DEBUG] _send_files_to_running_instance: 送信成功、終了します")
        _startup_log("child sent file paths to existing instance")
        return True
    except Exception as e:
        _debug_log(f"[DEBUG] _send_files_to_running_instance: 送信エラー ({e})")
        return False
    finally:
        local_socket.disconnectFromServer()


def _handle_new_connection(server, window):
    """QLocalServerの新しい接続を処理"""
    local_socket = server.nextPendingConnection()
    if local_socket is None:
        return

    _debug_log("[DEBUG] _handle_new_connection: 接続を受け付けました")

    def _read_data():
        """ソケットからデータを読み取る"""
        try:
            if local_socket.bytesAvailable() > 0:
                data = local_socket.readAll()
                text = data.data().decode("utf-8", errors="ignore")

                # 改行区切りでパスを復元
                paths = [line.strip() for line in text.splitlines() if line.strip()]
                if not paths:
                    _debug_log("[DEBUG] _handle_new_connection: パスが空")
                    local_socket.disconnectFromServer()
                    return

                _debug_log(
                    f"[DEBUG] _handle_new_connection: {len(paths)}個のファイルパスを受信しました: {paths}"
                )
                # GUI スレッドでファイル追加を実行（シグナルを使用）
                try:
                    window.ipc_files_received.emit(paths)
                except Exception as e:
                    _debug_log(
                        f"[DEBUG] _handle_new_connection: ipc_files_received.emit でエラー発生: {e}"
                    )

                local_socket.disconnectFromServer()
        except Exception as e:
            _debug_log(f"[DEBUG] _handle_new_connection: データ読み取りエラー ({e})")
            try:
                local_socket.disconnectFromServer()
            except Exception:
                pass

    # 接続時に既にデータが利用可能な場合をチェック
    if local_socket.bytesAvailable() > 0:
        _read_data()
    else:
        # データが利用可能になったら読み取る
        local_socket.readyRead.connect(_read_data)

    # エラー処理
    local_socket.errorOccurred.connect(
        lambda error: _debug_log(
            f"[DEBUG] _handle_new_connection: ソケットエラー ({error})"
        )
    )


if __name__ == "__main__":
    import os as os_module

    process_id = os_module.getpid()
    # コマンドライン引数からファイルパスを取得（最初の引数はスクリプト名なので除外）
    initial_file_paths = sys.argv[1:] if len(sys.argv) > 1 else []
    _startup_log(
        f"process start [PID:{process_id}], file_args={len(initial_file_paths)}"
    )
    _debug_log(
        f"[DEBUG] main: 起動 [PID:{process_id}]、引数={len(initial_file_paths)}個のファイル: {initial_file_paths}"
    )
    _debug_log(f"[DEBUG] main: デバッグログファイル = {_DEBUG_LOG_PATH}")

    # --- 親インスタンス判定（Windows Named Mutex を使用） ---
    is_main_instance = True
    mutex_handle = None

    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes

            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            CreateMutexW = kernel32.CreateMutexW
            CreateMutexW.argtypes = (
                wintypes.LPVOID,
                wintypes.BOOL,
                wintypes.LPCWSTR,
            )
            CreateMutexW.restype = wintypes.HANDLE

            mutex_name = "Global\\OfficePDFBinder_SingleInstance"
            mutex_handle = CreateMutexW(None, False, mutex_name)
            last_error = ctypes.get_last_error()

            if not mutex_handle:
                _debug_log(
                    "[DEBUG] main: Named mutex の作成に失敗しました（暫定的にメインとして起動）"
                )
            else:
                ERROR_ALREADY_EXISTS = 183  # ERROR_ALREADY_EXISTS
                if last_error == ERROR_ALREADY_EXISTS:
                    is_main_instance = False
                    _startup_log("mutex checked: child instance")
                    _debug_log(
                        f"[DEBUG] main: [PID:{process_id}] 既にメインインスタンスが存在するため、子インスタンスとして動作します"
                    )
                else:
                    _startup_log("mutex checked: main instance")
                    _debug_log(
                        f"[DEBUG] main: [PID:{process_id}] Named mutex を取得しました（メインインスタンス）"
                    )
        except Exception as e:
            _startup_log("mutex check failed; continue as main instance")
            _debug_log(
                f"[DEBUG] main: Named mutex 初期化に失敗しました（メインとして続行）: {e}"
            )
            is_main_instance = True
    else:
        _debug_log(
            "[DEBUG] main: 非Windows環境のため、常にメインインスタンスとして動作します"
        )

    # --- 子インスタンス側の処理 ---
    if not is_main_instance:
        if initial_file_paths:
            # 既存インスタンスへ自分のファイルパスを送信して終了
            for retry_count in range(5):
                _startup_log(f"child send attempt {retry_count + 1}")
                if _send_files_to_running_instance(initial_file_paths):
                    _debug_log(
                        f"[DEBUG] main: [PID:{process_id}] 既存インスタンスにファイルを送信して終了します"
                    )
                    _startup_log("child exit after successful send")
                    sys.exit(0)
                time.sleep(0.3)

            _debug_log(
                f"[DEBUG] main: [PID:{process_id}] 既存インスタンスへの送信に失敗したため、何もせず終了します"
            )
            _startup_log("child exit after send failure")
        else:
            _debug_log(
                f"[DEBUG] main: [PID:{process_id}] 子インスタンス（引数なし）のため、何もせず終了します"
            )
            _startup_log("child exit without file arguments")
        sys.exit(0)

    # ここから先はメインインスタンスのみ

    # --- アプリケーション起動 ---
    app = QApplication(sys.argv)
    app.setApplicationName("Office PDF Binder")
    app.setWindowIcon(_get_app_icon())
    _startup_log("main QApplication created")
    startup_geometry, startup_maximized = _load_startup_window_geometry()

    window = OfficePDFBinderApp(startup_geometry, startup_maximized)
    _startup_log("main window created")

    # IPC 用ローカルサーバーを開始（ファイル追加要求を受け付ける）
    local_server = QLocalServer()
    try:
        # 既存のサーバーが残っている可能性があるので、削除を試みる
        if local_server.listen(_SINGLE_INSTANCE_SERVER_NAME):
            _startup_log("local server listen started")
            _debug_log(
                f"[DEBUG] main: [PID:{process_id}] ローカルサーバー開始（{_SINGLE_INSTANCE_SERVER_NAME}）"
            )
            # 新しい接続を受け取るシグナルを接続
            local_server.newConnection.connect(
                lambda: _handle_new_connection(local_server, window)
            )
        else:
            # 既存のサーバーが残っている場合は削除して再試行
            QLocalServer.removeServer(_SINGLE_INSTANCE_SERVER_NAME)
            if local_server.listen(_SINGLE_INSTANCE_SERVER_NAME):
                _startup_log("local server listen started after removeServer")
                _debug_log(
                    f"[DEBUG] main: [PID:{process_id}] ローカルサーバー開始（再試行成功: {_SINGLE_INSTANCE_SERVER_NAME}）"
                )
                local_server.newConnection.connect(
                    lambda: _handle_new_connection(local_server, window)
                )
            else:
                _startup_log("local server listen failed")
                _debug_log(
                    f"[DEBUG] main: [PID:{process_id}] ローカルサーバーの起動に失敗しました（IPCは無効）: {local_server.errorString()}"
                )
    except Exception as e:
        _startup_log("local server setup failed")
        _debug_log(
            f"[DEBUG] main: [PID:{process_id}] ローカルサーバーの起動に失敗しました（IPCは無効）: {e}"
        )

    window.show()
    _startup_log("main window shown")

    # 起動時にファイルが指定されていた場合は、自分自身でも読み込む
    if initial_file_paths:
        _debug_log(
            f"[DEBUG] main: [PID:{process_id}] メインインスタンスとして起動時のファイルを読み込みます: {initial_file_paths}"
        )
        _startup_log("initial file loading scheduled")
        # GUI 初期化待ちのために少し遅延させる
        QTimer.singleShot(300, lambda: window._add_files_from_paths(initial_file_paths))

    _startup_log("event loop start")
    sys.exit(app.exec())
