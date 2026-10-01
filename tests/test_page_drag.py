"""Check hold gestures without entering a native drag-and-drop loop."""
import pytest
from PySide6.QtCore import QEvent, QPoint, QSize, Qt
from PySide6.QtGui import QMouseEvent
from PySide6.QtWidgets import QApplication, QListWidget

from OfficePDFBinder_Main import DropListWidget


@pytest.fixture
def page_list(qtbot):
    widget = DropListWidget()
    widget.setViewMode(QListWidget.IconMode)
    widget.setMovement(QListWidget.Static)
    widget.setSelectionMode(QListWidget.ExtendedSelection)
    widget.setGridSize(QSize(100, 140))
    widget.resize(600, 400)
    widget.addItems([str(i) for i in range(8)])
    qtbot.addWidget(widget)
    widget.show()
    QApplication.processEvents()
    return widget


def move_while_pressed(widget, position):
    event = QMouseEvent(
        QEvent.MouseMove, position, widget.viewport().mapToGlobal(position),
        Qt.NoButton, Qt.LeftButton, Qt.NoModifier,
    )
    QApplication.sendEvent(widget.viewport(), event)


@pytest.mark.parametrize("offset,armed", [(QPoint(5, 5), True), (QPoint(10, 0), False)])
def test_hold_tolerates_diagonal_jitter_but_preserves_range_selection(page_list, qtbot, offset, armed):
    position = page_list.visualItemRect(page_list.item(0)).center()
    qtbot.mousePress(page_list.viewport(), Qt.LeftButton, pos=position)
    try:
        assert page_list._drag_timer.interval() == 350
        move_while_pressed(page_list, position + offset)
        if armed:
            qtbot.waitUntil(lambda: page_list._drag_mode_enabled, timeout=1000)
            assert page_list.viewport().cursor().shape() == Qt.OpenHandCursor
        else:
            qtbot.wait(450)
            assert not page_list._drag_mode_enabled
    finally:
        qtbot.mouseRelease(page_list.viewport(), Qt.LeftButton, pos=position + offset)
    assert not page_list._drag_mode_enabled
    assert page_list.viewport().cursor().shape() != Qt.OpenHandCursor


@pytest.mark.parametrize("modifier", [Qt.ControlModifier, Qt.ShiftModifier])
def test_modifier_selection_does_not_arm_drag(page_list, qtbot, modifier):
    position = page_list.visualItemRect(page_list.item(0)).center()
    qtbot.mousePress(page_list.viewport(), Qt.LeftButton, modifier, pos=position)
    try:
        assert not page_list._drag_timer.isActive()
    finally:
        qtbot.mouseRelease(page_list.viewport(), Qt.LeftButton, modifier, pos=position)


def test_hold_preserves_multiple_selected_pages(page_list, qtbot, monkeypatch):
    for row in (0, 1):
        page_list.item(row).setSelected(True)
    captured = []
    monkeypatch.setattr(page_list, "_start_drag_for_selected_items",
                        lambda items, position: captured.extend(page_list.row(item) for item in items))
    position = page_list.visualItemRect(page_list.item(0)).center()
    qtbot.mousePress(page_list.viewport(), Qt.LeftButton, pos=position)
    try:
        qtbot.waitUntil(lambda: page_list._drag_mode_enabled, timeout=1000)
        move_while_pressed(page_list, position + QPoint(1, 0))
        assert captured == [0, 1]
    finally:
        qtbot.mouseRelease(page_list.viewport(), Qt.LeftButton, pos=position)
