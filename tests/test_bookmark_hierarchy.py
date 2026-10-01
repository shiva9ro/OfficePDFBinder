import fitz
import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QListWidgetItem

from OfficePDFBinder_Main import AppWorker


TOC = [[1, "Root", 1], [2, "Child", 1], [3, "Grandchild", 3],
       [2, "Sibling", 2], [1, "Other root", 4]]


def load_tree(window, source, order=(0, 1, 2, 3)):
    window.auto_bookmarks_enabled = False
    for page in order:
        item = QListWidgetItem()
        item.setData(Qt.UserRole, dict(type="pdf", path=str(source),
                     original_path=str(source), page_num=page, rotation=0))
        window.page_list_widget.addItem(item)
    worker = AppWorker("add_files")
    worker.signals.bookmarks_ready.connect(window._load_bookmarks_from_pdf)
    worker._run_add_files([str(source)])
    window._update_bookmark_tree()


def save_tree(window, output, selected=None):
    items = [window.page_list_widget.item(i).data(Qt.UserRole)
             for i in range(window.page_list_widget.count())
             if selected is None or i in selected]
    worker = AppWorker("merge_save")
    errors = []
    worker.signals.error.connect(lambda *args: errors.append(args))
    worker._run_merge_save(items, str(output),
                           bookmarks=window._prepare_bookmarks_for_export(items))
    assert not errors
    with fitz.open(output) as document:
        return document.get_toc()


def test_tree_roundtrip_preserves_same_page_nodes(main_window, pdf_factory, tmp_path):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source)
    root = main_window.bookmark_tree.topLevelItem(0)
    assert root.childCount() == 2
    assert root.child(0).child(0).text(0) == "Grandchild"
    assert save_tree(main_window, tmp_path / "saved.pdf") == TOC


@pytest.mark.parametrize("target,expected", [
    ("Child", [[1, "Root", 1], [2, "Grandchild", 3], [2, "Sibling", 2], [1, "Other root", 4]]),
    ("Grandchild", [[1, "Root", 1], [2, "Child", 1], [2, "Sibling", 2], [1, "Other root", 4]]),
    ("Root", [[1, "Child", 1], [2, "Grandchild", 3], [1, "Sibling", 2], [1, "Other root", 4]]),
])
def test_delete_node_promotes_children_without_restoring_source(
    main_window, pdf_factory, tmp_path, target, expected
):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source)
    item = main_window.bookmark_tree.findItems(target, Qt.MatchExactly | Qt.MatchRecursive)[0]
    main_window.bookmark_tree.setCurrentItem(item)
    main_window._delete_selected_bookmark()
    assert save_tree(main_window, tmp_path / "saved.pdf") == expected


def test_delete_all_stays_empty_in_full_and_selected_output(main_window, pdf_factory, tmp_path):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source)
    while main_window.bookmark_tree.topLevelItemCount():
        main_window.bookmark_tree.setCurrentItem(main_window.bookmark_tree.topLevelItem(0))
        main_window._delete_selected_bookmark()
    assert save_tree(main_window, tmp_path / "full.pdf") == []
    assert save_tree(main_window, tmp_path / "selected.pdf", [0, 2]) == []


def test_selected_output_promotes_orphans_as_siblings(main_window, pdf_factory, tmp_path):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source)
    assert save_tree(main_window, tmp_path / "selected.pdf", [1, 2]) == [
        [1, "Grandchild", 2], [1, "Sibling", 1]]


def test_reordered_pages_keep_tree_and_update_destinations(main_window, pdf_factory, tmp_path):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source, order=(2, 1, 0, 3))
    assert save_tree(main_window, tmp_path / "reordered.pdf") == [
        [1, "Root", 3], [2, "Child", 3], [3, "Grandchild", 1],
        [2, "Sibling", 2], [1, "Other root", 4]]


def test_worker_without_explicit_bookmarks_preserves_tree(pdf_factory, tmp_path):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    output = tmp_path / "saved.pdf"
    AppWorker("merge_save")._run_merge_save(
        [dict(type="pdf", path=str(source), original_path=str(source), page_num=i, rotation=0)
         for i in range(4)], str(output))
    with fitz.open(output) as document:
        assert document.get_toc() == TOC


@pytest.mark.parametrize("auto", [True, False])
def test_batch_keeps_same_page_parent_and_child(pdf_factory, tmp_path, auto):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    case = tmp_path / "input" / "case"
    case.mkdir(parents=True)
    (case / "tree.pdf").write_bytes(source.read_bytes())
    AppWorker("batch_merge_subfolders")._run_batch_merge_subfolders(
        str(case.parent), str(tmp_path / "output"), auto_bookmarks_enabled=auto)
    with fitz.open(tmp_path / "output" / "case.pdf") as document:
        assert document.get_toc() == ([[1, "tree", 1]] if auto else []) + TOC


def test_delete_undo_redo_preserves_tree(main_window, pdf_factory, tmp_path):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source)
    main_window._record_history_change(initial=True)
    main_window.bookmark_tree.setCurrentItem(main_window.bookmark_tree.topLevelItem(0).child(0))
    main_window._delete_selected_bookmark()
    main_window._undo()
    assert save_tree(main_window, tmp_path / "undo.pdf") == TOC
    main_window._redo()
    assert save_tree(main_window, tmp_path / "redo.pdf") == [
        [1, "Root", 1], [2, "Grandchild", 3], [2, "Sibling", 2], [1, "Other root", 4]]


def test_removed_parent_page_promotes_descendants(main_window, pdf_factory, tmp_path):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source)
    main_window.page_list_widget.takeItem(0)
    main_window._update_bookmark_tree()
    assert save_tree(main_window, tmp_path / "pruned.pdf") == [
        [1, "Sibling", 1], [1, "Grandchild", 2], [1, "Other root", 3]]


@pytest.mark.parametrize("auto", [False, True])
def test_delete_to_last_bookmark_by_mouse(main_window, pdf_factory, qtbot, tmp_path, auto):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source)
    main_window.auto_bookmarks_enabled = auto
    main_window.show()
    main_window.bookmark_dock.show()
    for remaining in range(len(TOC), 0, -1):
        tree = main_window.bookmark_tree
        item = tree.topLevelItem(0)
        qtbot.mouseClick(tree.viewport(), Qt.LeftButton,
                        pos=tree.visualItemRect(item).center())
        assert main_window.bookmark_delete_button.isEnabled(), remaining
        qtbot.mouseClick(main_window.bookmark_delete_button, Qt.LeftButton)
        assert len(main_window.bookmarks) == remaining - 1
    assert save_tree(main_window, tmp_path / "empty.pdf") == []


def test_bulk_bookmark_delete_and_undo(main_window, pdf_factory, qtbot, tmp_path):
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source)
    main_window._record_history_change(initial=True)
    main_window.show()
    main_window.bookmark_dock.show()
    tree = main_window.bookmark_tree
    for title in ("Root", "Child"):
        item = tree.findItems(title, Qt.MatchExactly | Qt.MatchRecursive)[0]
        qtbot.mouseClick(tree.viewport(), Qt.LeftButton, Qt.ControlModifier,
                        pos=tree.visualItemRect(item).center())
    assert len(tree.selectedItems()) == 2
    assert not main_window.bookmark_rename_button.isEnabled()
    qtbot.mouseClick(main_window.bookmark_delete_button, Qt.LeftButton)
    assert save_tree(main_window, tmp_path / "partial.pdf") == [
        [1, "Sibling", 2], [1, "Grandchild", 3], [1, "Other root", 4]]
    main_window._undo()
    assert save_tree(main_window, tmp_path / "undo.pdf") == TOC
    tree.setFocus()
    qtbot.keyClick(tree, Qt.Key_A, Qt.ControlModifier)
    assert len(tree.selectedItems()) == len(TOC)
    qtbot.mouseClick(main_window.bookmark_delete_button, Qt.LeftButton)
    assert save_tree(main_window, tmp_path / "empty.pdf") == []
    main_window._undo()
    assert save_tree(main_window, tmp_path / "restored.pdf") == TOC


def test_rename_uses_selection_after_ctrl_deselect(main_window, pdf_factory, qtbot, monkeypatch):
    from PySide6.QtWidgets import QInputDialog
    source = pdf_factory("tree.pdf", ["1", "2", "3", "4"], toc=TOC)
    load_tree(main_window, source)
    main_window.show()
    main_window.bookmark_dock.show()
    tree = main_window.bookmark_tree
    for title in ("Root", "Child", "Child"):
        item = tree.findItems(title, Qt.MatchExactly | Qt.MatchRecursive)[0]
        qtbot.mouseClick(tree.viewport(), Qt.LeftButton, Qt.ControlModifier,
                        pos=tree.visualItemRect(item).center())
    assert [item.text(0) for item in tree.selectedItems()] == ["Root"]
    monkeypatch.setattr(QInputDialog, "getText", lambda *a, **kw: ("Renamed root", True))
    qtbot.mouseClick(main_window.bookmark_rename_button, Qt.LeftButton)
    assert main_window.bookmarks[0]["title"] == "Renamed root"
    assert main_window.bookmarks[1]["title"] == "Child"
