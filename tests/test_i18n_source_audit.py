import ast
from pathlib import Path
import re


JAPANESE_PATTERN = re.compile(r"[ぁ-んァ-ヶ一-龠]")
USER_FACING_CALLS = {
    "QAction",
    "QCheckBox",
    "QDockWidget",
    "QGroupBox",
    "QInputDialog.getText",
    "QLabel",
    "QMessageBox.critical",
    "QMessageBox.information",
    "QMessageBox.question",
    "QMessageBox.warning",
    "QProgressDialog",
    "QPushButton",
    "error.emit",
    "finished.emit",
    "getExistingDirectory",
    "getOpenFileNames",
    "getSaveFileName",
    "non_cancellable_started.emit",
    "progress.emit",
    "setCancelButtonText",
    "setLabelText",
    "setPlaceholderText",
    "setText",
    "setToolTip",
    "setWindowTitle",
}


def call_name(node):
    parts = []
    while isinstance(node, ast.Attribute):
        parts.append(node.attr)
        node = node.value
    if isinstance(node, ast.Name):
        parts.append(node.id)
    return ".".join(reversed(parts))


def japanese_outside_translation(node):
    if isinstance(node, ast.Call) and call_name(node.func) == "translate":
        return []
    found = []
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        if JAPANESE_PATTERN.search(node.value):
            found.append(node.value)
    for child in ast.iter_child_nodes(node):
        found.extend(japanese_outside_translation(child))
    return found


def test_user_facing_calls_do_not_contain_untranslated_japanese():
    source_path = Path(__file__).parents[1] / "OfficePDFBinder_Main.py"
    tree = ast.parse(source_path.read_text(encoding="utf-8"))
    violations = []

    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        name = call_name(node.func)
        if not any(name == target or name.endswith(f".{target}") for target in USER_FACING_CALLS):
            continue
        for argument in [*node.args, *[keyword.value for keyword in node.keywords]]:
            for text in japanese_outside_translation(argument):
                violations.append((node.lineno, name, text))

    assert violations == []
