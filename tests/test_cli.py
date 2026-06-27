from pathlib import Path

import fitz
import pytest

import OfficePDFBinder_Main as app_module


@pytest.fixture(autouse=True)
def use_shared_qt_application(qapp):
    """後続GUIテストと同じQApplicationをCLIテストでも使用する。"""
    return qapp


def test_cli_invalid_arguments_exit_with_code_2():
    with pytest.raises(SystemExit) as exc_info:
        app_module.run_cli_batch(["--batch-subfolders"])

    assert exc_info.value.code == 2


def test_cli_passes_all_options_to_existing_batch_worker(
    tmp_path, monkeypatch, capsys
):
    captured = {}

    def fake_batch(self, **kwargs):
        captured.update(kwargs)
        return {
            "status": "completed",
            "message": "",
            "counts": {"Success": 1, "Warning": 0, "Skipped": 0, "Error": 0},
            "log_path": str(tmp_path / "output" / "batch_log.csv"),
            "input_root": kwargs["input_root"],
            "output_root": kwargs["output_root"],
        }

    monkeypatch.setattr(
        app_module.AppWorker,
        "_run_batch_merge_subfolders",
        fake_batch,
    )

    exit_code = app_module.run_cli_batch(
        [
            "--batch-subfolders",
            "--input-root",
            str(tmp_path / "input"),
            "--output-root",
            str(tmp_path / "output"),
            "--no-auto-bookmarks",
            "--no-show-bookmarks-on-open",
            "--remove-pdf-annotations",
            "--suppress-office-markup",
            "--disable-image-upscaling",
            "--overwrite",
        ]
    )

    assert exit_code == 0
    assert captured == {
        "input_root": str(tmp_path / "input"),
        "output_root": str(tmp_path / "output"),
        "overwrite": True,
        "continue_on_error": True,
        "auto_bookmarks_enabled": False,
        "show_bookmarks_on_open": False,
        "remove_pdf_annotations": True,
        "disable_image_upscaling": True,
        "suppress_office_markup": True,
    }
    assert "Success: 1" in capsys.readouterr().out


def test_cli_creates_pdf_and_utf8_sig_log_for_general_subfolder_name(
    pdf_factory, tmp_path
):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    case_dir = input_root / "任意の案件名"
    case_dir.mkdir(parents=True)
    source = pdf_factory("source.pdf", ["BODY"])
    (case_dir / "001.pdf").write_bytes(source.read_bytes())

    exit_code = app_module.run_cli_batch(
        [
            "--batch-subfolders",
            "--input-root",
            str(input_root),
            "--output-root",
            str(output_root),
        ]
    )

    output_pdf = output_root / "任意の案件名.pdf"
    assert exit_code == 0
    assert output_pdf.exists()
    with fitz.open(output_pdf) as document:
        assert "BODY" in document[0].get_text()
    log_path = next(output_root.glob("batch_log_*.csv"))
    assert log_path.read_bytes().startswith(b"\xef\xbb\xbf")


def test_cli_returns_code_3_when_input_root_does_not_exist(tmp_path):
    exit_code = app_module.run_cli_batch(
        [
            "--batch-subfolders",
            "--input-root",
            str(tmp_path / "missing"),
            "--output-root",
            str(tmp_path / "output"),
        ]
    )

    assert exit_code == 3


def test_cli_returns_code_4_when_input_has_no_subfolders(tmp_path):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    input_root.mkdir()

    exit_code = app_module.run_cli_batch(
        [
            "--batch-subfolders",
            "--input-root",
            str(input_root),
            "--output-root",
            str(output_root),
        ]
    )

    assert exit_code == 4
    assert next(output_root.glob("batch_log_*.csv")).exists()


def test_cli_returns_code_5_when_log_cannot_be_written(
    pdf_factory, tmp_path, monkeypatch
):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    case_dir = input_root / "case"
    case_dir.mkdir(parents=True)
    source = pdf_factory("source.pdf", ["BODY"])
    (case_dir / "001.pdf").write_bytes(source.read_bytes())

    def fail_log(_self, log_path, _rows):
        raise app_module.BatchLogWriteError(f"cannot write {Path(log_path).name}")

    monkeypatch.setattr(app_module.AppWorker, "_write_batch_log", fail_log)

    exit_code = app_module.run_cli_batch(
        [
            "--batch-subfolders",
            "--input-root",
            str(input_root),
            "--output-root",
            str(output_root),
        ]
    )

    assert exit_code == 5


def test_cli_returns_code_1_when_batch_is_not_fully_successful(
    tmp_path, monkeypatch
):
    result_name = "Error"

    def fake_batch(self, **kwargs):
        counts = {"Success": 0, "Warning": 0, "Skipped": 0, "Error": 0}
        counts[result_name] = 1
        return {
            "status": "completed",
            "message": "",
            "counts": counts,
            "log_path": str(tmp_path / "output" / "batch_log.csv"),
            "input_root": kwargs["input_root"],
            "output_root": kwargs["output_root"],
        }

    monkeypatch.setattr(
        app_module.AppWorker,
        "_run_batch_merge_subfolders",
        fake_batch,
    )

    for result_name in ("Warning", "Skipped", "Error"):
        exit_code = app_module.run_cli_batch(
            [
                "--batch-subfolders",
                "--input-root",
                str(tmp_path / "input"),
                "--output-root",
                str(tmp_path / "output"),
            ]
        )
        assert exit_code == 1


def test_cli_returns_code_9_for_unexpected_error(tmp_path, monkeypatch):
    def fail_batch(self, **_kwargs):
        raise RuntimeError("unexpected")

    monkeypatch.setattr(
        app_module.AppWorker,
        "_run_batch_merge_subfolders",
        fail_batch,
    )

    exit_code = app_module.run_cli_batch(
        [
            "--batch-subfolders",
            "--input-root",
            str(tmp_path / "input"),
            "--output-root",
            str(tmp_path / "output"),
        ]
    )

    assert exit_code == 9
