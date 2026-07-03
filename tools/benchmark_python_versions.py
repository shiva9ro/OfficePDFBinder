#!/usr/bin/env python3
"""Compare Office PDF Binder workloads between two Python environments.

This benchmark is intentionally separate from pytest because timing thresholds
are unstable on shared CI runners. Lower elapsed time is better for every
reported metric.
"""

from __future__ import annotations

import argparse
import importlib.metadata
import json
import os
import re
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _package_version(distribution_name: str) -> str:
    try:
        return importlib.metadata.version(distribution_name)
    except importlib.metadata.PackageNotFoundError:
        return "not installed"


def _median_elapsed(operation, rounds: int) -> float:
    values = []
    for round_index in range(rounds):
        started = time.perf_counter()
        operation(round_index)
        values.append(time.perf_counter() - started)
    return statistics.median(values)


def _run_worker(rounds: int, pages: int, dpi: int) -> dict:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    os.environ.setdefault("OFFICEPDFBINDER_LANGUAGE", "ja")
    sys.path.insert(0, str(PROJECT_ROOT))

    import_started = time.perf_counter()
    import fitz
    from PySide6.QtCore import QCoreApplication

    from OfficePDFBinder_Main import AppWorker

    import_seconds = time.perf_counter() - import_started

    app = QCoreApplication.instance()
    if app is None:
        app = QCoreApplication([sys.argv[0]])

    with tempfile.TemporaryDirectory(prefix="officepdfbinder_benchmark_") as temp:
        temp_root = Path(temp)
        source_pdf = temp_root / "source.pdf"
        document = fitz.open()
        for page_number in range(pages):
            page = document.new_page(width=595, height=842)
            page.insert_text(
                (36, 72),
                f"Office PDF Binder benchmark page {page_number + 1}",
                fontsize=12,
            )
            page.draw_rect(fitz.Rect(36, 100, 559, 806), width=0.5)
        document.save(source_pdf)
        document.close()

        items_data = [
            {
                "type": "pdf",
                "path": str(source_pdf),
                "original_path": str(source_pdf),
                "page_num": page_number,
                "rotation": 0,
            }
            for page_number in range(pages)
        ]

        names = [
            f"申請_{index % 137:03d}_資料_{50000 - index:05d}.pdf"
            for index in range(50000)
        ]
        splitter = re.compile(r"(\d+)")

        def python_workload(_round_index: int):
            sorted(
                names,
                key=lambda value: [
                    int(part) if part.isdigit() else part
                    for part in splitter.split(value)
                ],
            )

        def merge_workload(round_index: int):
            output = temp_root / f"merged_{round_index}.pdf"
            worker = AppWorker("merge_save")
            result = worker._run_merge_save(
                items_data,
                str(output),
                bookmarks=None,
                show_outlines=False,
                emit_completion=False,
                emit_progress=False,
            )
            if not result or not output.is_file():
                raise RuntimeError("PDF merge benchmark did not create an output PDF")

        def export_workload(round_index: int):
            output_dir = temp_root / f"images_{round_index}"
            worker = AppWorker("export_images")
            worker._run_export_images(
                items_data,
                str(output_dir),
                dpi=dpi,
                image_format="JPEG",
            )
            image_count = len(list(output_dir.glob("*.jpg")))
            if image_count != pages:
                raise RuntimeError(
                    f"Image export benchmark created {image_count}/{pages} images"
                )

        metrics = {
            "import_seconds": import_seconds,
            "python_workload_seconds": _median_elapsed(python_workload, rounds),
            "pdf_merge_seconds": _median_elapsed(merge_workload, rounds),
            "image_export_seconds": _median_elapsed(export_workload, rounds),
        }

    return {
        "python": sys.version.split()[0],
        "executable": sys.executable,
        "packages": {
            "PyMuPDF": _package_version("PyMuPDF"),
            "PySide6": _package_version("PySide6"),
            "Pillow": _package_version("Pillow"),
            "pywin32": _package_version("pywin32"),
        },
        "metrics": metrics,
    }


def _run_sample(
    python_executable: Path,
    rounds: int,
    pages: int,
    dpi: int,
) -> dict:
    command = [
        str(python_executable),
        str(Path(__file__).resolve()),
        "--worker",
        "--rounds",
        str(rounds),
        "--pages",
        str(pages),
        "--dpi",
        str(dpi),
    ]
    environment = os.environ.copy()
    environment.setdefault("QT_QPA_PLATFORM", "offscreen")
    environment.setdefault("OFFICEPDFBINDER_LANGUAGE", "ja")

    started = time.perf_counter()
    completed = subprocess.run(
        command,
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    process_seconds = time.perf_counter() - started
    if completed.returncode != 0:
        raise RuntimeError(
            f"Benchmark failed: {python_executable}\n"
            f"stdout:\n{completed.stdout}\n"
            f"stderr:\n{completed.stderr}"
        )

    output_lines = [line for line in completed.stdout.splitlines() if line.strip()]
    if not output_lines:
        raise RuntimeError(f"Benchmark returned no JSON: {python_executable}")
    result = json.loads(output_lines[-1])
    result["metrics"]["process_total_seconds"] = process_seconds
    return result


def _summarize(samples: list[dict]) -> dict:
    metric_names = samples[0]["metrics"]
    return {
        "python": samples[0]["python"],
        "executable": samples[0]["executable"],
        "packages": samples[0]["packages"],
        "metrics": {
            name: statistics.median(sample["metrics"][name] for sample in samples)
            for name in metric_names
        },
    }


def _print_summary(labels: tuple[str, str], summaries: dict[str, dict]):
    baseline_label, candidate_label = labels
    baseline = summaries[baseline_label]
    candidate = summaries[candidate_label]

    print()
    print("Environment versions")
    for label in labels:
        summary = summaries[label]
        packages = ", ".join(
            f"{name} {version}" for name, version in summary["packages"].items()
        )
        print(f"- {label}: Python {summary['python']} ({packages})")

    print()
    print("Median elapsed time in seconds (lower is better)")
    print(
        f"{'Metric':30} {baseline_label:>14} {candidate_label:>14} "
        f"{'Change':>10}"
    )
    print("-" * 72)
    for metric_name, baseline_seconds in baseline["metrics"].items():
        candidate_seconds = candidate["metrics"][metric_name]
        change = (
            ((candidate_seconds / baseline_seconds) - 1) * 100
            if baseline_seconds
            else 0
        )
        print(
            f"{metric_name:30} {baseline_seconds:14.4f} "
            f"{candidate_seconds:14.4f} {change:9.1f}%"
        )
    print()
    print("Negative Change means the candidate was faster.")


def _parse_args():
    parser = argparse.ArgumentParser(
        description="Compare Office PDF Binder performance between Python environments."
    )
    parser.add_argument("--baseline", type=Path)
    parser.add_argument("--candidate", type=Path)
    parser.add_argument("--baseline-label", default="Python 3.12")
    parser.add_argument("--candidate-label", default="Python 3.13")
    parser.add_argument("--samples", type=int, default=3)
    parser.add_argument("--warmups", type=int, default=1)
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--pages", type=int, default=20)
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--json-output", type=Path)
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    if args.worker:
        print(
            json.dumps(
                _run_worker(args.rounds, args.pages, args.dpi),
                ensure_ascii=False,
            )
        )
        return 0

    if args.baseline is None or args.candidate is None:
        raise SystemExit("--baseline and --candidate are required")
    for executable in (args.baseline, args.candidate):
        if not executable.is_file():
            raise SystemExit(f"Python executable does not exist: {executable}")
    if args.samples < 1 or args.warmups < 0 or args.rounds < 1:
        raise SystemExit("samples and rounds must be positive; warmups cannot be negative")

    labels = (args.baseline_label, args.candidate_label)
    executables = {
        args.baseline_label: args.baseline.resolve(),
        args.candidate_label: args.candidate.resolve(),
    }
    results = {label: [] for label in labels}
    total_iterations = args.warmups + args.samples

    for iteration in range(total_iterations):
        order = labels if iteration % 2 == 0 else tuple(reversed(labels))
        for label in order:
            phase = "warmup" if iteration < args.warmups else "sample"
            number = (
                iteration + 1
                if phase == "warmup"
                else iteration - args.warmups + 1
            )
            print(f"[{phase} {number}] {label}", flush=True)
            result = _run_sample(
                executables[label],
                args.rounds,
                args.pages,
                args.dpi,
            )
            if phase == "sample":
                results[label].append(result)

    summaries = {label: _summarize(results[label]) for label in labels}
    _print_summary(labels, summaries)

    if args.json_output:
        output = {
            "settings": {
                "samples": args.samples,
                "warmups": args.warmups,
                "rounds": args.rounds,
                "pages": args.pages,
                "dpi": args.dpi,
            },
            "summary": summaries,
            "samples": results,
        }
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(
            json.dumps(output, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"JSON: {args.json_output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
