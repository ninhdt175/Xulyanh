import numpy as np
import subprocess
import sys
from pathlib import Path

from watershed_btl.experiments import run_experiments
from watershed_btl.metrics import binary_metrics


def test_perfect_match_has_unit_scores():
    """Identical non-empty masks receive unit overlap and classification scores."""
    mask = np.array([[0, 1], [1, 0]], dtype=bool)

    scores = binary_metrics(mask, mask)

    assert scores == {
        "iou": 1.0,
        "dice": 1.0,
        "precision": 1.0,
        "recall": 1.0,
        "f1": 1.0,
    }


def test_empty_masks_receive_unit_scores():
    """Matching absence of foreground is a perfect binary prediction."""
    mask = np.zeros((2, 2), dtype=bool)

    scores = binary_metrics(mask, mask)

    assert scores == {
        "iou": 1.0,
        "dice": 1.0,
        "precision": 1.0,
        "recall": 1.0,
        "f1": 1.0,
    }


def test_missing_prediction_scores_zero_against_foreground_reference():
    """An empty prediction does not earn overlap or recall for true foreground."""
    prediction = np.zeros((2, 2), dtype=bool)
    reference = np.array([[0, 1], [1, 0]], dtype=bool)

    scores = binary_metrics(prediction, reference)

    assert scores == {
        "iou": 0.0,
        "dice": 0.0,
        "precision": 0.0,
        "recall": 0.0,
        "f1": 0.0,
    }


def test_overlap_scores_match_confusion_counts():
    """Metrics are computed from true-positive, false-positive and false-negative counts."""
    prediction = np.array([[1, 1], [0, 0]], dtype=bool)
    reference = np.array([[1, 0], [1, 0]], dtype=bool)

    scores = binary_metrics(prediction, reference)

    assert scores == {
        "iou": 1 / 3,
        "dice": 1 / 2,
        "precision": 1 / 2,
        "recall": 1 / 2,
        "f1": 1 / 2,
    }


def test_experiment_runner_writes_metrics_benchmark_and_summary(tmp_path):
    """A one-seed controlled run records all methods in reusable artifacts."""
    output_dir = tmp_path / "experiment"

    metrics = run_experiments(output_dir, seeds=[381])

    assert set(metrics["method"]) == {
        "full_watershed",
        "no_morphology",
        "connected_components",
    }
    assert set(metrics.columns) >= {
        "seed",
        "method",
        "iou",
        "dice",
        "precision",
        "recall",
        "f1",
        "count_error",
    }
    assert (output_dir / "metrics.csv").is_file()
    assert (output_dir / "benchmark.csv").is_file()
    assert (output_dir / "summary.png").is_file()


def test_experiment_cli_help_runs_from_project_root():
    """The documented direct-script invocation can import the src-layout package."""
    project_root = Path(__file__).parents[1]

    completed = subprocess.run(
        [sys.executable, "scripts/run_experiments.py", "--help"],
        cwd=project_root,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
