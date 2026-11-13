"""Smoke coverage for the real-image demonstration command."""

from pathlib import Path
import json
import os
import subprocess
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_real_image_command_exposes_help():
    """Removing the CLI entry point must make its public help command fail."""
    completed = subprocess.run(
        [sys.executable, "scripts/run_real_image.py", "--help"],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
    assert "--output" in completed.stdout


def test_real_image_command_writes_named_stages_and_parameter_record(tmp_path):
    """Dropping a stage or its recorded segmentation parameters must fail this test."""
    output_dir = tmp_path / "real"
    completed = subprocess.run(
        [sys.executable, "scripts/run_real_image.py", "--output", str(output_dir)],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
    assert {
        path.name for path in output_dir.glob("*.png")
    } == {
        "real_input.png",
        "real_binary_mask.png",
        "real_distance_map.png",
        "real_markers.png",
        "real_labels.png",
        "real_overlay.png",
    }
    metadata = json.loads((output_dir / "real_run.json").read_text(encoding="utf-8"))
    assert metadata["parameters"] == {
        "threshold": None,
        "radius": 2,
        "peak_fraction": 0.45,
    }


def test_real_image_command_uses_local_fallback_when_coins_import_fails(tmp_path):
    """A missing optional coins loader must still yield a recorded local run."""
    (tmp_path / "sitecustomize.py").write_text(
        """import builtins
original_import = builtins.__import__
def guarded_import(name, globals=None, locals=None, fromlist=(), level=0):
    if name == 'skimage' and 'data' in fromlist:
        raise ImportError('simulated unavailable coins dataset')
    return original_import(name, globals, locals, fromlist, level)
builtins.__import__ = guarded_import
""",
        encoding="utf-8",
    )
    output_dir = tmp_path / "fallback"
    environment = os.environ.copy()
    environment["PYTHONPATH"] = str(tmp_path) + os.pathsep + environment.get(
        "PYTHONPATH", ""
    )

    completed = subprocess.run(
        [sys.executable, "scripts/run_real_image.py", "--output", str(output_dir)],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
    metadata = json.loads((output_dir / "real_run.json").read_text(encoding="utf-8"))
    assert metadata["image_source"] == (
        "watershed_btl.synthetic.generate_synthetic_sample"
    )
