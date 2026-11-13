"""Command-line entry point for reproducible watershed experiments."""

import argparse
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from watershed_btl.experiments import run_experiments


DEFAULT_SEEDS = [381, 633, 731, 940, 101]


def main() -> None:
    """Parse arguments and run the fixed-seed benchmark suite."""
    parser = argparse.ArgumentParser(description="Run watershed ablation experiments.")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("artifacts/run_001"),
        help="Directory that will receive CSV, PNG and stage-image artifacts.",
    )
    arguments = parser.parse_args()
    metrics = run_experiments(arguments.output, seeds=DEFAULT_SEEDS)
    print(f"Wrote {len(metrics)} metric rows to {arguments.output}")


if __name__ == "__main__":
    main()
