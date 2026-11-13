"""Run the watershed pipeline on scikit-image's bundled coins photograph."""

import argparse
import json
from pathlib import Path
import sys

import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from watershed_btl.config import DEFAULT_SEED
from watershed_btl.pipeline import save_stage_images, segment
from watershed_btl.synthetic import generate_synthetic_sample


def main() -> None:
    """Parse CLI options, save all stage images, and record run provenance."""
    parser = argparse.ArgumentParser(
        description="Segment scikit-image's coins photograph with the watershed pipeline."
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("artifacts/real"),
        help="Directory that will receive real_input.png, stage images, and real_run.json.",
    )
    arguments = parser.parse_args()

    try:
        from skimage import data

        image = np.asarray(data.coins())
        image_source = "skimage.data.coins"
    except (ImportError, OSError):
        image = generate_synthetic_sample(DEFAULT_SEED).image
        image_source = "watershed_btl.synthetic.generate_synthetic_sample"

    result = segment(image)
    paths = save_stage_images(result, image, arguments.output, "real")
    metadata = {
        "image_source": image_source,
        "parameters": {"threshold": None, "radius": 2, "peak_fraction": 0.45},
        "stage_images": {name: str(path) for name, path in paths.items()},
    }
    metadata_path = arguments.output / "real_run.json"
    metadata_path.write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {len(paths)} stage images and {metadata_path}")


if __name__ == "__main__":
    main()
