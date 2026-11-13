import numpy as np

from watershed_btl.pipeline import save_stage_images, segment
from watershed_btl.synthetic import generate_synthetic_sample


def test_pipeline_returns_integer_labels_and_rgb_overlay():
    """Watershed output must cover the input and have an RGB visualization."""
    sample = generate_synthetic_sample(381, (160, 160), 7)

    result = segment(sample.image)

    assert result.labels.dtype.name == "int32"
    assert result.labels.shape == sample.image.shape
    assert result.labels.max() >= 1
    assert result.overlay.shape == (160, 160, 3)


def test_save_stage_images_writes_all_pngs_without_mutating_input(tmp_path):
    """Artifact export must create every named stage and preserve the source image."""
    sample = generate_synthetic_sample(381, (80, 80), 4)
    original = sample.image.copy()
    result = segment(sample.image)

    paths = save_stage_images(result, sample.image, tmp_path / "run", "sample")

    assert set(paths) == {
        "input",
        "binary_mask",
        "distance_map",
        "markers",
        "labels",
        "overlay",
    }
    assert all(path.suffix == ".png" and path.is_file() for path in paths.values())
    assert np.array_equal(sample.image, original)
