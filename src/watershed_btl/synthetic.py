"""Deterministic synthetic microscopy-like data for segmentation experiments."""

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class SyntheticSample:
    """A grayscale input image paired with per-instance ground-truth labels."""

    image: np.ndarray
    instance_labels: np.ndarray


def generate_synthetic_sample(
    seed: int,
    shape: tuple[int, int] = (512, 512),
    n_objects: int = 18,
) -> SyntheticSample:
    """Generate a repeatable noisy image with overlapping elliptical objects."""
    if len(shape) != 2 or any(dimension <= 0 for dimension in shape):
        raise ValueError("shape must contain two positive dimensions")
    if n_objects <= 0:
        raise ValueError("n_objects must be positive")

    rng = np.random.default_rng(seed)
    height, width = shape
    rows, columns = np.ogrid[:height, :width]
    labels = np.zeros(shape, dtype=np.int32)

    minimum_radius = max(3, min(height, width) // 24)
    maximum_radius = max(minimum_radius + 1, min(height, width) // 9)
    for label in range(1, n_objects + 1):
        center_row = rng.uniform(0, height)
        center_column = rng.uniform(0, width)
        row_radius = rng.integers(minimum_radius, maximum_radius + 1)
        column_radius = rng.integers(minimum_radius, maximum_radius + 1)
        ellipse = (
            ((rows - center_row) / row_radius) ** 2
            + ((columns - center_column) / column_radius) ** 2
            <= 1
        )
        labels[ellipse] = label

    row_gradient = np.linspace(0.0, 22.0, height, dtype=np.float64)[:, None]
    column_gradient = np.linspace(0.0, 14.0, width, dtype=np.float64)[None, :]
    background = 32.0 + row_gradient + column_gradient
    intensities = background + rng.normal(0.0, 5.0, size=shape)

    foreground = labels > 0
    object_brightness = rng.uniform(105.0, 175.0, size=n_objects + 1)
    intensities[foreground] = (
        background[foreground]
        + object_brightness[labels[foreground]]
        + rng.normal(0.0, 7.0, size=foreground.sum())
    )
    image = np.clip(intensities, 0, 255).astype(np.uint8)
    return SyntheticSample(image=image, instance_labels=labels)
