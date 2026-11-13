"""Reproducible segmentation ablations and timing benchmarks."""

from pathlib import Path
from time import perf_counter

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import ndimage as ndi
from skimage.filters import gaussian, threshold_otsu
from skimage.segmentation import watershed

from watershed_btl.metrics import binary_metrics
from watershed_btl.morphology import binary_closing, binary_opening, make_markers
from watershed_btl.pipeline import SegmentationResult, save_stage_images, segment
from watershed_btl.synthetic import generate_synthetic_sample


METHODS = ("full_watershed", "no_morphology", "connected_components")


def _binary_mask(image: np.ndarray, radius: int) -> np.ndarray:
    """Apply the pipeline preprocessing through binary morphology."""
    denoised = gaussian(image, sigma=1.0, preserve_range=True)
    mask = denoised > threshold_otsu(denoised)
    return binary_closing(binary_opening(mask, radius=radius), radius=radius)


def _connected_components(image: np.ndarray) -> np.ndarray:
    """Label the preprocessed foreground directly, without a watershed split."""
    labels, _ = ndi.label(_binary_mask(image, radius=2))
    return labels.astype(np.int32, copy=False)


def _run_method(image: np.ndarray, method: str) -> tuple[np.ndarray, SegmentationResult | None, float]:
    """Run one named condition and return labels, full result where available, and time."""
    started = perf_counter()
    if method == "full_watershed":
        result = segment(image)
        labels = result.labels
    elif method == "no_morphology":
        result = segment(image, radius=0)
        labels = result.labels
    elif method == "connected_components":
        result = None
        labels = _connected_components(image)
    else:
        raise ValueError(f"unsupported method: {method}")
    return labels, result, perf_counter() - started


def _summary_figure(metrics: pd.DataFrame, benchmark: pd.DataFrame, destination: Path) -> None:
    """Render compact quality and runtime comparisons from measured data."""
    method_order = list(METHODS)
    metric_means = metrics.groupby("method")[["iou", "dice", "f1"]].mean().reindex(method_order)
    timing_means = benchmark.groupby("method")["elapsed_seconds"].mean().reindex(method_order)

    figure, axes = plt.subplots(1, 2, figsize=(12, 4.8), constrained_layout=True)
    metric_means.plot.bar(ax=axes[0], ylim=(0, 1), rot=15)
    axes[0].set_title("Mean binary-mask quality")
    axes[0].set_xlabel("")
    axes[0].set_ylabel("Score")
    axes[0].legend(title="Metric")
    timing_means.plot.bar(ax=axes[1], color="#4c78a8", rot=15)
    axes[1].set_title("Mean elapsed runtime")
    axes[1].set_xlabel("")
    axes[1].set_ylabel("Seconds")
    figure.savefig(destination, dpi=180)
    plt.close(figure)


def run_experiments(output_dir: Path, seeds: list[int]) -> pd.DataFrame:
    """Run full and ablated segmentation conditions and save reusable artifacts.

    ``metrics.csv`` stores segmentation quality per seed and method,
    ``benchmark.csv`` contains separately measured elapsed runtimes, and
    ``summary.png`` visualizes their means.
    """
    if not seeds:
        raise ValueError("seeds must contain at least one value")

    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    metrics_rows: list[dict[str, float | int | str]] = []
    benchmark_rows: list[dict[str, float | int | str]] = []
    for seed in seeds:
        sample = generate_synthetic_sample(seed)
        reference = sample.instance_labels > 0
        reference_count = int(sample.instance_labels.max())
        for method in METHODS:
            labels, result, elapsed_seconds = _run_method(sample.image, method)
            predicted = labels > 0
            scores = binary_metrics(predicted, reference)
            predicted_count = int(labels.max())
            metrics_rows.append(
                {
                    "seed": seed,
                    "method": method,
                    **scores,
                    "predicted_instance_count": predicted_count,
                    "reference_instance_count": reference_count,
                    "count_error": abs(predicted_count - reference_count),
                }
            )
            benchmark_rows.append(
                {"seed": seed, "method": method, "elapsed_seconds": elapsed_seconds}
            )
            if result is not None:
                save_stage_images(result, sample.image, destination / "stages", f"{method}_seed_{seed}")

    metrics = pd.DataFrame(metrics_rows)
    benchmark = pd.DataFrame(benchmark_rows)
    metrics.to_csv(destination / "metrics.csv", index=False)
    benchmark.to_csv(destination / "benchmark.csv", index=False)
    _summary_figure(metrics, benchmark, destination / "summary.png")
    return metrics
