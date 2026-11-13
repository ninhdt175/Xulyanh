"""Binary segmentation quality metrics."""

import numpy as np


def _safe_divide(numerator: int, denominator: int, both_empty: bool) -> float:
    """Divide counts, treating only mutually empty masks as perfect."""
    if denominator == 0:
        return 1.0 if both_empty else 0.0
    return numerator / denominator


def binary_metrics(prediction: np.ndarray, reference: np.ndarray) -> dict[str, float]:
    """Return overlap and classification metrics for two binary masks.

    A metric with a zero denominator is one only when both masks are empty;
    otherwise it is zero.
    """
    predicted = np.asarray(prediction, dtype=bool)
    expected = np.asarray(reference, dtype=bool)
    if predicted.shape != expected.shape:
        raise ValueError("prediction and reference must have the same shape")

    true_positive = int(np.count_nonzero(predicted & expected))
    false_positive = int(np.count_nonzero(predicted & ~expected))
    false_negative = int(np.count_nonzero(~predicted & expected))
    predicted_positive = true_positive + false_positive
    reference_positive = true_positive + false_negative
    both_empty = predicted_positive == 0 and reference_positive == 0

    precision = _safe_divide(true_positive, predicted_positive, both_empty)
    recall = _safe_divide(true_positive, reference_positive, both_empty)
    dice = _safe_divide(
        2 * true_positive, 2 * true_positive + false_positive + false_negative, both_empty
    )
    iou = _safe_divide(
        true_positive, true_positive + false_positive + false_negative, both_empty
    )
    f1 = _safe_divide(
        2 * true_positive, 2 * true_positive + false_positive + false_negative, both_empty
    )
    return {
        "iou": iou,
        "dice": dice,
        "precision": precision,
        "recall": recall,
        "f1": f1,
    }
