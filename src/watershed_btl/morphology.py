"""Binary morphology and distance-transform marker helpers."""

import numpy as np
from scipy import ndimage as ndi
from skimage.morphology import disk, reconstruction


def _binary_image(image: np.ndarray) -> np.ndarray:
    """Validate a two-dimensional binary image and return it as Booleans."""
    array = np.asarray(image)
    if array.ndim != 2:
        raise ValueError("image must be a two-dimensional array")
    if array.dtype != np.bool_:
        raise ValueError("image must have Boolean dtype")
    return array


def _footprint(radius: int) -> np.ndarray:
    """Return the flat disk structuring element for a non-negative radius."""
    if not isinstance(radius, (int, np.integer)) or radius < 0:
        raise ValueError("radius must be a non-negative integer")
    return disk(radius).astype(bool)


def binary_erosion(image: np.ndarray, radius: int) -> np.ndarray:
    """Erode a binary image with a disk-shaped structuring element."""
    return ndi.binary_erosion(_binary_image(image), structure=_footprint(radius))


def binary_dilation(image: np.ndarray, radius: int) -> np.ndarray:
    """Dilate a binary image with a disk-shaped structuring element."""
    return ndi.binary_dilation(_binary_image(image), structure=_footprint(radius))


def binary_opening(image: np.ndarray, radius: int) -> np.ndarray:
    """Open a binary image with a disk-shaped structuring element."""
    return ndi.binary_opening(_binary_image(image), structure=_footprint(radius))


def binary_closing(image: np.ndarray, radius: int) -> np.ndarray:
    """Close a binary image with a disk-shaped structuring element."""
    return ndi.binary_closing(_binary_image(image), structure=_footprint(radius))


def reconstruct_by_dilation(marker: np.ndarray, mask: np.ndarray) -> np.ndarray:
    """Geodesically dilate ``marker`` until it reaches the Boolean ``mask``."""
    marker_array = _binary_image(marker)
    mask_array = _binary_image(mask)
    if marker_array.shape != mask_array.shape:
        raise ValueError("marker and mask must have the same shape")
    bounded_marker = np.logical_and(marker_array, mask_array)
    result = reconstruction(
        bounded_marker.astype(np.uint8), mask_array.astype(np.uint8), method="dilation"
    )
    return result.astype(bool)


def make_markers(mask: np.ndarray, peak_fraction: float) -> tuple[np.ndarray, np.ndarray]:
    """Create connected-component markers from high distance-transform peaks."""
    mask_array = _binary_image(mask)
    if not 0.0 < peak_fraction <= 1.0:
        raise ValueError("peak_fraction must be in the interval (0, 1]")

    distance_map = ndi.distance_transform_edt(mask_array).astype(np.float64, copy=False)
    if not mask_array.any():
        return distance_map, np.zeros(mask_array.shape, dtype=np.int32)

    peaks = mask_array & (distance_map >= peak_fraction * distance_map.max())
    markers, _ = ndi.label(peaks)
    return distance_map, markers.astype(np.int32, copy=False)
