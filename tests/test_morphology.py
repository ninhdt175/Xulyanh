import numpy as np

from watershed_btl.morphology import (
    binary_closing,
    binary_dilation,
    binary_erosion,
    binary_opening,
    make_markers,
    reconstruct_by_dilation,
)


def test_cross_dilates_to_full_three_by_three_square():
    """A radius-one disk should fill the diagonal corners around a cross."""
    cross = np.array([[0, 1, 0], [1, 1, 1], [0, 1, 0]], dtype=bool)

    assert binary_dilation(cross, radius=1).all()


def test_cross_erodes_to_its_center_pixel():
    """Erosion must retain only positions containing the full disk footprint."""
    cross = np.array([[0, 1, 0], [1, 1, 1], [0, 1, 0]], dtype=bool)
    expected = np.array([[0, 0, 0], [0, 1, 0], [0, 0, 0]], dtype=bool)

    assert np.array_equal(binary_erosion(cross, radius=1), expected)


def test_opening_removes_an_isolated_pixel():
    """Opening should remove foreground smaller than the structuring element."""
    image = np.zeros((5, 5), dtype=bool)
    image[2, 2] = True

    assert not binary_opening(image, radius=1).any()


def test_closing_fills_a_single_pixel_hole():
    """Closing should fill a hole enclosed by a radius-one disk."""
    image = np.zeros((7, 7), dtype=bool)
    image[1:6, 1:6] = True
    image[3, 3] = False
    expected = np.zeros((7, 7), dtype=bool)
    expected[1:6, 1:6] = True

    assert np.array_equal(binary_closing(image, radius=1), expected)


def test_reconstruction_grows_marker_only_within_mask():
    """Geodesic dilation must stop at the mask boundary."""
    marker = np.zeros((5, 5), dtype=bool)
    marker[2, 2] = True
    mask = np.zeros((5, 5), dtype=bool)
    mask[1:4, 1:4] = True

    assert np.array_equal(reconstruct_by_dilation(marker, mask), mask)


def test_markers_are_inside_foreground():
    """Each separated object should receive one marker wholly inside the mask."""
    mask = np.zeros((21, 21), dtype=bool)
    mask[3:9, 3:9] = True
    mask[12:18, 12:18] = True

    distance_map, markers = make_markers(mask, peak_fraction=0.45)

    assert distance_map.dtype == np.float64
    assert markers.dtype == np.int32
    assert np.all(~(markers > 0) | mask)
    assert markers.max() == 2
