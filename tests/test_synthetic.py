from watershed_btl.synthetic import generate_synthetic_sample


def test_same_seed_produces_identical_image_and_labels():
    """A fixed seed must reproduce the same sample across calls."""
    a = generate_synthetic_sample(seed=381, shape=(128, 128), n_objects=6)
    b = generate_synthetic_sample(seed=381, shape=(128, 128), n_objects=6)

    assert a.image.shape == (128, 128)
    assert a.image.dtype.name == "uint8"
    assert (a.image == b.image).all()
    assert (a.instance_labels == b.instance_labels).all()
    assert a.instance_labels.max() >= 1
