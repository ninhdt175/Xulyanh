"""Run morphological watershed and connected components on Gonzalez & Woods textbook images."""

import argparse
import json
from pathlib import Path
import sys
import time

import imageio.v3 as iio
import numpy as np
from skimage.color import label2rgb
from skimage.filters import gaussian, threshold_otsu
from skimage.measure import label
from skimage.segmentation import watershed

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from watershed_btl.morphology import binary_closing, binary_opening, make_markers
from watershed_btl.pipeline import SegmentationResult, save_stage_images, _to_uint8


def segment_course_image(
    image: np.ndarray,
    invert: bool = False,
    sigma: float = 1.0,
    radius: int = 2,
    peak_fraction: float = 0.35,
) -> tuple[SegmentationResult, np.ndarray, dict]:
    """Run both Connected Components and Marker-controlled Watershed on a course image."""
    source = np.asarray(image, dtype=np.float64)
    if source.ndim == 3:
        source = source[..., 0]
    
    if invert:
        source = source.max() - source

    # Normalize to 0-255
    src_norm = ((source - source.min()) / (source.max() - source.min() + 1e-8) * 255.0).astype(np.uint8)

    # 1. Denoise and threshold
    denoised = gaussian(src_norm, sigma=sigma, preserve_range=True)
    cutoff = threshold_otsu(denoised)
    binary_mask = denoised > cutoff

    # 2. Mathematical morphology (Opening & Closing)
    binary_mask = binary_opening(binary_mask, radius=radius)
    binary_mask = binary_closing(binary_mask, radius=radius)

    # 3. Baseline Connected Components (fails when objects touch)
    t0_cc = time.perf_counter()
    cc_labels, num_cc = label(binary_mask, return_num=True)
    cc_time = time.perf_counter() - t0_cc

    # 4. Marker-controlled Watershed
    t0_ws = time.perf_counter()
    distance_map, markers = make_markers(binary_mask, peak_fraction=peak_fraction)
    num_markers = int(markers.max())
    ws_labels = watershed(-distance_map, markers, mask=binary_mask).astype(np.int32)
    num_ws = int(ws_labels.max())
    ws_time = time.perf_counter() - t0_ws

    overlay = label2rgb(ws_labels, image=src_norm, bg_label=0, kind="overlay")
    cc_overlay = label2rgb(cc_labels, image=src_norm, bg_label=0, kind="overlay")

    result = SegmentationResult(
        binary_mask=binary_mask,
        distance_map=distance_map,
        markers=markers,
        labels=ws_labels,
        overlay=overlay,
    )

    stats = {
        "num_connected_components": num_cc,
        "num_watershed_objects": num_ws,
        "num_markers": num_markers,
        "cc_runtime_seconds": cc_time,
        "ws_runtime_seconds": ws_time,
        "parameters": {
            "sigma": sigma,
            "radius": radius,
            "peak_fraction": peak_fraction,
            "cutoff": float(cutoff),
        },
    }

    return result, cc_overlay, stats


def main() -> None:
    parser = argparse.ArgumentParser(description="Segment course textbook images with Watershed & CC.")
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "artifacts" / "course_images",
        help="Directory where course experiment stage images and results will be saved.",
    )
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)

    # Course image paths
    course_dir = Path(r"D:\Dowload\Giáo trình\Kì 1 năm 4\Xử lý ảnh\Bài tập")
    ch10_dir = course_dir / "DIP3E_Original_Images_CH10"
    ch09_dir = course_dir / "DIP3E_Original_Images_CH09"

    datasets = [
        {
            "id": "blobs",
            "name": "FigP1036(blobs).tif",
            "path": ch10_dir / "FigP1036(blobs).tif",
            "description": "Các hình tròn dính liền nhau (Gonzalez & Woods Ch10)",
            "invert": False,
            "sigma": 1.0,
            "radius": 2,
            "peak_fraction": 0.35,
        },
        {
            "id": "polymersomes",
            "name": "Fig1039(a)(polymersomes).tif",
            "path": ch10_dir / "Fig1039(a)(polymersomes).tif",
            "description": "Túi tế bào polymersomes chụp hiển vi huỳnh quang (Gonzalez & Woods Ch10)",
            "invert": False,
            "sigma": 1.5,
            "radius": 2,
            "peak_fraction": 0.25,
        },
        {
            "id": "dowels",
            "name": "Fig0941(a)(wood_dowels).tif",
            "path": ch09_dir / "Fig0941(a)(wood_dowels).tif",
            "description": "Tiết diện các thanh gỗ tròn tiếp xúc nhau (Gonzalez & Woods Ch09)",
            "invert": False,
            "sigma": 1.5,
            "radius": 2,
            "peak_fraction": 0.40,
        },
    ]

    all_results = {}

    for item in datasets:
        img_path = item["path"]
        if not img_path.exists():
            print(f"[WARN] Image not found: {img_path}")
            continue

        raw_img = iio.imread(img_path)
        print(f"Processing {item['id']} ({item['name']}): shape={raw_img.shape}")

        result, cc_overlay, stats = segment_course_image(
            raw_img,
            invert=item["invert"],
            sigma=item["sigma"],
            radius=item["radius"],
            peak_fraction=item["peak_fraction"],
        )

        stem = item["id"]
        stage_paths = save_stage_images(result, raw_img, args.output, stem)

        # Also save CC overlay for visual ablation comparison
        cc_overlay_path = args.output / f"{stem}_cc_overlay.png"
        iio.imwrite(cc_overlay_path, _to_uint8(cc_overlay))
        stage_paths["cc_overlay"] = cc_overlay_path

        all_results[stem] = {
            "name": item["name"],
            "description": item["description"],
            "stats": stats,
            "paths": {k: str(v) for k, v in stage_paths.items()},
        }

        print(
            f"  -> CC objects: {stats['num_connected_components']} | "
            f"Watershed objects: {stats['num_watershed_objects']} (Markers: {stats['num_markers']})"
        )

    meta_file = args.output / "course_run.json"
    meta_file.write_text(json.dumps(all_results, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"\n[SUCCESS] Completed course images benchmark. Metadata saved to {meta_file}")


if __name__ == "__main__":
    main()
