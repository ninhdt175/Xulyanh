"""Run automated headless Fiji/ImageJ benchmark on synthetic seeds and course images."""

import csv
from pathlib import Path
import subprocess
import sys
import time

import imageio.v3 as iio
import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

SEEDS = (381, 633, 731, 940, 101)
from watershed_btl.metrics import binary_metrics
from watershed_btl.synthetic import generate_synthetic_sample


def main() -> None:
    ij_dir = Path(r"D:\Dowload\Giáo trình\Kì 1 năm 4\Xử lý ảnh\Bài tập\ij153-win-java8\ImageJ")
    java_exe = ij_dir / "jre" / "bin" / "java.exe"
    ij_jar = ij_dir / "ij.jar"

    if not java_exe.exists() or not ij_jar.exists():
        print(f"[ERROR] ImageJ runtime not found at {ij_dir}")
        return

    output_dir = PROJECT_ROOT / "artifacts" / "fiji_benchmarks"
    output_dir.mkdir(parents=True, exist_ok=True)

    stages_dir = PROJECT_ROOT / "artifacts" / "run_001" / "stages"
    stages_dir.mkdir(parents=True, exist_ok=True)

    print("=== RUNNING FIJI / IMAGEJ AUTOMATED BENCHMARK ===")
    results = []

    for seed in SEEDS:
        sample = generate_synthetic_sample(seed)
        input_path = stages_dir / f"full_watershed_seed_{seed}_input.png"
        if not input_path.exists():
            iio.imwrite(input_path, sample.image)

        fiji_mask_path = output_dir / f"fiji_seed_{seed}_mask.png"
        fiji_csv_path = output_dir / f"fiji_seed_{seed}_results.csv"

        macro_lines = [
            f'open("{input_path.as_posix()}");',
            'run("8-bit");',
            'setAutoThreshold("Otsu dark");',
            'run("Convert to Mask");',
            'run("Open");',
            'run("Close");',
            'run("Watershed");',
            f'saveAs("PNG", "{fiji_mask_path.as_posix()}");',
            'run("Analyze Particles...", "size=0-Infinity show=Nothing display clear");',
            f'saveAs("Results", "{fiji_csv_path.as_posix()}");',
            'close();',
        ]
        macro_file = ij_dir / f"macro_seed_{seed}.ijm"
        macro_file.write_text("\n".join(macro_lines) + "\n", encoding="utf-8")

        t0 = time.perf_counter()
        proc = subprocess.run(
            [str(java_exe), "-jar", "ij.jar", "-batch", macro_file.name],
            cwd=str(ij_dir),
            capture_output=True,
            text=True,
        )
        elapsed = time.perf_counter() - t0
        macro_file.unlink(missing_ok=True)

        if proc.returncode != 0:
            print(f"Seed {seed}: Fiji error: {proc.stderr}")
            continue

        # Read results
        particle_count = 0
        if fiji_csv_path.exists():
            with open(fiji_csv_path, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)
                particle_count = max(0, len(rows) - 1)

        # Compute metrics against ground truth
        pred_mask = iio.imread(fiji_mask_path) > 128
        gt_mask = sample.binary_mask

        m = binary_metrics(pred_mask, gt_mask)
        m_iou = m["iou"]
        m_dice = m["dice"]
        m_prec = m["precision"]
        m_rec = m["recall"]
        m_f1 = m["f1"]
        m_cerr = abs(particle_count - sample.instance_count)

        print(
            f"Seed {seed:3d}: IoU={m_iou:.4f}, Dice={m_dice:.4f}, "
            f"Fiji Objects={particle_count:2d} (GT={sample.instance_count:2d}, Err={m_cerr:.1f}), "
            f"Time={elapsed*1000:.1f}ms"
        )

        results.append(
            {
                "seed": seed,
                "iou": m_iou,
                "dice": m_dice,
                "precision": m_prec,
                "recall": m_rec,
                "f1": m_f1,
                "predicted_instances": particle_count,
                "ground_truth_instances": sample.instance_count,
                "count_error": m_cerr,
                "elapsed_seconds": elapsed,
            }
        )

    # Save summary CSV
    summary_path = output_dir / "fiji_metrics.csv"
    with open(summary_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(results[0].keys()))
        writer.writeheader()
        writer.writerows(results)

    # Print averages
    mean_iou = np.mean([r["iou"] for r in results])
    mean_dice = np.mean([r["dice"] for r in results])
    mean_prec = np.mean([r["precision"] for r in results])
    mean_rec = np.mean([r["recall"] for r in results])
    mean_count = np.mean([r["predicted_instances"] for r in results])
    mean_err = np.mean([r["count_error"] for r in results])
    mean_time = np.mean([r["elapsed_seconds"] for r in results])

    print("\n--- FIJI / IMAGEJ 5-SEED AVERAGE ---")
    print(
        f"IoU: {mean_iou:.4f} | Dice: {mean_dice:.4f} | Prec: {mean_prec:.4f} | "
        f"Rec: {mean_rec:.4f} | Mean Objects: {mean_count:.1f} | Mean Count Error: {mean_err:.1f} | "
        f"Time: {mean_time*1000:.1f}ms"
    )
    print(f"Results written to {summary_path}")


if __name__ == "__main__":
    main()
