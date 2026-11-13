# Morphological Watershed Coursework Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Deliver a reproducible morphology and marker-controlled watershed study with Python, Fiji/MorphoLibJ reference workflow, DOCX report, and Canva-compatible slide deck.

**Architecture:** A small Python package creates deterministic synthetic images and labels, executes a configurable segmentation pipeline, calculates metrics, and exports figures/tables. The Fiji macro consumes exported inputs and records comparable output. Report and slides are generated only after the experiment artifacts are final, so no figure or metric is invented.

**Tech Stack:** Python 3.11+, NumPy, SciPy, scikit-image, pandas, matplotlib, python-docx, python-pptx, pytest; Fiji/ImageJ + MorphoLibJ; Pandoc is not required.

**Spec:** `docs/superpowers/specs/2026-09-18-morphological-watershed-design.md`

## Global Constraints

- Use the hybrid data design: labelled synthetic images for primary metrics and a small public real image for qualitative evaluation.
- Set and record a fixed random seed; never overwrite source data.
- Export every report figure/table from a reproducible pipeline run.
- Treat Fiji measurements as blank/reproducible until Fiji and MorphoLibJ are actually run; do not fabricate them.
- DOCX: Times New Roman 14 pt; body margins top/bottom/left 3 cm, right 2 cm; cover has no number; TOC begins page 1 in a centred footer.
- Create a PPTX with editable text/shapes and approximately 10 slides for Canva import.
- No Git repository currently exists; omit commit steps unless one is initialized later.

---

## Planned file structure

```text
requirements.txt
README.md
src/watershed_btl/__init__.py
src/watershed_btl/config.py
src/watershed_btl/synthetic.py
src/watershed_btl/morphology.py
src/watershed_btl/pipeline.py
src/watershed_btl/metrics.py
src/watershed_btl/experiments.py
scripts/run_experiments.py
fiji/morphological_watershed.ijm
fiji/README.md
tests/test_synthetic.py
tests/test_morphology.py
tests/test_metrics.py
tests/test_pipeline.py
artifacts/ (generated, ignored by no tool)
report/generate_report.py
report/Bao_cao_BTL_Watershed.docx (generated)
slides/generate_slides.py
slides/BTL_Watershed_Canva.pptx (generated)
```

### Task 1: Project setup and deterministic synthetic data

**Files:**
- Create: `requirements.txt`, `src/watershed_btl/__init__.py`, `src/watershed_btl/config.py`, `src/watershed_btl/synthetic.py`, `tests/test_synthetic.py`, `README.md`

**Interfaces:**
- Produces `SyntheticSample(image: np.ndarray, instance_labels: np.ndarray)` and `generate_synthetic_sample(seed: int, shape: tuple[int, int] = (512, 512), n_objects: int = 18) -> SyntheticSample`.
- `instance_labels` is `int32`, background is zero, and every foreground instance is a positive integer.

- [ ] **Step 1: Write the failing deterministic-data test.**

```python
from watershed_btl.synthetic import generate_synthetic_sample

def test_same_seed_produces_identical_image_and_labels():
    a = generate_synthetic_sample(seed=381, shape=(128, 128), n_objects=6)
    b = generate_synthetic_sample(seed=381, shape=(128, 128), n_objects=6)
    assert a.image.shape == (128, 128)
    assert a.image.dtype.name == "uint8"
    assert (a.image == b.image).all()
    assert (a.instance_labels == b.instance_labels).all()
    assert a.instance_labels.max() >= 1
```

- [ ] **Step 2: Run `pytest tests/test_synthetic.py -v`; verify it fails because the package is absent.**
- [ ] **Step 3: Add dependencies and minimal package configuration.** Set `DEFAULT_SEED = 381` in `config.py`; implement non-overwriting folder creation in the runner only, not this module.
- [ ] **Step 4: Implement `generate_synthetic_sample`.** Use `np.random.default_rng(seed)`, draw overlapping disks/ellipses into a labelled mask, derive grayscale intensity, add deterministic Gaussian noise and a smooth background gradient, then clip to `[0, 255]` and convert to `uint8`.
- [ ] **Step 5: Run `pytest tests/test_synthetic.py -v`; verify PASS.**
- [ ] **Step 6: Write README setup and run commands.** Include Windows PowerShell commands using `.venv`, `pip install -r requirements.txt`, and `python scripts/run_experiments.py`.

### Task 2: Morphology primitives and marker creation

**Files:**
- Create: `src/watershed_btl/morphology.py`, `tests/test_morphology.py`

**Interfaces:**
- Consumes a two-dimensional binary `np.ndarray` and `radius: int`.
- Produces `binary_erosion`, `binary_dilation`, `binary_opening`, `binary_closing`, `reconstruct_by_dilation`, and `make_markers(mask: np.ndarray, peak_fraction: float) -> tuple[np.ndarray, np.ndarray]`.
- `make_markers` returns `distance_map: float64` and `markers: int32`, with zero outside `mask`.

- [ ] **Step 1: Write failing tests for exact 3×3 dilation/erosion and marker invariants.**

```python
import numpy as np
from watershed_btl.morphology import binary_dilation, binary_erosion, make_markers

def test_cross_dilates_to_full_three_by_three_square():
    cross = np.array([[0,1,0],[1,1,1],[0,1,0]], dtype=bool)
    assert binary_dilation(cross, radius=1).all()

def test_markers_are_inside_foreground():
    mask = np.zeros((21, 21), dtype=bool); mask[3:9, 3:9] = True; mask[12:18, 12:18] = True
    _, markers = make_markers(mask, peak_fraction=0.45)
    assert np.all(~(markers > 0) | mask)
    assert markers.max() == 2
```

- [ ] **Step 2: Run `pytest tests/test_morphology.py -v`; verify FAIL.**
- [ ] **Step 3: Implement primitives using `skimage.morphology.disk`, `scipy.ndimage` and `skimage.morphology.reconstruction`.** Return Boolean results for binary operations; use `scipy.ndimage.distance_transform_edt` and `ndi.label` for markers.
- [ ] **Step 4: Run `pytest tests/test_morphology.py -v`; verify PASS.**

### Task 3: Segmentation pipeline and image artifact export

**Files:**
- Create: `src/watershed_btl/pipeline.py`, `tests/test_pipeline.py`

**Interfaces:**
- Defines `SegmentationResult(binary_mask, distance_map, markers, labels, overlay)`.
- Defines `segment(image: np.ndarray, threshold: int | None = None, radius: int = 2, peak_fraction: float = 0.45) -> SegmentationResult`.
- Defines `save_stage_images(result: SegmentationResult, image: np.ndarray, output_dir: Path, stem: str) -> dict[str, Path]`.

- [ ] **Step 1: Write a failing pipeline test.**

```python
from watershed_btl.pipeline import segment
from watershed_btl.synthetic import generate_synthetic_sample

def test_pipeline_returns_integer_labels_and_rgb_overlay():
    sample = generate_synthetic_sample(381, (160, 160), 7)
    result = segment(sample.image)
    assert result.labels.dtype.name == "int32"
    assert result.labels.shape == sample.image.shape
    assert result.labels.max() >= 1
    assert result.overlay.shape == (160, 160, 3)
```

- [ ] **Step 2: Run `pytest tests/test_pipeline.py -v`; verify FAIL.**
- [ ] **Step 3: Implement preprocessing and watershed.** Apply Gaussian denoise, Otsu threshold when no threshold is supplied, opening then closing, `make_markers`, and `skimage.segmentation.watershed(-distance_map, markers, mask=binary_mask)`. Use `label2rgb` for an RGB overlay.
- [ ] **Step 4: Implement PNG exports.** Write `input`, `binary_mask`, `distance_map`, `markers`, `labels` and `overlay` images under a caller-provided run directory without changing the source image.
- [ ] **Step 5: Run `pytest tests/test_pipeline.py -v`; verify PASS.**

### Task 4: Metrics and controlled ablation/benchmark experiments

**Files:**
- Create: `src/watershed_btl/metrics.py`, `src/watershed_btl/experiments.py`, `scripts/run_experiments.py`, `tests/test_metrics.py`

**Interfaces:**
- Defines `binary_metrics(prediction: np.ndarray, reference: np.ndarray) -> dict[str, float]` with keys `iou`, `dice`, `precision`, `recall`, `f1`.
- Defines `run_experiments(output_dir: Path, seeds: list[int]) -> pd.DataFrame` and writes `metrics.csv`, `benchmark.csv`, `summary.png`.

- [ ] **Step 1: Write failing metric tests.**

```python
import numpy as np
from watershed_btl.metrics import binary_metrics

def test_perfect_match_has_unit_scores():
    mask = np.array([[0, 1], [1, 0]], dtype=bool)
    scores = binary_metrics(mask, mask)
    assert scores == {"iou": 1.0, "dice": 1.0, "precision": 1.0, "recall": 1.0, "f1": 1.0}
```

- [ ] **Step 2: Run `pytest tests/test_metrics.py -v`; verify FAIL.**
- [ ] **Step 3: Implement safe metric division.** Interpret no prediction/no reference as score `1.0` only when both are empty; otherwise use `0.0` when denominator is zero.
- [ ] **Step 4: Implement experiments.** For seeds `[381, 633, 731, 940, 101]`, generate data, run full pipeline and two ablations (no morphology and connected-components without watershed), time each run using `time.perf_counter`, calculate metrics against `instance_labels > 0`, and create CSV/PNG tables and charts.
- [ ] **Step 5: Run `pytest tests/test_metrics.py -v` and `python scripts/run_experiments.py --output artifacts/run_001`; verify all tests PASS and listed artifacts exist.**

### Task 5: Real-image demonstration and Fiji/MorphoLibJ workflow

**Files:**
- Create: `scripts/run_real_image.py`, `fiji/morphological_watershed.ijm`, `fiji/README.md`
- Modify: `README.md`

**Interfaces:**
- `run_real_image.py --output artifacts/real` creates `real_input.png` and the same stages emitted by `save_stage_images`.
- The macro reads an image path supplied through `getDirectory`, writes Fiji result files to an explicitly chosen output directory, and never modifies the input file.

- [ ] **Step 1: Add a smoke test that invokes `run_real_image.py --help` and asserts exit code 0.**
- [ ] **Step 2: Run it and verify it fails until the script exists.**
- [ ] **Step 3: Implement the real-image script using `skimage.data.coins()` with a local fallback generated by Task 1.** Save input and stage images using the Task 3 interface and record parameters in `real_run.json`.
- [ ] **Step 4: Write the ImageJ macro.** Prompt for input/output directories; duplicate rather than alter the opened image; convert to 8-bit; run threshold and binary morphology; document exact MorphoLibJ menu alternatives in comments; save binary/label images and a CSV Results table.
- [ ] **Step 5: Write Fiji README.** State Fiji and MorphoLibJ installation steps, click-by-click workflow, parameter mapping to Python and a procedure to add observed Fiji metrics to the report.
- [ ] **Step 6: Run `pytest -q` and `python scripts/run_real_image.py --output artifacts/real`; verify PASS and stage images exist.**

### Task 6: Generate and validate the DOCX report

**Files:**
- Create: `report/generate_report.py`, `report/Bao_cao_BTL_Watershed.docx`
- Modify: `README.md`

**Interfaces:**
- `python report/generate_report.py --artifacts artifacts/run_001 --output report/Bao_cao_BTL_Watershed.docx` reads CSV/PNG experiment artifacts and produces the report.
- Uses two sections: `cover` (no footer page field) and `main` (page-number field reset to 1 in centered footer).

- [ ] **Step 1: Read the documents skill instructions and bundled dependency paths before writing DOCX code.**
- [ ] **Step 2: Write the report generator with helper functions `add_cover`, `add_heading`, `add_figure`, `add_metric_table`, `add_references`, and `add_page_number_field`.** Use `python-docx` OOXML for a TOC field and footer `PAGE` field. Set centimetre margins only on the main section.
- [ ] **Step 3: Populate substantive Vietnamese sections.** Cover the definitions/equations of erosion, dilation, opening, closing, reconstruction, distance transform, connected components, watershed; describe MorphoLibJ; state experiment protocol and limitations; cite primary/manual sources with URLs and access dates.
- [ ] **Step 4: Generate the DOCX from `artifacts/run_001`.** Check it opens with `python-docx` and includes heading/table/figure counts expected by the script.
- [ ] **Step 5: Render and visually inspect every DOCX page using the documents skill’s required renderer.** Confirm the cover lacks visible page number and the TOC’s main-section footer begins at 1; correct overflow, missing image, font, or margin defects and re-render.

### Task 7: Generate and validate Canva-compatible slides

**Files:**
- Create: `slides/generate_slides.py`, `slides/BTL_Watershed_Canva.pptx`

**Interfaces:**
- `python slides/generate_slides.py --artifacts artifacts/run_001 --output slides/BTL_Watershed_Canva.pptx` reads report-ready images/metrics and writes an editable presentation.

- [ ] **Step 1: Read the presentations skill instructions and dependency paths before writing PPTX code.**
- [ ] **Step 2: Implement 10-slide deck with editable title/text boxes and native shapes.** Use a consistent 16:9 light theme: title, problem/objective, morphology, watershed, Fiji workflow, Python architecture, data/setup, quantitative results, visual results/discussion, conclusion.
- [ ] **Step 3: Insert only artifacts generated by Tasks 4–5.** Use concise Vietnamese speaker-facing copy and cite the artifact/source on the appropriate slide.
- [ ] **Step 4: Render the PPTX via the presentations skill workflow.** Inspect all slides for clipping, low contrast, image aspect distortion, or absent content; fix the generator and re-render if needed.

### Task 8: End-to-end verification and handoff

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Run `pytest -q`; expected: all tests PASS.**
- [ ] **Step 2: Run `python scripts/run_experiments.py --output artifacts/final` and `python scripts/run_real_image.py --output artifacts/real_final`; expected: metric CSVs, benchmark CSV, figures, stage images and JSON exist.**
- [ ] **Step 3: Generate final DOCX and PPTX from `artifacts/final`; expected: both files exist and are non-empty.**
- [ ] **Step 4: Render and inspect final DOCX and PPTX once more using their required skills.**
- [ ] **Step 5: Update README’s deliverables table with absolute/relative locations, exact commands, Fiji steps and the limitation that Fiji values require a local Fiji run.**
- [ ] **Step 6: Report the final paths and exact verification outputs to the user.**
