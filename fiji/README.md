# Fiji / MorphoLibJ workflow

## Install

1. Install Fiji from <https://fiji.sc/> and start it once.
2. Use `Help > Update...`, choose `Manage update sites`, enable **MorphoLibJ**,
   then apply changes and restart Fiji.
3. Put one grayscale image in a dedicated input directory. Use a separate,
   existing output directory; the macro never writes to the input directory.

## Run the macro

1. In Fiji, choose `Plugins > New > Macro`, open
   `fiji/morphological_watershed.ijm`, then click `Run`.
2. Select the input directory, then select the output directory.
3. The macro opens the first PNG, TIFF, or JPEG it finds, duplicates it, and
   processes only the duplicate. It saves `fiji_binary_mask.png`,
   `fiji_labels.tif`, and `fiji_results.csv` in the output directory.
4. For marker-controlled watershed, replace the built-in watershed step with
   `Plugins > MorphoLibJ > Segmentation > Marker-Controlled Watershed` and use
   a marker image created from the distance-map maxima. The macro comments name
   the relevant MorphoLibJ menu entries.

## Mapping to Python

| Python pipeline | Fiji equivalent |
| --- | --- |
| Gaussian smoothing (`sigma=1.0`) | `Process > Filters > Gaussian Blur...`, radius 1.0 |
| Otsu threshold | `Image > Adjust > Auto Threshold...`, Otsu |
| opening / closing (`radius=2`) | `Process > Binary > Open` / `Close` |
| distance-map maxima markers | MorphoLibJ marker image for Marker-Controlled Watershed |
| `watershed` labels | MorphoLibJ Marker-Controlled Watershed, then Connected Components Labeling |

The Python defaults are threshold computed by Otsu, morphology radius `2`, and
marker peak fraction `0.45`. Fiji's pixel radius and connectivity should be
recorded alongside those values because the applications do not expose an
identical peak-fraction control.

## Add observed Fiji metrics to the report

Open `fiji_results.csv` after `Analyze Particles` finishes. Record the image
name, object count, area mean/standard deviation, and the threshold/morphology
settings in a clearly labeled **Fiji observed results** table. Keep these values
separate from Python metrics and cite the exact input image and plugin version;
do not invent values when Fiji has not been run.
