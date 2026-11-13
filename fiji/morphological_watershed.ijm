// Morphological watershed workflow for Fiji/ImageJ.
// MorphoLibJ alternatives: Plugins > MorphoLibJ > Segmentation >
// Marker-Controlled Watershed, and Plugins > MorphoLibJ > Binary Images >
// Connected Components Labeling. This macro uses Fiji's built-in binary
// watershed plus connected-component labeling so it remains inspectable.

inputDir = getDirectory("Choose the directory containing the input image");
if (inputDir == "") exit("No input directory selected.");
outputDir = getDirectory("Choose an existing output directory");
if (outputDir == "") exit("No output directory selected.");

files = getFileList(inputDir);
inputName = "";
for (i = 0; i < files.length; i++) {
    lower = toLowerCase(files[i]);
    if (endsWith(lower, ".png") || endsWith(lower, ".tif") || endsWith(lower, ".tiff") || endsWith(lower, ".jpg") || endsWith(lower, ".jpeg")) {
        inputName = files[i];
        break;
    }
}
if (inputName == "") exit("No supported image was found in the input directory.");

open(inputDir + inputName);
run("Duplicate...", "title=watershed_working_copy"); // Original remains unchanged.
run("8-bit");
setAutoThreshold("Otsu dark");
run("Convert to Mask");
run("Open");
run("Close");
saveAs("PNG", outputDir + "fiji_binary_mask.png");

// Built-in: Process > Binary > Watershed. For marker-controlled processing,
// use MorphoLibJ's Segmentation > Marker-Controlled Watershed dialog instead.
run("Watershed");
run("Connected Components Labeling", "connectivity=4 type=[16 bits]");
saveAs("Tiff", outputDir + "fiji_labels.tif");

run("Analyze Particles...", "size=0-Infinity show=Nothing display clear");
saveAs("Results", outputDir + "fiji_results.csv");
close();
