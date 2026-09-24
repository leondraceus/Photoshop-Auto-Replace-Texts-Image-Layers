#target photoshop

// =======================================================
// My PSD Batch autorun Replace
// =======================================================

function main() {
    if (app.documents.length === 0) {
        alert("Please open a Photoshop template file first.");
        return;
    }

    var doc = app.activeDocument;

    // Build & Display GUI Dialog
    var dialog = createDialog(doc.name.replace(/\.[^\.]+$/, ''));
    if (dialog.show() !== 1) return; // User cancelled

    var config = dialog.config;

    // Open & Read CSV File
    var csvFile = File(config.csvPath);
    if (!csvFile.exists) {
        alert("Selected CSV file does not exist.");
        return;
    }

    csvFile.open('r');
    var csvRaw = csvFile.read();
    csvFile.close();

    var lines = csvRaw.split(/\r\n|\n|\r/);
    if (lines.length < 2) {
        alert("CSV file is empty or missing data rows.");
        return;
    }

    // Parse CSV Headers
    var delimiter = config.delimiter;
    var headers = parseCSVLine(lines[0], delimiter);

    // Store initial history state for clean resets on each iteration
    var initialState = doc.activeHistoryState;
    var successCount = 0;

    // Base document name
    var docBaseName = doc.name.replace(/\.[^\.]+$/, '');

    // Check for missing/mismatched layer names before running
    var missingLayers = [];
    for (var h = 0; h < headers.length; h++) {
        var hName = trim(headers[h]);
        
        // Skip checking the Naming column if it's strictly used for filenames
        if (config.namingMode > 0 && hName.toLowerCase() === config.nameHeader.toLowerCase()) {
            continue;
        }

        if (!getLayerByName(doc, hName)) {
            missingLayers.push(hName);
        }
    }

    // Check if the Naming Header exists in CSV
    if (config.namingMode > 0) {
        var namingHeaderFound = false;
        for (var h = 0; h < headers.length; h++) {
            if (trim(headers[h]).toLowerCase() === config.nameHeader.toLowerCase()) {
                namingHeaderFound = true;
                break;
            }
        }
        if (!namingHeaderFound) {
            alert("Error: The CSV Header for Name ('" + config.nameHeader + "') was not found in your CSV file.");
            return;
        }
    }

    // Display warning box for missing layers and ask to continue or stop
    if (missingLayers.length > 0) {
        var warnMsg = "Warning: The following CSV header(s) do not match any layer name in your PSD document:\n\n" +
                      "• " + missingLayers.join("\n• ") + 
                      "\n\nDo you want to continue processing anyway?";
        if (!confirm(warnMsg)) {
            return; // Cancel the run
        }
    }
    
    // Process Rows
    for (var i = 1; i < lines.length; i++) {
        if (trim(lines[i]) === "") continue;

        var rowData = parseCSVLine(lines[i], delimiter);

        // Reset document to template original state
        doc.activeHistoryState = initialState;

        // Loop through CSV headers to replace contents
        for (var j = 0; j < headers.length; j++) {
            var headerName = trim(headers[j]);
            var cellValue = trim(rowData[j]);

            if (!cellValue) continue;

            var targetLayer = getLayerByName(doc, headerName);
            if (!targetLayer) continue;

            // 1. Process Text Layers
            if ((config.mode === 0 || config.mode === 1) && targetLayer.kind === LayerKind.TEXT) {
                targetLayer.textItem.contents = cellValue;
            }

            // 2. Process Image Layers (Smart Objects)
            if ((config.mode === 0 || config.mode === 2) && targetLayer.kind === LayerKind.SMARTOBJECT) {
                replaceAndScaleSmartObject(targetLayer, cellValue, config.placementMode);
            }
        }

        // Color Profile Conversion (JPEG, PSD, TIFF)
        if (config.convertProfile && config.format !== "PNG") {
            try {
                doc.convertProfile(config.profileName, Intent.RELATIVECOLORIMETRIC, true, true);
            } catch (e) {
                // Ignore if profile match/not found
            }
        }

        // Handle Flatten / Merged Layer Options
        if (config.format !== "JPEG") {
            if (config.flatten) {
                doc.flatten();
            } else if (config.merged) {
                doc.mergeVisibleLayers();
            }
        }

        // ---------------------------------------------------
        // Determine Output Filename
        // ---------------------------------------------------
        var fileName = "";
        var currentSerial = config.startNum + (i - 1);

        if (config.namingMode === 0) {
            // Option 1: [Document Name] + Numbers
            fileName = docBaseName + "_" + currentSerial;
        } else {
            // Search CSV row for specified Header name
            var rawValue = getCellValueByHeader(headers, rowData, config.nameHeader);
            var safeValue = sanitizeFileName(rawValue);

            if (safeValue === "") {
                safeValue = "Item_" + currentSerial; // Fallback if header not found or blank
            }

            if (config.namingMode === 1) {
                // Option 2: [Document Name] + CSV Value
                fileName = docBaseName + "_" + safeValue;
            } else if (config.namingMode === 2) {
                // Option 3: CSV Value Only
                fileName = safeValue;
            }
        }

        // Save File
        var savePath = File(config.outputPath + "/" + fileName);
        saveFileByFormat(doc, savePath, config);

        successCount++;
    }

    // Restore original history state
    doc.activeHistoryState = initialState;
    alert("Batch process completed!\nSuccessfully exported " + successCount + " files.");
}

// =======================================================
// ScriptUI Dialog Construction
// =======================================================
function createDialog(docName) {
    var win = new Window("dialog", "Variable Replace From CSV");
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 10;
    win.margins = 15;

    // ---------------------------------------------------
    // 1. Data File (CSV) Section
    // ---------------------------------------------------
    var csvPanel = win.add("panel", undefined, "Data file (CSV)");
    csvPanel.alignChildren = ["left", "top"];
    csvPanel.spacing = 8;

    var delimGrp = csvPanel.add("group");
    delimGrp.add("statictext", undefined, "Delimiter:");
    var radioComma = delimGrp.add("radiobutton", undefined, "Comma");
    var radioSemicolon = delimGrp.add("radiobutton", undefined, "Semicolon");
    radioComma.value = true;

    var csvFileGrp = csvPanel.add("group");
    var csvBtn = csvFileGrp.add("button", undefined, "File...");
    var csvPathTxt = csvFileGrp.add("statictext", undefined, "No file selected...", { truncate: "middle" });
    csvPathTxt.preferredSize.width = 270;
    csvPathTxt.graphics.font = ScriptUI.newFont("dialog", "ITALIC", 10);

    csvBtn.onClick = function () {
        var f = File.openDialog("Select CSV Data File", "*.csv");
        if (f) csvPathTxt.text = f.fsName;
    };

    // ---------------------------------------------------
    // 2. Input Text & Images Section
    // ---------------------------------------------------
    var inputPanel = win.add("panel", undefined, "Input Text & Images");
    inputPanel.alignChildren = ["left", "top"];
    inputPanel.spacing = 8;

    var modeGrp = inputPanel.add("group");
    modeGrp.add("statictext", undefined, "Mode:");
    var modeDropdown = modeGrp.add("dropdownlist", undefined, ["Texts & Images", "Texts only", "Images only"]);
    modeDropdown.selection = 0;

    var alignGrp = inputPanel.add("group");
    alignGrp.add("statictext", undefined, "Image placement:");
    var radioFill = alignGrp.add("radiobutton", undefined, "Fill");
    var radioFit = alignGrp.add("radiobutton", undefined, "Fit");
    var radioStretch = alignGrp.add("radiobutton", undefined, "Stretch");
    radioFill.value = true;

    modeDropdown.onChange = function () {
        alignGrp.enabled = (modeDropdown.selection.index !== 1);
    };

    // ---------------------------------------------------
    // 3. Output File Naming Options
    // ---------------------------------------------------
    var namePanel = win.add("panel", undefined, "File Naming Options");
    namePanel.alignChildren = ["left", "top"];
    namePanel.spacing = 4;

    // Dropdown row
    var nameDropdownGrp = namePanel.add("group");
    nameDropdownGrp.add("statictext", undefined, "Naming Pattern:");
    var nameDropdown = nameDropdownGrp.add("dropdownlist", undefined, [
        "[Document Name] + Numbers",
        "[Document Name] + CSV Value",
        "CSV Value Only"
    ]);

    // Dynamic example text directly below the dropdown
    var exGrp = namePanel.add("group");
    exGrp.margins = [100, 0, 0, 4];
    var exTxt = exGrp.add("statictext", undefined, "");
    exTxt.preferredSize.width = 320;
    exTxt.graphics.font = ScriptUI.newFont("dialog", "ITALIC", 10);

    // Map example strings to match dropdown indices [0, 1, 2]
    var namingExamples = [
        "Example: " + docName + "_1",
        "Example: " + docName + "_Hello",
        "Example: Hello"
    ];

    // Sub-options row
    var subOptsGrp = namePanel.add("group");
    subOptsGrp.alignChildren = ["left", "center"];

    subOptsGrp.add("statictext", undefined, "Starting Serial #:");
    var startNumTxt = subOptsGrp.add("edittext", undefined, "1");
    startNumTxt.preferredSize.width = 40;

    var csvColLabel = subOptsGrp.add("statictext", undefined, "   CSV Header for Name:");
    var csvColTxt = subOptsGrp.add("edittext", undefined, "layer1");
    csvColTxt.preferredSize.width = 100;

    // Update UI logic on change
    function updateNamingUI() {
        var idx = nameDropdown.selection ? nameDropdown.selection.index : 0;

        // Change small italic text dynamically
        exTxt.text = namingExamples[idx];

        // Toggle sub-option inputs
        startNumTxt.enabled = (idx === 0);
        csvColLabel.enabled = (idx !== 0);
        csvColTxt.enabled = (idx !== 0);
    }

    nameDropdown.onChange = updateNamingUI;

    // Initialize selection
    nameDropdown.selection = 0;
    updateNamingUI();

    // ---------------------------------------------------
    // 4. Output Folder & Format Section
    // ---------------------------------------------------
    var outPanel = win.add("panel", undefined, "Output Folder & Format");
    outPanel.alignChildren = ["left", "top"];
    outPanel.spacing = 8;

    var outFolderGrp = outPanel.add("group");
    var outFolderBtn = outFolderGrp.add("button", undefined, "Folder...");
    var outFolderPathTxt = outFolderGrp.add("statictext", undefined, "No folder selected...", { truncate: "middle" });
    outFolderPathTxt.preferredSize.width = 270;
    outFolderPathTxt.graphics.font = ScriptUI.newFont("dialog", "ITALIC", 10);

    outFolderBtn.onClick = function () {
        var f = Folder.selectDialog("Select Output Folder");
        if (f) outFolderPathTxt.text = f.fsName;
    };

    // Format & Layer Options Row (Single Horizontal Container)
    var formatRowGrp = outPanel.add("group");
    formatRowGrp.orientation = "row";
    formatRowGrp.alignChildren = ["left", "center"];
    formatRowGrp.spacing = 20; // Space between dropdown and checkboxes

    // Left side: Format Selector
    var fmtGrp = formatRowGrp.add("group");
    fmtGrp.add("statictext", undefined, "Format:");
    var fmtDropdown = fmtGrp.add("dropdownlist", undefined, ["PNG", "JPEG", "PSD", "TIFF"]);

    // Right side: Layer options (PNG, PSD, TIFF)
    var layerOptsGrp = formatRowGrp.add("group");
    var chkFlatten = layerOptsGrp.add("checkbox", undefined, "Flatten");
    var chkMerged = layerOptsGrp.add("checkbox", undefined, "Single merged layer");

    // JPEG Quality selection
    var qualityGrp = outPanel.add("group");
    qualityGrp.add("statictext", undefined, "Quality:");
    var qualityDropdown = qualityGrp.add("dropdownlist", undefined, ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]);
    qualityDropdown.selection = 9;

    // Color Profile Option
    var profileGrp = outPanel.add("group");
    var chkConvertProfile = profileGrp.add("checkbox", undefined, "Convert to profile:");
    var profileDropdown = profileGrp.add("dropdownlist", undefined, ["sRGB IEC61966-2.1", "Adobe RGB (1998)", "ProPhoto RGB"]);
    profileDropdown.selection = 0;

    // Format Change Logic Handler
    fmtDropdown.onChange = function () {
        var fmt = fmtDropdown.selection.text;

        if (fmt === "JPEG") {
            layerOptsGrp.enabled = false;
            qualityGrp.enabled = true;
            profileGrp.enabled = true;
        } else if (fmt === "PNG") {
            layerOptsGrp.enabled = true;
            qualityGrp.enabled = false;
            profileGrp.enabled = false;
            chkConvertProfile.value = false;
        } else { // PSD, TIFF
            layerOptsGrp.enabled = true;
            qualityGrp.enabled = false;
            profileGrp.enabled = true;
        }
    };

    fmtDropdown.selection = 0;

    chkFlatten.onClick = function () { if (chkFlatten.value) chkMerged.value = false; };
    chkMerged.onClick = function () { if (chkMerged.value) chkFlatten.value = false; };

    // ---------------------------------------------------
    // 5. Footer Section
    // ---------------------------------------------------
    var footerGroup = win.add("group");
    footerGroup.orientation = "row";
    footerGroup.alignment = ["fill", "center"];
    footerGroup.alignChildren = ["left", "center"];

    // Version & Credits Text
    var versionTxt = footerGroup.add("statictext", undefined, "Version 1.0 | Leon Ha");
    versionTxt.graphics.font = ScriptUI.newFont("dialog", "REGULAR", 10);

    // Flexible Spacer to push the button to the right side
    var spacer = footerGroup.add("group");
    spacer.alignment = ["fill", "fill"];

    // GitHub Button
    var btnGithub = footerGroup.add("button", undefined, "GitHub");
    btnGithub.preferredSize.height = 22;

    // GitHub Redirect Click Event
    btnGithub.onClick = function() {
        var githubUrl = "https://github.com/leondraceus/LeonDraceus"; // Replace with your repository URL
        openURL(githubUrl);
    };

    function openURL(url) {
        var tempFile = new File(Folder.temp + "/open_github.html");
        tempFile.open("w");
        tempFile.write('<html><head><meta http-equiv="refresh" content="0;url=' + url + '"></head><body></body></html>');
        tempFile.close();
        tempFile.execute();
    }

    // ---------------------------------------------------
    // OK / Cancel Buttons
    // ---------------------------------------------------
    var btnGrp = win.add("group");
    btnGrp.alignment = "center";
    var okBtn = btnGrp.add("button", undefined, "OK", { name: "ok" });
    var cancelBtn = btnGrp.add("button", undefined, "Cancel", { name: "cancel" });

    okBtn.onClick = function () {
        if (!csvPathTxt.text || csvPathTxt.text === "No file selected...") {
            alert("Please select a valid CSV file.");
            return;
        }
        if (!outFolderPathTxt.text || outFolderPathTxt.text === "No folder selected...") {
            alert("Please select an output folder.");
            return;
        }

        var startNumVal = parseInt(startNumTxt.text, 10);
        if (isNaN(startNumVal)) startNumVal = 1;

        var namingMode = nameDropdown.selection.index;

        win.config = {
            csvPath: csvPathTxt.text,
            delimiter: radioComma.value ? "," : ";",
            mode: modeDropdown.selection.index,
            placementMode: radioFill.value ? "fill" : (radioFit.value ? "fit" : "stretch"),
            namingMode: namingMode,
            startNum: startNumVal,
            nameHeader: trim(csvColTxt.text),
            outputPath: outFolderPathTxt.text,
            format: fmtDropdown.selection.text,
            flatten: chkFlatten.value,
            merged: chkMerged.value,
            quality: parseInt(qualityDropdown.selection.text, 10),
            convertProfile: chkConvertProfile.value,
            profileName: profileDropdown.selection.text
        };

        win.close(1);
    };

    cancelBtn.onClick = function () {
        win.close(0);
    };

    return win;
}

// =======================================================
// Helper & Engine Functions
// =======================================================

function replaceAndScaleSmartObject(layer, filePath, placementMode) {
    var imgFile = new File(filePath);
    if (!imgFile.exists) return;

    var origBounds = layer.bounds;
    var origW = Number(origBounds[2] - origBounds[0]);
    var origH = Number(origBounds[3] - origBounds[1]);

    app.activeDocument.activeLayer = layer;
    var idplacedLayerReplaceContents = stringIDToTypeID("placedLayerReplaceContents");
    var desc = new ActionDescriptor();
    desc.putPath(charIDToTypeID("null"), imgFile);
    executeAction(idplacedLayerReplaceContents, desc, DialogModes.NO);

    var newBounds = layer.bounds;
    var newW = Number(newBounds[2] - newBounds[0]);
    var newH = Number(newBounds[3] - newBounds[1]);

    if (newW === 0 || newH === 0) return;

    var scaleX = (origW / newW) * 100;
    var scaleY = (origH / newH) * 100;

    if (placementMode === "stretch") {
        layer.resize(scaleX, scaleY, AnchorPosition.MIDDLECENTER);
    } else if (placementMode === "fit") {
        var fitScale = Math.min(scaleX, scaleY);
        layer.resize(fitScale, fitScale, AnchorPosition.MIDDLECENTER);
    } else if (placementMode === "fill") {
        var fillScale = Math.max(scaleX, scaleY);
        layer.resize(fillScale, fillScale, AnchorPosition.MIDDLECENTER);
    }
}

function saveFileByFormat(doc, savePath, config) {
    var fmt = config.format;

    if (fmt === "PNG") {
        var pngOpts = new PNGSaveOptions();
        doc.saveAs(new File(savePath + ".png"), pngOpts, true, Extension.LOWERCASE);
    } else if (fmt === "JPEG") {
        var jpgOpts = new JPEGSaveOptions();
        jpgOpts.quality = config.quality;
        doc.saveAs(new File(savePath + ".jpg"), jpgOpts, true, Extension.LOWERCASE);
    } else if (fmt === "PSD") {
        var psdOpts = new PhotoshopSaveOptions();
        doc.saveAs(new File(savePath + ".psd"), psdOpts, true, Extension.LOWERCASE);
    } else if (fmt === "TIFF") {
        var tiffOpts = new TiffSaveOptions();
        doc.saveAs(new File(savePath + ".tif"), tiffOpts, true, Extension.LOWERCASE);
    }
}

function getLayerByName(parent, name) {
    for (var i = 0; i < parent.layers.length; i++) {
        var l = parent.layers[i];
        if (l.name === name) return l;
        if (l.typename === "LayerSet") {
            var found = getLayerByName(l, name);
            if (found) return found;
        }
    }
    return null;
}

function getCellValueByHeader(headers, rowData, targetHeader) {
    for (var k = 0; k < headers.length; k++) {
        if (trim(headers[k]).toLowerCase() === targetHeader.toLowerCase()) {
            return trim(rowData[k]);
        }
    }
    return "";
}

function sanitizeFileName(str) {
    if (!str) return "";
    var cleanStr = str.replace(/\\/g, '/');
    if (cleanStr.indexOf('/') !== -1) {
        cleanStr = cleanStr.substring(cleanStr.lastIndexOf('/') + 1);
    }
    cleanStr = cleanStr.replace(/\.[^\.]+$/, '');
    return cleanStr.replace(/[\/\\:\*\?"<>\|]/g, "_");
}

function parseCSVLine(line, delimiter) {
    var values = [];
    var current = "";
    var inQuotes = false;

    for (var i = 0; i < line.length; i++) {
        var c = line.charAt(i);
        if (c === '"') {
            inQuotes = !inQuotes;
        } else if (c === delimiter && !inQuotes) {
            values.push(current);
            current = "";
        } else {
            current += c;
        }
    }
    values.push(current);
    return values;
}

function trim(str) {
    return str ? str.replace(/^\s+|\s+$/g, '') : "";
}

// Run Script
main();