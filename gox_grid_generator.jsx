// Adobe Illustrator Script: Gox Grid Generator
// Creates a grid with multiple width columns and single height

function showDialog() {
    var dialog = new Window("dialog", "Gox Grid Generator");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 10;
    dialog.margins = 20;

    // Width input group
    var widthGroup = dialog.add("group");
    widthGroup.orientation = "column";
    widthGroup.alignChildren = "fill";
    widthGroup.spacing = 5;
    
    var widthLabel = widthGroup.add("statictext", undefined, "Width (A,B,C..)");
    widthLabel.characters = 20;
    
    var widthInput = widthGroup.add("edittext", undefined, "20,30,40,20,60");
    widthInput.characters = 30;
    widthInput.active = true;

    // Height input group
    var heightGroup = dialog.add("group");
    heightGroup.orientation = "column";
    heightGroup.alignChildren = "fill";
    heightGroup.spacing = 5;
    
    var heightLabel = heightGroup.add("statictext", undefined, "Height");
    heightLabel.characters = 20;
    
    var heightInput = heightGroup.add("edittext", undefined, "20");
    heightInput.characters = 10;

    // Dimension label controls
    var dimPanel = dialog.add("panel", undefined, "Dimension Labels");
    dimPanel.orientation = "column";
    dimPanel.alignChildren = "fill";
    dimPanel.margins = 10;
    dimPanel.spacing = 8;

    // Fonts list
    var fontsArray = [];
    var fontNames = [];
    try {
        for (var fi = 0; fi < app.textFonts.length; fi++) {
            var tf = app.textFonts[fi];
            fontsArray.push(tf);
            fontNames.push(tf.name);
        }
    } catch (eFonts) {}

    var fontRow = dimPanel.add("group");
    fontRow.add("statictext", undefined, "Font");
    var fontDropdown = fontRow.add("dropdownlist", undefined, fontNames);
    fontDropdown.minimumSize.width = 260;
    if (fontNames.length > 0) fontDropdown.selection = 0;

    // Font size (pt)
    var sizeRow = dimPanel.add("group");
    sizeRow.add("statictext", undefined, "Font size (pt)");
    var fontSizeInput = sizeRow.add("edittext", undefined, "12");
    fontSizeInput.characters = 6;

    // Gap (document units)
    var gapRow = dimPanel.add("group");
    gapRow.add("statictext", undefined, "Gap");
    var gapInput = gapRow.add("edittext", undefined, "10");
    gapInput.characters = 6;

    // Load saved settings if available
    var saved = loadSettings();
    if (saved) {
        if (saved.widthText) widthInput.text = saved.widthText;
        if (saved.heightText) heightInput.text = saved.heightText;
        if (!isNaN(saved.fontSizePt)) fontSizeInput.text = saved.fontSizePt;
        if (!isNaN(saved.gapValue)) gapInput.text = saved.gapValue;
        if (saved.fontName && fontNames.length > 0) {
            for (var sfi = 0; sfi < fontNames.length; sfi++) {
                if (fontNames[sfi] == saved.fontName) {
                    fontDropdown.selection = sfi;
                    break;
                }
            }
        }
    }

    // Button group
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    
    var createButton = buttonGroup.add("button", undefined, "Create");
    createButton.preferredSize.width = 80;
    
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    cancelButton.preferredSize.width = 80;

    // Event handlers
    createButton.onClick = function() {
        var widthText = widthInput.text;
        var heightText = heightInput.text;
        
        if (widthText == "" || heightText == "") {
            alert("Please enter both width and height values.");
            return;
        }
        
        try {
            var widths = parseWidths(widthText);
            var height = parseFloat(heightText);
            var fontSizePt = parseFloat(fontSizeInput.text);
            if (isNaN(fontSizePt) || fontSizePt <= 0) fontSizePt = 12;
            var gapValue = parseFloat(gapInput.text);
            if (isNaN(gapValue)) gapValue = 10;
            
            if (widths.length == 0) {
                alert("Please enter valid width values separated by commas.");
                return;
            }
            
            if (isNaN(height) || height <= 0) {
                alert("Please enter a valid height value.");
                return;
            }
            
            var selectedFont = null;
            try {
                if (fontDropdown && fontDropdown.selection) {
                    selectedFont = fontsArray[fontDropdown.selection.index];
                }
            } catch (eSel) {}
            
            var options = {
                textFont: selectedFont,
                fontSizePt: fontSizePt,
                gapValue: gapValue
            };
            
            createGoxGrid(widths, height, options);
            // Save settings for next time
            saveSettings({
                widthText: widthText,
                heightText: heightText,
                fontName: (options && options.textFont) ? options.textFont.name : (fontDropdown.selection ? fontNames[fontDropdown.selection.index] : null),
                fontSizePt: fontSizePt,
                gapValue: gapValue
            });
            dialog.close();
            
        } catch (e) {
            alert("Error: " + e.message);
        }
    };

    cancelButton.onClick = function() {
        dialog.close();
    };

    // Show dialog
    dialog.show();
}

function parseWidths(widthText) {
    var widths = [];
    var parts = widthText.split(",");
    
    for (var i = 0; i < parts.length; i++) {
        // Remove whitespace using replace instead of trim()
        var trimmed = parts[i].replace(/^\s+|\s+$/g, "");
        if (trimmed != "") {
            var value = parseFloat(trimmed);
            if (!isNaN(value) && value > 0) {
                widths.push(value);
            }
        }
    }
    
    return widths;
}

// Helper: map RulerUnits to UnitValue strings
function getDocUnitString(doc) {
    var units;
    try {
        units = doc.rulerUnits;
    } catch (e) {
        try {
            units = app.preferences.rulerUnits;
        } catch (e2) {
            units = RulerUnits.Points;
        }
    }
    try {
        if (units == RulerUnits.Points) return 'pt';
        if (units == RulerUnits.Picas) return 'pc';
        if (units == RulerUnits.Inches) return 'in';
        if (units == RulerUnits.Millimeters) return 'mm';
        if (units == RulerUnits.Centimeters) return 'cm';
        if (units == RulerUnits.Pixels) return 'px';
    } catch (e3) {}
    return 'pt';
}

// Convert a numeric value from the document's units to points using UnitValue
function convertDocUnitsToPoints(doc, value) {
    var u = getDocUnitString(doc);
    try {
        var uv = new UnitValue(value, u);
        return uv.as('pt');
    } catch (e) {
        // Fallback: assume entered value is already in points
        return value;
    }
}

// Persist settings between runs using a JSON file next to the script
// Settings file helpers (prefer userData folder; fallback to script folder)
function getUserDataSettingsFile() {
    try {
        var folder = new Folder(Folder.userData + "/GoxGrid");
        if (!folder.exists) folder.create();
        return new File(folder.fsName + "/gox_grid_settings.json");
    } catch (e) {
        return null;
    }
}

function getScriptFolderSettingsFile() {
    try {
        var scriptFile = new File($.fileName);
        var folder = scriptFile.parent;
        return new File(folder.fsName + "/gox_grid_settings.json");
    } catch (e) {
        return null;
    }
}

function safeStringify(obj) {
    try {
        if (typeof JSON !== 'undefined' && JSON.stringify) {
            return JSON.stringify(obj);
        }
    } catch (e) {}
    try {
        if (obj && obj.toSource) return obj.toSource();
    } catch (e2) {}
    // Minimal key=value fallback
    var parts = [];
    for (var k in obj) {
        if (!obj.hasOwnProperty(k)) continue;
        parts.push(k + "=" + ("" + obj[k]).replace(/\n/g, " "));
    }
    return parts.join("\n");
}

function safeParse(content) {
    // Try JSON
    try {
        if (typeof JSON !== 'undefined' && JSON.parse) return JSON.parse(content);
    } catch (e) {}
    // Try eval of toSource()
    try {
        return eval(content);
    } catch (e2) {}
    // Try key=value fallback
    try {
        var obj = {};
        var lines = content.split(/\r?\n/);
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            var idx = line.indexOf("=");
            if (idx > -1) {
                var key = line.substring(0, idx);
                var val = line.substring(idx + 1);
                obj[key] = val;
            }
        }
        return obj;
    } catch (e3) {}
    return null;
}

function loadSettings() {
    var f1 = getUserDataSettingsFile();
    var f2 = getScriptFolderSettingsFile();
    var f = null;
    if (f1 && f1.exists) f = f1; else if (f2 && f2.exists) f = f2;
    if (!f) return null;
    try {
        f.encoding = 'UTF-8';
        f.open('r');
        var content = f.read();
        f.close();
        if (!content) return null;
        return safeParse(content);
    } catch (e) {
        try { if (f && f.opened) f.close(); } catch (e2) {}
        return null;
    }
}

function saveSettings(obj) {
    var f = getUserDataSettingsFile() || getScriptFolderSettingsFile();
    if (!f) return;
    try {
        f.encoding = 'UTF-8';
        f.open('w');
        f.write(safeStringify(obj));
        f.close();
    } catch (e) {
        try { if (f && f.opened) f.close(); } catch (e2) {}
    }
}

function createGoxGrid(widths, height, options) {
    // Check if we have a document open
    if (app.documents.length == 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    
    // Get artboard bounds safely
    var artboardRect;
    try {
        var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        artboardRect = artboard.artboardRect;
    } catch (e) {
        // Fallback to document bounds if artboard access fails
        artboardRect = [0, 0, 1000, 1000];
    }
    
    // Convert user inputs (document units) to points using UnitValue
    var widthsPt = [];
    for (var i = 0; i < widths.length; i++) {
        widthsPt.push(convertDocUnitsToPoints(doc, widths[i]));
    }
    var heightPt = convertDocUnitsToPoints(doc, height);

    // Calculate starting position (center of artboard)
    var totalWidth = 0;
    for (var i = 0; i < widthsPt.length; i++) {
        totalWidth += widthsPt[i];
    }
    
    var startX = artboardRect[0] + (artboardRect[2] - artboardRect[0] - totalWidth) / 2;
    // Vertically center the grid correctly
    var centerY = (artboardRect[1] + artboardRect[3]) / 2;
    var rectTop = centerY + (heightPt / 2);
    
    // Create grid rectangles
    var currentX = startX;
    var createdRects = [];
    var createdTexts = [];
    
    for (var i = 0; i < widthsPt.length; i++) {
        var rect = doc.pathItems.rectangle(
            rectTop,  // top (correctly centered)
            currentX,         // left
            widthsPt[i],      // width
            heightPt          // height
        );
        
        // Style the rectangle safely
        rect.filled = true;
        rect.stroked = true;
        rect.strokeWidth = 1;
        
        // Set colors safely
        try {
            rect.fillColor = doc.swatches["White"].color;
            rect.strokeColor = doc.swatches["Black"].color;
        } catch (e) {
            // Use default colors if swatches don't exist
            rect.fillColor = new RGBColor();
            rect.fillColor.red = 255;
            rect.fillColor.green = 255;
            rect.fillColor.blue = 255;
            
            rect.strokeColor = new RGBColor();
            rect.strokeColor.red = 0;
            rect.strokeColor.green = 0;
            rect.strokeColor.blue = 0;
        }
        
        createdRects.push(rect);
        
        // Add width label above the rectangle
        var textFrame = doc.textFrames.add();
        textFrame.contents = widths[i].toString();
        
        // Set text properties safely
        try {
            var fontSizeToUse = (options && options.fontSizePt) ? options.fontSizePt : 12;
            textFrame.textRange.characterAttributes.size = fontSizeToUse;
            if (options && options.textFont) {
                textFrame.textRange.characterAttributes.textFont = options.textFont;
            }
            textFrame.textRange.characterAttributes.fillColor = doc.swatches["Black"].color;
        } catch (e) {
            // Use default text color
            var fontSizeFallback = (options && options.fontSizePt) ? options.fontSizePt : 12;
            textFrame.textRange.characterAttributes.size = fontSizeFallback;
            if (options && options.textFont) {
                try { textFrame.textRange.characterAttributes.textFont = options.textFont; } catch (eFont) {}
            }
            var textColor = new RGBColor();
            textColor.red = 0;
            textColor.green = 0;
            textColor.blue = 0;
            textFrame.textRange.characterAttributes.fillColor = textColor;
        }
        
        // Position the text above the rectangle (simplified positioning)
        var gapPt = convertDocUnitsToPoints(doc, (options && options.gapValue != null) ? options.gapValue : 10);
        textFrame.position = [
            currentX + widthsPt[i]/2,  // center horizontally
            rectTop + gapPt  // above the rectangle using top of rect
        ];
        
        createdTexts.push(textFrame);
        currentX += widthsPt[i];
    }
    
    // Add height label to the right of the grid
    var heightTextFrame = doc.textFrames.add();
    heightTextFrame.contents = height.toString();
    
    try {
        var heightFontSize = (options && options.fontSizePt) ? options.fontSizePt : 12;
        heightTextFrame.textRange.characterAttributes.size = heightFontSize;
        if (options && options.textFont) {
            heightTextFrame.textRange.characterAttributes.textFont = options.textFont;
        }
        heightTextFrame.textRange.characterAttributes.fillColor = doc.swatches["Black"].color;
    } catch (e) {
        var heightFontSizeFallback = (options && options.fontSizePt) ? options.fontSizePt : 12;
        heightTextFrame.textRange.characterAttributes.size = heightFontSizeFallback;
        if (options && options.textFont) {
            try { heightTextFrame.textRange.characterAttributes.textFont = options.textFont; } catch (eFont2) {}
        }
        var heightTextColor = new RGBColor();
        heightTextColor.red = 0;
        heightTextColor.green = 0;
        heightTextColor.blue = 0;
        heightTextFrame.textRange.characterAttributes.fillColor = heightTextColor;
    }
    
    var gapPtRight = convertDocUnitsToPoints(doc, (options && options.gapValue != null) ? options.gapValue : 10);
    heightTextFrame.position = [
        startX + totalWidth + gapPtRight,  // to the right of the grid
        centerY  // vertically centered
    ];

    // Draw a vertical line next to the grid with arrowheads to show height
    try {
        var dimLine = doc.pathItems.add();
        dimLine.stroked = true;
        dimLine.filled = false;
        dimLine.strokeWidth = 1;
        var dimLeft = startX + totalWidth + (gapPtRight / 2);
        var topPt = [dimLeft, rectTop];
        var bottomPt = [dimLeft, rectTop - heightPt];
        dimLine.setEntirePath([topPt, bottomPt]);
        createdRects.push(dimLine);
    } catch (eDim) {}
    
    createdTexts.push(heightTextFrame);
    
    // Select all created objects
    var selection = createdRects.concat(createdTexts);
    if (selection.length > 0) {
        doc.selection = selection;
    }
}

// Run the script
showDialog();
