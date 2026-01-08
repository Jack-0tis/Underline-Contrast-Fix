/*
    UnderlineContrastFix.jsx
    ------------------------
    Adjusts underline tint (100% → 0%) until contrast
    between text color and underline color passes a threshold.

    Treats underline as background.
    CMYK only.
*/

// ================= CONFIG =================
var MIN_CONTRAST = 5;     // Adjust later after testing
var TINT_STEP    = 0.01;    // 1% steps
// =========================================


// ---------- COLOR MATH ----------

function cmykToRgb(c, m, y, k) {
    c /= 100; m /= 100; y /= 100; k /= 100;
    return [
        255 * (1 - c) * (1 - k),
        255 * (1 - m) * (1 - k),
        255 * (1 - y) * (1 - k)
    ];
}

function channelToLinear(c) {
    c /= 255;
    return (c <= 0.03928)
        ? c / 12.92
        : Math.pow((c + 0.055) / 1.055, 2.4);
}

function relativeLuminance(rgb) {
    return 0.2126 * channelToLinear(rgb[0]) +
           0.7152 * channelToLinear(rgb[1]) +
           0.0722 * channelToLinear(rgb[2]);
}

function contrastRatio(rgb1, rgb2) {
    var l1 = relativeLuminance(rgb1);
    var l2 = relativeLuminance(rgb2);
    return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05);
}


// ---------- TINT HANDLING ----------

function applyTint(cmyk, tint) {
    return [
        Math.round(cmyk[0] * tint),
        Math.round(cmyk[1] * tint),
        Math.round(cmyk[2] * tint),
        Math.round(cmyk[3] * tint)
    ];
}


// ---------- STYLE UTILITIES ----------

function getRootParagraphStyle(paraStyle) {
    // Follow the basedOn chain to find the root style
    var current = paraStyle;
    var visited = {}; // Prevent infinite loops
    
    while (current.basedOn && current.basedOn.isValid && current.basedOn.name !== "[No Paragraph Style]") {
        if (visited[current.name]) {
            break; // Circular reference detected
        }
        visited[current.name] = true;
        current = current.basedOn;
    }
    
    return current;
}


// ---------- CORE LOGIC ----------

function fixUnderlineContrastInStyle(paraStyle) {
    
    if (!paraStyle.underline) return null;

    // Ensure CMYK process colors
    if (!paraStyle.fillColor || !paraStyle.underlineColor) return null;
    if (!paraStyle.fillColor.hasOwnProperty("colorValue")) return null;
    if (!paraStyle.underlineColor.hasOwnProperty("colorValue")) return null;

    var textCMYK = paraStyle.fillColor.colorValue;
    var underlineCMYK = paraStyle.underlineColor.colorValue;

    var textRGB = cmykToRgb(
        textCMYK[0], textCMYK[1], textCMYK[2], textCMYK[3]
    );

    // Get initial tint (default to 100 if not set)
    var initialTint = paraStyle.underlineTint;
    if (typeof initialTint !== "number" || isNaN(initialTint)) {
        initialTint = 100;
    }
    
    var tint = initialTint / 100;
    var MIN_TINT = 0.075; // Minimum tint of 7.5%

    while (tint >= MIN_TINT) {

        var visibleUnderline = applyTint(underlineCMYK, tint);
        var underlineRGB = cmykToRgb(
            visibleUnderline[0],
            visibleUnderline[1],
            visibleUnderline[2],
            visibleUnderline[3]
        );

        if (contrastRatio(textRGB, underlineRGB) >= MIN_CONTRAST) {
            var newTint = Math.round(tint * 100);
            paraStyle.underlineTint = newTint;
            return {
                initialTint: initialTint,
                newTint: newTint,
                adjusted: (newTint !== initialTint)
            };
        }

        tint -= TINT_STEP;
    }

    // Fallback: set to minimum tint (7.5%)
    var newTint = 7.5;
    paraStyle.underlineTint = newTint;
    return {
        initialTint: initialTint,
        newTint: newTint,
        adjusted: (newTint !== initialTint)
    };
}


// ---------- DOCUMENT WALK ----------

function processDocument(doc) {

    var adjustments = [];
    var paraStyles = doc.paragraphStyles;
    var processedRoots = {}; // Track which root styles we've already processed

    // Find all paragraph styles with underlines
    for (var i = 0; i < paraStyles.length; i++) {
        var paraStyle = paraStyles[i];
        
        // Skip [None] and [No Paragraph Style]
        if (paraStyle.name.substring(0, 1) === '[') continue;
        
        if (paraStyle.underline) {
            // Get the root style
            var rootStyle = getRootParagraphStyle(paraStyle);
            
            // Only process each root style once
            if (!processedRoots[rootStyle.name]) {
                processedRoots[rootStyle.name] = true;
                
                var result = fixUnderlineContrastInStyle(rootStyle);
                if (result && result.adjusted) {
                    adjustments.push({
                        styleName: rootStyle.name,
                        initialTint: result.initialTint,
                        newTint: result.newTint
                    });
                }
            }
        }
    }

    // Build report message
    var message = "Underline Contrast Fix Complete\n\n";
    
    if (adjustments.length === 0) {
        message += "No contrast adjustments were needed.\nAll underlined paragraph styles already meet the contrast threshold.";
    } else {
        message += "Contrast Errors\n";
        for (var j = 0; j < adjustments.length; j++) {
            var adj = adjustments[j];
            message += "Paragraph Style Name: " + adj.styleName + 
                      " Adjustment: " + adj.initialTint + "% -> " + adj.newTint + "%\n";
        }
    }

    alert(message);
}


// ---------- ENTRY POINT ----------

if (app.documents.length === 0) {
    alert("Please open a document first.");
} else {
    app.doScript(
        function () { processDocument(app.activeDocument); },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "Fix Underline Contrast"
    );
}
