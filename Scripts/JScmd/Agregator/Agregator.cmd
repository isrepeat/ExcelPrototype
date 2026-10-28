@if (@X)==(@Y) @end /*
@echo off
rem This is a Batch/JScript polyglot: both interpreters read the same file.
rem Batch: the first-line IF compares different strings and skips its command.
rem JScript: @if uses conditional compilation; the false branch ends at @end.
rem The following /* opens a JScript comment containing this Batch launcher.
rem CScript runs this file as JScript; %~f0 is this file and %* forwards arguments.
cscript.exe //nologo //E:JScript "%~f0" %*
rem Preserve the failure code and keep errors visible when launched by double-click.
if not errorlevel 1 exit /b 0
set "agregatorExitCode=%errorlevel%"
pause
rem Exit before Batch reaches the JScript section.
exit /b %agregatorExitCode%
*/

// The closing comment above ends the Batch-only section for JScript.
// Use Windows Script Host COM automation; this is not Node.js JavaScript.
var fso = new ActiveXObject("Scripting.FileSystemObject");
var folder = fso.GetParentFolderName(WScript.ScriptFullName);
var args = WScript.Arguments;
// Keep the editable UTF-8 configuration beside this polyglot script.
function readConfig() {
    var path = fso.BuildPath(folder, "Agregator.xml");
    if (!fso.FileExists(path)) throw new Error("Configuration not found: " + path);
    var xml = new ActiveXObject("Msxml2.DOMDocument.6.0");
    xml.async = false;
    xml.resolveExternals = false;
    xml.setProperty("ProhibitDTD", true);
    if (!xml.load(path)) {
        throw new Error("Invalid configuration: " + path + ": " + xml.parseError.reason);
    }
    var root = xml.documentElement;
    if (root.nodeName !== "aggregator") throw new Error("Expected aggregator configuration root.");
    var config = {
        sourceFolder: requiredAttribute(root, "sourceFolder"),
        output: requiredAttribute(root, "output"),
        sources: [],
        excludedSheets: []
    };
    var excluded = root.selectNodes("excludedSheets/sheet");
    for (var e = 0; e < excluded.length; e++) {
        config.excludedSheets.push(requiredAttribute(excluded.item(e), "name").toLowerCase());
    }
    var workbooks = root.selectNodes("sources/workbook");
    if (workbooks.length === 0) throw new Error("No source workbooks configured.");
    for (var w = 0; w < workbooks.length; w++) {
        var workbook = workbooks.item(w);
        var entry = { file: requiredAttribute(workbook, "file"), sheets: [] };
        var sheets = workbook.selectNodes("sheet");
        for (var s = 0; s < sheets.length; s++) {
            entry.sheets.push(requiredAttribute(sheets.item(s), "name"));
        }
        config.sources.push(entry);
    }
    return config;
}

function requiredAttribute(node, name) {
    var value = node.getAttribute(name);
    if (typeof value !== "string" || !/\S/.test(value)) {
        throw new Error("Missing configuration attribute: " + node.nodeName + "/@" + name);
    }
    return value;
}

function resolvePath(path) {
    // Resolve relative paths beside this script, not in the caller's directory.
    if (!/^(?:[A-Za-z]:[\\/]|\\\\)/.test(path)) {
        path = fso.BuildPath(folder, path);
    }
    return fso.GetAbsolutePathName(path);
}

function main() {
    // Parse launch options before validating the configured input selection.
    var validate = false;
    var paths = [];
    for (var a = 0; a < args.length; a++) {
        if (args.Item(a) === "--help") {
            WScript.Echo("Agregator.cmd [--validate] [output.xlsx]");
            WScript.Echo("Relative paths are resolved beside Agregator.cmd. Existing output is never overwritten.");
            return 0;
        }
        if (args.Item(a) === "--validate") {
            validate = true;
        } else {
            paths.push(args.Item(a));
        }
    }

    if (paths.length > 1) {
        throw new Error("Specify only an optional output path; edit Agregator.xml to select inputs.");
    }

    var config = readConfig();
    var sources = config.sources;
    // Reject existing output and missing inputs before creating an Excel instance.
    var output = resolvePath(paths.length === 0 ? config.output : paths[0]);
    if (fso.GetExtensionName(output).toLowerCase() !== "xlsx") {
        throw new Error("Output must have the .xlsx extension: " + output);
    }
    if (fso.FileExists(output) || fso.FolderExists(output)) {
        throw new Error("Output already exists: " + output);
    }
    if (!fso.FolderExists(fso.GetParentFolderName(output))) {
        throw new Error("Output folder does not exist: " + output);
    }
    var inputs = [];
    for (var p = 0; p < sources.length; p++) {
        var input = resolvePath(fso.BuildPath(resolvePath(config.sourceFolder), sources[p].file));
        if (!fso.FileExists(input)) {
            throw new Error("Input workbook does not exist: " + input);
        }
        if (input.toLowerCase() === output.toLowerCase()) {
            throw new Error("Input and output paths must differ: " + input);
        }
        if (sources[p].sheets.length === 0) {
            throw new Error("No worksheets selected: " + input);
        }
        for (var s = 0; s < sources[p].sheets.length; s++) {
            // Reject explicitly excluded names even if they appear in the allowlist.
            for (var e = 0; e < config.excludedSheets.length; e++) {
                if (sources[p].sheets[s].toLowerCase() === config.excludedSheets[e]) {
                    throw new Error("Excluded worksheet must not be selected: " + sources[p].sheets[s]);
                }
            }
        }
        inputs.push(input);
    }
    if (validate) {
        // Path validation does not open workbooks or verify worksheet existence.
        WScript.Echo("Paths validated. Input workbooks: " + inputs.length);
        WScript.Echo("Output: " + output);
        return 0;
    }

    var excel = null;
    var source = null;
    var result = null;
    try {
        // Own a separate Excel instance and suppress macros, events and recalculation.
        excel = new ActiveXObject("Excel.Application");
        excel.Visible = false;
        excel.ScreenUpdating = false;
        excel.DisplayAlerts = false;
        excel.EnableEvents = false;
        excel.AutomationSecurity = 3;
        result = excel.Workbooks.Add(-4167); // xlWBATWorksheet: one initial sheet.
        excel.Calculation = -4135; // xlCalculationManual.
        excel.CalculateBeforeSave = false;
        var placeholder = result.Worksheets.Item(1);
        var count = 0;

        for (var i = 0; i < inputs.length; i++) {
            var started = new Date().getTime();
            WScript.Echo("Opening: " + inputs[i]);
            source = excel.Workbooks.Open(inputs[i], 0, true); // No link updates; read-only.
            WScript.Echo("Opened in " + ((new Date().getTime() - started) / 1000) + " s");
            for (var j = 0; j < sources[i].sheets.length; j++) {
                var sheetName = sources[i].sheets[j];
                var sheetStarted = new Date().getTime();
                WScript.Echo("Copying worksheet: " + sheetName);
                var sourceSheet;
                try {
                    sourceSheet = source.Worksheets.Item(sheetName);
                } catch (sheetError) {
                    throw new Error("Selected worksheet not found: " + inputs[i] + " / " + sheetName);
                }
                // Copy the complete worksheet before the placeholder, preserving layout.
                // Passing Before directly avoids undefined optional COM arguments in WSH.
                sourceSheet.Copy(placeholder);
                var targetSheet = result.Worksheets.Item(placeholder.Index - 1);
                targetSheet.Visible = -1; // xlSheetVisible.
                count++;
                // A numeric prefix preserves uniqueness after Excel's 31-character limit.
                var name = count + "_" + fso.GetBaseName(inputs[i]) + "_" + sourceSheet.Name;
                targetSheet.Name = name.replace(/[\\\/\?\*\[\]:']/g, "_").substr(0, 31);

                // Convert tables to formatted ranges before freezing calculated columns.
                while (targetSheet.ListObjects.Count > 0) {
                    targetSheet.ListObjects.Item(1).Unlist();
                }
                // Freeze results in the copied sheet while retaining its cell formatting.
                var target = targetSheet.UsedRange;
                target.Copy();
                target.PasteSpecial(12); // xlPasteValuesAndNumberFormats: no formulas.
                // Remove rules that may accompany formatting; their appearance is not frozen.
                target.FormatConditions.Delete();
                target.Validation.Delete();
                excel.CutCopyMode = false;
                WScript.Echo("Copied in " + ((new Date().getTime() - sheetStarted) / 1000) + " s");
            }
            source.Close(false);
            source = null;
        }

        if (count === 0) {
            throw new Error("No worksheets found in the input workbooks.");
        }
        // Remove the initial empty sheet and save a macro-free .xlsx workbook.
        placeholder.Delete();
        result.SaveAs(output, 51); // xlOpenXMLWorkbook.
        WScript.Echo("Created: " + output + " (worksheets: " + count + ")");
    } finally {
        // Attempt every cleanup step, even if an earlier close operation fails.
        var cleanupError = null;
        try {
            if (source !== null) source.Close(false);
        } catch (sourceError) { cleanupError = sourceError; }
        try {
            if (result !== null) result.Close(false);
        } catch (resultError) { cleanupError = resultError; }
        try {
            if (excel !== null) excel.Quit();
        } catch (excelError) { cleanupError = excelError; }
        if (cleanupError !== null) {
            throw new Error("Excel cleanup failed: " + cleanupError.message);
        }
    }
    return 0;
}

// Report errors to CScript output and propagate failure to the Batch launcher.
var exitCode = 1;
try {
    exitCode = main();
} catch (error) {
    WScript.Echo("Error: " + error.message);
}
WScript.Quit(exitCode);