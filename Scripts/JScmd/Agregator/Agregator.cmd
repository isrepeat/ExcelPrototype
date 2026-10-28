@if (@X)==(@Y) @end /*
@echo off
rem This is a Batch/JScript polyglot: both interpreters read the same file.
rem Batch: the first-line IF compares different strings and skips its command.
rem JScript: @if uses conditional compilation; the false branch ends at @end.
rem The following /* opens a JScript comment containing this Batch launcher.
rem CScript runs this file as JScript; %~f0 is this file and %* forwards arguments.
cscript.exe //nologo //E:JScript "%~f0" %*
rem Preserve the exit code and keep both success and error output visible.
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
        output: requiredAttribute(root, "output"),
        copyMode: requiredAttribute(root, "copyMode"),
        sources: []
    };
    if (config.copyMode !== "values" && config.copyMode !== "sheet" && config.copyMode !== "xmlStyles") {
        throw new Error("Invalid copyMode: " + config.copyMode + ". Expected values, sheet or xmlStyles.");
    }
    var directories = root.selectNodes("sources/source");
    if (directories.length === 0) throw new Error("No source directories configured.");
    for (var d = 0; d < directories.length; d++) {
        var directory = directories.item(d);
        var dir = requiredAttribute(directory, "dir");
        var workbooks = directory.selectNodes("workbook");
        if (workbooks.length === 0) throw new Error("No source workbooks configured: " + dir);
        for (var w = 0; w < workbooks.length; w++) {
            var workbook = workbooks.item(w);
            var entry = { dir: dir, file: requiredAttribute(workbook, "file"), sheets: [] };
            var sheets = workbook.selectNodes("sheet");
            for (var s = 0; s < sheets.length; s++) {
                entry.sheets.push(requiredAttribute(sheets.item(s), "name"));
            }
            config.sources.push(entry);
        }
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

// Recreate table structure after the bulk value transfer, without source styles.
function copyTables(sourceSheet, targetSheet) {
    for (var t = 1; t <= sourceSheet.ListObjects.Count; t++) {
        var original = sourceSheet.ListObjects.Item(t);
        var address = original.Range.Address;
        // Exclude the totals row when defining the table body.
        var bodyRows = original.Range.Rows.Count - (original.ShowTotals ? 1 : 0);
        var tableRange = targetSheet.Range(address).Resize(bodyRows, original.Range.Columns.Count);
        var table = targetSheet.ListObjects.Add(1, tableRange, false, 1);
        table.TableStyle = "";
        table.ShowTotals = original.ShowTotals;
        table.ShowAutoFilter = original.ShowAutoFilter;
        // Enabling totals may insert aggregate formulas; restore their snapshot values.
        if (original.ShowTotals) {
            table.TotalsRowRange.NumberFormat = "@";
            table.TotalsRowRange.Value2 = original.TotalsRowRange.Value2;
            table.TotalsRowRange.NumberFormat = "General";
        }
    }
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
    if (config.copyMode === "xmlStyles") {
        for (var h = 0; h < 2; h++) {
            var helper = fso.BuildPath(folder, h === 0 ? "ApplyStyles.ps1" : "SnapshotStyles.cs");
            if (!fso.FileExists(helper)) throw new Error("Formatting helper not found: " + helper);
        }
    }
    WScript.Echo("Copy mode: " + config.copyMode);
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
        var input = resolvePath(fso.BuildPath(resolvePath(sources[p].dir), sources[p].file));
        if (!fso.FileExists(input)) {
            throw new Error("Input workbook does not exist: " + input);
        }
        if (input.toLowerCase() === output.toLowerCase()) {
            throw new Error("Input and output paths must differ: " + input);
        }
        if (sources[p].sheets.length === 0) {
            throw new Error("No worksheets selected: " + input);
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
                var targetSheet;
                if (config.copyMode !== "sheet") {
                    // Transfer the whole value array without clipboard or formatting rules.
                    placeholder.Activate();
                    targetSheet = result.Worksheets.Add();
                    var sourceRange = sourceSheet.UsedRange;
                    var destination = targetSheet.Range(sourceRange.Address);
                    // Text format prevents strings beginning with '=' from becoming formulas.
                    // Restore General afterwards; numeric variants retain their numeric type.
                    destination.NumberFormat = "@";
                    destination.Value2 = sourceRange.Value2;
                    destination.NumberFormat = "General";
                    copyTables(sourceSheet, targetSheet);
                } else {
                    // Passing Before directly avoids undefined optional COM arguments in WSH.
                    sourceSheet.Copy(placeholder);
                    targetSheet = result.Worksheets.Item(placeholder.Index - 1);
                    // Preserve ListObjects and their styles while freezing cell results.
                    var target = targetSheet.UsedRange;
                    target.Copy();
                    target.PasteSpecial(12); // xlPasteValuesAndNumberFormats.
                    target.FormatConditions.Delete();
                    target.Validation.Delete();
                    excel.CutCopyMode = false;
                }
                targetSheet.Visible = -1; // xlSheetVisible.
                count++;
                // A numeric prefix preserves uniqueness after Excel's 31-character limit.
                var name = count + "_" + fso.GetBaseName(inputs[i]) + "_" + sourceSheet.Name;
                targetSheet.Name = name.replace(/[\\\/\?\*\[\]:']/g, "_").substr(0, 31);

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
        WScript.Echo("Saving output...");
        result.SaveAs(output, 51); // xlOpenXMLWorkbook.
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
    if (config.copyMode === "xmlStyles") {
        WScript.Echo("Applying ordinary formatting from source XML...");
        var shell = new ActiveXObject("WScript.Shell");
        var powershell = shell.ExpandEnvironmentStrings("%SystemRoot%") + "\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
        var command = quoteArgument(powershell) + " -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " +
            quoteArgument(fso.BuildPath(folder, "ApplyStyles.ps1")) + " -ConfigPath " +
            quoteArgument(fso.BuildPath(folder, "Agregator.xml")) + " -OutputPath " + quoteArgument(output);
        // The helper runs only after Excel has closed and released the output package.
        var process = shell.Exec(command);
        while (!process.StdOut.AtEndOfStream) WScript.Echo(process.StdOut.ReadLine());
        var details = process.StdErr.ReadAll();
        while (process.Status === 0) WScript.Sleep(50);
        if (process.ExitCode !== 0) {
            // Never leave a value-only file presented as a successful formatted snapshot.
            fso.DeleteFile(output);
            throw new Error("XML formatting failed: " + details);
        }
    }
    WScript.Echo("Created: " + output + " (worksheets: " + count + ")");
    return 0;
}

function quoteArgument(value) {
    if (/["\r\n]/.test(value)) throw new Error("Invalid command argument.");
    return '"' + value + '"';
}

// Report errors to CScript output and propagate failure to the Batch launcher.
var exitCode = 1;
var aggregationStarted = new Date().getTime();
try {
    exitCode = main();
} catch (error) {
    WScript.Echo("Error: " + error.message);
}
// Include startup, opening, copying, saving and cleanup; exclude the Batch pause.
WScript.Echo("Total elapsed: " + ((new Date().getTime() - aggregationStarted) / 1000).toFixed(3) + " s");
WScript.Quit(exitCode);