@if (@X)==(@Y) @end /*
@echo off
rem The first line is valid in both batch and Microsoft JScript.
rem Batch: IF compares the literal strings (@X) and (@Y), which differ.
rem Its false condition skips the rest of that line; the leading @ hides it.
rem JScript: @if (@X) tests an undefined conditional-compilation variable.
rem Its NaN value is false, so ==(@Y) is skipped until @end on the same line.
rem After @end, the slash-star opens a JScript comment hiding this batch block.
rem The star-slash below closes that comment before the JScript code.
rem CScript reads this same file as JScript; batch waits for it to finish.
rem Both batch exit paths stop before reaching the JScript section.
cscript.exe //nologo //E:JScript "%~f0" %*
if not errorlevel 1 exit /b 0
pause
exit /b 1
*/

// Copy this file under the desired workspace name beside ExcelWorkspaces.json.
// Paths in the UTF-8 JSON may be absolute or relative to that JSON file.
// Use --validate to check configuration without starting Excel.
var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

// Parse JSON without eval or dependencies on a browser's JSON implementation.
function parseJson(text) {
    var pos = 0;
    function fail() { throw new Error("Invalid JSON at character " + (pos + 1) + "."); }
    function space() { while (/[\x20\t\r\n]/.test(text.charAt(pos)) && pos < text.length) pos++; }
    function string() {
        var result = "";
        if (text.charAt(pos++) !== '"') fail();
        while (pos < text.length) {
            var ch = text.charAt(pos++);
            if (ch === '"') return result;
            if (ch === "\\") {
                ch = text.charAt(pos++);
                if (ch === "u") {
                    var hex = text.substr(pos, 4);
                    if (!/^[0-9a-fA-F]{4}$/.test(hex)) fail();
                    result += String.fromCharCode(parseInt(hex, 16));
                    pos += 4;
                } else {
                    var escapes = {'"':'"', '\\':'\\', '/':'/', b:'\b', f:'\f', n:'\n', r:'\r', t:'\t'};
                    if (!Object.prototype.hasOwnProperty.call(escapes, ch)) fail();
                    result += escapes[ch];
                }
            } else {
                if (ch.charCodeAt(0) < 32) fail();
                result += ch;
            }
        }
        fail();
    }
    function value() {
        space();
        var ch = text.charAt(pos);
        if (ch === '"') return string();
        if (ch === "{" || ch === "[") {
            var object = ch === "{";
            var result = object ? {} : [];
            var end = object ? "}" : "]";
            pos++;
            space();
            if (text.charAt(pos) === end) { pos++; return result; }
            while (true) {
                space();
                if (object) {
                    var key = string();
                    if (Object.prototype.hasOwnProperty.call(result, key)) fail();
                    space();
                    if (text.charAt(pos++) !== ":") fail();
                    result[key] = value();
                } else result.push(value());
                space();
                ch = text.charAt(pos++);
                if (ch === end) return result;
                if (ch !== ",") fail();
            }
        }
        var match = /^(true|false|null|-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)/.exec(text.substr(pos));
        if (!match) fail();
        pos += match[0].length;
        if (match[0] === "true") return true;
        if (match[0] === "false") return false;
        if (match[0] === "null") return null;
        return Number(match[0]);
    }
    var result = value();
    space();
    if (pos !== text.length) fail();
    return result;
}

try {
    var folder = fso.GetParentFolderName(WScript.ScriptFullName);
    var workspace = fso.GetBaseName(WScript.ScriptFullName);
    var configPath = fso.BuildPath(folder, "ExcelWorkspaces.json");
    if (!fso.FileExists(configPath)) throw new Error("Configuration not found: " + configPath);
    var stream = new ActiveXObject("ADODB.Stream");
    var text;
    try {
        stream.Type = 2;
        stream.Charset = "utf-8";
        stream.Open();
        stream.LoadFromFile(configPath);
        text = stream.ReadText(-1).replace(/^\uFEFF/, "");
    } finally { stream.Close(); }
    var config = parseJson(text);
    if (!config || config instanceof Array || typeof config !== "object") {
        throw new Error("Configuration must be a JSON object.");
    }
    if (!Object.prototype.hasOwnProperty.call(config, workspace)) {
        throw new Error("Workspace key not found: " + workspace);
    }
    var entries = config[workspace];
    if (!(entries instanceof Array) || entries.length === 0) {
        throw new Error("Workspace must contain a nonempty array of workbook paths: " + workspace);
    }
    var paths = [];
    var seen = {};
    for (var i = 0; i < entries.length; i++) {
        if (typeof entries[i] !== "string" || /^\s*$/.test(entries[i])) {
            throw new Error("Invalid workbook path at index " + i + ".");
        }
        var path = shell.ExpandEnvironmentStrings(entries[i]).replace(/\//g, "\\");
        if (!/^(?:[a-zA-Z]:\\|\\\\)/.test(path)) {
            if (/^(?:[a-zA-Z]:|\\)/.test(path)) throw new Error("Use a full or relative path: " + path);
            path = fso.BuildPath(folder, path);
        }
        path = fso.GetAbsolutePathName(path);
        if (!fso.FileExists(path)) throw new Error("Workbook not found: " + path);
        var identity = "$" + path.toLowerCase();
        if (Object.prototype.hasOwnProperty.call(seen, identity)) throw new Error("Duplicate workbook: " + path);
        seen[identity] = true;
        paths.push(path);
    }
    if (WScript.Arguments.length > 1 || (WScript.Arguments.length === 1 && WScript.Arguments(0) !== "--validate")) {
        throw new Error("Only the optional --validate argument is supported.");
    }
    if (WScript.Arguments.length === 1) {
        WScript.Echo("Validated " + workspace + ": " + paths.length + " workbook(s).");
        WScript.Quit(0);
    }

    // Keep already opened workbooks visible if a later workbook fails to open.
    var excel = new ActiveXObject("Excel.Application");
    excel.Visible = true;
    excel.UserControl = true;
    for (var j = 0; j < paths.length; j++) excel.Workbooks.Open(paths[j]);
} catch (error) {
    WScript.Echo("Cannot open Excel workspace: " + error.message);
    WScript.Quit(1);
}