# Workbook snapshot aggregator

Requires 64-bit Windows, desktop Microsoft Excel and .NET Framework 4.5 or later. The launcher uses the Windows .NET Framework C# compiler; no .NET SDK, PowerShell, Python or Node.js is required.

Run Agregator.cmd in the parent directory. It compiles sources/Agregator.cs and sources/SnapshotStyles.cs into Agregator.exe beside the launcher, starts it and pauses on success or failure. Source code and this documentation stay in sources; the launcher, executable and runtime configuration Agregator.xml are top-level files. Build errors stop execution; an old executable is never launched after a failed build. After the first successful build, Agregator.exe can also run directly for unattended use without a pause.

```bat
Agregator.cmd
Agregator.cmd "C:\Reports\Snapshot.xlsx"
Agregator.cmd --validate
Agregator.cmd --help
```

The UTF-8 configuration is Agregator.xml beside the executable. The hierarchy is aggregator / sources / source / workbook / sheet. Each source has a dir attribute and contains explicitly selected workbooks and sheets. Source directories and the output path (including a command-line override) are resolved relative to the configuration directory, regardless of the caller's current directory. Workbook paths are resolved relative to their source directory. The default result is Aggregated.xlsx beside the executable. Multiple source directories are supported. Missing files, directories, selected sheets or required attributes cause an error. Existing output is never overwritten.

The current selection is Накази.xlsx (Накази), Установи.xlsx (Військові установи, Лікувальні Заклади), and ШПО.xlsx (ОС). Each sheet gets a unique numeric prefix and a name limited to 31 characters. Selected hidden worksheets become visible in the result.

The required copyMode attribute supports:

- values: bulk UsedRange.Value2 transfer to blank sheets, preserving value types and long text, with smart tables recreated without source formatting. Dates appear as numeric Excel serial values. Literal strings beginning with '=' remain text. Source table names and active filter criteria are not preserved.
- xmlStyles (configured): the same COM value transfer followed by direct ZIP/XML formatting. Fonts, fills, borders, alignment, number formats, row heights, column widths, hidden row/column states, merged cells, views, margins, headers/footers and smart table styles are restored without importing conditional rules or validation.
- sheet: full Worksheet.Copy, followed by replacing formulas with results and deleting conditional formatting and validation. Smart tables remain intact. This can be very slow for worksheets with thousands of conditional rules. Charts, shapes, comments and workbook dependencies may be carried along; chart formulas and external references are not sanitized.

In xmlStyles mode, source theme colors and theme font names are resolved explicitly for imported ordinary and custom table styles. Column widths are adjusted for differing Normal fonts using Windows font metrics; source fonts should be installed. Built-in table styles follow the output workbook theme. This mode does not copy drawings, pictures, comments, named styles, named ranges, print areas, printer-specific page setup or active table filter criteria. Conditional formatting appearance is not frozen. Sources must be unencrypted .xlsx ZIP packages and must not change during aggregation.

Sources are opened read-only in a separate Excel instance with macros and events disabled, external workbook link updates disabled and manual calculation selected. No explicit refresh or recalculation is requested. Connections set to refresh on open may still refresh. Sources are closed without saving. Generated snapshots are macro-free .xlsx files.

The application builds a temporary snapshot beside the destination and publishes it only after successful processing and cleanup. Failed runs remove temporary snapshots. It reports opening, copying, XML processing and total application time, including Excel startup and cleanup; compilation and the launcher's pause are outside the measured application time. Exit codes are 0 for success and 1 for errors. The --validate option checks configuration and paths without starting Excel; it does not inspect workbook contents or selected worksheet existence.

Console output and redirected logs use UTF-8 without a BOM, so Ukrainian workbook and sheet names are not restricted by the default Windows OEM code page. Use a console font with Cyrillic support.