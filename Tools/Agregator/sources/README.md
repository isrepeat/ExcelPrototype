# Workbook snapshot aggregator

Requires 64-bit Windows, desktop Microsoft Excel and .NET Framework 4.5 or later. The build script uses the Windows .NET Framework C# compiler; no .NET SDK, PowerShell, Python or Node.js is required.

Run CompileAgregator.cmd in the parent directory. It only compiles sources/Agregator.cs and sources/SnapshotStyles.cs into Agregator.exe, reports the build result and pauses on success or failure. It never runs the program or aggregates workbooks. Source code and this documentation stay in sources; the build script, executable and runtime configuration Agregator.xml are top-level files. Run Agregator.exe separately after a successful build. The executable reports SUCCESS or FAILED and pauses before closing when console input is interactive. Use --no-pause for unattended runs; redirected input also disables the pause. Errors and stack traces are appended to UTF-8 Agregator.log beside the executable. Failure to write the log is reported in the console.

```bat
CompileAgregator.cmd
Agregator.exe
Agregator.exe "C:\Reports\Snapshot.xlsx"
Agregator.exe --validate
Agregator.exe --help
Agregator.exe --no-pause
Agregator.exe --config Agregator.Network.xml
Agregator.exe --config Agregator.Network.xml --validate
```

The UTF-8 configuration is Agregator.xml beside the executable. The hierarchy is aggregator / sources / source / workbook / sheet. Each source has a dir attribute and contains explicitly selected workbooks and sheets. Source directories and the output path (including a command-line override) are resolved relative to the configuration directory, regardless of the caller's current directory. Workbook paths are resolved relative to their source directory. The default result is Aggregated.xlsx beside the executable. Multiple source directories are supported. Missing files, directories, selected sheets or required attributes cause an error. Input workbooks cannot be used as the output path. Close the result workbook in Excel before replacing it.

Publication uses same-directory renames compatible with SMB, without File.Replace. An existing result is first renamed to a unique previous-copy file; the completed snapshot is then renamed to the output name. A publication failure triggers restoration of the previous result. If restoration also fails, the recovery path is printed and recorded in the error log, and the previous copy is retained. This two-step replacement is not atomic: avoid concurrent runs or changing the destination during publication. XML formatting only overwrites a private intermediate snapshot. Failed processing never intentionally overwrites the prior user result.

The current selection is Накази.xlsx (Накази), Установи.xlsx (Військові установи, Лікувальні Заклади), and ШПО.xlsx (ОС). Original worksheet names are preserved. If a name is already used (case-insensitive), the workbook name is prefixed; further conflicts get numeric suffixes. Generated names are limited to 31 characters. Selected hidden worksheets become visible in the result.

Use --config to select another UTF-8 XML configuration; its path is relative to the executable, or absolute. Source directories and output paths are relative to the selected configuration directory. Without --config, Agregator.xml beside the executable is used. Agregator.Network.xml is prepared for deployment to \\athenaps\Облік\Tools\Agregator: its source dir ../.. resolves to \\athenaps\Облік. It selects ШПО.xlsx (ОС), Виплати.xlsx (Виплати) and Порушення.xlsx (Порушення), in that order. Its result is Aggregated.xlsx beside the network configuration. Keep the network configuration beside the deployed executable. Validate network access with --config Agregator.Network.xml --validate before aggregation; worksheet existence is checked during aggregation.

The required copyMode attribute supports:

- values: bulk UsedRange.Value2 transfer to blank sheets, preserving value types and long text, with smart tables recreated without source formatting. Dates appear as numeric Excel serial values. Literal strings beginning with '=' remain text. Source table names and active filter criteria are not preserved.
- xmlStyles (configured): the same COM value transfer followed by direct ZIP/XML formatting. Fonts, fills, borders, alignment, number formats, row heights, column widths, hidden row/column states, merged cells, views, margins, headers/footers and smart table styles are restored without importing conditional rules or validation.
- sheet: full Worksheet.Copy, followed by replacing formulas with results and deleting conditional formatting and validation. Smart tables remain intact. This can be very slow for worksheets with thousands of conditional rules. Charts, shapes, comments and workbook dependencies may be carried along; chart formulas and external references are not sanitized.

In xmlStyles mode, source theme colors and theme font names are resolved explicitly for imported ordinary and custom table styles. Column widths are adjusted for differing Normal fonts using Windows font metrics; source fonts should be installed. Built-in table styles follow the output workbook theme. This mode does not copy drawings, pictures, comments, named styles, named ranges, print areas, printer-specific page setup or active table filter criteria. Conditional formatting appearance is not frozen. Sources must be unencrypted .xlsx ZIP packages and must not change during aggregation.

Sources are opened read-only in a separate Excel instance with macros and events disabled, external workbook link updates disabled and manual calculation selected. No explicit refresh or recalculation is requested. Connections set to refresh on open may still refresh. Sources are closed without saving. Generated snapshots are macro-free .xlsx files.

The application builds a temporary snapshot beside the destination and publishes it only after successful processing and cleanup. Failed runs remove temporary snapshots. It reports opening, copying, XML processing and total application time, including Excel startup and cleanup. Compilation and the build script's pause are separate from aggregation. Exit codes are 0 for success and 1 for errors. The --validate option checks configuration and paths without starting Excel; it does not inspect workbook contents or selected worksheet existence.

Console output and redirected logs use UTF-8 without a BOM, so Ukrainian workbook and sheet names are not restricted by the default Windows OEM code page. Use a console font with Cyrillic support.