# Workbook snapshot aggregator

Requires Windows Script Host (JScript) and installed desktop Microsoft Excel. The xmlStyles mode also uses Windows PowerShell 5.1 and .NET Framework, supplied with Windows; keep ApplyStyles.ps1 and SnapshotStyles.cs beside the launcher. Python and Node.js are not required.

Run Agregator.cmd to create Aggregated.xlsx beside the script. Existing output is never overwritten. An optional argument specifies another output path:

```bat
Agregator.cmd
Agregator.cmd "C:\Reports\Snapshot.xlsx"
Agregator.cmd --validate
Agregator.cmd --help
```

Input selection is explicit in the UTF-8 Agregator.xml file beside Agregator.cmd. The hierarchy is aggregator / sources / source / workbook / sheet. Each source has a dir attribute and contains its own workbooks; multiple source directories are supported. Edit output on aggregator and file/name on workbook/sheet. Source directories and the output path are relative to the configuration/script directory. Workbook file paths are resolved against their source directory. The default source directory is PROJECTS_EXCEL/workbook/Dependencies in this repository.

| Workbook | Selected worksheets, in order |
| --- | --- |
| Накази.xlsx | Накази |
| Установи.xlsx | Військові установи; Лікувальні Заклади |
| ШПО.xlsx | ОС |

Each workbook entry has a file attribute and explicit sheet entries with name attributes. Only listed worksheets are included. New sheets are never selected automatically. Missing selected sheets stop the operation. Names are written as ordinary Unicode text. The XML parser reads the declared UTF-8 encoding independently of the Windows ANSI code page. Use XML escaping for special characters in attribute values, such as &amp; for an ampersand.

The required aggregator attribute copyMode selects values, xmlStyles (currently configured), or sheet. All modes produce visible sheets with unique numeric name prefixes, limited to 31 characters.

In xmlStyles mode, values and smart table structures are transferred through COM just as in values mode. After Excel closes, the .NET helper reads the source ZIP/XML packages and merges ordinary formatting into the output: cell styles, fonts, fills, borders, alignment, number formats, row heights, column widths, hidden row/column states, merged cells, worksheet views, margins and headers/footers. Theme colors and theme font names are resolved per source workbook. Built-in and custom smart table styles are restored. Conditional formatting and validation are never imported. Dates retain their source number formats. The helper writes a temporary package beside the output and replaces the output only after successful processing. Helper failures remove the newly generated value-only output and report an error.

Column widths are adjusted for different source Normal fonts using Windows font metrics; source fonts should be installed. Built-in table styles follow the destination workbook theme, while imported custom table styles have their source theme colors resolved explicitly.

This mode does not copy charts, pictures, comments, drawings, named styles, named ranges, print areas, printer-specific page setup or active table filter criteria. It preserves ordinary formatting, not the visual result of conditional rules. Source files must be unencrypted .xlsx ZIP packages. It reads sources again after COM closes; avoid changing them during aggregation.

In values mode, UsedRange.Value2 is transferred as a single array to a blank worksheet at the original coordinates. No clipboard, source formatting, conditional rules, validation, shapes or comments are copied. Smart tables are recreated at their original ranges with headers, filter buttons and optional totals rows, without source table styles or active filter criteria. Excel assigns new table names. Numeric values remain numeric; dates appear as Excel serial numbers because date formats are not copied. Literal strings beginning with '=' remain text. Formula results are transferred without their formulas. Long text does not pass through a SQL driver.

In sheet mode, each selected worksheet is copied in full using Worksheet.Copy. Excel preserves layout, merged cells, ordinary formatting, row heights, column widths, print settings and smart tables with their styles. Formulas are replaced with their results using PasteSpecial; tables are not converted to ordinary ranges.

In sheet mode, cell formulas, validation and conditional formatting rules are removed from the copied used range. Conditional formatting appearance is not frozen. Full-sheet copying also carries comments, charts, shapes and potentially workbook dependencies; these are not stripped. Chart formulas, defined names and external links are not sanitized. Workbook-level themes can affect appearance. The macro-free .xlsx output does not retain VBA.

Sources are opened read-only in a separate Excel instance with macros and events disabled, external workbook link updates disabled and manual calculation selected. No explicit refresh or recalculation is requested. Connections configured to refresh on open may still refresh; disable those in saved sources when a strictly saved-state snapshot is required. Sources are closed without saving. Excel cleanup is attempted on success and failure.

Screen updating is disabled in the owned Excel instance. Opening and worksheet-copy durations are printed separately. Total elapsed time includes startup, opening, copying, saving and cleanup. The Batch launcher pauses on success and failure, preserving the script exit code; the pause is excluded from the measured time. For unattended execution, invoke cscript.exe //nologo //E:JScript Agregator.cmd directly to bypass the Batch pause.

The --validate option checks paths and selection configuration without starting Excel; it does not inspect worksheet existence or workbook contents. Exit code is 0 on success and 1 on failure. Large UsedRange areas can slow down copying; protected workbooks may fail. Verify visual fidelity before relying on the output.