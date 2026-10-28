# Workbook snapshot aggregator

Requires Windows Script Host (JScript) and installed desktop Microsoft Excel.

Run Agregator.cmd to create Aggregated.xlsx beside the script. Existing output is never overwritten. An optional argument specifies another output path:

```bat
Agregator.cmd
Agregator.cmd "C:\Reports\Snapshot.xlsx"
Agregator.cmd --validate
Agregator.cmd --help
```

Input selection is explicit in the UTF-8 Agregator.xml file beside Agregator.cmd. Edit sourceFolder, output, sources/workbook and its sheet entries there. Paths are relative to the configuration/script directory. Workbook file paths are resolved against sourceFolder. The default folder is PROJECTS_EXCEL/workbook/Dependencies in this repository.

| Workbook | Selected worksheets, in order |
| --- | --- |
| Накази.xlsx | Накази |
| Установи.xlsx | Військові установи; Лікувальні Заклади |

Each workbook entry has a file attribute and explicit sheet entries with name attributes. Only listed worksheets are included. New sheets are never selected automatically. Missing selected sheets stop the operation. Selecting a name from excludedSheets is rejected; the configuration excludes Переліки and Перелики. Names are written as ordinary Unicode text. The XML parser reads the declared UTF-8 encoding independently of the Windows ANSI code page. Use XML escaping for special characters in attribute values, such as &amp; for an ampersand.

Each selected worksheet is copied in full using Worksheet.Copy and made visible. Output names have a unique numeric prefix and are limited to 31 characters. Excel preserves the worksheet layout, merged cells, ordinary formatting, row heights, column widths and print settings. Tables are converted to ordinary ranges with their formatting retained, and cell formulas are replaced with their results using PasteSpecial. Values keep their Excel types.

Cell formulas, validation and conditional formatting rules are removed from the copied used range. Conditional formatting appearance is not frozen. Full-sheet copying also carries comments, charts, shapes and potentially workbook dependencies; these are not stripped. Chart formulas, defined names and external links are not sanitized. Workbook-level themes can affect appearance. The macro-free .xlsx output does not retain VBA.

Sources are opened read-only in a separate Excel instance with macros and events disabled, external workbook link updates disabled and manual calculation selected. No explicit refresh or recalculation is requested. Connections configured to refresh on open may still refresh; disable those in saved sources when a strictly saved-state snapshot is required. Sources are closed without saving. Excel cleanup is attempted on success and failure.

Screen updating is disabled in the owned Excel instance. Full-sheet copying preserves row heights without a separate row loop. Opening and worksheet-copy durations are printed separately.

The --validate option checks paths and selection configuration without starting Excel; it does not inspect worksheet existence or workbook contents. Exit code is 0 on success and 1 on failure. Large UsedRange areas can slow down copying; protected workbooks may fail. Verify visual fidelity before relying on the output.