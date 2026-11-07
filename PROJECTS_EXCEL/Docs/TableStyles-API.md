# Native Excel table styles

`ex_UiTableStyles.fn_InstallStyle(workbook, xamlPath, styleId, styleName)` loads a separate XAML catalog and creates a native style in the destination workbook's `TableStyles`. The returned name can be assigned to one or more `ListObject.TableStyle` properties. Install each display name once per destination workbook; a conflicting existing name is an error, never a silent replacement.

The root is `tableStyleCatalog` in `urn:excelprototype:profiles`, with `version="1"`. Each `tableStyle` has a technical `name`; its display name comes from the caller's configuration. For an example, see `..\ui\PersonnelAtDisposalDaysCalculation\TableStyles.xaml`.

Each `element` requires `type`: `WholeTable`, `HeaderRow`, `TotalRow`, `FirstColumn`, `LastColumn`, `RowStripe1`, `RowStripe2`, `ColumnStripe1` or `ColumnStripe2`. Supported optional attributes are `backColor` and `textColor` in `#RRGGBB`, `fontBold` and `fontItalic` as `true`/`false`, and `stripeSize` as an integer from 1 to 9 on stripe elements only. The importer rejects unsupported element attributes, duplicate element types, missing selected styles and invalid values before creating the destination style. If style installation fails, the partially created style is deleted and the original error is raised.

The importer disables DTDs and external resolution. It does not open or save workbooks, change cells, select a style or enable table options. The exporter controls row/column stripes and first/last-column emphasis. Native table styles cannot express arbitrary individual-column accents, cell alignment, column widths, date formats or worksheet backgrounds; keep those separate. Direct table fills and font colors override gallery styles, so clear decorative direct formatting before applying the returned style and restore required number formats afterward.

PersonnelAtDisposalDaysCalculation resolves `export.styleSource` relative to the configured UI root, selects `export.styleId` from the catalog and uses `export.styleName` as the gallery label. Both export modes embed the same declaration. The saved XLSX contains the native style; the XAML file is required only during generation, not when opening the result.

Source/XML validation does not establish Excel runtime correctness. Workbook import, compilation and macro execution require explicit user authorization.