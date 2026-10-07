# Raw and Smart table presentation

`table` and `tableList` keep using `obj_UiRawTable` / `obj_IUiTableSource`. No data ID is required. Raw is the default and writes one two-dimensional buffer through `Range.Value2`. Smart creates an Excel `ListObject` over the written headers and body, before the existing style pipeline runs. Section titles remain outside the ListObject.

## XAML

Table policies are explicitly configured in XAML. The policy elements use the controls namespace, including the property wrapper; the standard page/grid elements use the profiles namespace.

```xml
<controls:tableList name="Results"
                    itemsSource="{Binding Source=Data; Path=Tables}">
    <controls:tableList.tablePolicy>
        <controls:tablePolicy defaultRepresentation="Raw">
            <controls:rule index="2"
                           representation="Smart"
                           excelTableName="tbSummary" />
            <controls:rule index="5, 8-12"
                           representation="Smart" />
            <controls:rule index="10"
                           representation="Raw" />
        </controls:tablePolicy>
    </controls:tableList.tablePolicy>
</controls:tableList>
```

Indices are one-based and local to one TableList. Commas separate indices or inclusive ascending ranges. Only positive decimal integers up to 2147483647 are accepted. Missing positions select nothing. Rules run in document order: later matching rules override representation and, when supplied, the name. A Raw rule clears the inherited name. Rule parsing occurs during control configuration; ranges are stored as intervals rather than expanded into arrays.

`excelTableName` on a rule requires Smart and exactly one index (a singleton interval such as `2-2` is also accepted). One fixed name cannot represent multiple Excel tables. The TableList's optional `representation="Raw|Smart"` establishes its default; a nested `defaultRepresentation` overrides it.

For a single table:

```xml
<controls:table name="Summary"
                itemsSource="{Binding Source=Data; Path=Summary}"
                representation="Smart"
                excelTableName="tbSummary" />
```

Explicit names are restricted to ASCII letters, digits and underscores, with a letter or underscore first, a maximum of 255 characters, and no cell-reference-like names. Name collisions anywhere in the workbook fail with diagnostics. Automatic names look like `tbGenerated_<sheet>_<control>_<index>_<hash>`; components are sanitized and truncated, and a numeric suffix resolves table-name collisions. The hash distinguishes sanitized names and lists on different sheets. These names identify presentation slots, not records or stable business entities.

## Convert an already rendered table

```vba
Dim context As obj_UiRenderContext
Dim diagnostic As String

If Not ex_UiRuntime.fn_TryGetContext(ActiveSheet, context) Then Exit Sub
If Not context.TryEnsureSmart("Results", 2, diagnostic, "tbSummary") Then
    MsgBox diagnostic, vbExclamation
    Exit Sub
End If
```

Omit the last argument for an automatic name. Repeating the operation reuses the owned ListObject and preserves an existing name when the argument is omitted. Conversion uses the saved body/header boundaries and does not write the values again. It requires a successfully rendered table with visible headers and at least one data row. The operation affects the current presentation; the next full render reapplies XAML rules. Do not insert/delete worksheet rows or move the output manually between rendering and this positional API call.

For callers outside a rendered page, `ex_UiTables.fn_TryEnsureSmart(range, owner, diagnostic, optionalName)` accepts an explicit contiguous header/body range. The owner is a worksheet-local presentation key, such as `Results|2`; it is not a required ID on the data model. Use a dedicated output sheet: full UI rendering owns the managed table slots on that sheet and removes slots absent from its new plan.

## Lifecycle and behavior

- Plans are resolved during layout, then checked for ownership, duplicate names and overlapping tables before table-object changes.
- Hidden workbook names with the reserved `_pxTable_` prefix record ownership. Their keys include the sheet CodeName; their comments store the generated ListObject name. They persist with the workbook and survive context disposal or VBA reset. Do not delete or edit these markers manually.
- A table at the same header position is reused and resized; obsolete cells are cleared on shrink. A moved table is unlisted and recreated, so external structured references to moving temporary tables are not guaranteed to survive. Use fixed-position tables for persistent formula references.
- A removed table or a Smart-to-Raw transition unlists the owned object. Raw data is then written from the current source. Foreign ListObjects are never adopted by name; overlapping UI clearing/writes fail.
- Full rendering resets filters and overwrites displayed data from the source. A user sort does not change the source-array order. Smart tables reject `onSelect` / `selectedItem` bindings because the existing selection model depends on physical row order.
- Smart requires unique, nonempty text headers of at most 255 characters and `showHeaders="true"`. Empty Smart sources reserve one blank body row in layout; the source still reports zero data rows. A raw empty table cannot be promoted through the positional API without a new layout.
- The writer clears the built-in TableStyle and leaves presentation to the existing style pipeline. Styles that merge cells inside a Smart table are unsupported by Excel.
- Direct conversion restores the previous ScreenUpdating and EnableEvents values on success and failure. Normal page rendering uses its existing guard. Rendering/conversion also temporarily disable AutoExpandListRange and AutoFillFormulasInLists and restore their previous values. Calculation mode is unchanged.

## Validation

`WorkbookUpdater/Test-UiElements.ps1` contains opt-in Excel integration probes for selectors, rule precedence, mixed rendering, explicit/generated names, title boundaries, repeated rendering, positional conversion, resizing, context recreation, stale-table cleanup and name/header collisions. It starts Excel and executes VBA; run only with explicit authorization. Source/XML checks alone do not establish VBA runtime correctness. The new class/module are covered by the existing `modules.json` wildcards.