# External table query API

The API reads external Excel tables independently of the UI. All components are included by the existing `vba/modules.json` wildcard rules. Import or reload the project sources to make the API available in a workbook; this change does not update the tracked workbooks.

## Usage

```vba
Dim source As New obj_TableSource
Dim query As New obj_TableQuery
Dim service As New obj_TableQueryService
Dim result As obj_DataTable
Dim diagnostic As String

If Not source.Initialize() Then Exit Sub
If Not query.Initialize() Then Exit Sub
If Not service.Initialize() Then Exit Sub

source.WorkbookPath = "C:\Data\Personnel.xlsx"
source.TableName = "tblPersonnel"
query.AddColumn "ID"
query.AddColumn "Name"
query.AddCondition "Amount", QueryGreaterOrEqual, 10, QueryNumber
query.AddOrderBy "Amount", True, QueryNumber
query.AddOrderBy "ID", False, QueryNumber
query.Limit = 20

If Not service.TryExecute(source, query, result, diagnostic) Then
    MsgBox diagnostic, vbExclamation
    service.Dispose
    Exit Sub
End If

' Успешный запрос может вернуть ноль строк, сохранив заголовки.
Debug.Print result.RowCount
service.Dispose
```

For a plain range, omit `TableName` and specify `SheetName` and `RangeAddress`, for example `A1:F500`. `RangeAddress="auto"` discovers a range starting at A1, ending at the last populated row and the last header column in row 1; it includes newly added rows and ignores formatting-only extent. Auto discovery requires headers in row 1. It uses the same workbook-session metadata path for both backends. The first row contains headers. Ranges must be contiguous; every header must be nonempty and unique after trimming and case-insensitive comparison. With a ListObject, the header must be shown; the totals row is excluded. Queries read all data rows, including rows hidden by worksheet filters.

## Contracts

| Component | Responsibility |
| --- | --- |
| `obj_TableSource` | Workbook path and ListObject name or explicit range |
| `obj_TableQuery` | Projection, filter tree, ordered sort keys, limit |
| `obj_QueryCondition` | Typed comparison and independent text options |
| `obj_QueryGroup` | AND by default; OR when `MatchAny=True`; nested groups |
| `obj_QueryOrder` | Column, type, direction, text options |
| `obj_DataTable` | Zero-based headers, one-based two-dimensional values, dimensions |
| `obj_ITableQueryExecutor` | `TryExecute` only; lifecycle belongs to the concrete executor |
| `obj_ExcelQueryExecutor` | Live or read-only Excel object-model access |
| `obj_AdoQueryExecutor` | Saved-file access through ACE OLE DB |
| `obj_TableQueryService` | Backend selection and resource lifecycle |

Leaving the selected columns empty means select all. `Limit=0` means unlimited; negative limits fail. Sorting is stable and runs before the limit. Specify a unique final sort key when deterministic selection among equal values matters. There is no physical-row-order guarantee for ADO without explicit ordering.

Supported operators: `QueryEquals`, `QueryNotEquals`, `QueryContains`, `QueryStartsWith`, `QueryEndsWith`, `QueryIsEmpty`, `QueryIsNotEmpty`, `QueryGreater`, `QueryGreaterOrEqual`, `QueryLess`, `QueryLessOrEqual`. Contains/prefix/suffix treat `%`, `_`, `[` and other characters literally.

Types are explicit: `QueryText`, `QueryNumber`, `QueryDate`, `QueryBoolean`. Numeric strings follow the current VBA locale. Date comparisons accept VBA Date values and Excel serial values in the 1900 date system. For dates, use `DateSerial` rather than ambiguous date strings. Values retain the types supplied by the reader; ADO may return Date where Excel Value2 returns a serial Double.

`NormalizeText=True` replaces NBSP with a space, removes CR/LF/tab and trims outer spaces. It does not change case sensitivity. `CaseSensitive=False` is the default and uses the VBA text comparison rules. Null, Empty and empty strings are blank; whitespace becomes blank only with normalization enabled. Blank values fail ordinary comparisons, including not-equal; use explicit empty predicates. Excel error values in filter or sort columns fail the query with diagnostics. Empty values sort first ascending and last descending.

For OR, create `obj_QueryGroup`, set `MatchAny=True`, add conditions, then call `query.Filter.AddGroup group`. Empty AND groups match every row; empty OR groups match none. Cyclic group references are rejected. The API is synchronous; do not mutate query objects during execution.

## Backend and resource behavior

`service.Backend=QueryAuto` reads a book open in the current Excel Application through the Excel executor, including unsaved changes. Otherwise it selects ADO. Paths resolve against the current working directory; absolute paths are recommended.

`QueryExcel` opens a closed source in an owned hidden Excel instance, read-only, with macros/events and link updates disabled. It closes only owned books. User books are never saved or closed. `QueryAdo` reads the saved file even if the current Excel instance has an unsaved copy. Books open in another Excel instance are not considered live by Auto.

The current ADO implementation obtains saved table/range metadata using a temporary read-only Excel instance. It then closes that instance and reads data using ACE SQL, selecting only output/filter/sort columns. Both executors apply the filter tree, stable sorting and limit through the same in-memory evaluator. SQL WHERE/ORDER BY/TOP pushdown is deliberately absent from this first implementation: all source rows in the requested columns are fetched. This avoids a second set of comparison rules but costs memory and transfer time for large sources. There is no connection or result cache; each request releases its resources.

ADO requires the matching installed `Microsoft.ACE.OLEDB.12.0` provider and supports `.xls`, `.xlsx`, `.xlsm`, `.xlsb`. ACE inference can change types or lose mixed-column/long-text data before the shared evaluator receives it. For exact Excel values, explicitly select `QueryExcel`. There is no silent fallback after an ADO failure. Password-protected workbooks and interactive credential prompts are outside this API.

`TryExecute=False` returns `Nothing` and an error description. The caller must display the diagnostic and stop the operation. `True` with zero rows is a valid result. `result.TryCreateUiTable` converts the result into the existing `obj_UiRawTable`, including zero rows; query execution has no dependency on page controllers or bindings.

## Validation

Run `WorkbookUpdater/Test-TableQueries.ps1` on Windows with Excel, ACE and trusted access to the VBA project object model. It creates and preserves temporary fixture workbooks, imports the API into a separate runner, and executes the same assertions through both backends. It does not modify repository workbooks.

The checks cover typed filters, AND/OR, normalization, case-sensitive equality, literal text search, sorting, limits, exclusion of totals, empty results and UI adaptation, missing-column diagnostics, unsaved Auto versus saved ADO reads, scalar ranges, and execution through the interface, rejection before initialization, rejection of duplicate initialization and rejection of initialization after disposal. All concrete API classes have Lifecycle callbacks and public Initialize/Dispose methods; interface classes contain only the necessary contract. Constructors create only minimal internal state and never call public Initialize or open resources. Call Initialize explicitly and check its Boolean result before use; services/executors reject calls before initialization. Both m_isInitialized and m_isDisposed start False. Successful Initialize sets m_isInitialized=True; Dispose sets it False and m_isDisposed=True. Initialize rejects already initialized or disposed instances; create a new instance for another lifecycle. Dispose is safe to repeat, and Class_Terminate calls Me.Dispose. DataTable.Initialize retains its schema/data arguments.