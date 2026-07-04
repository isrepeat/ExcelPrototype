VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ExporterToDailyScope"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IDataExporter

Private m_IsDisposed As Boolean
Private m_TargetWorkbookPath As String
Private m_TargetSheetName As String
Private m_TargetRangeStartMarker As String
Private m_TargetRangeEndMarker As String

Private Const SAVE_ALREADY_OPEN_WORKBOOK As Boolean = False
Private Const EXPORT_CONFIG_PREFIX As String = "Export."
Private Const EXPORT_CLASS_SUFFIX As String = ".ExporterClass"
Private Const EXPORT_FILE_PATH_SUFFIX As String = ".FilePath"
Private Const EXPORT_SHEET_NAME_SUFFIX As String = ".SheetName"
Private Const EXPORT_RANGE_START_MARKER_SUFFIX As String = ".RangeStartMarker"
Private Const EXPORT_RANGE_END_MARKER_SUFFIX As String = ".RangeEndMarker"

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
    
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IDataExporter_Export( _
    ByVal sourceTable As obj_TableDynamic _
) As Boolean
    obj_IDataExporter_Export = Me.Export(sourceTable)
End Function

' //
' // API
' //
Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    Dim cfgParserBase As obj_CfgParserBase
    Dim configEntries As Collection
    Dim cfgMap As Object
    Dim exportAlias As String

    m_IsDisposed = False
    m_TargetWorkbookPath = VBA.vbNullString
    m_TargetSheetName = VBA.vbNullString
    m_TargetRangeStartMarker = VBA.vbNullString
    m_TargetRangeEndMarker = VBA.vbNullString

    If configTable Is Nothing Then
        Initialize = True
        Exit Function
    End If

    exportAlias = private_TryResolveExportAliasFromConfigTable(configTable)
    If VBA.Len(exportAlias) = 0 Then Exit Function

    Set cfgParserBase = New obj_CfgParserBase
    If Not cfgParserBase.Initialize(configTable) Then Exit Function
    If Not cfgParserBase.TryGetConfigEntries(configEntries) Then Exit Function
    If Not cfgParserBase.BuildConfigDictionary(configEntries, cfgMap) Then Exit Function

    m_TargetWorkbookPath = VBA.Trim$(cfgParserBase.GetOptionalConfigValue(cfgMap, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_FILE_PATH_SUFFIX, VBA.vbNullString))
    m_TargetSheetName = VBA.Trim$(cfgParserBase.GetOptionalConfigValue(cfgMap, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_SHEET_NAME_SUFFIX, VBA.vbNullString))
    m_TargetRangeStartMarker = VBA.Trim$(cfgParserBase.GetOptionalConfigValue(cfgMap, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_START_MARKER_SUFFIX, VBA.vbNullString))
    m_TargetRangeEndMarker = VBA.Trim$(cfgParserBase.GetOptionalConfigValue(cfgMap, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_END_MARKER_SUFFIX, VBA.vbNullString))

    Initialize = True
End Function

Public Function TryGetSectionTypeOptions(ByRef outSectionTypeOptions As Collection) As Boolean
    Set outSectionTypeOptions = private_BuildSectionTypeOptions()
    If outSectionTypeOptions Is Nothing Then Exit Function
    TryGetSectionTypeOptions = (outSectionTypeOptions.Count > 0)
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next

    On Error GoTo 0
End Sub

Public Function Export( _
    ByVal sourceTable As obj_TableDynamic _
) As Boolean
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim targetTable As ListObject
    Dim insertedRow As ListRow
    Dim targetRowRange As Range
    Dim targetSectionCaption As String
    Dim targetSheetName As String
    Dim openedByExporter As Boolean
    Dim fastModeStarted As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalculation As XlCalculation
    Dim undoSnapshot As Object

    On Error GoTo EH
    If m_IsDisposed Then
        VBA.MsgBox "PrototypeNew: DailyScope exporter is disposed.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If
    If Not private_ValidateSourceTable(sourceTable) Then Exit Function
    If Not private_TryResolveTargetSectionCaption(sourceTable, targetSectionCaption) Then Exit Function

    private_BeginFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    fastModeStarted = True

    If Not private_TryOpenTargetWorkbook(targetWb, openedByExporter) Then GoTo CleanFail
    targetSheetName = private_ResolveTargetWorksheetName()
    If VBA.Len(targetSheetName) = 0 Then GoTo CleanFail
    If Not private_TryGetWorksheet(targetWb, targetSheetName, targetWs) Then GoTo CleanFail
    If Not private_TryFindConfiguredTargetTable(targetWs, targetTable) Then GoTo CleanFail

    If Not private_TryGetSectionWriteRowRange(targetTable, targetSectionCaption, targetRowRange, insertedRow) Then GoTo CleanFail
    If Not private_TryBuildUndoSnapshot(targetWb, targetWs, targetTable, targetRowRange, insertedRow, sourceTable, undoSnapshot) Then GoTo CleanFail

    If Not private_TryWriteSourceRow(sourceTable, targetTable, targetRowRange) Then GoTo CleanFail
    If Not rt_ExportUndo.fn_PushDailyScopeExportUndo(undoSnapshot) Then
        VBA.MsgBox "PrototypeNew: failed to register DailyScope undo action.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        GoTo CleanFail
    End If

    If Not openedByExporter And SAVE_ALREADY_OPEN_WORKBOOK Then targetWb.Save
    Export = True
    GoTo CleanExit

CleanFail:
    Export = False
    If Not insertedRow Is Nothing Then
        On Error Resume Next
        insertedRow.Delete
        On Error GoTo 0
    End If

CleanExit:
    If openedByExporter Then
        On Error Resume Next
        targetWb.Close SaveChanges:=Export
        On Error GoTo 0
    End If
    If fastModeStarted Then private_RestoreFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    Exit Function

EH:
    VBA.MsgBox "PrototypeNew: DailyScope test export failed. " & Err.Description, VBA.vbExclamation, "PrototypeNew / DailyScope export"
    On Error Resume Next
    If openedByExporter Then targetWb.Close SaveChanges:=False
    If fastModeStarted Then private_RestoreFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    On Error GoTo 0
End Function

Private Function private_TryBuildUndoSnapshot( _
    ByVal targetWb As Workbook, _
    ByVal targetWs As Worksheet, _
    ByVal targetTable As ListObject, _
    ByVal targetRowRange As Range, _
    ByVal insertedRow As ListRow, _
    ByVal sourceTable As obj_TableDynamic, _
    ByRef outSnapshot As Object _
) As Boolean
    Dim rowIndex As Long
    Dim isInsertedRow As Boolean
    Dim columnIndexes As Collection
    Dim oldValues As Collection

    Set outSnapshot = Nothing
    If targetWb Is Nothing Then Exit Function
    If targetWs Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If targetRowRange Is Nothing Then Exit Function
    If sourceTable Is Nothing Then Exit Function

    isInsertedRow = Not (insertedRow Is Nothing)

    If isInsertedRow Then
        rowIndex = insertedRow.Index
    Else
        If Not private_TryResolveListRowIndexByRange(targetTable, targetRowRange, rowIndex) Then Exit Function
        If Not private_TryCaptureRowUndoValues(sourceTable, targetTable, targetRowRange, columnIndexes, oldValues) Then Exit Function
    End If

    Set outSnapshot = ex_Helpers.fn_CreateDictionaryTextCompare()
    If outSnapshot Is Nothing Then Exit Function

    outSnapshot("WorkbookPath") = targetWb.FullName
    outSnapshot("WorkbookName") = targetWb.Name
    outSnapshot("SheetName") = targetWs.Name
    outSnapshot("TableName") = targetTable.Name
    outSnapshot("RowIndex") = rowIndex
    outSnapshot("InsertedRow") = isInsertedRow

    If Not isInsertedRow Then
        Set outSnapshot("ColumnIndexes") = columnIndexes
        Set outSnapshot("OldValues") = oldValues
    End If

    private_TryBuildUndoSnapshot = True
End Function

Private Function private_TryResolveListRowIndexByRange( _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByRef outRowIndex As Long _
) As Boolean
    outRowIndex = 0
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function
    If targetTable.DataBodyRange Is Nothing Then Exit Function

    outRowIndex = rowRange.Row - targetTable.DataBodyRange.Row + 1
    If outRowIndex <= 0 Then Exit Function
    If outRowIndex > targetTable.ListRows.Count Then Exit Function
    private_TryResolveListRowIndexByRange = True
End Function

Private Function private_TryCaptureRowUndoValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByRef outColumnIndexes As Collection, _
    ByRef outOldValues As Collection _
) As Boolean
    Dim sourceColIndex As Long
    Dim sourceColumn As obj_Column
    Dim sourceColumnName As String
    Dim targetColumnIndex As Long
    Dim targetColumnIndexByName As Object

    Set outColumnIndexes = Nothing
    Set outOldValues = Nothing
    If sourceTable Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    Set outColumnIndexes = New Collection
    Set outOldValues = New Collection

    If Not private_TryBuildTargetColumnIndexByName(targetTable, targetColumnIndexByName) Then Exit Function

    For sourceColIndex = 1 To sourceTable.ColumnCount
        Set sourceColumn = sourceTable.Columns.Item(sourceColIndex)
        If sourceColumn Is Nothing Then GoTo ContinueSourceColumn

        sourceColumnName = VBA.Trim$(VBA.CStr(sourceColumn.Name))
        If VBA.Len(sourceColumnName) = 0 Then GoTo ContinueSourceColumn
        If Not targetColumnIndexByName.exists(sourceColumnName) Then GoTo ContinueSourceColumn

        targetColumnIndex = VBA.CLng(targetColumnIndexByName(sourceColumnName))
        If targetColumnIndex <= 0 Then GoTo ContinueSourceColumn

        outColumnIndexes.Add targetColumnIndex
        outOldValues.Add rowRange.Cells(1, targetColumnIndex).Value2

ContinueSourceColumn:
    Next sourceColIndex

    private_TryCaptureRowUndoValues = True
End Function

' //
' // Internal
' //
Private Function private_TryOpenTargetWorkbook( _
    ByRef outWorkbook As Workbook, _
    ByRef outOpenedByExporter As Boolean _
) As Boolean
    Dim wb As Workbook
    Dim resolvedPath As String
    Dim targetWorkbookName As String

    Set outWorkbook = Nothing
    outOpenedByExporter = False
    resolvedPath = VBA.Trim$(m_TargetWorkbookPath)

    If VBA.Len(resolvedPath) > 0 Then
        Set outWorkbook = private_FindOpenWorkbookByPath(resolvedPath)
        If Not outWorkbook Is Nothing Then
            private_TryOpenTargetWorkbook = True
            Exit Function
        End If
    End If

    targetWorkbookName = private_ExtractWorkbookNameFromPath(resolvedPath)
    If VBA.Len(targetWorkbookName) > 0 Then
        For Each wb In Application.Workbooks
            If VBA.StrComp(wb.Name, targetWorkbookName, VBA.vbTextCompare) = 0 Then
                Set outWorkbook = wb
                private_TryOpenTargetWorkbook = True
                Exit Function
            End If
        Next wb
    End If

    If VBA.Len(resolvedPath) = 0 Then
        VBA.MsgBox "PrototypeNew: target workbook path is not configured for DailyScope export.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    If VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: DailyScope target workbook was not found: " & resolvedPath, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    Set outWorkbook = Application.Workbooks.Open(resolvedPath)
    outOpenedByExporter = True
    private_TryOpenTargetWorkbook = Not outWorkbook Is Nothing
End Function

Private Function private_TryGetWorksheet( _
    ByVal wb As Workbook, _
    ByVal worksheetName As String, _
    ByRef outWorksheet As Worksheet _
) As Boolean
    Set outWorksheet = Nothing
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set outWorksheet = wb.Worksheets(worksheetName)
    On Error GoTo 0

    If outWorksheet Is Nothing Then
        VBA.MsgBox "PrototypeNew: DailyScope worksheet was not found: " & worksheetName, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    private_TryGetWorksheet = True
End Function

Private Function private_TryFindConfiguredTargetTable( _
    ByVal ws As Worksheet, _
    ByRef outTable As ListObject _
) As Boolean
    Dim tableObj As ListObject
    Dim targetRange As Range
    Dim startCell As Range

    Set outTable = Nothing
    If ws Is Nothing Then Exit Function
    If Not private_TryResolveConfiguredTargetRange(ws, targetRange) Then Exit Function
    If targetRange Is Nothing Then Exit Function
    Set startCell = targetRange.Cells(1, 1)

    For Each tableObj In ws.ListObjects
        If tableObj Is Nothing Then GoTo ContinueTable
        If tableObj.Range Is Nothing Then GoTo ContinueTable
        If tableObj.Range.Row = startCell.Row And tableObj.Range.Column = startCell.Column Then
            Set outTable = tableObj
            private_TryFindConfiguredTargetTable = True
            Exit Function
        End If
        If Not Application.Intersect(tableObj.Range, targetRange) Is Nothing Then
            Set outTable = tableObj
            private_TryFindConfiguredTargetTable = True
            Exit Function
        End If

ContinueTable:
    Next tableObj

    VBA.MsgBox "PrototypeNew: target table was not found for configured markers on sheet '" & ws.Name & "'.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
End Function

Private Function private_ResolveTargetWorksheetName() As String
    private_ResolveTargetWorksheetName = private_ExtractSheetNameToken(m_TargetSheetName)
    If VBA.Len(private_ResolveTargetWorksheetName) = 0 Then
        VBA.MsgBox "PrototypeNew: target worksheet is not configured for DailyScope export.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
    End If
End Function

Private Function private_TryResolveConfiguredTargetRange( _
    ByVal ws As Worksheet, _
    ByRef outRange As Range _
) As Boolean
    Dim startCell As Range
    Dim endCell As Range
    Dim leftCol As Long
    Dim topRow As Long
    Dim rightCol As Long
    Dim bottomRow As Long
    Dim markerErrorText As String

    Set outRange = Nothing
    If ws Is Nothing Then Exit Function

    If VBA.Len(VBA.Trim$(m_TargetRangeStartMarker)) = 0 Or VBA.Len(VBA.Trim$(m_TargetRangeEndMarker)) = 0 Then
        VBA.MsgBox "PrototypeNew: target range markers are not configured for DailyScope export.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    If Not private_TryResolveMarkerCell(ws, m_TargetRangeStartMarker, startCell, markerErrorText) Then
        VBA.MsgBox "PrototypeNew: failed to resolve RangeStartMarker. " & markerErrorText, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If
    If Not private_TryResolveEndMarkerCell(ws, m_TargetRangeEndMarker, startCell, endCell, markerErrorText) Then
        VBA.MsgBox "PrototypeNew: failed to resolve RangeEndMarker. " & markerErrorText, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    topRow = startCell.Row
    If endCell.Row < topRow Then topRow = endCell.Row
    bottomRow = startCell.Row
    If endCell.Row > bottomRow Then bottomRow = endCell.Row
    leftCol = startCell.Column
    If endCell.Column < leftCol Then leftCol = endCell.Column
    rightCol = startCell.Column
    If endCell.Column > rightCol Then rightCol = endCell.Column

    On Error Resume Next
    Set outRange = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(bottomRow, rightCol))
    On Error GoTo 0
    private_TryResolveConfiguredTargetRange = Not outRange Is Nothing
End Function

Private Function private_TryResolveEndMarkerCell( _
    ByVal ws As Worksheet, _
    ByVal markerText As String, _
    ByVal startCell As Range, _
    ByRef outCell As Range, _
    ByRef outErrorText As String _
) As Boolean
    outErrorText = VBA.vbNullString
    Set outCell = Nothing
    If ws Is Nothing Then Exit Function
    If startCell Is Nothing Then Exit Function

    If private_IsCellReferenceMarker(markerText) Then
        private_TryResolveEndMarkerCell = private_TryResolveMarkerCell(ws, markerText, outCell, outErrorText)
        Exit Function
    End If

    Set outCell = private_FindMarkerTextCellAfterAnchor(ws, markerText, startCell)
    If outCell Is Nothing Then
        outErrorText = "End marker '" & markerText & "' was not found on sheet '" & ws.Name & "'."
        Exit Function
    End If

    private_TryResolveEndMarkerCell = True
End Function

Private Function private_TryResolveMarkerCell( _
    ByVal ws As Worksheet, _
    ByVal markerText As String, _
    ByRef outCell As Range, _
    ByRef outErrorText As String _
) As Boolean
    outErrorText = VBA.vbNullString
    Set outCell = Nothing
    markerText = VBA.Trim$(markerText)
    If VBA.Len(markerText) = 0 Then
        outErrorText = "Marker is empty."
        Exit Function
    End If

    If private_IsCellReferenceMarker(markerText) Then
        If Not private_TryGetCellByMarkerAddress(ws, markerText, outCell) Then
            outErrorText = "Cell marker '" & markerText & "' is invalid for worksheet '" & ws.Name & "'."
            Exit Function
        End If
        private_TryResolveMarkerCell = True
        Exit Function
    End If

    Set outCell = private_FindFirstMarkerTextCell(ws, markerText)
    If outCell Is Nothing Then
        outErrorText = "Text marker '" & markerText & "' was not found on worksheet '" & ws.Name & "'."
        Exit Function
    End If

    private_TryResolveMarkerCell = True
End Function

Private Function private_IsCellReferenceMarker(ByVal markerText As String) As Boolean
    markerText = VBA.Trim$(markerText)
    private_IsCellReferenceMarker = (VBA.Left$(markerText, 1) = "$")
End Function

Private Function private_TryGetCellByMarkerAddress( _
    ByVal ws As Worksheet, _
    ByVal markerText As String, _
    ByRef outCell As Range _
) As Boolean
    On Error GoTo EH_CELL_ADDR
    Set outCell = ws.Range(markerText)
    private_TryGetCellByMarkerAddress = Not outCell Is Nothing
    Exit Function

EH_CELL_ADDR:
    Set outCell = Nothing
End Function

Private Function private_FindFirstMarkerTextCell(ByVal ws As Worksheet, ByVal markerText As String) As Range
    Dim searchRange As Range

    If ws Is Nothing Then Exit Function
    markerText = VBA.Trim$(markerText)
    If VBA.Len(markerText) = 0 Then Exit Function

    Set searchRange = ws.UsedRange
    If searchRange Is Nothing Then Exit Function

    Set private_FindFirstMarkerTextCell = searchRange.Find(What:=markerText, After:=searchRange.Cells(searchRange.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
End Function

Private Function private_FindMarkerTextCellAfterAnchor( _
    ByVal ws As Worksheet, _
    ByVal markerText As String, _
    ByVal anchorCell As Range _
) As Range
    Dim searchRange As Range
    Dim firstFound As Range
    Dim currentFound As Range
    Dim firstAddress As String
    Dim bestWeight As Double
    Dim currentWeight As Double

    If ws Is Nothing Then Exit Function
    If anchorCell Is Nothing Then Exit Function
    markerText = VBA.Trim$(markerText)
    If VBA.Len(markerText) = 0 Then Exit Function

    Set searchRange = ws.UsedRange
    If searchRange Is Nothing Then Exit Function

    Set firstFound = searchRange.Find(What:=markerText, After:=searchRange.Cells(searchRange.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If firstFound Is Nothing Then Exit Function

    bestWeight = -1
    firstAddress = firstFound.Address
    Set currentFound = firstFound

    Do
        If currentFound.Row > anchorCell.Row Or (currentFound.Row = anchorCell.Row And currentFound.Column > anchorCell.Column) Then
            currentWeight = VBA.CDbl(currentFound.Row) * 100000# + VBA.CDbl(currentFound.Column)
            If bestWeight < 0 Or currentWeight < bestWeight Then
                bestWeight = currentWeight
                Set private_FindMarkerTextCellAfterAnchor = currentFound
            End If
        End If
        Set currentFound = searchRange.FindNext(currentFound)
        If currentFound Is Nothing Then Exit Do
    Loop While currentFound.Address <> firstAddress
End Function

Private Function private_FindOpenWorkbookByPath(ByVal sourcePath As String) As Workbook
    Dim wb As Workbook
    Dim normalizedPath As String

    normalizedPath = VBA.LCase$(VBA.Trim$(sourcePath))
    If VBA.Len(normalizedPath) = 0 Then Exit Function

    For Each wb In Application.Workbooks
        If VBA.StrComp(VBA.LCase$(VBA.Trim$(wb.FullName)), normalizedPath, VBA.vbBinaryCompare) = 0 Then
            Set private_FindOpenWorkbookByPath = wb
            Exit Function
        End If
    Next wb
End Function

Private Function private_ExtractWorkbookNameFromPath(ByVal sourcePath As String) As String
    sourcePath = VBA.Trim$(sourcePath)
    If VBA.Len(sourcePath) = 0 Then Exit Function
    private_ExtractWorkbookNameFromPath = VBA.Trim$(VBA.Dir$(sourcePath))
End Function

Private Function private_ExtractSheetNameToken(ByVal configuredSheetName As String) As String
    Dim token As String
    Dim dollarPos As Long

    token = VBA.Trim$(configuredSheetName)
    If VBA.Len(token) = 0 Then Exit Function

    If VBA.Left$(token, 1) = "[" And VBA.Right$(token, 1) = "]" Then
        token = VBA.Mid$(token, 2, VBA.Len(token) - 2)
    End If

    dollarPos = VBA.InStr(1, token, "$", VBA.vbBinaryCompare)
    If dollarPos > 0 Then token = VBA.Left$(token, dollarPos - 1)

    private_ExtractSheetNameToken = VBA.Trim$(token)
End Function

Private Function private_TryResolveExportAliasFromConfigTable(ByVal configTable As obj_ConfigTable) As String
    Dim configEntries As Collection
    Dim entryObj As Variant
    Dim configEntry As obj_ConfigEntry
    Dim keyText As String
    Dim keySuffix As String
    Dim suffixPos As Long

    If configTable Is Nothing Then Exit Function
    If configTable.Items Is Nothing Then Exit Function
    Set configEntries = configTable.Items.AsCollection
    If configEntries Is Nothing Then Exit Function

    For Each entryObj In configEntries
        Set configEntry = Nothing
        On Error Resume Next
        Set configEntry = entryObj
        On Error GoTo 0
        If configEntry Is Nothing Then GoTo ContinueEntry
        keyText = VBA.Trim$(configEntry.Key)
        suffixPos = VBA.InStr(VBA.Len(EXPORT_CONFIG_PREFIX) + 1, keyText, ".", VBA.vbTextCompare)
        If suffixPos <= VBA.Len(EXPORT_CONFIG_PREFIX) + 1 Then GoTo ContinueEntry
        keySuffix = VBA.Mid$(keyText, suffixPos)
        If VBA.StrComp(keySuffix, EXPORT_CLASS_SUFFIX, VBA.vbTextCompare) = 0 Or _
           VBA.StrComp(keySuffix, EXPORT_FILE_PATH_SUFFIX, VBA.vbTextCompare) = 0 Or _
           VBA.StrComp(keySuffix, EXPORT_SHEET_NAME_SUFFIX, VBA.vbTextCompare) = 0 Or _
           VBA.StrComp(keySuffix, EXPORT_RANGE_START_MARKER_SUFFIX, VBA.vbTextCompare) = 0 Or _
           VBA.StrComp(keySuffix, EXPORT_RANGE_END_MARKER_SUFFIX, VBA.vbTextCompare) = 0 Then
            private_TryResolveExportAliasFromConfigTable = VBA.Trim$(VBA.Mid$(keyText, VBA.Len(EXPORT_CONFIG_PREFIX) + 1, suffixPos - VBA.Len(EXPORT_CONFIG_PREFIX) - 1))
            Exit Function
        End If
ContinueEntry:
    Next entryObj
End Function

Private Function private_ValidateSourceTable(ByVal sourceTable As obj_TableDynamic) As Boolean
    If sourceTable Is Nothing Then
        VBA.MsgBox "PrototypeNew: export source table is not specified.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If
    If sourceTable.RowCount <= 0 Then
        VBA.MsgBox "PrototypeNew: export source table has no rows.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If
    If sourceTable.ColumnCount <= 0 Then
        VBA.MsgBox "PrototypeNew: export source table has no columns.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    private_ValidateSourceTable = True
End Function

Private Function private_TryWriteSourceRow( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range _
) As Boolean
    Dim sourceRow As obj_Row
    Dim sourceColumn As obj_Column
    Dim sourceColIndex As Long
    Dim targetColumnIndex As Long
    Dim writtenCount As Long
    Dim targetColumnIndexByName As Object
    Dim sourceColumnName As String
    Dim sourceColumnKey As String

    If sourceTable Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    If Not private_TryBuildTargetColumnIndexByName(targetTable, targetColumnIndexByName) Then Exit Function

    For sourceColIndex = 1 To sourceTable.ColumnCount
        Set sourceColumn = sourceTable.Columns.Item(sourceColIndex)
        If sourceColumn Is Nothing Then GoTo ContinueSourceColumn
        sourceColumnName = VBA.Trim$(VBA.CStr(sourceColumn.Name))
        If VBA.Len(sourceColumnName) = 0 Then GoTo ContinueSourceColumn

        targetColumnIndex = 0
        If targetColumnIndexByName.exists(sourceColumnName) Then
            targetColumnIndex = VBA.CLng(targetColumnIndexByName(sourceColumnName))
        Else
            sourceColumnKey = private_NormalizeColumnNameForMatch(sourceColumnName)
            If VBA.Len(sourceColumnKey) > 0 Then
                If targetColumnIndexByName.exists(sourceColumnKey) Then
                    targetColumnIndex = VBA.CLng(targetColumnIndexByName(sourceColumnKey))
                End If
            End If
        End If
        If targetColumnIndex <= 0 Then GoTo ContinueSourceColumn

        ' Paste-values semantics: write only scalar values to mapped cells,
        ' do not copy styles and do not overwrite unmapped columns.
        rowRange.Cells(1, targetColumnIndex).Value2 = sourceRow.GetCellValue(sourceColIndex)
        writtenCount = writtenCount + 1

ContinueSourceColumn:
    Next sourceColIndex

    If writtenCount <= 0 Then
        VBA.MsgBox "PrototypeNew: DailyScope export found no matching target columns for source table: " & sourceTable.HeaderText, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    private_TryWriteSourceRow = True
End Function

Private Function private_TryBuildTargetColumnIndexByName( _
    ByVal targetTable As ListObject, _
    ByRef outColumnIndexByName As Object _
) As Boolean
    Dim tableColumn As ListColumn
    Dim columnName As String
    Dim normalizedKey As String

    Set outColumnIndexByName = Nothing
    If targetTable Is Nothing Then Exit Function

    Set outColumnIndexByName = ex_Helpers.fn_CreateDictionaryTextCompare()
    For Each tableColumn In targetTable.ListColumns
        columnName = VBA.Trim$(VBA.CStr(tableColumn.Name))
        If VBA.Len(columnName) > 0 Then
            outColumnIndexByName(columnName) = tableColumn.Index

            normalizedKey = private_NormalizeColumnNameForMatch(columnName)
            If VBA.Len(normalizedKey) > 0 Then
                If Not outColumnIndexByName.Exists(normalizedKey) Then
                    outColumnIndexByName(normalizedKey) = tableColumn.Index
                End If
            End If
        End If
    Next tableColumn

    private_TryBuildTargetColumnIndexByName = Not outColumnIndexByName Is Nothing
End Function

Private Function private_NormalizeColumnNameForMatch(ByVal columnName As String) As String
    columnName = private_NormalizeText(columnName)
    columnName = VBA.Replace(columnName, "№", "")
    columnName = VBA.Replace(columnName, "n", "")
    Do While VBA.InStr(1, columnName, "  ", VBA.vbBinaryCompare) > 0
        columnName = VBA.Replace(columnName, "  ", " ")
    Loop
    private_NormalizeColumnNameForMatch = VBA.Trim$(columnName)
End Function

Private Function private_TryResolveTargetSectionCaption( _
    ByVal sourceTable As obj_TableDynamic, _
    ByRef outSectionCaption As String _
) As Boolean
    Dim sourceRow As obj_Row
    Dim sectionKey As String

    outSectionCaption = VBA.vbNullString
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, sectionKey, "meta_SectionType") Then
        VBA.MsgBox "PrototypeNew: DailyScope export requires meta_SectionType.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    If private_TryMapSectionKeyToCaption(sectionKey, outSectionCaption) Then
        private_TryResolveTargetSectionCaption = True
        Exit Function
    End If

    VBA.MsgBox "PrototypeNew: unknown DailyScope section key: " & sectionKey, VBA.vbExclamation, "PrototypeNew / DailyScope export"
End Function

Private Function private_TryGetSectionWriteRowRange( _
    ByVal targetTable As ListObject, _
    ByVal sectionCaption As String, _
    ByRef outRowRange As Range, _
    ByRef outInsertedRow As ListRow _
) As Boolean
    Dim sectionRowIndex As Long
    Dim nextSectionRowIndex As Long
    Dim sectionFirstDataIndex As Long
    Dim sectionLastDataIndex As Long
    Dim writeRowIndex As Long
    Dim insertPosition As Long

    Set outRowRange = Nothing
    Set outInsertedRow = Nothing
    If targetTable Is Nothing Then Exit Function
    If targetTable.DataBodyRange Is Nothing Then Exit Function

    If Not private_TryFindSectionRowIndex(targetTable, sectionCaption, sectionRowIndex) Then
        VBA.MsgBox "PrototypeNew: DailyScope section was not found: " & sectionCaption, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    nextSectionRowIndex = private_FindNextSectionRowIndex(targetTable, sectionRowIndex + 1)
    sectionFirstDataIndex = sectionRowIndex + 1
    If nextSectionRowIndex > 0 Then
        sectionLastDataIndex = nextSectionRowIndex - 1
    Else
        sectionLastDataIndex = targetTable.ListRows.Count
    End If

    writeRowIndex = private_FindNextEmptySectionRowIndex(targetTable, sectionFirstDataIndex, sectionLastDataIndex)
    If writeRowIndex > 0 Then
        Set outRowRange = targetTable.ListRows.Item(writeRowIndex).Range
        private_TryGetSectionWriteRowRange = Not outRowRange Is Nothing
        Exit Function
    End If

    If nextSectionRowIndex > 0 Then
        insertPosition = nextSectionRowIndex
    Else
        insertPosition = targetTable.ListRows.Count + 1
    End If

    Set outInsertedRow = targetTable.ListRows.Add(Position:=insertPosition)
    If outInsertedRow Is Nothing Then
        VBA.MsgBox "PrototypeNew: failed to insert row into DailyScope section: " & sectionCaption, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    If Not private_TryApplyInsertedRowSectionFormat(targetTable, outInsertedRow, sectionFirstDataIndex, sectionLastDataIndex) Then Exit Function

    Set outRowRange = outInsertedRow.Range
    private_TryGetSectionWriteRowRange = Not outRowRange Is Nothing
End Function

Private Function private_TryApplyInsertedRowSectionFormat( _
    ByVal targetTable As ListObject, _
    ByVal insertedRow As ListRow, _
    ByVal sectionFirstDataIndex As Long, _
    ByVal sectionLastDataIndex As Long _
) As Boolean
    Dim templateRowIndex As Long
    Dim templateRange As Range

    private_TryApplyInsertedRowSectionFormat = True
    If targetTable Is Nothing Then Exit Function
    If insertedRow Is Nothing Then Exit Function
    If insertedRow.Range Is Nothing Then Exit Function

    templateRowIndex = private_FindSectionTemplateRowIndex(targetTable, sectionFirstDataIndex, sectionLastDataIndex)
    If templateRowIndex <= 0 Then Exit Function

    Set templateRange = targetTable.ListRows.Item(templateRowIndex).Range
    If templateRange Is Nothing Then Exit Function

    On Error GoTo EH_APPLY_FORMAT
    ' Keep export as values-only while still preserving visual section template.
    templateRange.Copy
    insertedRow.Range.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    private_TryApplyInsertedRowSectionFormat = True
    Exit Function

EH_APPLY_FORMAT:
    Application.CutCopyMode = False
    VBA.MsgBox "PrototypeNew: failed to apply section row format for inserted DailyScope row.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
    private_TryApplyInsertedRowSectionFormat = False
End Function

Private Function private_FindSectionTemplateRowIndex( _
    ByVal targetTable As ListObject, _
    ByVal sectionFirstDataIndex As Long, _
    ByVal sectionLastDataIndex As Long _
) As Long
    Dim rowIndex As Long

    If targetTable Is Nothing Then Exit Function
    If sectionFirstDataIndex <= 0 Then Exit Function
    If sectionLastDataIndex < sectionFirstDataIndex Then Exit Function
    If targetTable.ListRows.Count <= 0 Then Exit Function

    For rowIndex = sectionFirstDataIndex To sectionLastDataIndex
        If rowIndex >= 1 And rowIndex <= targetTable.ListRows.Count Then
            private_FindSectionTemplateRowIndex = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function private_TryFindSectionRowIndex( _
    ByVal targetTable As ListObject, _
    ByVal sectionCaption As String, _
    ByRef outRowIndex As Long _
) As Boolean
    Dim rowIndex As Long
    Dim rowText As String
    Dim expectedText As String

    outRowIndex = 0
    expectedText = private_NormalizeText(sectionCaption)
    If VBA.Len(expectedText) = 0 Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If targetTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To targetTable.ListRows.Count
        rowText = private_GetSectionTextFromRow(targetTable.ListRows.Item(rowIndex).Range)
        If VBA.StrComp(private_NormalizeText(rowText), expectedText, VBA.vbTextCompare) = 0 Then
            outRowIndex = rowIndex
            private_TryFindSectionRowIndex = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function private_FindNextSectionRowIndex(ByVal targetTable As ListObject, ByVal firstRowIndex As Long) As Long
    Dim rowIndex As Long
    Dim rowText As String

    If targetTable Is Nothing Then Exit Function
    If targetTable.DataBodyRange Is Nothing Then Exit Function
    If firstRowIndex < 1 Then firstRowIndex = 1

    For rowIndex = firstRowIndex To targetTable.ListRows.Count
        rowText = private_GetSectionTextFromRow(targetTable.ListRows.Item(rowIndex).Range)
        If private_IsKnownSectionCaption(rowText) Then
            private_FindNextSectionRowIndex = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function private_FindNextEmptySectionRowIndex( _
    ByVal targetTable As ListObject, _
    ByVal firstRowIndex As Long, _
    ByVal lastRowIndex As Long _
) As Long
    Dim rowIndex As Long
    Dim lastUsedRowIndex As Long
    Dim candidateRowIndex As Long

    If targetTable Is Nothing Then Exit Function
    If firstRowIndex <= 0 Then Exit Function
    If lastRowIndex < firstRowIndex Then Exit Function

    lastUsedRowIndex = firstRowIndex - 1
    For rowIndex = firstRowIndex To lastRowIndex
        If Not private_IsRowEmpty(targetTable.ListRows.Item(rowIndex).Range) Then lastUsedRowIndex = rowIndex
    Next rowIndex

    candidateRowIndex = lastUsedRowIndex + 1
    If candidateRowIndex <= lastRowIndex Then
        If private_IsRowEmpty(targetTable.ListRows.Item(candidateRowIndex).Range) Then private_FindNextEmptySectionRowIndex = candidateRowIndex
    End If
End Function

Private Function private_IsRowEmpty(ByVal rowRange As Range) As Boolean
    Dim cellObj As Range
    Dim valueText As String

    If rowRange Is Nothing Then Exit Function

    For Each cellObj In rowRange.Cells
        valueText = VBA.Trim$(VBA.CStr(cellObj.Value2))
        If VBA.Len(valueText) > 0 Then Exit Function
    Next cellObj

    private_IsRowEmpty = True
End Function

Private Function private_GetSectionTextFromRow(ByVal rowRange As Range) As String
    If rowRange Is Nothing Then Exit Function

    On Error Resume Next
    private_GetSectionTextFromRow = VBA.CStr(rowRange.Cells(1, 1).Value2)
    On Error GoTo 0
End Function

Private Function private_TryGetSourceTextByAnyColumn( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByRef outText As String, _
    ParamArray columnNames() As Variant _
) As Boolean
    Dim columnName As Variant
    Dim columnIndex As Long

    outText = VBA.vbNullString
    If sourceTable Is Nothing Then Exit Function
    If sourceRow Is Nothing Then Exit Function

    For Each columnName In columnNames
        columnIndex = private_GetSourceColumnIndex(sourceTable, VBA.CStr(columnName))
        If columnIndex > 0 Then
            outText = VBA.Trim$(VBA.CStr(sourceRow.GetCellValue(columnIndex)))
            private_TryGetSourceTextByAnyColumn = VBA.Len(outText) > 0
            Exit Function
        End If
    Next columnName
End Function

Private Function private_GetSourceColumnIndex(ByVal sourceTable As obj_TableDynamic, ByVal columnName As String) As Long
    Dim sourceColIndex As Long
    Dim sourceColumn As obj_Column
    Dim expectedName As String

    If sourceTable Is Nothing Then Exit Function
    expectedName = private_NormalizeText(columnName)
    If VBA.Len(expectedName) = 0 Then Exit Function

    For sourceColIndex = 1 To sourceTable.ColumnCount
        Set sourceColumn = sourceTable.Columns.Item(sourceColIndex)
        If sourceColumn Is Nothing Then GoTo ContinueColumn
        If VBA.StrComp(private_NormalizeText(sourceColumn.Name), expectedName, VBA.vbTextCompare) = 0 Then
            private_GetSourceColumnIndex = sourceColIndex
            Exit Function
        End If

ContinueColumn:
    Next sourceColIndex
End Function

Private Function private_TryMapSectionKeyToCaption(ByVal sectionKey As String, ByRef outCaption As String) As Boolean
    Dim normalizedKey As String

    outCaption = VBA.vbNullString
    normalizedKey = private_NormalizeText(sectionKey)
    If VBA.Len(normalizedKey) = 0 Then Exit Function

    If Not private_IsSupportedSectionType(normalizedKey) Then
        Exit Function
    End If

    outCaption = private_GetKnownSectionCaptionByText(normalizedKey)

    private_TryMapSectionKeyToCaption = VBA.Len(outCaption) > 0
End Function

Private Function private_BuildSectionTypeOptions() As Collection
    Dim sectionTypes As Collection

    Set sectionTypes = New Collection
    sectionTypes.Add "з лікування"
    sectionTypes.Add "з відпустки для лікування"
    sectionTypes.Add "з щорічної основної відпустки"
    sectionTypes.Add "з відпустки за сімейними обставинами"
    sectionTypes.Add "з лікування медична рота"
    sectionTypes.Add "з амбулаторного обстеження влк"
    sectionTypes.Add "на лікування"
    sectionTypes.Add "у частину щорічної основної відпустки"
    sectionTypes.Add "у відпустку за сімейними обставинами"
    sectionTypes.Add "у відпустку для лікування"
    sectionTypes.Add "на лікування медична рота"
    sectionTypes.Add "на амбулаторне обстеження влк"
    sectionTypes.Add "зміна місця перебування лікування => відпустка для лік"
    sectionTypes.Add "зміна місця перебування відпустка для лік => відпустка для лік"
    sectionTypes.Add "зміна місця перебування відпустка для лік => лікування"
    sectionTypes.Add "зміна місця перебування відпустка для лік => влк"
    sectionTypes.Add "зміна місця перебування влк => відпустка для лік"
    sectionTypes.Add "зміна місця перебування влк => лікування"
    sectionTypes.Add "у відрядження"
    sectionTypes.Add "у відрядження сзч"

    Set private_BuildSectionTypeOptions = sectionTypes
End Function

Private Function private_IsSupportedSectionType(ByVal valueText As String) As Boolean
    Dim normalizedText As String
    Dim sectionTypes As Collection
    Dim sectionTypeValue As Variant

    normalizedText = private_NormalizeText(valueText)
    If VBA.Len(normalizedText) = 0 Then Exit Function

    Set sectionTypes = private_BuildSectionTypeOptions()
    If sectionTypes Is Nothing Then Exit Function

    For Each sectionTypeValue In sectionTypes
        If VBA.StrComp(private_NormalizeText(VBA.CStr(sectionTypeValue)), normalizedText, VBA.vbTextCompare) = 0 Then
            private_IsSupportedSectionType = True
            Exit Function
        End If
    Next sectionTypeValue
End Function

Private Function private_IsKnownSectionCaption(ByVal valueText As String) As Boolean
    Dim normalizedCandidate As String
    Dim knownCaptions As Collection
    Dim captionObj As Variant
    Dim normalizedKnownCaption As String

    normalizedCandidate = private_NormalizeText(valueText)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "exporter:section-caption-check start raw='" & VBA.Replace(VBA.CStr(valueText), "'", "''") & "' normalized='" & VBA.Replace(normalizedCandidate, "'", "''") & "'"
#End If
    If VBA.Len(normalizedCandidate) = 0 Then Exit Function

    Set knownCaptions = private_BuildKnownSectionCaptions()
    If knownCaptions Is Nothing Then Exit Function

    For Each captionObj In knownCaptions
        normalizedKnownCaption = private_NormalizeText(VBA.CStr(captionObj))
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "exporter:section-caption-check compare candidate='" & VBA.Replace(normalizedCandidate, "'", "''") & "' known='" & VBA.Replace(normalizedKnownCaption, "'", "''") & "'"
#End If
        If VBA.StrComp(normalizedCandidate, normalizedKnownCaption, VBA.vbTextCompare) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo "exporter:section-caption-check match known='" & VBA.Replace(normalizedKnownCaption, "'", "''") & "'"
#End If
            private_IsKnownSectionCaption = True
            Exit Function
        End If
    Next captionObj

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "exporter:section-caption-check no-match candidate='" & VBA.Replace(normalizedCandidate, "'", "''") & "'"
#End If
End Function

Private Function private_BuildKnownSectionCaptions() As Collection
    Dim sectionTypes As Collection
    Dim sectionTypeValue As Variant
    Dim captionText As String
    Dim captionKey As String
    Dim knownCaptionMap As Object
    Dim captions As Collection

    Set captions = New Collection
    Set knownCaptionMap = ex_Helpers.fn_CreateDictionaryTextCompare()
    If knownCaptionMap Is Nothing Then Exit Function

    Set sectionTypes = private_BuildSectionTypeOptions()
    If sectionTypes Is Nothing Then Exit Function

    For Each sectionTypeValue In sectionTypes
        captionText = private_GetKnownSectionCaptionByText(VBA.CStr(sectionTypeValue))
        captionKey = private_NormalizeText(captionText)
        If VBA.Len(captionKey) = 0 Then GoTo ContinueSectionType
        If knownCaptionMap.Exists(captionKey) Then GoTo ContinueSectionType

        knownCaptionMap(captionKey) = True
        captions.Add captionText

ContinueSectionType:
    Next sectionTypeValue

    Set private_BuildKnownSectionCaptions = captions
End Function

Private Function private_GetKnownSectionCaptionByText(ByVal valueText As String) As String
    Dim normalizedText As String

    normalizedText = private_NormalizeText(valueText)
    Select Case normalizedText
        Case "з лікування"
            private_GetKnownSectionCaptionByText = "З лікування:"
        Case "з відпустки для лікування"
            private_GetKnownSectionCaptionByText = "З відпустки для лікування:"
        Case "з щорічної основної відпустки"
            private_GetKnownSectionCaptionByText = "З щорічної основної відпустки:"
        Case "з відпустки за сімейними обставинами"
            private_GetKnownSectionCaptionByText = "З відпустки за сімейними обставинами:"
        Case "з лікування медична рота"
            private_GetKnownSectionCaptionByText = "З лікування (медична рота):"
        Case "з амбулаторного обстеження влк"
            private_GetKnownSectionCaptionByText = "З амбулаторного обстеження / ВЛК:"
        Case "на лікування"
            private_GetKnownSectionCaptionByText = "На лікування:" & VBA.vbLf & "(давальний відмінок)"
        Case "у частину щорічної основної відпустки"
            private_GetKnownSectionCaptionByText = "У частину щорічної основної відпустки:"
        Case "у відпустку за сімейними обставинами"
            private_GetKnownSectionCaptionByText = "У відпустку за сімейними обставинами:"
        Case "у відпустку для лікування"
            private_GetKnownSectionCaptionByText = "У відпустку для лікування:"
        Case "на лікування медична рота"
            private_GetKnownSectionCaptionByText = "На лікування (медична рота):"
        Case "на амбулаторне обстеження влк"
            private_GetKnownSectionCaptionByText = "На амбулаторне обстеження / ВЛК:"
        Case "зміна місця перебування лікування => відпустка для лік"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(Лікування / Відпустка лік.):"
        Case "зміна місця перебування відпустка для лік => відпустка для лік"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(Відпустка лік / Відпустка лік.):"
        Case "зміна місця перебування відпустка для лік => лікування"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(Відпустка лік. / Лікування):"
        Case "зміна місця перебування відпустка для лік => влк"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(Відпустка лік. / ВЛК):"
        Case "зміна місця перебування влк => відпустка для лік"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(ВЛК / Відпустка лік.):"
        Case "зміна місця перебування влк => лікування"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(ВЛК / Лікування):"
        Case "у відрядження"
            private_GetKnownSectionCaptionByText = "У відрядження"
        Case "у відрядження сзч"
            private_GetKnownSectionCaptionByText = "У відрядження (СЗЧ):"
    End Select
End Function

Private Function private_NormalizeText(ByVal valueText As String) As String
    valueText = VBA.LCase$(VBA.Trim$(VBA.CStr(valueText)))
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    valueText = VBA.Replace(valueText, ":", VBA.vbNullString)
    valueText = VBA.Replace(valueText, ".", VBA.vbNullString)
    valueText = VBA.Replace(valueText, "/", " ")
    valueText = VBA.Replace(valueText, "(", " ")
    valueText = VBA.Replace(valueText, ")", " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    private_NormalizeText = VBA.Trim$(valueText)
End Function

Private Sub private_BeginFastExcelMode( _
    ByRef outScreenUpdating As Boolean, _
    ByRef outEnableEvents As Boolean, _
    ByRef outDisplayAlerts As Boolean, _
    ByRef outCalculation As XlCalculation _
)
    outScreenUpdating = Application.ScreenUpdating
    outEnableEvents = Application.EnableEvents
    outDisplayAlerts = Application.DisplayAlerts
    outCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub private_RestoreFastExcelMode( _
    ByVal screenUpdating As Boolean, _
    ByVal enableEvents As Boolean, _
    ByVal displayAlerts As Boolean, _
    ByVal calculation As XlCalculation _
)
    On Error Resume Next
    Application.Calculation = calculation
    Application.DisplayAlerts = displayAlerts
    Application.EnableEvents = enableEvents
    Application.ScreenUpdating = screenUpdating
    On Error GoTo 0
End Sub
