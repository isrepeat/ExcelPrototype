VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_DataExporterBase"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False

Private Const EXPORT_CONFIG_PREFIX As String = "Export."
Private Const EXPORT_CLASS_SUFFIX As String = ".ExporterClass"
Private Const EXPORT_FILE_PATH_SUFFIX As String = ".FilePath"
Private Const EXPORT_SHEET_NAME_SUFFIX As String = ".SheetName"
Private Const EXPORT_RANGE_START_MARKER_SUFFIX As String = ".RangeStartMarker"
Private Const EXPORT_RANGE_END_MARKER_SUFFIX As String = ".RangeEndMarker"

Private m_ExporterName As String
Private m_DialogTitle As String
Private m_TargetWorkbookPath As String
Private m_TargetSheetName As String
Private m_TargetRangeStartMarker As String
Private m_TargetRangeEndMarker As String

Public Function Initialize( _
    ByVal configTable As obj_ConfigTable, _
    ByVal exporterName As String, _
    ByVal dialogTitle As String _
) As Boolean
    Dim cfgParserBase As obj_CfgParserBase
    Dim configEntries As Collection
    Dim cfgMap As Object
    Dim exportAlias As String

    m_ExporterName = VBA.Trim$(exporterName)
    m_DialogTitle = VBA.Trim$(dialogTitle)
    m_TargetWorkbookPath = VBA.vbNullString
    m_TargetSheetName = VBA.vbNullString
    m_TargetRangeStartMarker = VBA.vbNullString
    m_TargetRangeEndMarker = VBA.vbNullString

    If VBA.Len(m_ExporterName) = 0 Then m_ExporterName = "Data"
    If VBA.Len(m_DialogTitle) = 0 Then m_DialogTitle = "PrototypeNew / Export"

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

Public Sub Dispose()
    m_ExporterName = VBA.vbNullString
    m_DialogTitle = VBA.vbNullString
    m_TargetWorkbookPath = VBA.vbNullString
    m_TargetSheetName = VBA.vbNullString
    m_TargetRangeStartMarker = VBA.vbNullString
    m_TargetRangeEndMarker = VBA.vbNullString
End Sub

Public Property Get TargetWorkbookPath() As String
    TargetWorkbookPath = m_TargetWorkbookPath
End Property

Public Property Get TargetSheetName() As String
    TargetSheetName = m_TargetSheetName
End Property

Public Property Get TargetRangeStartMarker() As String
    TargetRangeStartMarker = m_TargetRangeStartMarker
End Property

Public Property Get TargetRangeEndMarker() As String
    TargetRangeEndMarker = m_TargetRangeEndMarker
End Property

Public Function TryOpenTargetWorkbook( _
    ByRef outWorkbook As Workbook, _
    ByRef outOpenedByExporter As Boolean _
) As Boolean
    Dim wb As Workbook
    Dim resolvedPath As String
    Dim targetWorkbookName As String

    private_LogMethodEntry "TryOpenTargetWorkbook"

    Set outWorkbook = Nothing
    outOpenedByExporter = False
    resolvedPath = VBA.Trim$(m_TargetWorkbookPath)

    If VBA.Len(resolvedPath) > 0 Then
        Set outWorkbook = private_FindOpenWorkbookByPath(resolvedPath)
        If Not outWorkbook Is Nothing Then
            TryOpenTargetWorkbook = True
            Exit Function
        End If
    End If

    targetWorkbookName = private_ExtractWorkbookNameFromPath(resolvedPath)
    If VBA.Len(targetWorkbookName) > 0 Then
        For Each wb In Application.Workbooks
            If VBA.StrComp(wb.Name, targetWorkbookName, VBA.vbTextCompare) = 0 Then
                Set outWorkbook = wb
                TryOpenTargetWorkbook = True
                Exit Function
            End If
        Next wb
    End If

    If VBA.Len(resolvedPath) = 0 Then
        VBA.MsgBox "PrototypeNew: target workbook path is not configured for " & m_ExporterName & " export.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If

    If VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: " & m_ExporterName & " target workbook was not found: " & resolvedPath, VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If

    Set outWorkbook = Application.Workbooks.Open(resolvedPath)
    outOpenedByExporter = True
    TryOpenTargetWorkbook = Not outWorkbook Is Nothing
End Function

Public Function ResolveTargetWorksheetName() As String
    private_LogMethodEntry "ResolveTargetWorksheetName"
    ResolveTargetWorksheetName = private_ExtractSheetNameToken(m_TargetSheetName)
    If VBA.Len(ResolveTargetWorksheetName) = 0 Then
        VBA.MsgBox "PrototypeNew: target worksheet is not configured for " & m_ExporterName & " export.", VBA.vbExclamation, m_DialogTitle
    End If
End Function

Public Function TryGetWorksheet( _
    ByVal wb As Workbook, _
    ByVal worksheetName As String, _
    ByRef outWorksheet As Worksheet _
) As Boolean
    private_LogMethodEntry "TryGetWorksheet"
    Set outWorksheet = Nothing
    If wb Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(worksheetName)) = 0 Then Exit Function

    On Error Resume Next
    Set outWorksheet = wb.Worksheets(worksheetName)
    On Error GoTo 0

    If outWorksheet Is Nothing Then
        VBA.MsgBox "PrototypeNew: " & m_ExporterName & " worksheet was not found: " & worksheetName, VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If

    TryGetWorksheet = True
End Function

Public Function TryFindConfiguredTargetTable( _
    ByVal ws As Worksheet, _
    ByRef outTable As ListObject _
) As Boolean
    Dim tableObj As ListObject
    Dim targetRange As Range
    Dim startCell As Range

    private_LogMethodEntry "TryFindConfiguredTargetTable"

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
            TryFindConfiguredTargetTable = True
            Exit Function
        End If
        If Not Application.Intersect(tableObj.Range, targetRange) Is Nothing Then
            Set outTable = tableObj
            TryFindConfiguredTargetTable = True
            Exit Function
        End If

ContinueTable:
    Next tableObj

    VBA.MsgBox "PrototypeNew: target table was not found for configured markers on sheet '" & ws.Name & "'.", VBA.vbExclamation, m_DialogTitle
End Function

Public Function ValidateSourceTable(ByVal sourceTable As obj_TableDynamic) As Boolean
    private_LogMethodEntry "ValidateSourceTable"

    If sourceTable Is Nothing Then
        VBA.MsgBox "PrototypeNew: export source table is not specified.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If
    If sourceTable.RowCount <= 0 Then
        VBA.MsgBox "PrototypeNew: export source table has no rows.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If
    If sourceTable.ColumnCount <= 0 Then
        VBA.MsgBox "PrototypeNew: export source table has no columns.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If

    ValidateSourceTable = True
End Function

Public Function TryGetMainSourceTable( _
    ByVal sourceTables As Collection, _
    ByRef outSourceTable As obj_TableDynamic _
) As Boolean
    Set outSourceTable = Nothing
    If sourceTables Is Nothing Then
        VBA.MsgBox "PrototypeNew: export source tables are not specified.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If
    If sourceTables.Count <= 0 Then
        VBA.MsgBox "PrototypeNew: export source tables list is empty.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If
    If Not IsObject(sourceTables.Item(1)) Then
        VBA.MsgBox "PrototypeNew: main export source table is not an object.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.TypeName(sourceTables.Item(1))), "obj_tabledynamic", VBA.vbTextCompare) <> 0 Then
        VBA.MsgBox "PrototypeNew: main export source table has unsupported type: " & VBA.TypeName(sourceTables.Item(1)), VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If

    Set outSourceTable = sourceTables.Item(1)
    If Not Me.ValidateSourceTable(outSourceTable) Then Exit Function
    TryGetMainSourceTable = True
End Function

Public Function TryRememberExportRow(ByVal context As Object, ByVal exporterKey As String, ByVal targetTable As ListObject, ByVal rowRange As Range) As Boolean
    Dim history As Object
    Dim entry As Object

    If context Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function
    If Not private_TryGetContextObject(context, "ExportHistory", history) Then Exit Function

    ' История живет в context контроллера и переживает создание нового экземпляра
    ' экспортера. Храним координаты, а не данные человека: строка может быть черновой.
    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("Workbook") = targetTable.Parent.Parent.FullName
    entry("Worksheet") = targetTable.Parent.Name
    entry("Table") = targetTable.Name
    entry("RowIndex") = rowRange.Row - targetTable.DataBodyRange.Row + 1
    Set history(VBA.Trim$(exporterKey)) = entry
    TryRememberExportRow = True
End Function

Public Function TryGetRememberedExportRow(ByVal context As Object, ByVal exporterKey As String, ByVal targetTable As ListObject, ByRef outRowRange As Range) As Boolean
    Dim history As Object
    Dim entry As Object
    Dim rowIndex As Long

    Set outRowRange = Nothing
    If context Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If Not private_TryGetContextObject(context, "ExportHistory", history) Then Exit Function
    exporterKey = VBA.Trim$(exporterKey)
    If Not history.Exists(exporterKey) Then
        VBA.MsgBox "PrototypeNew: Rewrite Last is unavailable because this exporter has no insertion history in the current session.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If
    Set entry = history(exporterKey)
    ' Не переносим Rewrite Last между разными физическими export targets.
    If VBA.StrComp(VBA.CStr(entry("Workbook")), targetTable.Parent.Parent.FullName, VBA.vbTextCompare) <> 0 Or _
       VBA.StrComp(VBA.CStr(entry("Worksheet")), targetTable.Parent.Name, VBA.vbTextCompare) <> 0 Or _
       VBA.StrComp(VBA.CStr(entry("Table")), targetTable.Name, VBA.vbTextCompare) <> 0 Then
        VBA.MsgBox "PrototypeNew: the remembered export location belongs to another workbook, worksheet, or table.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If
    rowIndex = VBA.CLng(entry("RowIndex"))
    If rowIndex < 1 Or rowIndex > targetTable.ListRows.Count Then
        VBA.MsgBox "PrototypeNew: the remembered export row no longer exists. Use Default mode to create a new history entry.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If
    Set outRowRange = targetTable.ListRows.Item(rowIndex).Range
    TryGetRememberedExportRow = Not outRowRange Is Nothing
End Function

Private Function private_TryGetContextObject(ByVal context As Object, ByVal keyText As String, ByRef outValue As Object) As Boolean
    Set outValue = Nothing
    On Error Resume Next
    If context.Exists(keyText) Then Set outValue = context(keyText)
    On Error GoTo 0
    private_TryGetContextObject = Not outValue Is Nothing
End Function

Public Sub BeginFastExcelMode( _
    ByRef outScreenUpdating As Boolean, _
    ByRef outEnableEvents As Boolean, _
    ByRef outDisplayAlerts As Boolean, _
    ByRef outCalculation As XlCalculation _
)
    private_LogMethodEntry "BeginFastExcelMode"
    outScreenUpdating = Application.ScreenUpdating
    outEnableEvents = Application.EnableEvents
    outDisplayAlerts = Application.DisplayAlerts
    outCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub RestoreFastExcelMode( _
    ByVal screenUpdating As Boolean, _
    ByVal enableEvents As Boolean, _
    ByVal displayAlerts As Boolean, _
    ByVal calculation As XlCalculation _
)
    private_LogMethodEntry "RestoreFastExcelMode"
    On Error Resume Next
    Application.Calculation = calculation
    Application.DisplayAlerts = displayAlerts
    Application.EnableEvents = enableEvents
    Application.ScreenUpdating = screenUpdating
    On Error GoTo 0
End Sub

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
        VBA.MsgBox "PrototypeNew: target range markers are not configured for " & m_ExporterName & " export.", VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If

    If Not private_TryResolveMarkerCell(ws, m_TargetRangeStartMarker, startCell, markerErrorText) Then
        VBA.MsgBox "PrototypeNew: failed to resolve RangeStartMarker. " & markerErrorText, VBA.vbExclamation, m_DialogTitle
        Exit Function
    End If
    If Not private_TryResolveEndMarkerCell(ws, m_TargetRangeEndMarker, startCell, endCell, markerErrorText) Then
        VBA.MsgBox "PrototypeNew: failed to resolve RangeEndMarker. " & markerErrorText, VBA.vbExclamation, m_DialogTitle
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

Private Sub private_LogMethodEntry(ByVal methodName As String)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "data-exporter-base:enter " & VBA.Trim$(methodName)
#End If
End Sub
