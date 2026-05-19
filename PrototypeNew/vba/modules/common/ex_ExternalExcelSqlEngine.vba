Attribute VB_Name = "ex_ExternalExcelSqlEngine"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const RUNTIME_ERROR_TITLE As String = "PrototypeNew / SQL runtime"
Private Const ADO_UNSUPPORTED_EXT_ERROR_CODE As Long = VBA.vbObjectError + 7312
Private Const RANGE_REF_CACHE_NAMESPACE As String = "SqlEngine.RangeRefsByMarkers"
Private Const RANGE_REF_CACHE_VERSION As String = "v1"

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_ExternalExcelSqlEngine.fn_Module_Dispose"
#End If
    On Error Resume Next
    Call ex_CacheRuntime.fn_ClearNamespace(RANGE_REF_CACHE_NAMESPACE)
    On Error GoTo 0
End Sub

Public Function fn_TrySqlRequest( _
    ByVal sqlParams As obj_SqlParams, _
    ByRef outTable As obj_TableDynamic, _
    Optional ByRef outRowsCount As Long = 0 _
) As Boolean
    Dim conn As Object
    Dim rsSchema As Object
    Dim rsData As Object
    Dim tableRef As String
    Dim sql As String
    Dim validationError As String
    Dim sourcePath As String
    Dim sourceColumnHeaders As Collection
    Dim mappedColumnHeaders As Collection
    Dim resolvedSourceColumnHeaders As Collection
    Dim sourceColumnOrdinals() As Long
    Dim tableObj As obj_TableDynamic
    Dim rowObj As obj_Row
    Dim colObj As obj_Column
    Dim markerErrorText As String
    Dim i As Long
    Dim rowIndex As Long
    Dim resolvedSourceColumnHeader As String
    Dim availableFields As String
    Dim hasGenericFields As Boolean

    On Error GoTo EH_QUERY

    Set outTable = Nothing
    outRowsCount = 0

    If sqlParams Is Nothing Then
        MsgBox "PrototypeNew: SQL params object is not specified.", vbExclamation, RUNTIME_ERROR_TITLE
        Exit Function
    End If

    If Not sqlParams.TryValidate(validationError) Then
        MsgBox "PrototypeNew: invalid SQL params. " & validationError, vbExclamation, RUNTIME_ERROR_TITLE
        Exit Function
    End If

    sourcePath = private_ResolvePathLocal(sqlParams.SourcePath)
    If VBA.Len(sourcePath) = 0 Then
        MsgBox "PrototypeNew: resolved source path is empty.", vbExclamation, RUNTIME_ERROR_TITLE
        Exit Function
    End If
    If VBA.Dir$(sourcePath) = VBA.vbNullString Then
        MsgBox "PrototypeNew: source file not found: " & sourcePath, vbExclamation, RUNTIME_ERROR_TITLE
        Exit Function
    End If

    If private_HasRangeMarkers(sqlParams) Then
        If Not private_TryBuildTableRefFromMarkers(sourcePath, sqlParams.SheetName, sqlParams.RangeStartMarker, sqlParams.RangeEndMarker, tableRef, markerErrorText) Then
            MsgBox "PrototypeNew: failed to resolve range by markers. " & markerErrorText, vbExclamation, RUNTIME_ERROR_TITLE
            Exit Function
        End If
    Else
        tableRef = private_BuildTableRefFromSheetName(sqlParams.SheetName)
        If VBA.Len(tableRef) = 0 Then
            MsgBox "PrototypeNew: failed to build SQL table reference from SheetName '" & sqlParams.SheetName & "'.", vbExclamation, RUNTIME_ERROR_TITLE
            Exit Function
        End If
    End If

    Set sourceColumnHeaders = sqlParams.SourceColumnHeaders
    Set mappedColumnHeaders = sqlParams.MappedColumnHeaders

    Set conn = CreateObject("ADODB.Connection")
    conn.Open private_BuildAdoConnectionString(sourcePath)

    Set rsSchema = CreateObject("ADODB.Recordset")
    rsSchema.Open "SELECT * FROM " & tableRef & " WHERE 1=0", conn, 0, 1
    availableFields = private_ListRecordsetFields(rsSchema, 40)
    hasGenericFields = private_RecordsetLooksLikeGenericFields(rsSchema)

    Set resolvedSourceColumnHeaders = New Collection
    For i = 1 To sourceColumnHeaders.Count
        resolvedSourceColumnHeader = VBA.vbNullString
        If Not private_TryResolveHeaderInRecordset(rsSchema, VBA.CStr(sourceColumnHeaders.Item(i)), resolvedSourceColumnHeader) Then
            MsgBox "PrototypeNew: mapped source header '" & VBA.CStr(sourceColumnHeaders.Item(i)) & "' is not found. Available fields: " & availableFields & private_GenericFieldsHint(hasGenericFields), vbExclamation, RUNTIME_ERROR_TITLE
            GoTo CleanupFail
        End If
        resolvedSourceColumnHeaders.Add resolvedSourceColumnHeader
    Next i

    rsSchema.Close
    Set rsSchema = Nothing

    sql = "SELECT " & private_BuildSelectColumnsClause(resolvedSourceColumnHeaders) & " FROM " & tableRef
    Set rsData = CreateObject("ADODB.Recordset")
    rsData.Open sql, conn, 0, 1

    If rsData.EOF Then
        rsData.Close
        Set rsData = Nothing
        fn_TrySqlRequest = True
        GoTo CleanupDone
    End If

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = "Query Result"

    For i = 1 To mappedColumnHeaders.Count
        Set colObj = New obj_Column
        colObj.Position = i
        colObj.Name = VBA.Trim$(VBA.CStr(mappedColumnHeaders.Item(i)))
        If VBA.Len(colObj.Name) = 0 Then colObj.Name = "Col" & VBA.CStr(i)
        If Not tableObj.AddColumn(colObj) Then GoTo CleanupFail
    Next i

    ReDim sourceColumnOrdinals(1 To resolvedSourceColumnHeaders.Count)
    For i = 1 To resolvedSourceColumnHeaders.Count
        sourceColumnOrdinals(i) = private_RecordsetGetFieldOrdinal(rsData, VBA.CStr(resolvedSourceColumnHeaders.Item(i)))
        If sourceColumnOrdinals(i) < 0 Then
            MsgBox "PrototypeNew: resolved header '" & VBA.CStr(resolvedSourceColumnHeaders.Item(i)) & "' is not available in data recordset.", vbExclamation, RUNTIME_ERROR_TITLE
            GoTo CleanupFail
        End If
    Next i

    rowIndex = 0
    Do While Not rsData.EOF
        rowIndex = rowIndex + 1
        Set rowObj = New obj_Row
        For i = 1 To UBound(sourceColumnOrdinals)
            rowObj.AddCell private_ToSafeText(rsData.Fields(sourceColumnOrdinals(i)).Value)
        Next i

        If Not tableObj.AddRow(rowObj) Then GoTo CleanupFail
        rsData.MoveNext
    Loop

    rsData.Close
    Set rsData = Nothing

    outRowsCount = rowIndex
    Set outTable = tableObj
    fn_TrySqlRequest = True

CleanupDone:
    On Error Resume Next
    If Not rsSchema Is Nothing Then If rsSchema.State <> 0 Then rsSchema.Close
    If Not rsData Is Nothing Then If rsData.State <> 0 Then rsData.Close
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    On Error GoTo 0
    Exit Function

CleanupFail:
    Set outTable = Nothing
    outRowsCount = 0
    GoTo CleanupDone

EH_QUERY:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: SQL query error [" & VBA.CStr(Err.Number) & "] " & Err.Description
#End If
    MsgBox "PrototypeNew: SQL query error [" & VBA.CStr(Err.Number) & "] " & Err.Description, vbExclamation, RUNTIME_ERROR_TITLE
    Resume CleanupFail
End Function

Private Function private_BuildSelectColumnsClause(ByVal headers As Collection) As String
    Dim i As Long

    If headers Is Nothing Then Exit Function
    If headers.Count <= 0 Then Exit Function

    For i = 1 To headers.Count
        If i > 1 Then private_BuildSelectColumnsClause = private_BuildSelectColumnsClause & ", "
        private_BuildSelectColumnsClause = private_BuildSelectColumnsClause & private_QuoteSqlIdentifier(VBA.CStr(headers.Item(i)))
    Next i
End Function

Private Function private_BuildAdoConnectionString(ByVal sourcePath As String) As String
    Dim ext As String
    Dim props As String

    sourcePath = VBA.Trim$(sourcePath)
    ext = VBA.LCase$(VBA.Mid$(sourcePath, VBA.InStrRev(sourcePath, ".") + 1))
    Select Case ext
        Case "xls"
            props = "Excel 8.0;HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"
        Case "xlsx"
            props = "Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"
        Case "xlsm"
            props = "Excel 12.0 Macro;HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"
        Case "xlsb"
            props = "Excel 12.0;HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"
        Case Else
            Err.Raise ADO_UNSUPPORTED_EXT_ERROR_CODE, "ex_ExternalExcelSqlEngine.private_BuildAdoConnectionString", _
                "Unsupported source file extension for ADO: ." & ext
    End Select

    private_BuildAdoConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePath & ";Extended Properties=""" & props & """;"
End Function

Private Function private_BuildTableRefFromSheetName(ByVal sheetName As String) As String
    sheetName = VBA.Trim$(sheetName)
    If VBA.Len(sheetName) = 0 Then Exit Function

    If VBA.Left$(sheetName, 1) = "[" And VBA.Right$(sheetName, 1) = "]" Then
        sheetName = VBA.Mid$(sheetName, 2, VBA.Len(sheetName) - 2)
    End If

    sheetName = private_CleanAdoSchemaObjectName(sheetName)
    If VBA.Len(sheetName) = 0 Then Exit Function

    If VBA.InStr(1, sheetName, "$", VBA.vbBinaryCompare) = 0 Then
        sheetName = sheetName & "$"
    End If

    private_BuildTableRefFromSheetName = private_QuoteSqlIdentifier(sheetName)
End Function

Private Function private_HasRangeMarkers(ByVal sqlParams As obj_SqlParams) As Boolean
    If sqlParams Is Nothing Then Exit Function
    private_HasRangeMarkers = (VBA.Len(VBA.Trim$(sqlParams.RangeStartMarker)) > 0 And VBA.Len(VBA.Trim$(sqlParams.RangeEndMarker)) > 0)
End Function

Private Function private_TryBuildTableRefFromMarkers( _
    ByVal sourcePath As String, _
    ByVal configuredSheetName As String, _
    ByVal rangeStartMarker As String, _
    ByVal rangeEndMarker As String, _
    ByRef outTableRef As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim startCell As Range
    Dim endCell As Range
    Dim firstHeaderCell As Range
    Dim markerColumn As Long
    Dim headerRow As Long
    Dim dataLastRow As Long
    Dim leftCol As Long
    Dim topRow As Long
    Dim rightCol As Long
    Dim bottomRow As Long
    Dim sheetToken As String
    Dim isStartCellAddress As Boolean
    Dim isEndCellAddress As Boolean
    Dim openedHere As Boolean
    Dim hiddenExcelApp As Object
    Dim cacheKey As String
    Dim cachedTableRef As Variant

    On Error GoTo EH_RANGE

    outTableRef = VBA.vbNullString
    outErrorText = VBA.vbNullString
    rangeStartMarker = VBA.Trim$(rangeStartMarker)
    rangeEndMarker = VBA.Trim$(rangeEndMarker)

    If (VBA.Len(rangeStartMarker) > 0 Xor VBA.Len(rangeEndMarker) > 0) Then
        outErrorText = "RangeStartMarker and RangeEndMarker must be provided together."
        Exit Function
    End If
    If VBA.Len(rangeStartMarker) = 0 Then
        outErrorText = "Markers are empty."
        Exit Function
    End If

    If VBA.Dir$(sourcePath) = VBA.vbNullString Then
        outErrorText = "Source file not found: " & sourcePath
        Exit Function
    End If

    cacheKey = private_BuildMarkerRangeCacheKey(sourcePath, configuredSheetName, rangeStartMarker, rangeEndMarker)
    If ex_CacheRuntime.fn_TryGetValue(RANGE_REF_CACHE_NAMESPACE, cacheKey, cachedTableRef) Then
        outTableRef = VBA.Trim$(VBA.CStr(cachedTableRef))
        If VBA.Len(outTableRef) > 0 Then
            private_TryBuildTableRefFromMarkers = True
            Exit Function
        End If
    End If

    Set wb = private_FindOpenWorkbookByPath(sourcePath)
    If wb Is Nothing Then
        Set hiddenExcelApp = CreateObject("Excel.Application")
        hiddenExcelApp.Visible = False
        hiddenExcelApp.ScreenUpdating = False
        hiddenExcelApp.DisplayAlerts = False
        hiddenExcelApp.EnableEvents = False

        Set wb = hiddenExcelApp.Workbooks.Open( _
            Filename:=sourcePath, _
            ReadOnly:=True, _
            UpdateLinks:=0, _
            AddToMru:=False)
        openedHere = True
    End If

    Set ws = private_FindWorksheetByConfiguredSheetName(wb, configuredSheetName)
    If ws Is Nothing Then
        outErrorText = "Worksheet was not found for SheetName '" & configuredSheetName & "'."
        GoTo CleanupFail
    End If

    isStartCellAddress = private_IsCellReferenceMarker(rangeStartMarker)
    isEndCellAddress = private_IsCellReferenceMarker(rangeEndMarker)

    If Not private_TryResolveMarkerCell(ws, rangeStartMarker, startCell, outErrorText) Then GoTo CleanupFail

    If isEndCellAddress Then
        If Not private_TryResolveMarkerCell(ws, rangeEndMarker, endCell, outErrorText) Then GoTo CleanupFail
    Else
        markerColumn = startCell.Column
        Set endCell = private_FindMarkerTextCellInColumnAfterRow(ws, markerColumn, rangeEndMarker, startCell.Row)
        If endCell Is Nothing Then Set endCell = private_FindMarkerTextCellAfterAnchor(ws, rangeEndMarker, startCell)
        If endCell Is Nothing Then
            outErrorText = "End marker '" & rangeEndMarker & "' was not found on sheet '" & ws.Name & "'."
            GoTo CleanupFail
        End If
    End If

    ' Legacy-режим:
    ' стартовый маркер текстовый, а start/end находятся в одной колонке.
    ' Тогда маркерная колонка считается правой границей таблицы, header берется строкой выше start.
    If (Not isStartCellAddress) And endCell.Column = startCell.Column And endCell.Row > startCell.Row Then
        markerColumn = startCell.Column
        headerRow = startCell.Row - 1
        dataLastRow = endCell.Row - 1

        If headerRow < 1 Then
            outErrorText = "Invalid marker layout: start marker row must be greater than 1."
            GoTo CleanupFail
        End If
        If dataLastRow < startCell.Row Then
            outErrorText = "Invalid marker layout: end marker is above data rows."
            GoTo CleanupFail
        End If

        Set firstHeaderCell = private_FindFirstNonEmptyHeaderCell(ws, headerRow, markerColumn)
        If firstHeaderCell Is Nothing Then
            outErrorText = "Header row " & VBA.CStr(headerRow) & " has no cells before marker column."
            GoTo CleanupFail
        End If

        leftCol = firstHeaderCell.Column
        topRow = headerRow
        rightCol = markerColumn
        bottomRow = dataLastRow
    Else
        topRow = startCell.Row
        If endCell.Row < topRow Then topRow = endCell.Row
        bottomRow = startCell.Row
        If endCell.Row > bottomRow Then bottomRow = endCell.Row

        leftCol = startCell.Column
        If endCell.Column < leftCol Then leftCol = endCell.Column
        rightCol = startCell.Column
        If endCell.Column > rightCol Then rightCol = endCell.Column
    End If

    If leftCol <= 0 Or rightCol <= 0 Or topRow <= 0 Or bottomRow <= 0 Then
        outErrorText = "Failed to calculate marker-based range bounds."
        GoTo CleanupFail
    End If
    If rightCol < leftCol Or bottomRow < topRow Then
        outErrorText = "Invalid marker-based range bounds."
        GoTo CleanupFail
    End If

    sheetToken = private_BuildAdoSheetTokenForRange(configuredSheetName)
    If VBA.Len(sheetToken) = 0 Then
        outErrorText = "Failed to build sheet token from SheetName '" & configuredSheetName & "'."
        GoTo CleanupFail
    End If

    outTableRef = private_QuoteSqlIdentifier( _
        sheetToken & private_ToColumnLetter(leftCol) & VBA.CStr(topRow) & ":" & private_ToColumnLetter(rightCol) & VBA.CStr(bottomRow))

    Call ex_CacheRuntime.fn_SetValue(RANGE_REF_CACHE_NAMESPACE, cacheKey, outTableRef)
    private_TryBuildTableRefFromMarkers = True

CleanupDone:
    On Error Resume Next
    If openedHere Then
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
        If Not hiddenExcelApp Is Nothing Then hiddenExcelApp.Quit
    End If
    Set ws = Nothing
    Set wb = Nothing
    Set hiddenExcelApp = Nothing
    On Error GoTo 0
    Exit Function

CleanupFail:
    outTableRef = VBA.vbNullString
    GoTo CleanupDone

EH_RANGE:
    outErrorText = "[" & Err.Source & " #" & VBA.CStr(Err.Number) & "] " & Err.Description
    Resume CleanupFail
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

Private Function private_FindWorksheetByConfiguredSheetName( _
    ByVal wb As Workbook, _
    ByVal configuredSheetName As String _
) As Worksheet
    Dim ws As Worksheet
    Dim needle As String
    Dim needleAlt As String

    If wb Is Nothing Then Exit Function

    needle = private_ExtractSheetNameToken(configuredSheetName)
    If VBA.Len(needle) = 0 Then Exit Function
    needleAlt = VBA.Replace$(needle, "#", ".")

    For Each ws In wb.Worksheets
        If VBA.StrComp(VBA.Trim$(ws.Name), needle, VBA.vbTextCompare) = 0 Then
            Set private_FindWorksheetByConfiguredSheetName = ws
            Exit Function
        End If
        If VBA.StrComp(VBA.Replace$(VBA.Trim$(ws.Name), ".", "#"), needle, VBA.vbTextCompare) = 0 Then
            Set private_FindWorksheetByConfiguredSheetName = ws
            Exit Function
        End If
        If VBA.StrComp(VBA.Trim$(ws.Name), needleAlt, VBA.vbTextCompare) = 0 Then
            Set private_FindWorksheetByConfiguredSheetName = ws
            Exit Function
        End If
    Next ws
End Function

Private Function private_ExtractSheetNameToken(ByVal configuredSheetName As String) As String
    Dim token As String
    Dim dollarPos As Long

    token = VBA.Trim$(configuredSheetName)
    If VBA.Len(token) = 0 Then Exit Function

    If VBA.Left$(token, 1) = "[" And VBA.Right$(token, 1) = "]" Then
        token = VBA.Mid$(token, 2, VBA.Len(token) - 2)
    End If
    token = private_CleanAdoSchemaObjectName(token)

    dollarPos = VBA.InStr(1, token, "$", VBA.vbBinaryCompare)
    If dollarPos > 0 Then token = VBA.Left$(token, dollarPos - 1)

    private_ExtractSheetNameToken = VBA.Trim$(token)
End Function

Private Function private_BuildAdoSheetTokenForRange(ByVal configuredSheetName As String) As String
    Dim token As String
    Dim dollarPos As Long

    token = VBA.Trim$(configuredSheetName)
    If VBA.Len(token) = 0 Then Exit Function

    If VBA.Left$(token, 1) = "[" And VBA.Right$(token, 1) = "]" Then
        token = VBA.Mid$(token, 2, VBA.Len(token) - 2)
    End If
    token = private_CleanAdoSchemaObjectName(token)
    token = VBA.Trim$(token)
    If VBA.Len(token) = 0 Then Exit Function

    dollarPos = VBA.InStr(1, token, "$", VBA.vbBinaryCompare)
    If dollarPos > 0 Then
        private_BuildAdoSheetTokenForRange = VBA.Left$(token, dollarPos)
    Else
        private_BuildAdoSheetTokenForRange = token & "$"
    End If
End Function

Private Function private_IsCellReferenceMarker(ByVal markerText As String) As Boolean
    markerText = VBA.Trim$(markerText)
    private_IsCellReferenceMarker = (VBA.Left$(markerText, 1) = "$")
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

    Set private_FindFirstMarkerTextCell = searchRange.Find( _
        What:=markerText, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
End Function

Private Function private_FindMarkerTextCellInColumnAfterRow( _
    ByVal ws As Worksheet, _
    ByVal markerColumn As Long, _
    ByVal markerText As String, _
    ByVal minExclusiveRow As Long _
) As Range
    Dim searchRange As Range
    Dim firstFound As Range
    Dim currentFound As Range
    Dim firstAddress As String
    Dim bestRow As Long

    If ws Is Nothing Then Exit Function
    If markerColumn <= 0 Then Exit Function
    markerText = VBA.Trim$(markerText)
    If VBA.Len(markerText) = 0 Then Exit Function

    On Error Resume Next
    Set searchRange = Intersect(ws.Columns(markerColumn), ws.UsedRange)
    On Error GoTo 0
    If searchRange Is Nothing Then Set searchRange = ws.Columns(markerColumn)

    Set firstFound = searchRange.Find( _
        What:=markerText, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
    If firstFound Is Nothing Then Exit Function

    bestRow = 0
    firstAddress = firstFound.Address
    Set currentFound = firstFound

    Do
        If currentFound.Row > minExclusiveRow Then
            If bestRow = 0 Or currentFound.Row < bestRow Then
                bestRow = currentFound.Row
                Set private_FindMarkerTextCellInColumnAfterRow = currentFound
            End If
        End If
        Set currentFound = searchRange.FindNext(currentFound)
        If currentFound Is Nothing Then Exit Do
    Loop While currentFound.Address <> firstAddress
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

    Set firstFound = searchRange.Find( _
        What:=markerText, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
    If firstFound Is Nothing Then Exit Function

    bestWeight = -1
    firstAddress = firstFound.Address
    Set currentFound = firstFound

    Do
        If currentFound.Row > anchorCell.Row Or (currentFound.Row = anchorCell.Row And currentFound.Column >= anchorCell.Column) Then
            currentWeight = CDbl(currentFound.Row) * 100000# + CDbl(currentFound.Column)
            If bestWeight < 0 Or currentWeight < bestWeight Then
                bestWeight = currentWeight
                Set private_FindMarkerTextCellAfterAnchor = currentFound
            End If
        End If

        Set currentFound = searchRange.FindNext(currentFound)
        If currentFound Is Nothing Then Exit Do
    Loop While currentFound.Address <> firstAddress
End Function

Private Function private_FindFirstNonEmptyHeaderCell( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal maxColumn As Long _
) As Range
    If ws Is Nothing Then Exit Function
    If headerRow <= 0 Then Exit Function
    If maxColumn <= 0 Then Exit Function

    Set private_FindFirstNonEmptyHeaderCell = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, maxColumn)).Find( _
        What:="*", _
        After:=ws.Cells(headerRow, maxColumn), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
End Function

Private Function private_ToColumnLetter(ByVal columnNumber As Long) As String
    Dim n As Long

    n = columnNumber
    If n <= 0 Then Exit Function

    Do While n > 0
        private_ToColumnLetter = VBA.Chr$(((n - 1) Mod 26) + 65) & private_ToColumnLetter
        n = (n - 1) \ 26
    Loop
End Function

Private Function private_BuildMarkerRangeCacheKey( _
    ByVal sourcePath As String, _
    ByVal configuredSheetName As String, _
    ByVal rangeStartMarker As String, _
    ByVal rangeEndMarker As String _
) As String
    private_BuildMarkerRangeCacheKey = _
        RANGE_REF_CACHE_VERSION & "|" & _
        private_NormalizeCacheKeyPart(sourcePath) & "|" & _
        private_BuildFileVersionToken(sourcePath) & "|" & _
        private_NormalizeCacheKeyPart(configuredSheetName) & "|" & _
        private_NormalizeCacheKeyPart(rangeStartMarker) & "|" & _
        private_NormalizeCacheKeyPart(rangeEndMarker)
End Function

Private Function private_BuildFileVersionToken(ByVal sourcePath As String) As String
    Dim fileDateToken As String
    Dim fileSizeToken As String

    On Error Resume Next
    fileDateToken = VBA.Format$(VBA.FileDateTime(sourcePath), "yyyymmddhhnnss")
    fileSizeToken = VBA.CStr(VBA.FileLen(sourcePath))
    On Error GoTo 0

    If VBA.Len(fileDateToken) = 0 Then fileDateToken = "0"
    If VBA.Len(fileSizeToken) = 0 Then fileSizeToken = "0"

    private_BuildFileVersionToken = fileDateToken & ":" & fileSizeToken
End Function

Private Function private_NormalizeCacheKeyPart(ByVal valueText As String) As String
    private_NormalizeCacheKeyPart = VBA.LCase$(VBA.Trim$(VBA.CStr(valueText)))
End Function

Private Function private_CleanAdoSchemaObjectName(ByVal value As String) As String
    Dim cleaned As String

    cleaned = VBA.Trim$(value)
    If VBA.Len(cleaned) = 0 Then Exit Function

    If VBA.Left$(cleaned, 1) = "[" And VBA.Right$(cleaned, 1) = "]" Then
        cleaned = VBA.Mid$(cleaned, 2, VBA.Len(cleaned) - 2)
    End If

    cleaned = VBA.Replace$(cleaned, "]]", "]")
    cleaned = VBA.Replace$(cleaned, "'", VBA.vbNullString)
    private_CleanAdoSchemaObjectName = VBA.Trim$(cleaned)
End Function

Private Function private_QuoteSqlIdentifier(ByVal valueText As String) As String
    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) >= 2 Then
        If VBA.Left$(valueText, 1) = "[" And VBA.Right$(valueText, 1) = "]" Then
            valueText = VBA.Mid$(valueText, 2, VBA.Len(valueText) - 2)
        End If
    End If

    private_QuoteSqlIdentifier = "[" & VBA.Replace$(valueText, "]", "]]" ) & "]"
End Function

Private Function private_TryResolveHeaderInRecordset( _
    ByVal rs As Object, _
    ByVal desiredHeader As String, _
    ByRef outResolvedHeader As String _
) As Boolean
    Dim desiredToken As String
    Dim i As Long
    Dim fieldName As String

    outResolvedHeader = VBA.vbNullString
    If rs Is Nothing Then Exit Function

    desiredToken = private_NormalizeHeader(desiredHeader)
    If VBA.Len(desiredToken) = 0 Then Exit Function

    For i = 0 To rs.Fields.Count - 1
        fieldName = VBA.CStr(rs.Fields(i).Name)
        If VBA.StrComp(private_NormalizeHeader(fieldName), desiredToken, VBA.vbTextCompare) = 0 Then
            outResolvedHeader = fieldName
            private_TryResolveHeaderInRecordset = True
            Exit Function
        End If
    Next i
End Function

Private Function private_RecordsetGetFieldOrdinal(ByVal rs As Object, ByVal fieldName As String) As Long
    Dim i As Long
    Dim targetToken As String

    private_RecordsetGetFieldOrdinal = -1
    If rs Is Nothing Then Exit Function

    targetToken = private_NormalizeHeader(fieldName)
    If VBA.Len(targetToken) = 0 Then Exit Function

    For i = 0 To rs.Fields.Count - 1
        If VBA.StrComp(private_NormalizeHeader(VBA.CStr(rs.Fields(i).Name)), targetToken, VBA.vbTextCompare) = 0 Then
            private_RecordsetGetFieldOrdinal = i
            Exit Function
        End If
    Next i
End Function

Private Function private_NormalizeHeader(ByVal valueText As String) As String
    Dim normalized As String

    normalized = VBA.CStr(valueText)
    normalized = VBA.Replace$(normalized, VBA.vbCr, " ")
    normalized = VBA.Replace$(normalized, VBA.vbLf, " ")
    normalized = VBA.Replace$(normalized, VBA.vbTab, " ")
    normalized = VBA.Replace$(normalized, VBA.ChrW$(160), " ")
    normalized = VBA.Replace$(normalized, "#", ".")
    normalized = VBA.Replace$(normalized, VBA.ChrW$(&H2019), "'")
    normalized = VBA.Replace$(normalized, VBA.ChrW$(&H2BC), "'")
    normalized = VBA.Replace$(normalized, VBA.ChrW$(&H60), "'")
    normalized = VBA.Replace$(normalized, VBA.ChrW$(&HB4), "'")
    normalized = VBA.Replace$(normalized, "  ", " ")
    normalized = VBA.Replace$(normalized, "  ", " ")
    normalized = VBA.Trim$(normalized)
    normalized = VBA.LCase$(normalized)

    private_NormalizeHeader = normalized
End Function

Private Function private_ListRecordsetFields(ByVal rs As Object, Optional ByVal maxCount As Long = 25) As String
    Dim i As Long
    Dim count As Long
    Dim fieldName As String

    If rs Is Nothing Then Exit Function
    If maxCount <= 0 Then maxCount = 25

    For i = 0 To rs.Fields.Count - 1
        fieldName = VBA.Trim$(VBA.CStr(rs.Fields(i).Name))
        If VBA.Len(fieldName) = 0 Then fieldName = "(empty)"
        If count > 0 Then private_ListRecordsetFields = private_ListRecordsetFields & ", "
        private_ListRecordsetFields = private_ListRecordsetFields & "[" & fieldName & "]"
        count = count + 1
        If count >= maxCount Then Exit For
    Next i

    If rs.Fields.Count > maxCount Then
        private_ListRecordsetFields = private_ListRecordsetFields & ", ..."
    End If
End Function

Private Function private_RecordsetLooksLikeGenericFields(ByVal rs As Object) As Boolean
    Dim i As Long
    Dim probeCount As Long
    Dim fieldName As String

    If rs Is Nothing Then Exit Function
    If rs.Fields.Count = 0 Then Exit Function

    probeCount = rs.Fields.Count
    If probeCount > 10 Then probeCount = 10

    For i = 0 To probeCount - 1
        fieldName = VBA.UCase$(VBA.Trim$(VBA.CStr(rs.Fields(i).Name)))
        If VBA.Len(fieldName) < 2 Then Exit Function
        If VBA.Left$(fieldName, 1) <> "F" Then Exit Function
        If Not VBA.IsNumeric(VBA.Mid$(fieldName, 2)) Then Exit Function
    Next i

    private_RecordsetLooksLikeGenericFields = True
End Function

Private Function private_GenericFieldsHint(ByVal hasGenericFields As Boolean) As String
    If Not hasGenericFields Then Exit Function
    private_GenericFieldsHint = " Hint: ADO returned generic fields (F1..Fn). Configure explicit range with real headers."
End Function

Private Function private_ResolvePathLocal(ByVal inputPath As String) As String
    Dim basePath As String

    inputPath = VBA.Trim$(inputPath)
    If VBA.Len(inputPath) = 0 Then Exit Function

    If VBA.Left$(inputPath, 2) = "\\" Or VBA.InStr(1, inputPath, ":\", VBA.vbTextCompare) > 0 Then
        private_ResolvePathLocal = inputPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If VBA.Len(basePath) = 0 Then basePath = CurDir$
    If VBA.Right$(basePath, 1) <> "\" Then basePath = basePath & "\"

    private_ResolvePathLocal = basePath & inputPath
End Function

Private Function private_ToSafeText(ByVal valueIn As Variant) As String
    On Error GoTo EH_SAFE_TEXT
    If VBA.IsError(valueIn) Then
        private_ToSafeText = "#ERR"
        Exit Function
    End If
    If VBA.IsNull(valueIn) Or VBA.IsEmpty(valueIn) Then
        private_ToSafeText = VBA.vbNullString
        Exit Function
    End If

    private_ToSafeText = VBA.CStr(valueIn)
    Exit Function

EH_SAFE_TEXT:
    private_ToSafeText = VBA.vbNullString
End Function
