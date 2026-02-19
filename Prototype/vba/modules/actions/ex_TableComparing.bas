Attribute VB_Name = "ex_TableComparing"
Option Explicit

Public Sub m_RunComparing()
    Dim resultTable As Variant
    Dim keyColumnName As String

    On Error GoTo EH
    If Not ex_SheetStylesXmlProvider.m_EnsureInitialized(ThisWorkbook) Then
        Err.Raise vbObjectError + 2201, "ex_TableComparing", "Failed to initialize style registry."
    End If

    keyColumnName = mp_GetKeyColumnFromConfig()
    If Len(keyColumnName) = 0 Then
        keyColumnName = "Id"
    End If

    resultTable = m_CompareConfiguredTables(keyColumnName)
    ex_ResultWriter.m_WriteTableToResultSheet resultTable
    Exit Sub
EH:
    MsgBox "Run error: " & Err.Description, vbExclamation
End Sub

Public Function m_CompareConfiguredTables(ByVal keyColumnName As String) As Variant
    Dim oldPath As String
    Dim oldTableName As String
    Dim newPath As String
    Dim newTableName As String

    oldPath = mp_ResolvePath(mp_GetConfigValueSafe("OldFilePath"))
    oldTableName = Trim$(mp_GetConfigValueSafe("OldTableName"))
    newPath = mp_ResolvePath(mp_GetConfigValueSafe("NewFilePath"))
    newTableName = Trim$(mp_GetConfigValueSafe("NewTableName"))

    If Len(oldPath) = 0 Or Len(oldTableName) = 0 Or Len(newPath) = 0 Or Len(newTableName) = 0 Then
        Err.Raise vbObjectError + 2202, "ex_TableComparing", _
            "Config keys OldFilePath/OldTableName/NewFilePath/NewTableName are required."
    End If

    If Dir(oldPath) = vbNullString Then
        Err.Raise vbObjectError + 2203, "ex_TableComparing", "OldFilePath not found: " & oldPath
    End If

    If Dir(newPath) = vbNullString Then
        Err.Raise vbObjectError + 2204, "ex_TableComparing", "NewFilePath not found: " & newPath
    End If

    m_CompareConfiguredTables = mp_CompareExternalTables(oldPath, oldTableName, newPath, newTableName, keyColumnName)
End Function

Private Function mp_CompareExternalTables( _
    ByVal oldWorkbookPath As String, _
    ByVal oldTableName As String, _
    ByVal newWorkbookPath As String, _
    ByVal newTableName As String, _
    ByVal keyColumnName As String _
) As Variant
    Dim wbOld As Workbook
    Dim wbNew As Workbook
    Dim loOld As ListObject
    Dim loNew As ListObject
    Dim hadError As Boolean
    Dim errText As String
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean

    On Error GoTo EH
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbOld = mp_OpenWorkbookHidden(oldWorkbookPath)
    Set wbNew = mp_OpenWorkbookHidden(newWorkbookPath)

    Set loOld = mp_FindListObjectByName(wbOld, oldTableName)
    If loOld Is Nothing Then
        Err.Raise vbObjectError + 2205, "ex_TableComparing", "Table '" & oldTableName & "' not found in file: " & oldWorkbookPath
    End If

    Set loNew = mp_FindListObjectByName(wbNew, newTableName)
    If loNew Is Nothing Then
        Err.Raise vbObjectError + 2206, "ex_TableComparing", "Table '" & newTableName & "' not found in file: " & newWorkbookPath
    End If

    mp_CompareExternalTables = mp_CompareRanges(loOld.Range, loNew.Range, keyColumnName)

Cleanup:
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    If Not wbOld Is Nothing Then wbOld.Close SaveChanges:=False
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    If hadError Then
        Err.Raise vbObjectError + 2207, "ex_TableComparing", errText
    End If
    Exit Function
EH:
    errText = Err.Description
    hadError = True
    Resume Cleanup
End Function

Private Function mp_OpenWorkbookHidden(ByVal filePath As String) As Workbook
    Dim wb As Workbook

    Set wb = Workbooks.Open( _
        Filename:=filePath, _
        ReadOnly:=True, _
        UpdateLinks:=0, _
        AddToMru:=False)

    On Error Resume Next
    If wb.Windows.Count > 0 Then wb.Windows(1).Visible = False
    On Error GoTo 0

    Set mp_OpenWorkbookHidden = wb
End Function

Private Function mp_FindListObjectByName(ByVal wbSrc As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wbSrc.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set mp_FindListObjectByName = lo
                Exit Function
            End If
        Next lo
    Next ws

    Set mp_FindListObjectByName = Nothing
End Function

Private Function mp_GetKeyColumnFromConfig() As String
    On Error GoTo Fallback
    mp_GetKeyColumnFromConfig = ex_ConfigProvider.m_GetConfigValue("KeyColumnName", "Id")
    Exit Function
Fallback:
    mp_GetKeyColumnFromConfig = "Id"
End Function

Private Function mp_ResolvePath(ByVal inputPath As String) As String
    Dim basePath As String

    inputPath = Trim$(inputPath)
    If Len(inputPath) = 0 Then Exit Function

    If Left$(inputPath, 2) = "\\" Or InStr(1, inputPath, ":\", vbTextCompare) > 0 Then
        mp_ResolvePath = inputPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        mp_ResolvePath = inputPath
        Exit Function
    End If

    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    mp_ResolvePath = basePath & inputPath
End Function

Private Function mp_GetConfigValueSafe(ByVal keyName As String) As String
    On Error GoTo Fallback
    mp_GetConfigValueSafe = ex_ConfigProvider.m_GetConfigValue(keyName, vbNullString)
    Exit Function
Fallback:
    mp_GetConfigValueSafe = vbNullString
End Function

Private Function mp_CompareRanges( _
    ByVal oldRange As Range, _
    ByVal newRange As Range, _
    ByVal keyColumnName As String _
) As Variant
    Dim oldHeaders As Variant
    Dim newHeaders As Variant
    Dim oldKeyCol As Long
    Dim newKeyCol As Long
    Dim oldDict As Object
    Dim newDict As Object
    Dim outData() As Variant
    Dim outRow As Long
    Dim outColCount As Long
    Dim i As Long
    Dim keyValue As String
    Dim key As Variant
    Dim rowArr As Variant
    Dim totalRows As Long

    oldHeaders = mp_ReadHeaderRow(oldRange)
    newHeaders = mp_ReadHeaderRow(newRange)

    oldKeyCol = mp_FindHeaderIndex(oldHeaders, keyColumnName)
    newKeyCol = mp_FindHeaderIndex(newHeaders, keyColumnName)

    If oldKeyCol = 0 Or newKeyCol = 0 Then
        Err.Raise vbObjectError + 1, "CompareRanges", "Key column not found: " & keyColumnName
    End If

    Set oldDict = mp_BuildRowDict(oldRange, oldKeyCol)
    Set newDict = mp_BuildRowDict(newRange, newKeyCol)

    totalRows = 1

    For Each key In newDict.Keys
        totalRows = totalRows + 1
    Next key

    For Each key In oldDict.Keys
        If Not newDict.Exists(CStr(key)) Then
            totalRows = totalRows + 1
        End If
    Next key

    outColCount = 2 + UBound(newHeaders)
    ReDim outData(1 To totalRows, 1 To outColCount)

    outData(1, 1) = keyColumnName
    outData(1, 2) = "Status"

    For i = 1 To UBound(newHeaders)
        outData(1, 2 + i) = newHeaders(i)
    Next i

    outRow = 1

    For Each key In newDict.Keys
        keyValue = CStr(key)
        outRow = outRow + 1

        outData(outRow, 1) = keyValue

        If Not oldDict.Exists(keyValue) Then
            outData(outRow, 2) = "Added"
        Else
            If mp_RowsAreDifferent(oldDict(keyValue), newDict(keyValue)) Then
                outData(outRow, 2) = "Changed"
            Else
                outData(outRow, 2) = "OK"
            End If
        End If

        rowArr = newDict(keyValue)
        For i = 1 To UBound(rowArr)
            outData(outRow, 2 + i) = rowArr(i)
        Next i
    Next key

    For Each key In oldDict.Keys
        keyValue = CStr(key)

        If Not newDict.Exists(keyValue) Then
            outRow = outRow + 1

            outData(outRow, 1) = keyValue
            outData(outRow, 2) = "Removed"

            rowArr = oldDict(keyValue)
            For i = 1 To UBound(rowArr)
                outData(outRow, 2 + i) = rowArr(i)
            Next i
        End If
    Next key

    mp_CompareRanges = outData
End Function

Private Function mp_ReadHeaderRow(ByVal dataRange As Range) As Variant
    Dim data As Variant
    Dim colCount As Long
    Dim headers() As Variant
    Dim c As Long

    data = dataRange.Value2
    colCount = dataRange.Columns.Count
    ReDim headers(1 To colCount)

    For c = 1 To colCount
        headers(c) = CStr(data(1, c))
    Next c

    mp_ReadHeaderRow = headers
End Function

Private Function mp_FindHeaderIndex(ByVal headers As Variant, ByVal name As String) As Long
    Dim i As Long

    For i = LBound(headers) To UBound(headers)
        If StrComp(CStr(headers(i)), name, vbTextCompare) = 0 Then
            mp_FindHeaderIndex = i
            Exit Function
        End If
    Next i

    mp_FindHeaderIndex = 0
End Function

Private Function mp_BuildRowDict(ByVal dataRange As Range, ByVal keyCol As Long) As Object
    Dim dict As Object
    Dim data As Variant
    Dim r As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim effectiveColCount As Long
    Dim keyValue As String
    Dim rowArr() As Variant
    Dim c As Long

    Set dict = CreateObject("Scripting.Dictionary")

    data = dataRange.Value2
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)

    effectiveColCount = 0
    For c = colCount To 1 Step -1
        If Len(CStr(data(1, c))) > 0 Then
            effectiveColCount = c
            Exit For
        End If
    Next c
    If effectiveColCount = 0 Then effectiveColCount = colCount

    For r = 2 To rowCount
        keyValue = CStr(data(r, keyCol))

        If Len(keyValue) > 0 Then
            ReDim rowArr(1 To effectiveColCount)
            For c = 1 To effectiveColCount
                rowArr(c) = CStr(data(r, c))
            Next c
            dict(keyValue) = rowArr
        End If
    Next r

    Set mp_BuildRowDict = dict
End Function

Private Function mp_RowsAreDifferent(ByVal oldRow As Variant, ByVal newRow As Variant) As Boolean
    Dim i As Long

    If UBound(oldRow) <> UBound(newRow) Then
        mp_RowsAreDifferent = True
        Exit Function
    End If

    For i = LBound(oldRow) To UBound(oldRow)
        If CStr(oldRow(i)) <> CStr(newRow(i)) Then
            mp_RowsAreDifferent = True
            Exit Function
        End If
    Next i

    mp_RowsAreDifferent = False
End Function
