Attribute VB_Name = "ex_TableComparer"
Option Explicit

Public Function CompareSheets( _
    ByVal oldSheetName As String, _
    ByVal newSheetName As String, _
    ByVal keyColumnName As String _
) As Variant
    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    
    Set wsOld = ThisWorkbook.Worksheets("g_" & oldSheetName)
    Set wsNew = ThisWorkbook.Worksheets("g_" & newSheetName)
    
    CompareSheets = CompareRanges( _
        wsOld.UsedRange, _
        wsNew.UsedRange, _
        keyColumnName _
    )
End Function

Private Function CompareRanges( _
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
    
    oldHeaders = ReadHeaderRow(oldRange)
    newHeaders = ReadHeaderRow(newRange)
    
    oldKeyCol = FindHeaderIndex(oldHeaders, keyColumnName)
    newKeyCol = FindHeaderIndex(newHeaders, keyColumnName)
    
    If oldKeyCol = 0 Or newKeyCol = 0 Then
        Err.Raise vbObjectError + 1, "CompareRanges", "Key column not found: " & keyColumnName
    End If
    
    Set oldDict = BuildRowDict(oldRange, oldKeyCol)
    Set newDict = BuildRowDict(newRange, newKeyCol)
    
    ' ------------------------------------------------------------------
    ' Allocate output array ONCE (no ReDim Preserve on 2D arrays!)
    ' Rows count:
    '   1 header
    '   + all New keys (OK/Changed/Added)
    '   + Old keys missing in New (Removed)
    ' ------------------------------------------------------------------
    Dim totalRows As Long
    
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
    
    ' Header row
    outData(1, 1) = keyColumnName
    outData(1, 2) = "Status"
    
    For i = 1 To UBound(newHeaders)
        outData(1, 2 + i) = newHeaders(i)
    Next i
    
    outRow = 1
    
    ' ------------------------------------------------------------------
    ' Added / Changed / OK (iterate New)
    ' ------------------------------------------------------------------
    For Each key In newDict.Keys
        keyValue = CStr(key)
        outRow = outRow + 1
        
        outData(outRow, 1) = keyValue
        
        If Not oldDict.Exists(keyValue) Then
            outData(outRow, 2) = "Added"
        Else
            If RowsAreDifferent(oldDict(keyValue), newDict(keyValue)) Then
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
    
    ' ------------------------------------------------------------------
    ' Removed (iterate Old keys not in New)
    ' ------------------------------------------------------------------
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
    
    CompareRanges = outData
End Function

' -----------------------------------------------------------------------------
' Helpers
' -----------------------------------------------------------------------------
Private Function ReadHeaderRow(ByVal dataRange As Range) As Variant
    Dim colCount As Long
    Dim headers() As Variant
    Dim c As Long
    
    colCount = dataRange.Columns.Count
    ReDim headers(1 To colCount)
    
    For c = 1 To colCount
        headers(c) = CStr(dataRange.Cells(1, c).Value)
    Next c
    
    ReadHeaderRow = headers
End Function

Private Function FindHeaderIndex(ByVal headers As Variant, ByVal name As String) As Long
    Dim i As Long
    
    For i = LBound(headers) To UBound(headers)
        If StrComp(CStr(headers(i)), name, vbTextCompare) = 0 Then
            FindHeaderIndex = i
            Exit Function
        End If
    Next i
    
    FindHeaderIndex = 0
End Function

Private Function BuildRowDict(ByVal dataRange As Range, ByVal keyCol As Long) As Object
    Dim dict As Object
    Dim r As Long
    Dim rowCount As Long
    Dim colCount As Long
    
    Dim keyValue As String
    Dim rowArr() As Variant
    Dim c As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    rowCount = dataRange.Rows.Count
    colCount = dataRange.Columns.Count
    
    For r = 2 To rowCount
        keyValue = CStr(dataRange.Cells(r, keyCol).Value)
        
        If Len(keyValue) > 0 Then
            ReDim rowArr(1 To colCount)
            
            For c = 1 To colCount
                rowArr(c) = CStr(dataRange.Cells(r, c).Value)
            Next c
            
            dict(keyValue) = rowArr
        End If
    Next r
    
    Set BuildRowDict = dict
End Function

Private Function RowsAreDifferent(ByVal oldRow As Variant, ByVal newRow As Variant) As Boolean
    Dim i As Long
    
    If UBound(oldRow) <> UBound(newRow) Then
        RowsAreDifferent = True
        Exit Function
    End If
    
    For i = LBound(oldRow) To UBound(oldRow)
        If CStr(oldRow(i)) <> CStr(newRow(i)) Then
            RowsAreDifferent = True
            Exit Function
        End If
    Next i
    
    RowsAreDifferent = False
End Function
