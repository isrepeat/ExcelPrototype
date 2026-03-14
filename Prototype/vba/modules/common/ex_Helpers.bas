Attribute VB_Name = "ex_Helpers"
Option Explicit

Public Function m_RegexIsMatch( _
    ByVal textValue As String, _
    ByVal regexPattern As String _
) As Boolean
    Dim rx As Object

    Set rx = mp_CreateRegex(regexPattern)
    m_RegexIsMatch = rx.Test(CStr(textValue))
End Function

Public Function m_RegexFirstMatch( _
    ByVal textValue As String, _
    ByVal regexPattern As String _
) As String
    Dim rx As Object
    Dim matches As Object

    Set rx = mp_CreateRegex(regexPattern)
    Set matches = rx.Execute(CStr(textValue))
    If matches.Count > 0 Then
        m_RegexFirstMatch = CStr(matches(0).Value)
    End If
End Function

Public Function m_RegexGetGroup( _
    ByVal textValue As String, _
    ByVal regexPattern As String, _
    Optional ByVal groupIndex As Long = 1 _
) As String
    Dim rx As Object
    Dim matches As Object
    Dim firstMatch As Object

    If groupIndex < 0 Then
        Err.Raise vbObjectError + 1811, "ex_Helpers", "Regex group index cannot be negative."
    End If

    Set rx = mp_CreateRegex(regexPattern)
    Set matches = rx.Execute(CStr(textValue))
    If matches.Count = 0 Then Exit Function

    Set firstMatch = matches(0)
    If groupIndex = 0 Then
        m_RegexGetGroup = CStr(firstMatch.Value)
        Exit Function
    End If

    If groupIndex > firstMatch.SubMatches.Count Then Exit Function
    m_RegexGetGroup = CStr(firstMatch.SubMatches(groupIndex - 1))
End Function

Public Function m_RowCellRegexIsMatch( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal regexPattern As String _
) As Boolean
    On Error GoTo SafeExit
    m_RowCellRegexIsMatch = m_RegexIsMatch(mp_GetRowCellLiveText(rowRef, columnRef), regexPattern)
    Exit Function

SafeExit:
    m_RowCellRegexIsMatch = False
End Function

Public Function m_RowCellRegexFirstMatch( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal regexPattern As String _
) As String
    On Error GoTo SafeExit
    m_RowCellRegexFirstMatch = m_RegexFirstMatch(mp_GetRowCellLiveText(rowRef, columnRef), regexPattern)
    Exit Function

SafeExit:
    m_RowCellRegexFirstMatch = vbNullString
End Function

Public Sub m_EmphasizeRowCellTextByRegex( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal regexPattern As String, _
    Optional ByVal fontColorHex As String = "#FF0000", _
    Optional ByVal uppercaseMatches As String = "false" _
)
    Dim targetCell As Range
    Dim targetCol As Long
    Dim originalText As String
    Dim transformedText As String
    Dim rx As Object
    Dim matches As Object
    Dim i As Long
    Dim matchObj As Object
    Dim colorValue As Long
    Dim makeUpper As Boolean
    Dim ws As Worksheet
    Dim rowIndex As Long

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1814, "ex_Helpers", "Row reference is required for regex text emphasis."
    End If

    If Not mp_TryResolveColumnIndexInRow(rowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1815, "ex_Helpers", "Unknown row cell reference '" & columnRef & "' for regex text emphasis."
    End If
    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1816, "ex_Helpers", "Active sheet is not available for regex text emphasis."
    End If
    rowIndex = mp_ResolveAnchoredRowIndex(ws, rowRef, "regex text emphasis")
    Set targetCell = mp_GetRowCellRange(ws, rowIndex, targetCol)
    originalText = CStr(targetCell.Value)

    If Len(Trim$(fontColorHex)) = 0 Then fontColorHex = "#FF0000"
    If Not ex_XmlCore.m_TryParseColor(fontColorHex, colorValue) Then
        Err.Raise vbObjectError + 1817, "ex_Helpers", "Invalid regex emphasis color: " & fontColorHex
    End If
    makeUpper = mp_ParseRequiredBoolean(uppercaseMatches, "uppercaseMatches")

    Set rx = mp_CreateRegex(regexPattern, True)
    Set matches = rx.Execute(originalText)
    If matches Is Nothing Or matches.Count = 0 Then Exit Sub

    If makeUpper Then
        transformedText = originalText
        For i = 0 To matches.Count - 1
            Set matchObj = matches(i)
            If matchObj.Length > 0 Then
                transformedText = Left$(transformedText, matchObj.FirstIndex) & UCase$(Mid$(transformedText, matchObj.FirstIndex + 1, matchObj.Length)) & Mid$(transformedText, matchObj.FirstIndex + matchObj.Length + 1)
            End If
        Next i
        targetCell.Value = transformedText
    End If

    For i = 0 To matches.Count - 1
        Set matchObj = matches(i)
        If matchObj.Length > 0 Then
            targetCell.Characters(matchObj.FirstIndex + 1, matchObj.Length).Font.Color = colorValue
            targetCell.Characters(matchObj.FirstIndex + 1, matchObj.Length).Font.Bold = True
        End If
    Next i
End Sub

Private Function mp_CreateRegex( _
    ByVal regexPattern As String, _
    Optional ByVal globalMatches As Boolean = False _
) As Object
    Dim rx As Object

    regexPattern = Trim$(regexPattern)
    If Len(regexPattern) = 0 Then
        Err.Raise vbObjectError + 1812, "ex_Helpers", "Regex pattern is empty."
    End If

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = globalMatches
    rx.IgnoreCase = True
    rx.MultiLine = True

    On Error GoTo PatternErr
    rx.Pattern = regexPattern
    On Error GoTo 0

    Set mp_CreateRegex = rx
    Exit Function

PatternErr:
    Err.Raise vbObjectError + 1813, "ex_Helpers", "Invalid regex pattern '" & regexPattern & "': " & Err.Description
End Function

Private Function mp_ParseRequiredBoolean(ByVal valueText As String, ByVal fieldName As String) As Boolean
    Dim parsedValue As Boolean

    valueText = Trim$(CStr(valueText))
    If Not ex_XmlCore.m_TryParseBoolean(valueText, parsedValue) Then
        Err.Raise vbObjectError + 1818, "ex_Helpers", "Invalid boolean for '" & fieldName & "': '" & valueText & "'."
    End If

    mp_ParseRequiredBoolean = parsedValue
End Function

Private Function mp_TryResolveColumnIndexInRow( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByRef outColumnIndex As Long _
) As Boolean
    Dim numericIndex As Long
    Dim columns As Collection
    Dim i As Long
    Dim colObj As obj_ResultColumn

    If rowRef Is Nothing Then Exit Function
    columnRef = Trim$(CStr(columnRef))
    If Len(columnRef) = 0 Then Exit Function

    If ex_XmlCore.m_TryParseLong(columnRef, numericIndex) Then
        If numericIndex < 1 Then Exit Function
        Set columns = rowRef.Columns
        If numericIndex > columns.Count Then Exit Function
        outColumnIndex = numericIndex
        mp_TryResolveColumnIndexInRow = True
        Exit Function
    End If

    Set columns = rowRef.Columns
    For i = 1 To columns.Count
        Set colObj = columns(i)
        If StrComp(colObj.Alias, columnRef, vbTextCompare) = 0 Then
            outColumnIndex = i
            mp_TryResolveColumnIndexInRow = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetRowCellLiveText( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String _
) As String
    Dim targetCol As Long
    Dim ws As Worksheet
    Dim rowIndex As Long

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1826, "ex_Helpers", "Row reference is required for live cell text parsing."
    End If

    columnRef = Trim$(CStr(columnRef))
    If Len(columnRef) = 0 Then
        Err.Raise vbObjectError + 1827, "ex_Helpers", "Column reference is empty for live cell text parsing."
    End If

    If Not mp_TryResolveColumnIndexInRow(rowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1828, "ex_Helpers", "Unknown row cell reference '" & columnRef & "' for live cell text parsing."
    End If

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1829, "ex_Helpers", "Active sheet is not available for live cell text parsing."
    End If

    rowIndex = mp_ResolveAnchoredRowIndex(ws, rowRef, "live cell text parsing")
    mp_GetRowCellLiveText = CStr(ws.Cells(rowIndex, targetCol).Value)
End Function

Private Function mp_GetRowCellRange( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long _
) As Range
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1819, "ex_Helpers", "Active sheet is not available for row cell resolve."
    End If
    If rowIndex < 1 Then
        Err.Raise vbObjectError + 1820, "ex_Helpers", "Row index must be >= 1 for row cell resolve."
    End If
    If columnIndex < 1 Then
        Err.Raise vbObjectError + 1821, "ex_Helpers", "Column index must be >= 1 for row cell resolve."
    End If

    Set mp_GetRowCellRange = ws.Cells(rowIndex, columnIndex)
End Function

Private Function mp_ResolveAnchoredRowIndex( _
    ByVal ws As Worksheet, _
    ByVal rowRef As obj_ResultRow, _
    ByVal operationName As String _
) As Long
    Dim anchorName As String
    Dim resolvedRowIndex As Long

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1822, "ex_Helpers", "Active sheet is not available for " & operationName & "."
    End If
    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1823, "ex_Helpers", "Row reference is required for " & operationName & "."
    End If

    anchorName = Trim$(rowRef.RowAnchorName)
    If Len(anchorName) = 0 Then
        Err.Raise vbObjectError + 1824, "ex_Helpers", "Row anchor is not defined for " & operationName & "."
    End If
    If Not mp_TryGetNamedRowAnchor(ws, anchorName, resolvedRowIndex) Then
        Err.Raise vbObjectError + 1825, "ex_Helpers", "Row anchor '" & anchorName & "' is not found for " & operationName & "."
    End If

    mp_ResolveAnchoredRowIndex = resolvedRowIndex
End Function

Private Function mp_TryGetNamedRowAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String, _
    ByRef outRowIndex As Long _
) As Boolean
    Dim namedEntry As Name
    Dim anchorCell As Range

    If ws Is Nothing Then Exit Function
    anchorName = Trim$(CStr(anchorName))
    If Len(anchorName) = 0 Then Exit Function

    On Error Resume Next
    Set namedEntry = ws.Names(anchorName)
    On Error GoTo 0
    If namedEntry Is Nothing Then Exit Function

    On Error Resume Next
    Set anchorCell = namedEntry.RefersToRange
    On Error GoTo 0
    If anchorCell Is Nothing Then Exit Function

    outRowIndex = anchorCell.Row
    If outRowIndex < 1 Or outRowIndex > ws.Rows.Count Then Exit Function
    mp_TryGetNamedRowAnchor = True
End Function
