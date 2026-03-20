Attribute VB_Name = "ex_obj_ResultTableDsl"
Option Explicit

Public Function m_TryParseSpecialRowColumnRef( _
    ByVal refText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outRowSelector As String, _
    ByRef outFieldAlias As String _
) As Boolean
    Dim sheetPos As Long
    Dim tableEndPos As Long
    Dim tableRef As String
    Dim tailText As String
    Dim lowerTail As String
    Dim memberName As String
    Dim prefixLen As Long

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function

    outSourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Len(outSourceAlias) = 0 Then Exit Function

    tableEndPos = InStr(sheetPos + Len(".Sheet["), refText, "]", vbBinaryCompare)
    If tableEndPos <= sheetPos + Len(".Sheet[") Then Exit Function
    outTableAlias = Trim$(Mid$(refText, sheetPos + Len(".Sheet["), tableEndPos - (sheetPos + Len(".Sheet["))))
    If Len(outTableAlias) = 0 Then Exit Function

    tableRef = outSourceAlias & ".Sheet[" & outTableAlias & "]"
    If Not ex_ScriptParserCore.m_IsSheetRef(tableRef) Then Exit Function

    tailText = Mid$(refText, tableEndPos + 1)
    lowerTail = LCase$(tailText)

    If Left$(lowerTail, Len(".lastrow.column[")) = ".lastrow.column[" Then
        outRowSelector = "lastrow"
        prefixLen = Len(".lastRow.column[")
    ElseIf Left$(lowerTail, Len(".prevrow.column[")) = ".prevrow.column[" Then
        outRowSelector = "prevrow"
        prefixLen = Len(".prevRow.column[")
    Else
        Exit Function
    End If

    memberName = outRowSelector
    If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_SHEET_REF, memberName) Then Exit Function

    If Right$(tailText, 1) <> "]" Then Exit Function
    outFieldAlias = Trim$(Mid$(tailText, prefixLen + 1, Len(tailText) - prefixLen - 1))
    If Len(outFieldAlias) = 0 Then Exit Function

    memberName = "column"
    If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_ROW, memberName) Then Exit Function

    m_TryParseSpecialRowColumnRef = True
End Function

Public Function m_TryParseTableSpecialRowRef( _
    ByVal refText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outRowSelector As String _
) As Boolean
    Dim sheetPos As Long
    Dim tableEndPos As Long
    Dim tableRef As String
    Dim tailText As String
    Dim memberName As String

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function

    outSourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Len(outSourceAlias) = 0 Then Exit Function

    tableEndPos = InStr(sheetPos + Len(".Sheet["), refText, "]", vbBinaryCompare)
    If tableEndPos <= sheetPos + Len(".Sheet[") Then Exit Function
    outTableAlias = Trim$(Mid$(refText, sheetPos + Len(".Sheet["), tableEndPos - (sheetPos + Len(".Sheet["))))
    If Len(outTableAlias) = 0 Then Exit Function

    tableRef = outSourceAlias & ".Sheet[" & outTableAlias & "]"
    If Not ex_ScriptParserCore.m_IsSheetRef(tableRef) Then Exit Function

    tailText = LCase$(Trim$(Mid$(refText, tableEndPos + 1)))
    If tailText = ".lastrow" Then
        outRowSelector = "lastrow"
    ElseIf tailText = ".prevrow" Then
        outRowSelector = "prevrow"
    Else
        Exit Function
    End If

    memberName = outRowSelector
    If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_SHEET_REF, memberName) Then Exit Function

    m_TryParseTableSpecialRowRef = True
End Function

Public Function m_TryParseTableScalarRef( _
    ByVal refText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outMemberName As String _
) As Boolean
    Dim sheetPos As Long
    Dim tableEndPos As Long
    Dim tableRef As String
    Dim tailText As String

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function

    outSourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Len(outSourceAlias) = 0 Then Exit Function

    tableEndPos = InStr(sheetPos + Len(".Sheet["), refText, "]", vbBinaryCompare)
    If tableEndPos <= sheetPos + Len(".Sheet[") Then Exit Function
    outTableAlias = Trim$(Mid$(refText, sheetPos + Len(".Sheet["), tableEndPos - (sheetPos + Len(".Sheet["))))
    If Len(outTableAlias) = 0 Then Exit Function

    tableRef = outSourceAlias & ".Sheet[" & outTableAlias & "]"
    If Not ex_ScriptParserCore.m_IsSheetRef(tableRef) Then Exit Function

    tailText = Trim$(Mid$(refText, tableEndPos + 1))
    If Left$(tailText, 1) <> "." Then Exit Function

    outMemberName = Trim$(Mid$(tailText, 2))
    If Len(outMemberName) = 0 Then Exit Function
    If Not ex_ScriptParserCore.m_IsIdentifier(outMemberName) Then Exit Function
    If Not mp_IsTableScalarMember(outMemberName) Then Exit Function
    If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_SHEET_REF, outMemberName) Then Exit Function

    m_TryParseTableScalarRef = True
End Function

Public Function m_TryParseTableRowsRef(ByVal refText As String, ByRef outTableRef As String) As Boolean
    Dim sheetPos As Long
    Dim tableEndPos As Long
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim tailText As String
    Dim memberName As String

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function

    sourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Len(sourceAlias) = 0 Then Exit Function

    tableEndPos = InStr(sheetPos + Len(".Sheet["), refText, "]", vbBinaryCompare)
    If tableEndPos <= sheetPos + Len(".Sheet[") Then Exit Function
    tableAlias = Trim$(Mid$(refText, sheetPos + Len(".Sheet["), tableEndPos - (sheetPos + Len(".Sheet["))))
    If Len(tableAlias) = 0 Then Exit Function

    outTableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
    If Not ex_ScriptParserCore.m_IsSheetRef(outTableRef) Then Exit Function

    tailText = LCase$(Trim$(Mid$(refText, tableEndPos + 1)))
    If tailText <> ".rows" Then Exit Function

    memberName = "rows"
    If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_SHEET_REF, memberName) Then Exit Function

    m_TryParseTableRowsRef = True
End Function

Public Function m_TryParseFullRowColumnRef( _
    ByVal refText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outRowIndex As Long, _
    ByRef outFieldAlias As String _
) As Boolean
    Dim sheetPos As Long
    Dim rowPos As Long
    Dim colPos As Long
    Dim rowText As String
    Dim tableRef As String
    Dim memberName As String

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    rowPos = InStr(1, refText, "].row[", vbTextCompare)
    colPos = InStr(1, refText, "].column[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function
    If rowPos <= sheetPos Then Exit Function
    If colPos <= rowPos Then Exit Function
    If Right$(refText, 1) <> "]" Then Exit Function

    outSourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Len(outSourceAlias) = 0 Then Exit Function

    outTableAlias = Trim$(Mid$(refText, sheetPos + Len(".Sheet["), rowPos - (sheetPos + Len(".Sheet["))))
    If Len(outTableAlias) = 0 Then Exit Function
    tableRef = outSourceAlias & ".Sheet[" & outTableAlias & "]"
    If Not ex_ScriptParserCore.m_IsSheetRef(tableRef) Then Exit Function

    memberName = "row"
    If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_SHEET_REF, memberName) Then Exit Function

    rowText = Trim$(Mid$(refText, rowPos + Len("].row["), colPos - (rowPos + Len("].row["))))
    If Len(rowText) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseLong(rowText, outRowIndex) Then Exit Function
    If outRowIndex < 0 Then Exit Function

    memberName = "column"
    If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_ROW, memberName) Then Exit Function

    outFieldAlias = Trim$(Mid$(refText, colPos + Len("].column["), Len(refText) - (colPos + Len("].column["))))
    If Len(outFieldAlias) = 0 Then Exit Function

    m_TryParseFullRowColumnRef = True
End Function

Public Function m_TryParseTableRowRef( _
    ByVal refText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outRowIndex As Long _
) As Boolean
    Dim sheetPos As Long
    Dim rowPos As Long
    Dim rowText As String
    Dim tableRef As String
    Dim memberName As String

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    rowPos = InStr(1, refText, "].row[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function
    If rowPos <= sheetPos Then Exit Function
    If Right$(refText, 1) <> "]" Then Exit Function

    outSourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Len(outSourceAlias) = 0 Then Exit Function

    outTableAlias = Trim$(Mid$(refText, sheetPos + Len(".Sheet["), rowPos - (sheetPos + Len(".Sheet["))))
    If Len(outTableAlias) = 0 Then Exit Function
    tableRef = outSourceAlias & ".Sheet[" & outTableAlias & "]"
    If Not ex_ScriptParserCore.m_IsSheetRef(tableRef) Then Exit Function

    memberName = "row"
    If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_SHEET_REF, memberName) Then Exit Function

    rowText = Trim$(Mid$(refText, rowPos + Len("].row["), Len(refText) - (rowPos + Len("].row["))))
    If Len(rowText) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseLong(rowText, outRowIndex) Then Exit Function
    If outRowIndex < 0 Then Exit Function

    m_TryParseTableRowRef = True
End Function

Public Function m_TryGetRowBySelector( _
    ByVal tableObj As obj_ResultTable, _
    ByVal rowSelector As String, _
    ByRef outRowObj As obj_ResultRow _
) As Boolean
    rowSelector = LCase$(Trim$(rowSelector))
    If tableObj Is Nothing Then Exit Function

    Select Case rowSelector
        Case "lastrow"
            Set outRowObj = tableObj.LastRow
        Case "prevrow"
            Set outRowObj = tableObj.PrevRow
        Case Else
            Exit Function
    End Select

    m_TryGetRowBySelector = Not (outRowObj Is Nothing)
End Function

Private Function mp_IsTableScalarMember(ByVal memberName As String) As Boolean
    memberName = LCase$(Trim$(memberName))
    Select Case memberName
        Case "count", "rowcount"
            mp_IsTableScalarMember = True
    End Select
End Function
