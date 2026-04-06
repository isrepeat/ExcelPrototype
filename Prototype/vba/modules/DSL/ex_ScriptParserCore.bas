Attribute VB_Name = "ex_ScriptParserCore"
Option Explicit

Public Function m_IsIdentifier(ByVal valueText As String) As Boolean
    Dim i As Long
    Dim ch As String

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
            Exit Function
        End If
        If i = 1 And (ch >= "0" And ch <= "9") Then Exit Function
    Next i

    m_IsIdentifier = True
End Function

Public Function m_IsSheetRef(ByVal refText As String) As Boolean
    Dim sheetPos As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim sourceAlias As String
    Dim tableAlias As String

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function

    sourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Not m_TryNormalizeSourceAliasToken(sourceAlias, sourceAlias) Then Exit Function

    openPos = sheetPos + Len(".Sheet[")
    closePos = InStr(openPos, refText, "]", vbBinaryCompare)
    If closePos <= openPos Then Exit Function
    If closePos <> Len(refText) Then Exit Function

    tableAlias = Trim$(Mid$(refText, openPos, closePos - openPos))
    If Len(tableAlias) = 0 Then Exit Function

    m_IsSheetRef = True
End Function

Public Function m_TryNormalizeSourceAliasToken( _
    ByVal rawAlias As String, _
    ByRef outAlias As String _
) As Boolean
    Dim normalized As String

    normalized = Trim$(rawAlias)
    If Len(normalized) = 0 Then Exit Function

    If Left$(normalized, 1) = """" Or Right$(normalized, 1) = """" Then
        If Len(normalized) < 2 Then Exit Function
        If Left$(normalized, 1) <> """" Then Exit Function
        If Right$(normalized, 1) <> """" Then Exit Function
        normalized = Mid$(normalized, 2, Len(normalized) - 2)
    ElseIf InStr(1, normalized, """", vbBinaryCompare) > 0 Then
        Exit Function
    End If

    normalized = Trim$(normalized)
    If Len(normalized) = 0 Then Exit Function

    outAlias = normalized
    m_TryNormalizeSourceAliasToken = True
End Function
