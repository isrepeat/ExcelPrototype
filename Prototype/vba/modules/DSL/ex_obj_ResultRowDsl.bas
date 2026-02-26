Attribute VB_Name = "ex_obj_ResultRowDsl"
Option Explicit

Public Function m_TryParseRowColumnRef( _
    ByVal refText As String, _
    ByRef outRowVar As String, _
    ByRef outFieldAlias As String _
) As Boolean
    Dim dotPos As Long
    Dim memberName As String

    refText = Trim$(refText)
    dotPos = InStr(1, refText, ".column[", vbTextCompare)
    If dotPos <= 1 Then Exit Function
    If Right$(refText, 1) <> "]" Then Exit Function

    memberName = Mid$(refText, dotPos + 1, Len("column"))
    If Not ex_PostProcessDslContracts.m_IsMemberAllowed(ex_PostProcessDslContracts.TYPE_ROW, memberName) Then Exit Function

    outRowVar = Trim$(Left$(refText, dotPos - 1))
    If Len(outRowVar) = 0 Then Exit Function
    If Not ex_PostProcessParserCore.m_IsIdentifier(outRowVar) Then Exit Function

    outFieldAlias = Mid$(refText, dotPos + Len(".column["), Len(refText) - (dotPos + Len(".column[")))
    outFieldAlias = Trim$(outFieldAlias)
    If Len(outFieldAlias) = 0 Then Exit Function

    m_TryParseRowColumnRef = True
End Function

Public Function m_TryParseRowColumnsRef( _
    ByVal refText As String, _
    ByRef outRowVar As String _
) As Boolean
    Dim dotPos As Long
    Dim memberName As String

    refText = Trim$(refText)
    dotPos = InStr(1, refText, ".columns", vbTextCompare)
    If dotPos <= 1 Then Exit Function
    If dotPos + Len(".columns") - 1 <> Len(refText) Then Exit Function

    memberName = Mid$(refText, dotPos + 1)
    If Not ex_PostProcessDslContracts.m_IsMemberAllowed(ex_PostProcessDslContracts.TYPE_ROW, memberName) Then Exit Function

    outRowVar = Trim$(Left$(refText, dotPos - 1))
    If Len(outRowVar) = 0 Then Exit Function
    If Not ex_PostProcessParserCore.m_IsIdentifier(outRowVar) Then Exit Function

    m_TryParseRowColumnsRef = True
End Function
