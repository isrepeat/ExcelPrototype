Attribute VB_Name = "ex_PostProcessDslContracts"
Option Explicit

Public Const TYPE_SHEET_REF As String = "sheetref"
Public Const TYPE_ROW As String = "row"

Private g_MembersByType As Object

Public Function m_IsMemberAllowed(ByVal objectType As String, ByVal memberName As String) As Boolean
    Dim membersByType As Object
    Dim memberSet As Object

    objectType = LCase$(Trim$(objectType))
    memberName = LCase$(Trim$(memberName))
    If Len(objectType) = 0 Or Len(memberName) = 0 Then Exit Function

    Set membersByType = mp_GetMembersByType()
    If membersByType Is Nothing Then Exit Function
    If Not membersByType.Exists(objectType) Then Exit Function

    Set memberSet = membersByType(objectType)
    m_IsMemberAllowed = memberSet.Exists(memberName)
End Function

Private Function mp_GetMembersByType() As Object
    If g_MembersByType Is Nothing Then
        Set g_MembersByType = CreateObject("Scripting.Dictionary")
        g_MembersByType.CompareMode = 1
        g_MembersByType.Add TYPE_SHEET_REF, mp_CreateMemberSet("rows,column,row,lastRow,prevRow,count,rowCount")
        g_MembersByType.Add TYPE_ROW, mp_CreateMemberSet("column")
    End If
    Set mp_GetMembersByType = g_MembersByType
End Function

Private Function mp_CreateMemberSet(ByVal csvMembers As String) As Object
    Dim result As Object
    Dim parts As Variant
    Dim i As Long
    Dim memberName As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    parts = Split(csvMembers, ",")
    For i = LBound(parts) To UBound(parts)
        memberName = Trim$(CStr(parts(i)))
        If Len(memberName) > 0 Then result(memberName) = True
    Next i

    Set mp_CreateMemberSet = result
End Function
