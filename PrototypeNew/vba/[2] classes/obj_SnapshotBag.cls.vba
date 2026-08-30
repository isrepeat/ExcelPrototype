VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SnapshotBag"
Option Explicit

Private m_Items As Object

Private Sub Class_Initialize()
    private_EnsureStorage
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    private_EnsureStorage
    Initialize = True
End Function

Public Function Count() As Long
    private_EnsureStorage
    Count = m_Items.Count
End Function

Public Sub Clear()
    private_EnsureStorage
    m_Items.RemoveAll
End Sub

Public Function Exists(ByVal keyText As String) As Boolean
    private_EnsureStorage
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Function
    Exists = m_Items.Exists(keyText)
End Function

Public Sub Remove(ByVal keyText As String)
    private_EnsureStorage
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Sub
    If m_Items.Exists(keyText) Then m_Items.Remove keyText
End Sub

Public Sub PutText(ByVal keyText As String, ByVal valueText As String)
    private_EnsureStorage
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Sub
    m_Items(keyText) = VBA.CStr(valueText)
End Sub

Public Sub PutLong(ByVal keyText As String, ByVal valueNumber As Long)
    private_EnsureStorage
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Sub
    m_Items(keyText) = valueNumber
End Sub

Public Sub PutBoolean(ByVal keyText As String, ByVal valueFlag As Boolean)
    private_EnsureStorage
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Sub
    m_Items(keyText) = VBA.CBool(valueFlag)
End Sub

Public Sub PutObject(ByVal keyText As String, ByVal valueObject As Object)
    private_EnsureStorage
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Sub
    If valueObject Is Nothing Then
        If m_Items.Exists(keyText) Then m_Items.Remove keyText
        Exit Sub
    End If
    ' Used for nested bags such as Locator / Before / After inside one undo snapshot.
    Set m_Items(keyText) = valueObject
End Sub

Public Function TryGetText(ByVal keyText As String, ByRef outValue As String) As Boolean
    Dim rawValue As Variant

    outValue = VBA.vbNullString
    If Not private_TryGetRawValue(keyText, rawValue) Then Exit Function
    If VBA.IsObject(rawValue) Then Exit Function
    outValue = VBA.CStr(rawValue)
    TryGetText = True
End Function

Public Function TryGetLong(ByVal keyText As String, ByRef outValue As Long) As Boolean
    Dim rawValue As Variant

    outValue = 0
    If Not private_TryGetRawValue(keyText, rawValue) Then Exit Function
    If VBA.IsObject(rawValue) Then Exit Function
    On Error Resume Next
    outValue = VBA.CLng(rawValue)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    TryGetLong = True
End Function

Public Function TryGetBoolean(ByVal keyText As String, ByRef outValue As Boolean) As Boolean
    Dim rawValue As Variant

    outValue = False
    If Not private_TryGetRawValue(keyText, rawValue) Then Exit Function
    If VBA.IsObject(rawValue) Then Exit Function
    On Error Resume Next
    outValue = VBA.CBool(rawValue)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    TryGetBoolean = True
End Function

Public Function TryGetObject(ByVal keyText As String, ByRef outObject As Object) As Boolean
    Set outObject = Nothing
    private_EnsureStorage
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Function
    If Not m_Items.Exists(keyText) Then Exit Function

    ' Read object values directly from the dictionary using Set.
    ' This avoids Variant/default-property coercion for nested snapshot objects.
    On Error GoTo EH_TRY_GET_OBJECT
    Set outObject = m_Items(keyText)
    TryGetObject = Not outObject Is Nothing
    Exit Function

EH_TRY_GET_OBJECT:
    Set outObject = Nothing
End Function

' //
' // Internal
' //
Private Sub private_EnsureStorage()
    If m_Items Is Nothing Then
        Set m_Items = VBA.CreateObject("Scripting.Dictionary")
        m_Items.CompareMode = 1
    End If
End Sub

Private Function private_TryGetRawValue(ByVal keyText As String, ByRef outValue As Variant) As Boolean
    Dim itemValue As Variant

    private_EnsureStorage
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Function
    If Not m_Items.Exists(keyText) Then Exit Function

    itemValue = m_Items(keyText)
    If VBA.IsObject(itemValue) Then
        Set outValue = itemValue
    Else
        outValue = itemValue
    End If
    private_TryGetRawValue = True
End Function
