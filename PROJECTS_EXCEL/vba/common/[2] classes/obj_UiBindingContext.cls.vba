VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiBindingContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sources As Object

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Initialize() As Boolean
    Set m_sources = VBA.CreateObject("Scripting.Dictionary")
    m_sources.CompareMode = VBA.vbTextCompare
    fn_Initialize = True
End Function

Public Function fn_SetValue( _
    ByVal sourceName As String, _
    ByVal keyName As String, _
    ByVal value As Variant _
) As Boolean
    Dim sourceMap As Object

    If Not private_TryGetOrCreateSource(sourceName, sourceMap) Then Exit Function
    keyName = VBA.Trim$(keyName)
    If VBA.Len(keyName) = 0 Then Exit Function
    sourceMap(keyName) = value
    fn_SetValue = True
End Function

Public Function fn_SetObject( _
    ByVal sourceName As String, _
    ByVal keyName As String, _
    ByVal sourceObject As Object _
) As Boolean
    Dim sourceMap As Object

    If sourceObject Is Nothing Then Exit Function
    If Not private_TryGetOrCreateSource(sourceName, sourceMap) Then Exit Function
    keyName = VBA.Trim$(keyName)
    If VBA.Len(keyName) = 0 Then Exit Function
    Set sourceMap(keyName) = sourceObject
    fn_SetObject = True
End Function

Public Function fn_TryGetValue( _
    ByVal sourceName As String, _
    ByVal bindingPath As String, _
    ByRef outValue As Variant, _
    ByRef outObject As Object, _
    ByRef outIsObject As Boolean _
) As Boolean
    Dim currentObject As Object
    Dim pathParts As Variant
    Dim pathIndex As Long

    Set outObject = Nothing
    outIsObject = False
    If m_sources Is Nothing Then Exit Function
    sourceName = VBA.Trim$(sourceName)
    bindingPath = VBA.Trim$(bindingPath)
    If VBA.Len(sourceName) = 0 Or VBA.Len(bindingPath) = 0 Then Exit Function
    If Not m_sources.Exists(sourceName) Then Exit Function
    Set currentObject = m_sources(sourceName)
    pathParts = VBA.Split(bindingPath, ".")
    For pathIndex = LBound(pathParts) To UBound(pathParts)
        If Not private_TryReadMember(currentObject, VBA.CStr(pathParts(pathIndex)), outValue, outObject, outIsObject) Then Exit Function
        If pathIndex < UBound(pathParts) Then
            If Not outIsObject Then Exit Function
            Set currentObject = outObject
        End If
    Next pathIndex
    fn_TryGetValue = True
End Function

Public Sub fn_Dispose()
    Set m_sources = Nothing
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_TryGetOrCreateSource(ByVal sourceName As String, ByRef outSourceMap As Object) As Boolean
    sourceName = VBA.Trim$(sourceName)
    If VBA.Len(sourceName) = 0 Then Exit Function
    If m_sources Is Nothing Then If Not fn_Initialize() Then Exit Function
    If Not m_sources.Exists(sourceName) Then
        Set outSourceMap = VBA.CreateObject("Scripting.Dictionary")
        outSourceMap.CompareMode = VBA.vbTextCompare
        Set m_sources(sourceName) = outSourceMap
    Else
        Set outSourceMap = m_sources(sourceName)
    End If
    private_TryGetOrCreateSource = True
End Function

Private Function private_TryReadMember( _
    ByVal sourceObject As Object, _
    ByVal memberName As String, _
    ByRef outValue As Variant, _
    ByRef outObject As Object, _
    ByRef outIsObject As Boolean _
) As Boolean
    memberName = VBA.Trim$(memberName)
    Set outObject = Nothing
    outIsObject = False
    If sourceObject Is Nothing Or VBA.Len(memberName) = 0 Then Exit Function
    If TypeName(sourceObject) = "Dictionary" Or TypeName(sourceObject) = "Scripting.Dictionary" Then
        If Not sourceObject.Exists(memberName) Then Exit Function
        On Error Resume Next
        Set outObject = sourceObject(memberName)
        outIsObject = (VBA.Err.Number = 0 And Not outObject Is Nothing)
        VBA.Err.Clear
        On Error GoTo 0
        If outIsObject Then
            private_TryReadMember = True
            Exit Function
        End If
        outValue = sourceObject(memberName)
        private_TryReadMember = True
        Exit Function
    End If
    On Error Resume Next
    Set outObject = VBA.CallByName(sourceObject, memberName, VbGet)
    outIsObject = (VBA.Err.Number = 0 And Not outObject Is Nothing)
    VBA.Err.Clear
    On Error GoTo 0
    If outIsObject Then
        private_TryReadMember = True
        Exit Function
    End If
    On Error Resume Next
    outValue = VBA.CallByName(sourceObject, memberName, VbGet)
    private_TryReadMember = (VBA.Err.Number = 0)
    On Error GoTo 0
End Function