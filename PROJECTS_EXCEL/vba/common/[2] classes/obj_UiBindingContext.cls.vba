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

Public Event ValueChanged(ByVal sourceName As String, ByVal bindingPath As String)

Private m_sources As Object
Private m_isDisposed As Boolean

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    m_isDisposed = False
    Set m_sources = VBA.CreateObject("Scripting.Dictionary")
    m_sources.CompareMode = VBA.vbTextCompare
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    Set m_sources = Nothing
End Sub

Public Function SetValue( _
    ByVal sourceName As String, _
    ByVal keyName As String, _
    ByVal value As Variant _
) As Boolean
    Dim sourceMap As Object

    If Not private_TryGetOrCreateSource(sourceName, sourceMap) Then Exit Function
    keyName = VBA.Trim$(keyName)
    If VBA.Len(keyName) = 0 Then Exit Function
    sourceMap(keyName) = value
    RaiseEvent ValueChanged(VBA.Trim$(sourceName), keyName)
    SetValue = True
End Function

Public Function SetObject( _
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
    RaiseEvent ValueChanged(VBA.Trim$(sourceName), keyName)
    SetObject = True
End Function

Public Function TrySetPathValue( _
    ByVal sourceName As String, _
    ByVal bindingPath As String, _
    ByVal value As Variant _
) As Boolean
    Dim sourceMap As Object
    Dim currentObject As Object
    Dim childObject As Object
    Dim pathParts As Variant
    Dim pathIndex As Long
    Dim memberName As String
    Dim childMap As Object

    If Not private_TryGetOrCreateSource(sourceName, sourceMap) Then Exit Function
    bindingPath = VBA.Trim$(bindingPath)
    If VBA.Len(bindingPath) = 0 Then Exit Function
    pathParts = VBA.Split(bindingPath, ".")
    Set currentObject = sourceMap
    For pathIndex = LBound(pathParts) To UBound(pathParts) - 1
        memberName = VBA.Trim$(VBA.CStr(pathParts(pathIndex)))
        If VBA.Len(memberName) = 0 Then Exit Function
        Set childObject = Nothing
        If ex_Helpers.fn_RTTI_IsDictionary(currentObject) Then
            If Not currentObject.Exists(memberName) Then
                Set childMap = VBA.CreateObject("Scripting.Dictionary")
                childMap.CompareMode = VBA.vbTextCompare
                Set currentObject(memberName) = childMap
                Set childObject = childMap
            Else
                On Error Resume Next
                Set childObject = currentObject(memberName)
                VBA.Err.Clear
                On Error GoTo 0
            End If
            If childObject Is Nothing Then Exit Function
        Else
            On Error Resume Next
            Set childObject = VBA.CallByName(currentObject, memberName, VbGet)
            If VBA.Err.Number <> 0 Then
                VBA.Err.Clear
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
        If childObject Is Nothing Then Exit Function
        Set currentObject = childObject
    Next pathIndex
    memberName = VBA.Trim$(VBA.CStr(pathParts(UBound(pathParts))))
    If VBA.Len(memberName) = 0 Then Exit Function
    If ex_Helpers.fn_RTTI_IsDictionary(currentObject) Then
        currentObject(memberName) = value
    Else
        On Error Resume Next
        VBA.CallByName currentObject, memberName, VbLet, value
        If VBA.Err.Number <> 0 Then
            VBA.Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    RaiseEvent ValueChanged(VBA.Trim$(sourceName), bindingPath)
    TrySetPathValue = True
End Function

Public Function TryGetValue( _
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
    TryGetValue = True
End Function

' //
' // Private
' //
Private Function private_TryGetOrCreateSource(ByVal sourceName As String, ByRef outSourceMap As Object) As Boolean
    sourceName = VBA.Trim$(sourceName)
    If VBA.Len(sourceName) = 0 Then Exit Function
    If m_sources Is Nothing Then If Not Initialize() Then Exit Function
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
    If ex_Helpers.fn_RTTI_IsDictionary(sourceObject) Then
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