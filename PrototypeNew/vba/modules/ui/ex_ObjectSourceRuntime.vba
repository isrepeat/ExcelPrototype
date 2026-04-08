Attribute VB_Name = "ex_ObjectSourceRuntime"
Option Explicit

Private g_ObjectSourceMap As Object
Private Const INTERNAL_RUNTIME_SOURCE_PREFIX As String = "__object_runtime_"

Public Sub m_ResetObjectSources()
    Set g_ObjectSourceMap = Nothing
End Sub

Public Function m_SetObjectSource( _
    ByVal objectSourceKey As String, _
    ByVal sourceObject As Object, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    Dim normalizedKey As String

    normalizedKey = LCase$(Trim$(objectSourceKey))
    If Len(normalizedKey) = 0 Then
        MsgBox "PrototypeNew: objectSource key is empty.", vbExclamation
        Exit Function
    End If
    If sourceObject Is Nothing Then
        MsgBox "PrototypeNew: objectSource object is not specified for key '" & normalizedKey & "'.", vbExclamation
        Exit Function
    End If

    mp_EnsureObjectSourceMap
    Set g_ObjectSourceMap(normalizedKey) = sourceObject

    If notifyChange Then
        If Not mp_IsInternalRuntimeSourceKey(normalizedKey) Then
            ex_SheetRenderer.m_TryRerenderLastRenderedPage "objectSource:" & normalizedKey
        End If
    End If

    m_SetObjectSource = True
End Function

Public Function m_RemoveObjectSource( _
    ByVal objectSourceKey As String, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    Dim normalizedKey As String

    normalizedKey = LCase$(Trim$(objectSourceKey))
    If Len(normalizedKey) = 0 Then
        MsgBox "PrototypeNew: objectSource key is empty.", vbExclamation
        Exit Function
    End If

    mp_EnsureObjectSourceMap
    If g_ObjectSourceMap.Exists(normalizedKey) Then
        g_ObjectSourceMap.Remove normalizedKey
    End If

    If notifyChange Then
        If Not mp_IsInternalRuntimeSourceKey(normalizedKey) Then
            ex_SheetRenderer.m_TryRerenderLastRenderedPage "objectSource:" & normalizedKey
        End If
    End If

    m_RemoveObjectSource = True
End Function

Public Function m_TryResolveObjectSource( _
    ByVal rawSource As String, _
    ByRef outObject As Object, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim resolvedValue As Variant
    Dim sourceText As String
    Dim sourceKey As String

    rawSource = Trim$(rawSource)
    If Len(rawSource) = 0 Then
        If allowMissing Then
            m_TryResolveObjectSource = True
            Exit Function
        End If
        MsgBox "PrototypeNew: itemControl objectSource is required.", vbExclamation
        Exit Function
    End If

    mp_EnsureObjectSourceMap
    If Not ex_BindingRuntime.m_TryResolveValueBinding(rawSource, g_ObjectSourceMap, resolvedValue) Then Exit Function

    If IsObject(resolvedValue) Then
        Set outObject = resolvedValue
        m_TryResolveObjectSource = True
        Exit Function
    End If

    sourceText = Trim$(CStr(resolvedValue))
    If Len(sourceText) = 0 Then
        If allowMissing Then
            m_TryResolveObjectSource = True
            Exit Function
        End If
        MsgBox "PrototypeNew: objectSource resolved to empty key.", vbExclamation
        Exit Function
    End If

    sourceKey = LCase$(sourceText)

    If g_ObjectSourceMap.Exists(sourceKey) Then
        Set outObject = g_ObjectSourceMap(sourceKey)
        m_TryResolveObjectSource = True
        Exit Function
    End If

    If allowMissing Then
        m_TryResolveObjectSource = True
        Exit Function
    End If

    MsgBox "PrototypeNew: objectSource '" & sourceText & "' is not registered.", vbExclamation
End Function

Private Sub mp_EnsureObjectSourceMap()
    If Not g_ObjectSourceMap Is Nothing Then Exit Sub

    Set g_ObjectSourceMap = CreateObject("Scripting.Dictionary")
    g_ObjectSourceMap.CompareMode = 1
End Sub

Private Function mp_IsInternalRuntimeSourceKey(ByVal normalizedKey As String) As Boolean
    normalizedKey = LCase$(Trim$(normalizedKey))
    If Len(normalizedKey) = 0 Then Exit Function

    mp_IsInternalRuntimeSourceKey = (Left$(normalizedKey, Len(INTERNAL_RUNTIME_SOURCE_PREFIX)) = INTERNAL_RUNTIME_SOURCE_PREFIX)
End Function
