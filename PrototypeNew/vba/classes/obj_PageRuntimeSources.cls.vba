VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageRuntimeSources"
Option Explicit

Private m_ItemsSourceMap As Object
Private m_ObjectSourceMap As Object

' //
' // API
' //
' Callstack[1]: obj_PageMain.private_PrepareDemoConfigRuntime -> m_Base.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
' Callstack[2]: obj_PageMain.private_RegisterDemoTableItems -> m_Base.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
' Callstack[3]: obj_PageMain.private_RegisterDemoSingleTableItems -> m_Base.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
' Callstack[4]: obj_PageMain.private_RegisterDemoTablePartStylesItems -> m_Base.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
' Callstack[5]: ex_Test.private_ResetItemsSources -> pageBase.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
Public Sub ResetItemsSources()
    Set m_ItemsSourceMap = Nothing
End Sub

' Callstack[1]: obj_PageMain.private_PrepareDemoConfigRuntime -> m_Base.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
' Callstack[2]: obj_PageMain.private_RegisterDemoTableItems -> m_Base.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
' Callstack[3]: obj_PageMain.private_RegisterDemoSingleTableItems -> m_Base.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
' Callstack[4]: obj_PageMain.private_RegisterDemoTablePartStylesItems -> m_Base.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
' Callstack[5]: ex_Test.private_ResetObjectSources -> pageBase.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
Public Sub ResetObjectSources()
    Set m_ObjectSourceMap = Nothing
End Sub

' Callstack[1]: obj_PageMain.private_RegisterDemoTableItems -> m_Base.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[2]: obj_PageMain.private_RegisterDemoSingleTableItems -> m_Base.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[3]: obj_PageMain.private_RegisterDemoTablePartStylesItems -> m_Base.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[4]: ex_Test.private_TrySetItemsSource -> pageBase.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[5]: ex_LayoutListRenderer.private_RegisterRuntimeListItemsSourceKey -> renderCtx.PageBase.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[6]: ex_LayoutItemControlRenderer.private_RegisterRuntimeListItemsSourceKey -> renderCtx.PageBase.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
Public Function SetItemsSource(ByVal itemsSourceKey As String, ByVal items As Collection) As Boolean
    Dim normalizedKey As String

    normalizedKey = VBA.LCase$(VBA.Trim$(itemsSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
        VBA.MsgBox "PrototypeNew: itemsSource key is empty.", VBA.vbExclamation
        Exit Function
    End If
    If items Is Nothing Then
        VBA.MsgBox "PrototypeNew: itemsSource collection is not specified for key '" & normalizedKey & "'.", VBA.vbExclamation
        Exit Function
    End If

    private_EnsureItemsSourceMap
    Set m_ItemsSourceMap(normalizedKey) = items

    SetItemsSource = True
End Function

' Callstack[1]: ex_Test.private_TrySetObjectSource -> pageBase.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
' Callstack[2]: ex_LayoutListRenderer.private_RegisterRuntimeObjectSourceKey -> renderCtx.PageBase.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
' Callstack[3]: ex_LayoutItemControlRenderer.private_RegisterRuntimeObjectSourceKey -> renderCtx.PageBase.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
' Callstack[4]: obj_SelectControlVM.private_TryPublishRuntimeSelectedItem -> m_Page.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
Public Function SetObjectSource(ByVal objectSourceKey As String, ByVal sourceObject As Object) As Boolean
    Dim normalizedKey As String

    normalizedKey = VBA.LCase$(VBA.Trim$(objectSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
        VBA.MsgBox "PrototypeNew: objectSource key is empty.", VBA.vbExclamation
        Exit Function
    End If
    If sourceObject Is Nothing Then
        VBA.MsgBox "PrototypeNew: objectSource object is not specified for key '" & normalizedKey & "'.", VBA.vbExclamation
        Exit Function
    End If

    private_EnsureObjectSourceMap
    Set m_ObjectSourceMap(normalizedKey) = sourceObject

    SetObjectSource = True
End Function

' Callstack[1]: ex_Test.private_TryRemoveObjectSource -> pageBase.RuntimeSources.RemoveObjectSource -> obj_PageRuntimeSources.RemoveObjectSource
Public Function RemoveObjectSource(ByVal objectSourceKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = VBA.LCase$(VBA.Trim$(objectSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
        VBA.MsgBox "PrototypeNew: objectSource key is empty.", VBA.vbExclamation
        Exit Function
    End If

    private_EnsureObjectSourceMap
    If m_ObjectSourceMap.Exists(normalizedKey) Then
        m_ObjectSourceMap.Remove normalizedKey
    End If

    RemoveObjectSource = True
End Function

' Callstack[1]: obj_TableSingleControlVM.obj_IControl_Configure -> currentPage.RuntimeSources.TryResolveItemsSource -> obj_PageRuntimeSources.TryResolveItemsSource
' Callstack[2]: obj_TableListControlVM.obj_IControl_Configure -> currentPage.RuntimeSources.TryResolveItemsSource -> obj_PageRuntimeSources.TryResolveItemsSource
' Callstack[3]: obj_ConfigControlVM.obj_IControl_Configure -> currentPage.RuntimeSources.TryResolveItemsSource -> obj_PageRuntimeSources.TryResolveItemsSource
' Callstack[4]: obj_SelectControlVM.obj_IControl_Configure -> currentPage.RuntimeSources.TryResolveItemsSource -> obj_PageRuntimeSources.TryResolveItemsSource
' Callstack[5]: ex_LayoutListRenderer.m_Render -> pageBase.RuntimeSources.TryResolveItemsSource -> obj_PageRuntimeSources.TryResolveItemsSource
' Callstack[6]: ex_LayoutListRenderer.m_TryMeasureContentSpan -> pageBase.RuntimeSources.TryResolveItemsSource -> obj_PageRuntimeSources.TryResolveItemsSource
Public Function TryResolveItemsSource(ByVal rawSource As String, ByRef outItems As Collection) As Boolean
    Dim resolvedValue As Variant
    Dim sourceText As String
    Dim sourceKey As String

    rawSource = VBA.Trim$(rawSource)
    If VBA.Len(rawSource) = 0 Then
        VBA.MsgBox "PrototypeNew: list itemsSource is required.", VBA.vbExclamation
        Exit Function
    End If

    private_EnsureItemsSourceMap
    If Not ex_BindingRuntime.m_TryResolveValueBinding(rawSource, m_ItemsSourceMap, resolvedValue) Then Exit Function

    If VBA.IsObject(resolvedValue) Then
        If VBA.TypeName(resolvedValue) <> "Collection" Then
            VBA.MsgBox "PrototypeNew: list itemsSource must resolve to Collection.", VBA.vbExclamation
            Exit Function
        End If

        Set outItems = resolvedValue
        TryResolveItemsSource = True
        Exit Function
    End If

    sourceText = VBA.Trim$(VBA.CStr(resolvedValue))
    If VBA.Len(sourceText) = 0 Then
        VBA.MsgBox "PrototypeNew: list itemsSource resolved to empty key.", VBA.vbExclamation
        Exit Function
    End If

    sourceKey = VBA.LCase$(sourceText)

    If m_ItemsSourceMap.Exists(sourceKey) Then
        Set outItems = m_ItemsSourceMap(sourceKey)
        TryResolveItemsSource = True
        Exit Function
    End If

    If private_TryParseInlineScalarList(sourceText, outItems) Then
        TryResolveItemsSource = True
        Exit Function
    End If

    VBA.MsgBox "PrototypeNew: list itemsSource '" & sourceText & "' is not registered.", VBA.vbExclamation
End Function

' Callstack[1]: ex_LayoutItemControlRenderer.m_Render -> pageBase.RuntimeSources.TryResolveObjectSource -> obj_PageRuntimeSources.TryResolveObjectSource
' Callstack[2]: ex_LayoutItemControlRenderer.m_TryMeasureContentSpan -> pageBase.RuntimeSources.TryResolveObjectSource -> obj_PageRuntimeSources.TryResolveObjectSource
Public Function TryResolveObjectSource( _
    ByVal rawSource As String, _
    ByRef outObject As Object, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim resolvedValue As Variant
    Dim sourceText As String
    Dim sourceKey As String

    rawSource = VBA.Trim$(rawSource)
    If VBA.Len(rawSource) = 0 Then
        If allowMissing Then
            TryResolveObjectSource = True
            Exit Function
        End If
        VBA.MsgBox "PrototypeNew: itemControl objectSource is required.", VBA.vbExclamation
        Exit Function
    End If

    private_EnsureObjectSourceMap
    If Not ex_BindingRuntime.m_TryResolveValueBinding(rawSource, m_ObjectSourceMap, resolvedValue) Then Exit Function

    If VBA.IsObject(resolvedValue) Then
        Set outObject = resolvedValue
        TryResolveObjectSource = True
        Exit Function
    End If

    sourceText = VBA.Trim$(VBA.CStr(resolvedValue))
    If VBA.Len(sourceText) = 0 Then
        If allowMissing Then
            TryResolveObjectSource = True
            Exit Function
        End If
        VBA.MsgBox "PrototypeNew: objectSource resolved to empty key.", VBA.vbExclamation
        Exit Function
    End If

    sourceKey = VBA.LCase$(sourceText)

    If m_ObjectSourceMap.Exists(sourceKey) Then
        Set outObject = m_ObjectSourceMap(sourceKey)
        TryResolveObjectSource = True
        Exit Function
    End If

    If allowMissing Then
        TryResolveObjectSource = True
        Exit Function
    End If

    VBA.MsgBox "PrototypeNew: objectSource '" & sourceText & "' is not registered.", VBA.vbExclamation
End Function

' //
' // Internal
' //
Private Sub private_EnsureItemsSourceMap()
    If Not m_ItemsSourceMap Is Nothing Then Exit Sub

    Set m_ItemsSourceMap = VBA.CreateObject("Scripting.Dictionary")
    m_ItemsSourceMap.CompareMode = 1
End Sub

Private Sub private_EnsureObjectSourceMap()
    If Not m_ObjectSourceMap Is Nothing Then Exit Sub

    Set m_ObjectSourceMap = VBA.CreateObject("Scripting.Dictionary")
    m_ObjectSourceMap.CompareMode = 1
End Sub

Private Function private_TryParseInlineScalarList(ByVal rawText As String, ByRef outItems As Collection) As Boolean
    Dim normalized As String
    Dim separator As String
    Dim chunks As Variant
    Dim i As Long
    Dim itemText As String

    normalized = VBA.Trim$(rawText)
    If VBA.Len(normalized) = 0 Then Exit Function

    If VBA.InStr(1, normalized, "|", VBA.vbBinaryCompare) > 0 Then
        separator = "|"
    ElseIf VBA.InStr(1, normalized, ";", VBA.vbBinaryCompare) > 0 Then
        separator = ";"
    Else
        Exit Function
    End If

    chunks = VBA.Split(normalized, separator)
    Set outItems = New Collection

    For i = LBound(chunks) To UBound(chunks)
        itemText = VBA.Trim$(VBA.CStr(chunks(i)))
        If VBA.Len(itemText) = 0 Then GoTo ContinueChunk
        outItems.Add itemText
ContinueChunk:
    Next i

    If outItems.Count = 0 Then
        Set outItems = Nothing
        Exit Function
    End If

    private_TryParseInlineScalarList = True
End Function
