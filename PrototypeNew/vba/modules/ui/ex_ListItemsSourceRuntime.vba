Attribute VB_Name = "ex_ListItemsSourceRuntime"
Option Explicit

Private g_ItemsSourceMap As Object
Private Const INTERNAL_RUNTIME_SOURCE_PREFIX As String = "__list_runtime_"

Public Sub m_ResetItemsSources()
    Set g_ItemsSourceMap = Nothing
End Sub

Public Function m_SetItemsSource( _
    ByVal itemsSourceKey As String, _
    ByVal items As Collection, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    Dim normalizedKey As String

    normalizedKey = LCase$(Trim$(itemsSourceKey))
    If Len(normalizedKey) = 0 Then
        MsgBox "PrototypeNew: itemsSource key is empty.", vbExclamation
        Exit Function
    End If
    If items Is Nothing Then
        MsgBox "PrototypeNew: itemsSource collection is not specified for key '" & normalizedKey & "'.", vbExclamation
        Exit Function
    End If

    mp_EnsureItemsSourceMap
    Set g_ItemsSourceMap(normalizedKey) = items

    If notifyChange Then
        If Not mp_IsInternalRuntimeSourceKey(normalizedKey) Then
            ex_SheetRenderer.m_TryRerenderLastRenderedPage "itemsSource:" & normalizedKey
        End If
    End If

    m_SetItemsSource = True
End Function

Public Function m_TryResolveItemsSource(ByVal rawSource As String, ByRef outItems As Collection) As Boolean
    Dim resolvedValue As Variant
    Dim sourceText As String
    Dim sourceKey As String

    rawSource = Trim$(rawSource)
    If Len(rawSource) = 0 Then
        MsgBox "PrototypeNew: list itemsSource is required.", vbExclamation
        Exit Function
    End If

    mp_EnsureItemsSourceMap
    If Not ex_BindingRuntime.m_TryResolveValueBinding(rawSource, g_ItemsSourceMap, resolvedValue) Then Exit Function

    If IsObject(resolvedValue) Then
        If TypeName(resolvedValue) <> "Collection" Then
            MsgBox "PrototypeNew: list itemsSource must resolve to Collection.", vbExclamation
            Exit Function
        End If

        Set outItems = resolvedValue
        m_TryResolveItemsSource = True
        Exit Function
    End If

    sourceText = Trim$(CStr(resolvedValue))
    If Len(sourceText) = 0 Then
        MsgBox "PrototypeNew: list itemsSource resolved to empty key.", vbExclamation
        Exit Function
    End If

    sourceKey = LCase$(sourceText)

    If g_ItemsSourceMap.Exists(sourceKey) Then
        Set outItems = g_ItemsSourceMap(sourceKey)
        m_TryResolveItemsSource = True
        Exit Function
    End If

    If mp_TryParseInlineScalarList(sourceText, outItems) Then
        m_TryResolveItemsSource = True
        Exit Function
    End If

    MsgBox "PrototypeNew: list itemsSource '" & sourceText & "' is not registered.", vbExclamation
End Function

Private Sub mp_EnsureItemsSourceMap()
    If Not g_ItemsSourceMap Is Nothing Then Exit Sub

    Set g_ItemsSourceMap = CreateObject("Scripting.Dictionary")
    g_ItemsSourceMap.CompareMode = 1
End Sub

Private Function mp_IsInternalRuntimeSourceKey(ByVal normalizedKey As String) As Boolean
    normalizedKey = LCase$(Trim$(normalizedKey))
    If Len(normalizedKey) = 0 Then Exit Function

    mp_IsInternalRuntimeSourceKey = (Left$(normalizedKey, Len(INTERNAL_RUNTIME_SOURCE_PREFIX)) = INTERNAL_RUNTIME_SOURCE_PREFIX)
End Function

Private Function mp_TryParseInlineScalarList(ByVal rawText As String, ByRef outItems As Collection) As Boolean
    Dim normalized As String
    Dim separator As String
    Dim chunks As Variant
    Dim i As Long
    Dim itemText As String

    normalized = Trim$(rawText)
    If Len(normalized) = 0 Then Exit Function

    If InStr(1, normalized, "|", vbBinaryCompare) > 0 Then
        separator = "|"
    ElseIf InStr(1, normalized, ";", vbBinaryCompare) > 0 Then
        separator = ";"
    Else
        Exit Function
    End If

    chunks = Split(normalized, separator)
    Set outItems = New Collection

    For i = LBound(chunks) To UBound(chunks)
        itemText = Trim$(CStr(chunks(i)))
        If Len(itemText) = 0 Then GoTo ContinueChunk
        outItems.Add itemText
ContinueChunk:
    Next i

    If outItems.Count = 0 Then
        Set outItems = Nothing
        Exit Function
    End If

    mp_TryParseInlineScalarList = True
End Function
