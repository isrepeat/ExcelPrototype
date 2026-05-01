Attribute VB_Name = "ex_ShapeMetaRuntime"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const META_BLOCK_BEGIN As String = "[[EX_PN_META]]"
Private Const META_BLOCK_END As String = "[[/EX_PN_META]]"

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_ShapeMetaRuntime.m_Module_Dispose"
#End If
End Sub
' //
' // API
' //
Public Function m_ReadShapeMetaMap(ByVal shp As Shape) As Object
    Dim meta As Object
    Dim altText As String
    Dim blockText As String
    Dim lines() As String
    Dim lineText As Variant
    Dim sepPos As Long
    Dim keyName As String
    Dim valueText As String

    Set meta = CreateObject("Scripting.Dictionary")
    meta.CompareMode = 1

    If shp Is Nothing Then
        Set m_ReadShapeMetaMap = meta
        Exit Function
    End If

    On Error Resume Next
    altText = VBA.CStr(shp.AlternativeText)
    On Error GoTo 0

    blockText = private_GetMetaBlockContent(altText)
    If VBA.Len(blockText) = 0 Then
        Set m_ReadShapeMetaMap = meta
        Exit Function
    End If

    blockText = VBA.Replace$(blockText, VBA.vbCrLf, VBA.vbLf)
    blockText = VBA.Replace$(blockText, VBA.vbCr, VBA.vbLf)
    lines = VBA.Split(blockText, VBA.vbLf)

    For Each lineText In lines
        lineText = VBA.Trim$(VBA.CStr(lineText))
        If VBA.Len(VBA.CStr(lineText)) = 0 Then GoTo ContinueLine

        sepPos = VBA.InStr(1, VBA.CStr(lineText), "=", VBA.vbBinaryCompare)
        If sepPos <= 1 Then GoTo ContinueLine

        keyName = VBA.Trim$(VBA.Left$(VBA.CStr(lineText), sepPos - 1))
        If VBA.Len(keyName) = 0 Then GoTo ContinueLine

        valueText = VBA.Mid$(VBA.CStr(lineText), sepPos + 1)
        meta(keyName) = valueText

ContinueLine:
    Next lineText

    Set m_ReadShapeMetaMap = meta
End Function


Public Function m_GetShapeMetaValue( _
    ByVal shp As Shape, _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = VBA.vbNullString _
) As String
    Dim meta As Object

    keyName = VBA.Trim$(keyName)
    If VBA.Len(keyName) = 0 Then
        m_GetShapeMetaValue = defaultValue
        Exit Function
    End If

    Set meta = m_ReadShapeMetaMap(shp)
    If meta Is Nothing Then
        m_GetShapeMetaValue = defaultValue
        Exit Function
    End If

    If meta.Exists(keyName) Then
        m_GetShapeMetaValue = VBA.CStr(meta(keyName))
    Else
        m_GetShapeMetaValue = defaultValue
    End If
End Function


Public Function m_TrySetShapeMetaValue( _
    ByVal shp As Shape, _
    ByVal keyName As String, _
    ByVal valueText As String _
) As Boolean
    Dim meta As Object

    If shp Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: shape is not specified for metadata write."
#End If
        Exit Function
    End If

    keyName = VBA.Trim$(keyName)
    If VBA.Len(keyName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: metadata key is empty."
#End If
        Exit Function
    End If

    Set meta = m_ReadShapeMetaMap(shp)
    private_SetMetaValue meta, keyName, valueText
    If Not private_TryWriteShapeMetaMap(shp, meta) Then Exit Function

    m_TrySetShapeMetaValue = True
End Function


Public Function m_TrySetShapeMetaValues( _
    ByVal shp As Shape, _
    ByVal values As Object _
) As Boolean
    Dim meta As Object
    Dim keyName As Variant

    If shp Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: shape is not specified for metadata write."
#End If
        Exit Function
    End If
    If values Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: metadata values map is not specified."
#End If
        Exit Function
    End If

    Set meta = m_ReadShapeMetaMap(shp)

    For Each keyName In values.Keys
        private_SetMetaValue meta, VBA.CStr(keyName), VBA.CStr(values(keyName))
    Next keyName

    If Not private_TryWriteShapeMetaMap(shp, meta) Then Exit Function
    m_TrySetShapeMetaValues = True
End Function

' //
' // Internal
' //
Private Sub private_SetMetaValue(ByVal meta As Object, ByVal keyName As String, ByVal valueText As String)
    keyName = VBA.Trim$(keyName)
    If VBA.Len(keyName) = 0 Then Exit Sub
    If meta Is Nothing Then Exit Sub

    If VBA.Len(VBA.CStr(valueText)) = 0 Then
        If meta.Exists(keyName) Then meta.Remove keyName
        Exit Sub
    End If

    meta(keyName) = VBA.CStr(valueText)
End Sub


Private Function private_TryWriteShapeMetaMap(ByVal shp As Shape, ByVal meta As Object) As Boolean
    Dim altText As String
    Dim baseText As String
    Dim blockText As String

    If shp Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: shape is not specified for metadata write."
#End If
        Exit Function
    End If

    On Error Resume Next
    altText = VBA.CStr(shp.AlternativeText)
    On Error GoTo 0

    baseText = VBA.Trim$(private_RemoveMetaBlock(altText))
    blockText = private_BuildMetaBlock(meta)

    On Error GoTo EH_WRITE
    If VBA.Len(blockText) = 0 Then
        shp.AlternativeText = baseText
    ElseIf VBA.Len(baseText) = 0 Then
        shp.AlternativeText = blockText
    Else
        shp.AlternativeText = baseText & VBA.vbLf & blockText
    End If

    private_TryWriteShapeMetaMap = True
    Exit Function

EH_WRITE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "PrototypeNew: failed to write metadata to shape '" & shp.Name & "': " & Err.Description
#End If
End Function


Private Function private_GetMetaBlockContent(ByVal altText As String) As String
    Dim beginPos As Long
    Dim endPos As Long
    Dim contentStart As Long

    beginPos = VBA.InStr(1, altText, META_BLOCK_BEGIN, VBA.vbTextCompare)
    If beginPos = 0 Then Exit Function

    contentStart = beginPos + VBA.Len(META_BLOCK_BEGIN)
    endPos = VBA.InStr(contentStart, altText, META_BLOCK_END, VBA.vbTextCompare)
    If endPos = 0 Then Exit Function

    private_GetMetaBlockContent = VBA.Mid$(altText, contentStart, endPos - contentStart)
End Function


Private Function private_RemoveMetaBlock(ByVal altText As String) As String
    Dim beginPos As Long
    Dim endPos As Long
    Dim beforeText As String
    Dim afterText As String

    beginPos = VBA.InStr(1, altText, META_BLOCK_BEGIN, VBA.vbTextCompare)
    If beginPos = 0 Then
        private_RemoveMetaBlock = altText
        Exit Function
    End If

    endPos = VBA.InStr(beginPos + VBA.Len(META_BLOCK_BEGIN), altText, META_BLOCK_END, VBA.vbTextCompare)
    If endPos = 0 Then
        private_RemoveMetaBlock = VBA.Left$(altText, beginPos - 1)
        Exit Function
    End If

    beforeText = VBA.Left$(altText, beginPos - 1)
    afterText = VBA.Mid$(altText, endPos + VBA.Len(META_BLOCK_END))
    private_RemoveMetaBlock = beforeText & afterText
End Function


Private Function private_BuildMetaBlock(ByVal meta As Object) As String
    Dim keyName As Variant
    Dim resultText As String

    If meta Is Nothing Then Exit Function
    If meta.Count = 0 Then Exit Function

    resultText = META_BLOCK_BEGIN & VBA.vbLf
    For Each keyName In meta.Keys
        resultText = resultText & VBA.CStr(keyName) & "=" & VBA.CStr(meta(keyName)) & VBA.vbLf
    Next keyName
    resultText = resultText & META_BLOCK_END

    private_BuildMetaBlock = resultText
End Function
