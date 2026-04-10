Attribute VB_Name = "ex_ShapeMetaRuntime"
Option Explicit

Private Const META_BLOCK_BEGIN As String = "[[EX_PN_META]]"
Private Const META_BLOCK_END As String = "[[/EX_PN_META]]"

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
    altText = CStr(shp.AlternativeText)
    On Error GoTo 0

    blockText = mp_GetMetaBlockContent(altText)
    If Len(blockText) = 0 Then
        Set m_ReadShapeMetaMap = meta
        Exit Function
    End If

    blockText = Replace$(blockText, vbCrLf, vbLf)
    blockText = Replace$(blockText, vbCr, vbLf)
    lines = Split(blockText, vbLf)

    For Each lineText In lines
        lineText = Trim$(CStr(lineText))
        If Len(CStr(lineText)) = 0 Then GoTo ContinueLine

        sepPos = InStr(1, CStr(lineText), "=", vbBinaryCompare)
        If sepPos <= 1 Then GoTo ContinueLine

        keyName = Trim$(Left$(CStr(lineText), sepPos - 1))
        If Len(keyName) = 0 Then GoTo ContinueLine

        valueText = Mid$(CStr(lineText), sepPos + 1)
        meta(keyName) = valueText

ContinueLine:
    Next lineText

    Set m_ReadShapeMetaMap = meta
End Function

Public Function m_GetShapeMetaValue( _
    ByVal shp As Shape, _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim meta As Object

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then
        m_GetShapeMetaValue = defaultValue
        Exit Function
    End If

    Set meta = m_ReadShapeMetaMap(shp)
    If meta Is Nothing Then
        m_GetShapeMetaValue = defaultValue
        Exit Function
    End If

    If meta.Exists(keyName) Then
        m_GetShapeMetaValue = CStr(meta(keyName))
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
        MsgBox "PrototypeNew: shape is not specified for metadata write.", vbExclamation
        Exit Function
    End If

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then
        MsgBox "PrototypeNew: metadata key is empty.", vbExclamation
        Exit Function
    End If

    Set meta = m_ReadShapeMetaMap(shp)
    mp_SetMetaValue meta, keyName, valueText
    If Not mp_TryWriteShapeMetaMap(shp, meta) Then Exit Function

    m_TrySetShapeMetaValue = True
End Function

Public Function m_TrySetShapeMetaValues( _
    ByVal shp As Shape, _
    ByVal values As Object _
) As Boolean
    Dim meta As Object
    Dim keyName As Variant

    If shp Is Nothing Then
        MsgBox "PrototypeNew: shape is not specified for metadata write.", vbExclamation
        Exit Function
    End If
    If values Is Nothing Then
        MsgBox "PrototypeNew: metadata values map is not specified.", vbExclamation
        Exit Function
    End If

    Set meta = m_ReadShapeMetaMap(shp)

    For Each keyName In values.Keys
        mp_SetMetaValue meta, CStr(keyName), CStr(values(keyName))
    Next keyName

    If Not mp_TryWriteShapeMetaMap(shp, meta) Then Exit Function
    m_TrySetShapeMetaValues = True
End Function

Private Sub mp_SetMetaValue(ByVal meta As Object, ByVal keyName As String, ByVal valueText As String)
    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Sub
    If meta Is Nothing Then Exit Sub

    If Len(CStr(valueText)) = 0 Then
        If meta.Exists(keyName) Then meta.Remove keyName
        Exit Sub
    End If

    meta(keyName) = CStr(valueText)
End Sub

Private Function mp_TryWriteShapeMetaMap(ByVal shp As Shape, ByVal meta As Object) As Boolean
    Dim altText As String
    Dim baseText As String
    Dim blockText As String

    If shp Is Nothing Then
        MsgBox "PrototypeNew: shape is not specified for metadata write.", vbExclamation
        Exit Function
    End If

    On Error Resume Next
    altText = CStr(shp.AlternativeText)
    On Error GoTo 0

    baseText = Trim$(mp_RemoveMetaBlock(altText))
    blockText = mp_BuildMetaBlock(meta)

    On Error GoTo EH_WRITE
    If Len(blockText) = 0 Then
        shp.AlternativeText = baseText
    ElseIf Len(baseText) = 0 Then
        shp.AlternativeText = blockText
    Else
        shp.AlternativeText = baseText & vbLf & blockText
    End If

    mp_TryWriteShapeMetaMap = True
    Exit Function

EH_WRITE:
    MsgBox "PrototypeNew: failed to write metadata to shape '" & shp.Name & "': " & Err.Description, vbExclamation
End Function

Private Function mp_GetMetaBlockContent(ByVal altText As String) As String
    Dim beginPos As Long
    Dim endPos As Long
    Dim contentStart As Long

    beginPos = InStr(1, altText, META_BLOCK_BEGIN, vbTextCompare)
    If beginPos = 0 Then Exit Function

    contentStart = beginPos + Len(META_BLOCK_BEGIN)
    endPos = InStr(contentStart, altText, META_BLOCK_END, vbTextCompare)
    If endPos = 0 Then Exit Function

    mp_GetMetaBlockContent = Mid$(altText, contentStart, endPos - contentStart)
End Function

Private Function mp_RemoveMetaBlock(ByVal altText As String) As String
    Dim beginPos As Long
    Dim endPos As Long
    Dim beforeText As String
    Dim afterText As String

    beginPos = InStr(1, altText, META_BLOCK_BEGIN, vbTextCompare)
    If beginPos = 0 Then
        mp_RemoveMetaBlock = altText
        Exit Function
    End If

    endPos = InStr(beginPos + Len(META_BLOCK_BEGIN), altText, META_BLOCK_END, vbTextCompare)
    If endPos = 0 Then
        mp_RemoveMetaBlock = Left$(altText, beginPos - 1)
        Exit Function
    End If

    beforeText = Left$(altText, beginPos - 1)
    afterText = Mid$(altText, endPos + Len(META_BLOCK_END))
    mp_RemoveMetaBlock = beforeText & afterText
End Function

Private Function mp_BuildMetaBlock(ByVal meta As Object) As String
    Dim keyName As Variant
    Dim resultText As String

    If meta Is Nothing Then Exit Function
    If meta.Count = 0 Then Exit Function

    resultText = META_BLOCK_BEGIN & vbLf
    For Each keyName In meta.Keys
        resultText = resultText & CStr(keyName) & "=" & CStr(meta(keyName)) & vbLf
    Next keyName
    resultText = resultText & META_BLOCK_END

    mp_BuildMetaBlock = resultText
End Function

