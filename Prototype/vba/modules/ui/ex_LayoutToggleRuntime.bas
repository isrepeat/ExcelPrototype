Attribute VB_Name = "ex_LayoutToggleRuntime"
Option Explicit

Private Const META_BLOCK_BEGIN As String = "[[EX_LAYOUT_TOGGLE_META]]"
Private Const META_BLOCK_END As String = "[[/EX_LAYOUT_TOGGLE_META]]"

Private Const TAG_KIND As String = "lt_kind"
Private Const TAG_SOURCE As String = "lt_source"
Private Const TAG_VARIANT_COUNT As String = "lt_variantCount"
Private Const TAG_ACTIVE_INDEX As String = "lt_activeIndex"
Private Const TAG_ON_TOGGLE_MACRO As String = "lt_onToggleMacro"
Private Const TAG_VARIANT_PREFIX As String = "lt_v"

Public Function m_ConfigureToggleShape( _
    ByVal shp As Shape, _
    ByVal toggleSource As String, _
    ByVal variants As Collection, _
    ByVal selectedIndex As Long, _
    Optional ByVal onToggleMacro As String = vbNullString) As Boolean

    Dim meta As Object
    Dim selectedVariant As Object
    Dim variantCount As Long

    If shp Is Nothing Then Exit Function
    toggleSource = Trim$(toggleSource)
    If Len(toggleSource) = 0 Then Exit Function
    If variants Is Nothing Then Exit Function

    variantCount = variants.Count
    If variantCount <= 0 Then Exit Function

    If selectedIndex < 1 Or selectedIndex > variantCount Then selectedIndex = 1

    Set selectedVariant = variants(selectedIndex)
    If selectedVariant Is Nothing Then Exit Function

    Set meta = mp_ReadShapeMetaMap(shp)
    mp_RemoveToggleKeys meta

    mp_SetMetaValue meta, TAG_KIND, "toggle"
    mp_SetMetaValue meta, TAG_SOURCE, toggleSource
    mp_SetMetaValue meta, TAG_VARIANT_COUNT, CStr(variantCount)
    mp_SetMetaValue meta, TAG_ACTIVE_INDEX, CStr(selectedIndex)
    mp_SetMetaValue meta, TAG_ON_TOGGLE_MACRO, Trim$(onToggleMacro)

    mp_WriteVariantsToMeta meta, variants
    mp_WriteShapeMetaMap shp, meta

    If Not mp_ApplyVariantToShape(shp, selectedVariant) Then Exit Function

    m_ConfigureToggleShape = True
End Function

Public Function m_TryAdvanceToggleByCaller(Optional ByVal wb As Workbook, Optional ByVal callerName As String = vbNullString) As Boolean
    Dim shp As Shape
    Dim ws As Worksheet
    Dim meta As Object
    Dim toggleSource As String
    Dim onToggleMacro As String
    Dim variantCount As Long
    Dim variants As Collection
    Dim nextVariant As Object
    Dim currentValue As String
    Dim currentIndex As Long
    Dim nextIndex As Long
    Dim nextValue As String

    Set shp = mp_ResolveCallerShape(wb, callerName, ws)
    If shp Is Nothing Then Exit Function

    Set meta = mp_ReadShapeMetaMap(shp)
    If StrComp(Trim$(mp_GetMetaValue(meta, TAG_KIND)), "toggle", vbTextCompare) <> 0 Then Exit Function

    toggleSource = Trim$(mp_GetMetaValue(meta, TAG_SOURCE))
    If Len(toggleSource) = 0 Then Exit Function

    variantCount = mp_ParseLongOrDefault(mp_GetMetaValue(meta, TAG_VARIANT_COUNT), 0)
    If variantCount <= 0 Then Exit Function

    Set variants = mp_ReadVariantsFromMeta(meta, variantCount)
    If variants Is Nothing Then Exit Function
    If variants.Count = 0 Then Exit Function

    currentValue = ex_ToggleStateRouter.m_GetToggleValue(toggleSource, CStr(variants(1)("Value")), ws)
    currentIndex = mp_FindVariantIndexByValue(variants, currentValue)
    If currentIndex = 0 Then
        currentIndex = mp_ParseLongOrDefault(mp_GetMetaValue(meta, TAG_ACTIVE_INDEX), 1)
    End If
    If currentIndex < 1 Or currentIndex > variants.Count Then currentIndex = 1

    nextIndex = currentIndex + 1
    If nextIndex > variants.Count Then nextIndex = 1

    Set nextVariant = variants(nextIndex)
    If nextVariant Is Nothing Then Exit Function

    nextValue = CStr(nextVariant("Value"))
    ex_ToggleStateRouter.m_SetToggleValue toggleSource, nextValue, ws

    If Not mp_ApplyVariantToShape(shp, nextVariant) Then Exit Function

    mp_SetMetaValue meta, TAG_ACTIVE_INDEX, CStr(nextIndex)
    mp_WriteShapeMetaMap shp, meta

    onToggleMacro = Trim$(mp_GetMetaValue(meta, TAG_ON_TOGGLE_MACRO))
    If Len(onToggleMacro) > 0 Then
        mp_RunToggleChangedMacro onToggleMacro, nextValue, CStr(nextVariant("Caption")), toggleSource
    End If

    m_TryAdvanceToggleByCaller = True
End Function

Private Sub mp_WriteVariantsToMeta(ByVal meta As Object, ByVal variants As Collection)
    Dim i As Long
    Dim variantMap As Object

    If meta Is Nothing Then Exit Sub
    If variants Is Nothing Then Exit Sub

    For i = 1 To variants.Count
        Set variantMap = variants(i)
        If variantMap Is Nothing Then GoTo ContinueVariant

        mp_SetMetaValue meta, TAG_VARIANT_PREFIX & CStr(i) & "_value", CStr(mp_GetVariantString(variantMap, "Value", vbNullString))
        mp_SetMetaValue meta, TAG_VARIANT_PREFIX & CStr(i) & "_caption", CStr(mp_GetVariantString(variantMap, "Caption", vbNullString))
        mp_SetMetaValue meta, TAG_VARIANT_PREFIX & CStr(i) & "_hasBack", IIf(mp_GetVariantBoolean(variantMap, "HasBackColor", False), "true", "false")
        mp_SetMetaValue meta, TAG_VARIANT_PREFIX & CStr(i) & "_back", CStr(mp_GetVariantLong(variantMap, "BackColor", 0)))
        mp_SetMetaValue meta, TAG_VARIANT_PREFIX & CStr(i) & "_hasText", IIf(mp_GetVariantBoolean(variantMap, "HasTextColor", False), "true", "false")
        mp_SetMetaValue meta, TAG_VARIANT_PREFIX & CStr(i) & "_text", CStr(mp_GetVariantLong(variantMap, "TextColor", 0)))
        mp_SetMetaValue meta, TAG_VARIANT_PREFIX & CStr(i) & "_hasBorder", IIf(mp_GetVariantBoolean(variantMap, "HasBorderColor", False), "true", "false")
        mp_SetMetaValue meta, TAG_VARIANT_PREFIX & CStr(i) & "_border", CStr(mp_GetVariantLong(variantMap, "BorderColor", 0)))
ContinueVariant:
    Next i
End Sub

Private Function mp_ReadVariantsFromMeta(ByVal meta As Object, ByVal variantCount As Long) As Collection
    Dim i As Long
    Dim variantMap As Object

    If meta Is Nothing Then Exit Function
    If variantCount <= 0 Then Exit Function

    Set mp_ReadVariantsFromMeta = New Collection

    For i = 1 To variantCount
        Set variantMap = CreateObject("Scripting.Dictionary")
        variantMap.CompareMode = 1

        variantMap("Value") = mp_GetMetaValue(meta, TAG_VARIANT_PREFIX & CStr(i) & "_value")
        variantMap("Caption") = mp_GetMetaValue(meta, TAG_VARIANT_PREFIX & CStr(i) & "_caption")
        variantMap("HasBackColor") = mp_ParseBooleanOrDefault(mp_GetMetaValue(meta, TAG_VARIANT_PREFIX & CStr(i) & "_hasBack"), False)
        variantMap("BackColor") = mp_ParseLongOrDefault(mp_GetMetaValue(meta, TAG_VARIANT_PREFIX & CStr(i) & "_back"), 0)
        variantMap("HasTextColor") = mp_ParseBooleanOrDefault(mp_GetMetaValue(meta, TAG_VARIANT_PREFIX & CStr(i) & "_hasText"), False)
        variantMap("TextColor") = mp_ParseLongOrDefault(mp_GetMetaValue(meta, TAG_VARIANT_PREFIX & CStr(i) & "_text"), 0)
        variantMap("HasBorderColor") = mp_ParseBooleanOrDefault(mp_GetMetaValue(meta, TAG_VARIANT_PREFIX & CStr(i) & "_hasBorder"), False)
        variantMap("BorderColor") = mp_ParseLongOrDefault(mp_GetMetaValue(meta, TAG_VARIANT_PREFIX & CStr(i) & "_border"), 0)

        If Len(Trim$(CStr(variantMap("Value")))) = 0 Then GoTo ContinueVariant
        mp_ReadVariantsFromMeta.Add variantMap
ContinueVariant:
    Next i
End Function

Private Function mp_ApplyVariantToShape(ByVal shp As Shape, ByVal variantMap As Object) As Boolean
    Dim captionText As String

    If shp Is Nothing Then Exit Function
    If variantMap Is Nothing Then Exit Function

    captionText = CStr(mp_GetVariantString(variantMap, "Caption", vbNullString))
    If Len(Trim$(captionText)) = 0 Then
        captionText = CStr(mp_GetVariantString(variantMap, "Value", vbNullString))
    End If

    On Error GoTo EH_APPLY
    shp.TextFrame.Characters.Text = captionText

    If mp_GetVariantBoolean(variantMap, "HasBackColor", False) Then
        shp.Fill.ForeColor.RGB = mp_GetVariantLong(variantMap, "BackColor", 0)
    End If
    If mp_GetVariantBoolean(variantMap, "HasTextColor", False) Then
        shp.TextFrame.Characters.Font.Color = mp_GetVariantLong(variantMap, "TextColor", 0)
    End If
    If mp_GetVariantBoolean(variantMap, "HasBorderColor", False) Then
        shp.Line.ForeColor.RGB = mp_GetVariantLong(variantMap, "BorderColor", 0)
    End If

    mp_ApplyVariantToShape = True
    Exit Function
EH_APPLY:
    mp_ApplyVariantToShape = False
End Function

Private Function mp_FindVariantIndexByValue(ByVal variants As Collection, ByVal valueText As String) As Long
    Dim i As Long
    Dim variantMap As Object
    Dim variantValue As String

    If variants Is Nothing Then Exit Function

    valueText = Trim$(valueText)
    For i = 1 To variants.Count
        Set variantMap = variants(i)
        If variantMap Is Nothing Then GoTo ContinueVariant

        variantValue = Trim$(CStr(mp_GetVariantString(variantMap, "Value", vbNullString)))
        If StrComp(variantValue, valueText, vbTextCompare) = 0 Then
            mp_FindVariantIndexByValue = i
            Exit Function
        End If
ContinueVariant:
    Next i
End Function

Private Sub mp_RunToggleChangedMacro( _
    ByVal macroName As String, _
    ByVal valueText As String, _
    ByVal captionText As String, _
    ByVal sourceText As String)

    On Error GoTo TryNoArgs
    mp_RunMacroByNameWithArgs macroName, valueText, captionText, sourceText
    Exit Sub
TryNoArgs:
    On Error GoTo EH
    mp_RunMacroByName macroName
    Exit Sub
EH:
End Sub

Private Sub mp_RunMacroByNameWithArgs( _
    ByVal macroName As String, _
    ByVal arg1 As String, _
    ByVal arg2 As String, _
    ByVal arg3 As String)

    Dim fullyQualified As String

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then Exit Sub

    If InStr(1, macroName, "!", vbTextCompare) > 0 Then
        fullyQualified = macroName
    Else
        fullyQualified = "'" & ThisWorkbook.Name & "'!" & macroName
    End If

    Application.Run fullyQualified, arg1, arg2, arg3
End Sub

Private Sub mp_RunMacroByName(ByVal macroName As String)
    Dim fullyQualified As String

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then Exit Sub

    If InStr(1, macroName, "!", vbTextCompare) > 0 Then
        fullyQualified = macroName
    Else
        fullyQualified = "'" & ThisWorkbook.Name & "'!" & macroName
    End If

    Application.Run fullyQualified
End Sub

Private Function mp_ResolveCallerShape(ByVal wb As Workbook, ByVal callerName As String, ByRef outSheet As Worksheet) As Shape
    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    callerName = Trim$(callerName)
    If Len(callerName) = 0 Then
        On Error Resume Next
        callerName = Trim$(CStr(Application.Caller))
        On Error GoTo 0
    End If
    If Len(callerName) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0

    If Not ws Is Nothing Then
        Set mp_ResolveCallerShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, callerName)
        If Not mp_ResolveCallerShape Is Nothing Then
            Set outSheet = ws
            Exit Function
        End If
    End If

    For Each ws In wb.Worksheets
        Set mp_ResolveCallerShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, callerName)
        If Not mp_ResolveCallerShape Is Nothing Then
            Set outSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Sub mp_RemoveToggleKeys(ByVal meta As Object)
    Dim keyName As Variant
    Dim keysToRemove As Collection

    If meta Is Nothing Then Exit Sub

    Set keysToRemove = New Collection
    For Each keyName In meta.Keys
        If StrComp(CStr(keyName), TAG_KIND, vbTextCompare) = 0 Then
            keysToRemove.Add CStr(keyName)
        ElseIf StrComp(CStr(keyName), TAG_SOURCE, vbTextCompare) = 0 Then
            keysToRemove.Add CStr(keyName)
        ElseIf StrComp(CStr(keyName), TAG_VARIANT_COUNT, vbTextCompare) = 0 Then
            keysToRemove.Add CStr(keyName)
        ElseIf StrComp(CStr(keyName), TAG_ACTIVE_INDEX, vbTextCompare) = 0 Then
            keysToRemove.Add CStr(keyName)
        ElseIf StrComp(CStr(keyName), TAG_ON_TOGGLE_MACRO, vbTextCompare) = 0 Then
            keysToRemove.Add CStr(keyName)
        ElseIf StrComp(Left$(CStr(keyName), Len(TAG_VARIANT_PREFIX)), TAG_VARIANT_PREFIX, vbTextCompare) = 0 Then
            keysToRemove.Add CStr(keyName)
        End If
    Next keyName

    For Each keyName In keysToRemove
        meta.Remove CStr(keyName)
    Next keyName
End Sub

Private Function mp_GetVariantString(ByVal variantMap As Object, ByVal keyName As String, ByVal defaultValue As String) As String
    If variantMap Is Nothing Then
        mp_GetVariantString = defaultValue
        Exit Function
    End If

    If variantMap.Exists(keyName) Then
        mp_GetVariantString = CStr(variantMap(keyName))
    Else
        mp_GetVariantString = defaultValue
    End If
End Function

Private Function mp_GetVariantBoolean(ByVal variantMap As Object, ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
    If variantMap Is Nothing Then
        mp_GetVariantBoolean = defaultValue
        Exit Function
    End If

    If Not variantMap.Exists(keyName) Then
        mp_GetVariantBoolean = defaultValue
        Exit Function
    End If

    mp_GetVariantBoolean = CBool(variantMap(keyName))
End Function

Private Function mp_GetVariantLong(ByVal variantMap As Object, ByVal keyName As String, ByVal defaultValue As Long) As Long
    If variantMap Is Nothing Then
        mp_GetVariantLong = defaultValue
        Exit Function
    End If

    If Not variantMap.Exists(keyName) Then
        mp_GetVariantLong = defaultValue
        Exit Function
    End If

    mp_GetVariantLong = CLng(variantMap(keyName))
End Function

Private Function mp_ReadShapeMetaMap(ByVal shp As Shape) As Object
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
        Set mp_ReadShapeMetaMap = meta
        Exit Function
    End If

    On Error Resume Next
    altText = CStr(shp.AlternativeText)
    On Error GoTo 0

    blockText = mp_GetMetaBlockContent(altText)
    If Len(blockText) = 0 Then
        Set mp_ReadShapeMetaMap = meta
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

    Set mp_ReadShapeMetaMap = meta
End Function

Private Sub mp_WriteShapeMetaMap(ByVal shp As Shape, ByVal meta As Object)
    Dim altText As String
    Dim baseText As String
    Dim blockText As String

    If shp Is Nothing Then Exit Sub

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
    Exit Sub
EH_WRITE:
End Sub

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

Private Function mp_GetMetaValue(ByVal meta As Object, ByVal keyName As String) As String
    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If meta Is Nothing Then Exit Function
    If meta.Exists(keyName) Then mp_GetMetaValue = CStr(meta(keyName))
End Function

Private Function mp_ParseLongOrDefault(ByVal valueText As String, ByVal defaultValue As Long) As Long
    If ex_XmlCore.m_TryParseLong(Trim$(valueText), mp_ParseLongOrDefault) Then Exit Function
    mp_ParseLongOrDefault = defaultValue
End Function

Private Function mp_ParseBooleanOrDefault(ByVal valueText As String, ByVal defaultValue As Boolean) As Boolean
    If ex_XmlCore.m_TryParseBoolean(Trim$(valueText), mp_ParseBooleanOrDefault) Then Exit Function
    mp_ParseBooleanOrDefault = defaultValue
End Function
