Attribute VB_Name = "ex_ManagedDropdownRuntime"
Option Explicit

Private Const META_BLOCK_BEGIN As String = "[[EX_DD_META]]"
Private Const META_BLOCK_END As String = "[[/EX_DD_META]]"

Private Const TAG_ROLE As String = "md_role"
Private Const TAG_ROLE_HEADER As String = "header"
Private Const TAG_ROLE_OPTION As String = "option"
Private Const TAG_SOURCE_CONTROL As String = "md_sourceControl"
Private Const TAG_HEADER_SHAPE As String = "md_headerShape"
Private Const TAG_OPTION_PREFIX As String = "md_optionPrefix"
Private Const TAG_OPTION_COUNT As String = "md_optionCount"
Private Const TAG_SELECTION_CHANGED_MACRO As String = "dd_selectionChangedMacro"
Private Const TAG_KEY As String = "dd_key"
Private Const TAG_SET_CONTEXT As String = "dd_setContext"
Private Const TAG_HEADER_SHOW_SELECTION As String = "dd_headerShowsSelection"
Private Const DEFAULT_OPTION_MACRO As String = "ex_UIActions.m_SelectDropdownOption_OnClick"

Public Sub m_SetHeaderMetadata( _
    ByVal headerShape As Shape, _
    ByVal sourceControlName As String, _
    ByVal optionPrefix As String, _
    ByVal optionCount As Long, _
    ByVal selectionChangedMacro As String, _
    ByVal headerShowsSelection As Boolean)

    Dim meta As Object

    If headerShape Is Nothing Then Exit Sub

    Set meta = mp_ReadShapeMetaMap(headerShape)
    mp_SetMetaValue meta, TAG_ROLE, TAG_ROLE_HEADER
    mp_SetMetaValue meta, TAG_SOURCE_CONTROL, sourceControlName
    mp_SetMetaValue meta, TAG_HEADER_SHAPE, headerShape.Name
    mp_SetMetaValue meta, TAG_OPTION_PREFIX, optionPrefix
    mp_SetMetaValue meta, TAG_OPTION_COUNT, CStr(optionCount)
    mp_SetMetaValue meta, TAG_SELECTION_CHANGED_MACRO, selectionChangedMacro
    mp_SetMetaValue meta, TAG_HEADER_SHOW_SELECTION, IIf(headerShowsSelection, "true", "false")
    mp_WriteShapeMetaMap headerShape, meta
End Sub

Public Sub m_SetOptionMetadata( _
    ByVal optionShape As Shape, _
    ByVal sourceControlName As String, _
    ByVal headerShapeName As String, _
    ByVal optionPrefix As String, _
    ByVal optionCount As Long, _
    ByVal keyText As String, _
    ByVal setContextText As String, _
    ByVal selectionChangedMacro As String, _
    ByVal headerShowsSelection As Boolean)

    Dim meta As Object

    If optionShape Is Nothing Then Exit Sub

    Set meta = mp_ReadShapeMetaMap(optionShape)
    mp_SetMetaValue meta, TAG_ROLE, TAG_ROLE_OPTION
    mp_SetMetaValue meta, TAG_SOURCE_CONTROL, sourceControlName
    mp_SetMetaValue meta, TAG_HEADER_SHAPE, headerShapeName
    mp_SetMetaValue meta, TAG_OPTION_PREFIX, optionPrefix
    mp_SetMetaValue meta, TAG_OPTION_COUNT, CStr(optionCount)
    mp_SetMetaValue meta, TAG_KEY, keyText
    mp_SetMetaValue meta, TAG_SET_CONTEXT, setContextText
    mp_SetMetaValue meta, TAG_SELECTION_CHANGED_MACRO, selectionChangedMacro
    mp_SetMetaValue meta, TAG_HEADER_SHOW_SELECTION, IIf(headerShowsSelection, "true", "false")
    mp_WriteShapeMetaMap optionShape, meta
End Sub

Public Function m_TryToggleByCaller(Optional ByVal wb As Workbook, Optional ByVal callerName As String = vbNullString) As Boolean
    Dim shp As Shape
    Dim ws As Worksheet
    Dim meta As Object
    Dim roleText As String
    Dim optionPrefix As String
    Dim optionCount As Long
    Dim shouldShow As Boolean

    Set shp = mp_ResolveCallerShape(wb, callerName, ws)
    If shp Is Nothing Then Exit Function

    Set meta = mp_ReadShapeMetaMap(shp)
    roleText = LCase$(Trim$(mp_GetMetaValue(meta, TAG_ROLE)))
    If StrComp(roleText, TAG_ROLE_HEADER, vbTextCompare) <> 0 Then Exit Function

    optionPrefix = Trim$(mp_GetMetaValue(meta, TAG_OPTION_PREFIX))
    optionCount = mp_ParseLongOrDefault(mp_GetMetaValue(meta, TAG_OPTION_COUNT), 0)
    If Len(optionPrefix) = 0 Or optionCount <= 0 Then
        m_TryToggleByCaller = True
        Exit Function
    End If

    shouldShow = Not mp_AnyOptionVisible(ws, optionPrefix, optionCount)
    mp_SetAllManagedOptionsVisible ws, False
    mp_SetOptionsVisible ws, optionPrefix, optionCount, shouldShow
    m_TryToggleByCaller = True
End Function

Public Function m_TrySelectByCaller(Optional ByVal wb As Workbook, Optional ByVal callerName As String = vbNullString) As Boolean
    Dim optionShape As Shape
    Dim ws As Worksheet
    Dim meta As Object
    Dim roleText As String
    Dim sourceControlName As String
    Dim keyText As String
    Dim captionText As String
    Dim setContextText As String
    Dim selectionChangedMacro As String
    Dim headerShapeName As String
    Dim optionPrefix As String
    Dim optionCount As Long
    Dim headerShape As Shape
    Dim headerShowsSelection As Boolean

    Set optionShape = mp_ResolveCallerShape(wb, callerName, ws)
    If optionShape Is Nothing Then Exit Function

    Set meta = mp_ReadShapeMetaMap(optionShape)
    roleText = LCase$(Trim$(mp_GetMetaValue(meta, TAG_ROLE)))
    If StrComp(roleText, TAG_ROLE_OPTION, vbTextCompare) <> 0 Then Exit Function

    sourceControlName = Trim$(mp_GetMetaValue(meta, TAG_SOURCE_CONTROL))
    keyText = Trim$(mp_GetMetaValue(meta, TAG_KEY))
    captionText = Trim$(optionShape.TextFrame.Characters.Text)
    If Len(captionText) = 0 Then captionText = keyText
    setContextText = Trim$(mp_GetMetaValue(meta, TAG_SET_CONTEXT))
    selectionChangedMacro = Trim$(mp_GetMetaValue(meta, TAG_SELECTION_CHANGED_MACRO))
    headerShapeName = Trim$(mp_GetMetaValue(meta, TAG_HEADER_SHAPE))
    optionPrefix = Trim$(mp_GetMetaValue(meta, TAG_OPTION_PREFIX))
    optionCount = mp_ParseLongOrDefault(mp_GetMetaValue(meta, TAG_OPTION_COUNT), 0)
    headerShowsSelection = mp_ParseBooleanOrDefault(mp_GetMetaValue(meta, TAG_HEADER_SHOW_SELECTION), True)

    If Len(setContextText) > 0 Then
        ex_UiXmlProvider.m_ApplyDropdownSetContext setContextText
    End If

    If headerShowsSelection And Len(headerShapeName) > 0 Then
        Set headerShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, headerShapeName)
        If Not headerShape Is Nothing Then
            On Error Resume Next
            headerShape.TextFrame.Characters.Text = captionText
            On Error GoTo 0
        End If
    End If

    If Len(optionPrefix) > 0 And optionCount > 0 Then
        mp_SetOptionsVisible ws, optionPrefix, optionCount, False
    End If

    If Len(selectionChangedMacro) > 0 Then
        If Not mp_RunSelectionChangedMacro(selectionChangedMacro, keyText, captionText, sourceControlName) Then Exit Function
    End If

    m_TrySelectByCaller = True
End Function

Public Function m_RebuildDropdownButton( _
    ByVal ws As Worksheet, _
    ByVal headerShape As Shape, _
    ByVal sourceControlName As String, _
    ByVal itemRecords As Variant, _
    ByVal itemStyleName As String, _
    ByVal marginLeft As Double, _
    ByVal firstGap As Double, _
    ByVal itemGap As Double, _
    ByVal itemHeight As Double, _
    ByVal itemMatchWidth As Boolean, _
    ByVal selectionChangedMacro As String, _
    ByVal optionMacro As String, _
    ByVal stylesMap As Object, _
    ByVal selectedItem As String, _
    ByVal headerShowsSelection As Boolean, _
    ByRef outErrorText As String) As Boolean

    Dim optionPrefix As String
    Dim lowerRow As Long
    Dim upperRow As Long
    Dim i As Long
    Dim optionIndex As Long
    Dim optionShape As Shape
    Dim optionName As String
    Dim optionCaption As String
    Dim optionKey As String
    Dim optionSetContext As String
    Dim optionSelectionMacro As String
    Dim rowMacro As String
    Dim currentTop As Double
    Dim optionWidth As Double
    Dim selectedIndex As Long

    If ws Is Nothing Then
        outErrorText = "DropdownButton rebuild requires a worksheet."
        Exit Function
    End If
    If headerShape Is Nothing Then
        outErrorText = "DropdownButton rebuild requires a header shape."
        Exit Function
    End If

    sourceControlName = Trim$(sourceControlName)
    If Len(sourceControlName) = 0 Then sourceControlName = Trim$(headerShape.Name)
    If Len(sourceControlName) = 0 Then
        outErrorText = "DropdownButton rebuild requires non-empty source control name."
        Exit Function
    End If

    optionPrefix = sourceControlName & "__opt_"
    mp_DeleteOptionShapesByPrefix ws, optionPrefix

    If Len(Trim$(optionMacro)) = 0 Then optionMacro = DEFAULT_OPTION_MACRO
    If itemHeight <= 0 Then
        outErrorText = "DropdownButton '" & sourceControlName & "' has invalid itemHeight <= 0."
        Exit Function
    End If
    If firstGap < 0 Then firstGap = 0
    If itemGap < 0 Then itemGap = 0

    If Not mp_HasRecords(itemRecords) Then
        m_SetHeaderMetadata headerShape, sourceControlName, optionPrefix, 0, selectionChangedMacro, headerShowsSelection
        m_RebuildDropdownButton = True
        Exit Function
    End If

    lowerRow = LBound(itemRecords, 1)
    upperRow = UBound(itemRecords, 1)
    currentTop = headerShape.Top + headerShape.Height + firstGap

    If itemMatchWidth Then
        optionWidth = headerShape.Width
    Else
        optionWidth = headerShape.Width
    End If
    If optionWidth <= 0 Then optionWidth = 1

    selectedItem = Trim$(selectedItem)
    If Len(selectedItem) > 0 Then
        selectedIndex = mp_FindRecordIndex(itemRecords, selectedItem)
        If selectedIndex = 0 Then
            outErrorText = "DropdownButton '" & sourceControlName & "' selectedItem '" & selectedItem & "' was not found in resolved items."
            Exit Function
        End If
    End If

    optionIndex = 0
    For i = lowerRow To upperRow
        optionKey = Trim$(CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY)))
        optionCaption = Trim$(CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION)))
        optionSetContext = Trim$(CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_SET_CONTEXT)))
        rowMacro = Trim$(CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_MACRO)))
        optionSelectionMacro = selectionChangedMacro
        If Len(rowMacro) > 0 Then optionSelectionMacro = rowMacro

        If Len(optionKey) = 0 And Len(optionCaption) = 0 Then GoTo ContinueRow
        If Len(optionCaption) = 0 Then optionCaption = optionKey
        If Len(optionKey) = 0 Then optionKey = optionCaption

        optionIndex = optionIndex + 1
        optionName = optionPrefix & CStr(optionIndex)

        On Error GoTo EH_CREATE
        Set optionShape = ws.Shapes.AddShape( _
            msoShapeRectangle, _
            headerShape.Left + marginLeft, _
            currentTop, _
            optionWidth, _
            itemHeight)
        optionShape.Name = optionName
        optionShape.TextFrame.Characters.Text = optionCaption
        optionShape.Visible = msoFalse
        optionShape.Placement = xlFreeFloating
        optionShape.OnAction = mp_BuildWorkbookMacroRef(optionMacro)
        On Error GoTo 0

        If Len(Trim$(itemStyleName)) > 0 Then
            If stylesMap Is Nothing Then Set stylesMap = ex_UiXmlProvider.m_ReadButtonStyles(ThisWorkbook)
            If stylesMap Is Nothing Then
                outErrorText = "DropdownButton '" & sourceControlName & "' requested itemStyle='" & itemStyleName & "', but styles map is unavailable."
                Exit Function
            End If
            If Not ex_UiXmlProvider.m_ApplyButtonStyleByName(optionShape, itemStyleName, stylesMap) Then
                outErrorText = "Failed to apply itemStyle '" & itemStyleName & "' for dropdown option '" & optionName & "'."
                Exit Function
            End If
        End If

        m_SetOptionMetadata optionShape, sourceControlName, headerShape.Name, optionPrefix, upperRow - lowerRow + 1, optionKey, optionSetContext, optionSelectionMacro, headerShowsSelection
        currentTop = currentTop + itemHeight + itemGap
ContinueRow:
    Next i

    m_SetHeaderMetadata headerShape, sourceControlName, optionPrefix, optionIndex, selectionChangedMacro, headerShowsSelection

    If headerShowsSelection And selectedIndex > 0 Then
        On Error Resume Next
        headerShape.TextFrame.Characters.Text = Trim$(CStr(itemRecords(selectedIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION)))
        If Len(Trim$(headerShape.TextFrame.Characters.Text)) = 0 Then
            headerShape.TextFrame.Characters.Text = Trim$(CStr(itemRecords(selectedIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY)))
        End If
        On Error GoTo 0
    End If

    m_RebuildDropdownButton = True
    Exit Function

EH_CREATE:
    outErrorText = "Failed to create dropdown option shape for '" & sourceControlName & "': " & Err.Description
End Function

Private Function mp_ParseBooleanOrDefault(ByVal valueText As String, ByVal defaultValue As Boolean) As Boolean
    Dim parsedValue As Boolean

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then
        mp_ParseBooleanOrDefault = defaultValue
        Exit Function
    End If

    If ex_XmlCore.m_TryParseBoolean(valueText, parsedValue) Then
        mp_ParseBooleanOrDefault = parsedValue
    Else
        mp_ParseBooleanOrDefault = defaultValue
    End If
End Function

Public Sub m_HideAllOptions(Optional ByVal wb As Workbook)
    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    For Each ws In wb.Worksheets
        mp_SetAllManagedOptionsVisible ws, False
    Next ws
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

Private Function mp_AnyOptionVisible(ByVal ws As Worksheet, ByVal optionPrefix As String, ByVal optionCount As Long) As Boolean
    Dim i As Long
    Dim shp As Shape

    If ws Is Nothing Then Exit Function
    If optionCount <= 0 Then Exit Function

    For i = 1 To optionCount
        Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, optionPrefix & CStr(i))
        If Not shp Is Nothing Then
            If shp.Visible = msoTrue Then
                mp_AnyOptionVisible = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub mp_SetOptionsVisible(ByVal ws As Worksheet, ByVal optionPrefix As String, ByVal optionCount As Long, ByVal isVisible As Boolean)
    Dim i As Long
    Dim shp As Shape

    If ws Is Nothing Then Exit Sub
    If optionCount <= 0 Then Exit Sub

    For i = 1 To optionCount
        Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, optionPrefix & CStr(i))
        If Not shp Is Nothing Then
            shp.Visible = IIf(isVisible, msoTrue, msoFalse)
            If isVisible Then shp.ZOrder msoBringToFront
        End If
    Next i
End Sub

Private Sub mp_SetAllManagedOptionsVisible(ByVal ws As Worksheet, ByVal isVisible As Boolean)
    Dim shp As Shape
    Dim meta As Object
    Dim roleText As String

    If ws Is Nothing Then Exit Sub

    For Each shp In ws.Shapes
        Set meta = mp_ReadShapeMetaMap(shp)
        roleText = LCase$(Trim$(mp_GetMetaValue(meta, TAG_ROLE)))
        If StrComp(roleText, TAG_ROLE_OPTION, vbTextCompare) = 0 Then
            shp.Visible = IIf(isVisible, msoTrue, msoFalse)
        End If
    Next shp
End Sub

Private Function mp_RunSelectionChangedMacro( _
    ByVal macroName As String, _
    ByVal keyText As String, _
    ByVal captionText As String, _
    ByVal sourceControlName As String) As Boolean

    On Error GoTo TryNoArgs
    mp_RunMacroByNameWithArgs macroName, keyText, captionText, sourceControlName
    mp_RunSelectionChangedMacro = True
    Exit Function
TryNoArgs:
    On Error GoTo EH
    mp_RunMacroByName macroName
    mp_RunSelectionChangedMacro = True
    Exit Function
EH:
    MsgBox "Failed to run selectionChangedMacro '" & macroName & "': " & Err.Description, vbExclamation
End Function

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

Private Function mp_BuildWorkbookMacroRef(ByVal macroName As String) As String
    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then Exit Function

    If InStr(1, macroName, "!", vbTextCompare) > 0 Then
        mp_BuildWorkbookMacroRef = macroName
    Else
        mp_BuildWorkbookMacroRef = "'" & ThisWorkbook.Name & "'!" & macroName
    End If
End Function

Private Sub mp_DeleteOptionShapesByPrefix(ByVal ws As Worksheet, ByVal optionPrefix As String)
    Dim i As Long
    Dim shp As Shape
    Dim shapeName As String

    If ws Is Nothing Then Exit Sub
    optionPrefix = Trim$(optionPrefix)
    If Len(optionPrefix) = 0 Then Exit Sub

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        shapeName = Trim$(shp.Name)
        If Len(shapeName) >= Len(optionPrefix) Then
            If StrComp(Left$(shapeName, Len(optionPrefix)), optionPrefix, vbTextCompare) = 0 Then
                On Error Resume Next
                shp.Delete
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

Private Function mp_HasRecords(ByVal records As Variant) As Boolean
    On Error GoTo EH
    If Not IsArray(records) Then Exit Function
    mp_HasRecords = (UBound(records, 1) >= LBound(records, 1))
    Exit Function
EH:
    mp_HasRecords = False
End Function

Private Function mp_FindRecordIndex(ByVal records As Variant, ByVal targetText As String) As Long
    Dim i As Long
    Dim lowerRow As Long
    Dim upperRow As Long
    Dim keyText As String
    Dim captionText As String
    Dim targetValue As String

    targetText = Trim$(targetText)
    If Len(targetText) = 0 Then Exit Function
    If Not mp_HasRecords(records) Then Exit Function

    lowerRow = LBound(records, 1)
    upperRow = UBound(records, 1)
    For i = lowerRow To upperRow
        keyText = Trim$(CStr(records(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY)))
        captionText = Trim$(CStr(records(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION)))
        targetValue = Trim$(CStr(records(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_TARGET)))

        If StrComp(keyText, targetText, vbTextCompare) = 0 Then
            mp_FindRecordIndex = i
            Exit Function
        End If
        If StrComp(captionText, targetText, vbTextCompare) = 0 Then
            mp_FindRecordIndex = i
            Exit Function
        End If
        If StrComp(targetValue, targetText, vbTextCompare) = 0 Then
            mp_FindRecordIndex = i
            Exit Function
        End If
    Next i
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
        If Len(CStr(lineText)) = 0 Then GoTo NextLine

        sepPos = InStr(1, CStr(lineText), "=", vbBinaryCompare)
        If sepPos <= 1 Then GoTo NextLine

        keyName = Trim$(Left$(CStr(lineText), sepPos - 1))
        If Len(keyName) = 0 Then GoTo NextLine

        valueText = Mid$(CStr(lineText), sepPos + 1)
        meta(keyName) = valueText
NextLine:
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

    On Error GoTo EH
    If Len(blockText) = 0 Then
        shp.AlternativeText = baseText
    ElseIf Len(baseText) = 0 Then
        shp.AlternativeText = blockText
    Else
        shp.AlternativeText = baseText & vbLf & blockText
    End If
    Exit Sub
EH:
    MsgBox "Failed to write dropdown metadata to shape '" & shp.Name & "': " & Err.Description, vbExclamation
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
