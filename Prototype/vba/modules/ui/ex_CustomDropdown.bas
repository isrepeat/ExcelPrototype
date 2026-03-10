Attribute VB_Name = "ex_CustomDropdown"
Option Explicit

Private Const DEV_SHEET_NAME As String = "Dev"
Private Const HEADER_SHAPE_NAME As String = "btnCustomMode"
Private Const OPTION_SHAPE_PREFIX As String = "btnCustomModeOption_"
Private Const STATE_DROPDOWN_EXPANDED_FLAG As String = "Settings.CustomModeDropdownExpanded"
Private Const PROFILE_HEADER_SHAPE_NAME As String = "btnCustomProfile"
Private Const PROFILE_OPTION_SHAPE_PREFIX As String = "btnCustomProfileOption_"
Private Const STATE_PROFILE_DROPDOWN_EXPANDED_FLAG As String = "Settings.CustomProfileDropdownExpanded"

Private Const TAG_KEY As String = "dd_key"
Private Const TAG_SET_CONTEXT As String = "dd_setContext"
Private Const TAG_SOURCE_CONTROL As String = "dd_sourceControl"
Private Const TAG_SELECTION_CHANGED_MACRO As String = "dd_selectionChangedMacro"
Private Const META_BLOCK_BEGIN As String = "[[EX_DD_META]]"
Private Const META_BLOCK_END As String = "[[/EX_DD_META]]"

Public Sub m_InitDevTestDropdown(Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim isExpanded As Boolean
    Dim isProfileExpanded As Boolean

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    If Not mp_EnsureShapesExist(ws, wb) Then Exit Sub
    If Not mp_EnsureProfileShapesExist(ws, wb) Then Exit Sub

    mp_SyncDropdownContext ws
    mp_RebuildModeOptions ws
    mp_RebuildProfileOptions ws

    isExpanded = mp_GetExpandedState()
    mp_SetOptionsVisible ws, OPTION_SHAPE_PREFIX, isExpanded

    isProfileExpanded = mp_GetProfileExpandedState()
    mp_SetOptionsVisible ws, PROFILE_OPTION_SHAPE_PREFIX, isProfileExpanded
End Sub

Public Sub m_ToggleDropdownButton(Optional ByVal wb As Workbook, Optional ByVal callerName As String = vbNullString)
    Dim ws As Worksheet
    Dim headerName As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    headerName = Trim$(callerName)
    If Len(headerName) = 0 Then
        On Error Resume Next
        headerName = Trim$(CStr(Application.Caller))
        On Error GoTo 0
    End If
    If Len(headerName) = 0 Then Exit Sub

    mp_SyncDropdownContext ws

    If StrComp(headerName, HEADER_SHAPE_NAME, vbTextCompare) = 0 Then
        If Not mp_EnsureShapesExist(ws, wb) Then Exit Sub
        mp_RebuildModeOptions ws
        mp_ToggleModeDropdown ws
        Exit Sub
    End If

    If StrComp(headerName, PROFILE_HEADER_SHAPE_NAME, vbTextCompare) = 0 Then
        If Not mp_EnsureProfileShapesExist(ws, wb) Then Exit Sub
        mp_RebuildProfileOptions ws
        mp_ToggleProfileDropdown ws
    End If
End Sub

Public Sub m_SelectDropdownOption(Optional ByVal wb As Workbook, Optional ByVal callerName As String = vbNullString)
    Dim ws As Worksheet
    Dim optionName As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    optionName = Trim$(callerName)
    If Len(optionName) = 0 Then
        On Error Resume Next
        optionName = Trim$(CStr(Application.Caller))
        On Error GoTo 0
    End If
    If Len(optionName) = 0 Then Exit Sub

    If StrComp(Left$(optionName, Len(OPTION_SHAPE_PREFIX)), OPTION_SHAPE_PREFIX, vbTextCompare) = 0 Then
        mp_HandleDynamicOptionSelect ws, OPTION_SHAPE_PREFIX
        mp_SetExpandedState False
        mp_SetOptionsVisible ws, OPTION_SHAPE_PREFIX, False
        Exit Sub
    End If

    If StrComp(Left$(optionName, Len(PROFILE_OPTION_SHAPE_PREFIX)), PROFILE_OPTION_SHAPE_PREFIX, vbTextCompare) = 0 Then
        mp_HandleDynamicOptionSelect ws, PROFILE_OPTION_SHAPE_PREFIX
        mp_SetProfileExpandedState False
        mp_SetOptionsVisible ws, PROFILE_OPTION_SHAPE_PREFIX, False
    End If
End Sub

Public Sub m_HideDevTestDropdown(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        Set ws = mp_GetDevSheet(ThisWorkbook)
    End If
    If ws Is Nothing Then Exit Sub

    m_HideCustomModeDropdown ws
    m_HideCustomProfileDropdown ws
End Sub

Public Sub m_HideCustomModeDropdown(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        Set ws = mp_GetDevSheet(ThisWorkbook)
    End If
    If ws Is Nothing Then Exit Sub

    mp_SetExpandedState False
    mp_SetOptionsVisible ws, OPTION_SHAPE_PREFIX, False
End Sub

Public Sub m_HideCustomProfileDropdown(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        Set ws = mp_GetDevSheet(ThisWorkbook)
    End If
    If ws Is Nothing Then Exit Sub

    mp_SetProfileExpandedState False
    mp_SetOptionsVisible ws, PROFILE_OPTION_SHAPE_PREFIX, False
End Sub

Public Sub m_OnManagedButtonClick(Optional ByVal callerName As String = vbNullString)
    Dim ws As Worksheet

    Set ws = mp_GetDevSheet(ThisWorkbook)
    If ws Is Nothing Then Exit Sub

    If Len(Trim$(callerName)) = 0 Then
        On Error Resume Next
        callerName = CStr(Application.Caller)
        On Error GoTo 0
    End If
    callerName = Trim$(callerName)

    If StrComp(callerName, HEADER_SHAPE_NAME, vbTextCompare) = 0 Then
        m_HideCustomProfileDropdown ws
        Exit Sub
    End If

    If StrComp(callerName, PROFILE_HEADER_SHAPE_NAME, vbTextCompare) = 0 Then
        m_HideCustomModeDropdown ws
        Exit Sub
    End If

    m_HideDevTestDropdown ws
End Sub

Public Sub m_StabilizeChooseModeAnchorX(Optional ByVal ws As Worksheet, Optional ByVal targetLeft As Double = -1)
    Dim stableStartCol As Long
    Dim bufferCol As Long
    Dim currentLeft As Double
    Dim delta As Double
    Dim bufferRange As Range
    Dim targetBufferWidth As Double
    Dim minBufferWidthUnits As Double
    Dim minBufferWidthPoints As Double

    If ws Is Nothing Then
        Set ws = mp_GetDevSheet(ThisWorkbook)
    End If
    If ws Is Nothing Then Exit Sub

    If Not mp_TryGetStableZoneColumns(ws, stableStartCol, bufferCol) Then Exit Sub

    currentLeft = ws.Cells(1, stableStartCol).Left
    If targetLeft < 0 Then
        targetLeft = currentLeft
    End If

    delta = targetLeft - currentLeft
    If Abs(delta) < 0.1 Then Exit Sub

    Set bufferRange = ws.Columns(bufferCol)
    targetBufferWidth = bufferRange.Width + delta
    minBufferWidthUnits = mp_GetStableZoneMinBufferWidthUnits()
    If minBufferWidthUnits <= 0 Then Exit Sub

    minBufferWidthPoints = mp_GetColumnWidthPointsForUnits(bufferRange, minBufferWidthUnits)
    If minBufferWidthPoints <= 0 Then minBufferWidthPoints = bufferRange.Width
    If targetBufferWidth < minBufferWidthPoints Then targetBufferWidth = minBufferWidthPoints

    If Not mp_SetColumnWidthByPoints(bufferRange, targetBufferWidth) Then
        MsgBox "Failed to stabilize Choose mode anchor: unable to set buffer column width.", vbExclamation
    ElseIf bufferRange.ColumnWidth < minBufferWidthUnits Then
        bufferRange.ColumnWidth = minBufferWidthUnits
    End If
    Exit Sub
End Sub

Public Function m_GetStableZoneStartLeft(Optional ByVal ws As Worksheet) As Double
    Dim stableStartCol As Long
    Dim bufferCol As Long

    If ws Is Nothing Then
        Set ws = mp_GetDevSheet(ThisWorkbook)
    End If
    If ws Is Nothing Then
        m_GetStableZoneStartLeft = -1
        Exit Function
    End If

    If Not mp_TryGetStableZoneColumns(ws, stableStartCol, bufferCol) Then
        m_GetStableZoneStartLeft = -1
        Exit Function
    End If

    m_GetStableZoneStartLeft = ws.Cells(1, stableStartCol).Left
End Function

Private Sub mp_RebuildModeOptions(ByVal ws As Worksheet)
    mp_RebuildDynamicOptions ws, HEADER_SHAPE_NAME, OPTION_SHAPE_PREFIX, HEADER_SHAPE_NAME, "ex_UIActions.m_SelectDropdownOption_OnClick"
End Sub

Private Sub mp_RebuildProfileOptions(ByVal ws As Worksheet)
    mp_RebuildDynamicOptions ws, PROFILE_HEADER_SHAPE_NAME, PROFILE_OPTION_SHAPE_PREFIX, PROFILE_HEADER_SHAPE_NAME, "ex_UIActions.m_SelectDropdownOption_OnClick"
End Sub

Private Sub mp_ToggleModeDropdown(ByVal ws As Worksheet)
    Dim isExpanded As Boolean

    isExpanded = Not mp_GetExpandedState()
    mp_SetExpandedState isExpanded
    mp_SetOptionsVisible ws, OPTION_SHAPE_PREFIX, isExpanded
End Sub

Private Sub mp_ToggleProfileDropdown(ByVal ws As Worksheet)
    Dim isExpanded As Boolean

    isExpanded = Not mp_GetProfileExpandedState()
    mp_SetProfileExpandedState isExpanded
    mp_SetOptionsVisible ws, PROFILE_OPTION_SHAPE_PREFIX, isExpanded
End Sub

Private Sub mp_RebuildDynamicOptions( _
    ByVal ws As Worksheet, _
    ByVal headerShapeName As String, _
    ByVal optionPrefix As String, _
    ByVal sourceControlName As String, _
    ByVal onClickMacro As String)

    Dim headerShape As Shape
    Dim templateShape As Shape
    Dim itemRecords As Variant
    Dim recordCount As Long
    Dim recordLower As Long
    Dim recordUpper As Long
    Dim i As Long
    Dim optionShape As Shape
    Dim marginLeft As Double
    Dim firstGap As Double
    Dim nextGap As Double
    Dim matchWidth As Boolean
    Dim optionHeight As Double
    Dim optionStyleName As String
    Dim currentTop As Double
    Dim sourceModeKey As String
    Dim selectionChangedMacro As String

    Set headerShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, headerShapeName)
    If headerShape Is Nothing Then
        MsgBox "Header shape '" & headerShapeName & "' was not found for dynamic dropdown.", vbExclamation
        Exit Sub
    End If

    mp_DeleteAllOptionShapes ws, optionPrefix, onClickMacro

    Set templateShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, optionPrefix & "1")

    sourceModeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey(ws))
    itemRecords = ex_UiXmlProvider.m_GetDropdownItemRecordsByControl(sourceControlName, ThisWorkbook, sourceModeKey)
    If Not mp_HasRecords(itemRecords) Then
        MsgBox "No items were resolved for dynamic dropdown control '" & sourceControlName & "'.", vbExclamation
        Exit Sub
    End If

    recordLower = LBound(itemRecords, 1)
    recordUpper = UBound(itemRecords, 1)
    recordCount = recordUpper - recordLower + 1

    If Not mp_GetOptionLayout(sourceControlName, marginLeft, firstGap, nextGap, matchWidth, optionHeight, optionStyleName) Then Exit Sub
    selectionChangedMacro = mp_GetOptionalTextAttr(sourceControlName, "selectionChangedMacro", vbNullString)

    currentTop = headerShape.Top + headerShape.Height + firstGap

    For i = recordLower To recordUpper
        Set optionShape = mp_GetOrCreateOptionShape( _
            ws, _
            optionPrefix, _
            i - recordLower + 1, _
            templateShape, _
            headerShape.Left + marginLeft, _
            currentTop, _
            headerShape.Width, _
            optionHeight, _
            optionStyleName)
        If optionShape Is Nothing Then Exit Sub

        mp_SetShapeOnAction optionShape, onClickMacro
        mp_SetOptionShapeMetadata optionShape, sourceControlName, itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY), itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_SET_CONTEXT), selectionChangedMacro

        optionShape.TextFrame.Characters.Text = CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION))

        optionShape.Left = headerShape.Left + marginLeft
        optionShape.Top = currentTop
        If matchWidth Then optionShape.Width = headerShape.Width
        optionShape.ZOrder msoBringToFront

        currentTop = optionShape.Top + optionShape.Height + nextGap
    Next i

    mp_HideExcessOptionShapes ws, optionPrefix, recordCount
End Sub

Private Sub mp_DeleteAllOptionShapes( _
    ByVal ws As Worksheet, _
    ByVal optionPrefix As String, _
    Optional ByVal onClickMacro As String = vbNullString)

    Dim i As Long
    Dim shp As Shape

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If mp_IsManagedOptionShape(shp, optionPrefix, onClickMacro) Then
            On Error Resume Next
            shp.Delete
            If Err.Number <> 0 Then
                Err.Clear
                ' Fallback for protected/grouped shapes: hide instead of failing startup.
                shp.Visible = msoFalse
                shp.OnAction = vbNullString
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function mp_IsManagedOptionShape( _
    ByVal shp As Shape, _
    ByVal optionPrefix As String, _
    ByVal onClickMacro As String) As Boolean

    Dim shapeName As String
    Dim onActionText As String

    If shp Is Nothing Then Exit Function

    shapeName = Trim$(shp.Name)
    If Len(optionPrefix) > 0 Then
        If StrComp(Left$(shapeName, Len(optionPrefix)), optionPrefix, vbTextCompare) = 0 Then
            mp_IsManagedOptionShape = True
            Exit Function
        End If
    End If

    onClickMacro = Trim$(onClickMacro)
    If Len(onClickMacro) = 0 Then Exit Function

    On Error Resume Next
    onActionText = Trim$(CStr(shp.OnAction))
    On Error GoTo 0
    If Len(onActionText) = 0 Then Exit Function

    If mp_OnActionMatchesMacro(onActionText, onClickMacro) Then
        mp_IsManagedOptionShape = True
    End If
End Function

Private Function mp_OnActionMatchesMacro(ByVal onActionText As String, ByVal macroName As String) As Boolean
    Dim suffix As String

    onActionText = Trim$(onActionText)
    macroName = Trim$(macroName)

    If Len(onActionText) = 0 Or Len(macroName) = 0 Then Exit Function

    If StrComp(onActionText, macroName, vbTextCompare) = 0 Then
        mp_OnActionMatchesMacro = True
        Exit Function
    End If

    suffix = "!" & macroName
    If Len(onActionText) >= Len(suffix) Then
        If StrComp(Right$(onActionText, Len(suffix)), suffix, vbTextCompare) = 0 Then
            mp_OnActionMatchesMacro = True
            Exit Function
        End If
    End If

    suffix = "." & macroName
    If Len(onActionText) >= Len(suffix) Then
        If StrComp(Right$(onActionText, Len(suffix)), suffix, vbTextCompare) = 0 Then
            mp_OnActionMatchesMacro = True
        End If
    End If
End Function

Private Function mp_GetOrCreateOptionShape( _
    ByVal ws As Worksheet, _
    ByVal optionPrefix As String, _
    ByVal optionIndex As Long, _
    ByVal templateShape As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal optionStyleName As String) As Shape

    Dim shapeName As String
    Dim duplicateRange As Object
    Dim createdShape As Shape
    Dim duplicateErr As String

    shapeName = optionPrefix & CStr(optionIndex)
    Set createdShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, shapeName)
    If Not createdShape Is Nothing Then
        Set mp_GetOrCreateOptionShape = createdShape
        Exit Function
    End If

    If Not templateShape Is Nothing Then
        On Error Resume Next
        Set duplicateRange = templateShape.Duplicate
        If Err.Number <> 0 Then
            duplicateErr = Err.Description
            Err.Clear
        Else
            Set createdShape = duplicateRange.Item(1)
            If Err.Number <> 0 Then
                duplicateErr = Err.Description
                Err.Clear
                Set createdShape = Nothing
            End If
        End If
        On Error GoTo 0
    End If

    If createdShape Is Nothing Then
        If widthVal <= 0 Then widthVal = 60
        If heightVal <= 0 Then heightVal = 16
        On Error GoTo EH_ADD
        Set createdShape = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, widthVal, heightVal)
        If Not templateShape Is Nothing Then
            On Error Resume Next
            templateShape.PickUp
            createdShape.Apply
            On Error GoTo EH_ADD
        ElseIf Not mp_ApplyOptionStyle(createdShape, optionStyleName) Then
            Exit Function
        End If
    End If

    createdShape.Name = shapeName
    createdShape.Visible = msoFalse
    createdShape.Placement = xlFreeFloating

    Set mp_GetOrCreateOptionShape = createdShape
    Exit Function
EH_ADD:
    If Len(duplicateErr) > 0 Then
        MsgBox "Failed to create dynamic option shape '" & shapeName & "'. Duplicate error: " & duplicateErr & ". Fallback error: " & Err.Description, vbExclamation
    Else
        MsgBox "Failed to create dynamic option shape '" & shapeName & "': " & Err.Description, vbExclamation
    End If
End Function

Private Sub mp_HideExcessOptionShapes(ByVal ws As Worksheet, ByVal optionPrefix As String, ByVal keepCount As Long)
    Dim shp As Shape
    Dim suffix As String
    Dim optionIndex As Long

    For Each shp In ws.Shapes
        If StrComp(Left$(shp.Name, Len(optionPrefix)), optionPrefix, vbTextCompare) = 0 Then
            suffix = Mid$(shp.Name, Len(optionPrefix) + 1)
            If IsNumeric(suffix) Then
                optionIndex = CLng(suffix)
                If optionIndex > keepCount Then shp.Visible = msoFalse
            Else
                shp.Visible = msoFalse
            End If
        End If
    Next shp
End Sub

Private Function mp_GetOptionLayout( _
    ByVal headerControlName As String, _
    ByRef marginLeft As Double, _
    ByRef firstGap As Double, _
    ByRef nextGap As Double, _
    ByRef matchWidth As Boolean, _
    ByRef optionHeight As Double, _
    ByRef optionStyleName As String) As Boolean

    marginLeft = 0
    firstGap = 2
    nextGap = 2
    matchWidth = True
    optionHeight = 16
    optionStyleName = "modeChooserOption"

    If Not mp_TryGetOptionalDoubleAttr(headerControlName, "itemMarginLeft", marginLeft) Then Exit Function
    If Not mp_TryGetOptionalDoubleAttr(headerControlName, "itemFirstGap", firstGap) Then Exit Function
    If Not mp_TryGetOptionalDoubleAttr(headerControlName, "itemGap", nextGap) Then Exit Function
    If Not mp_TryGetOptionalDoubleAttr(headerControlName, "itemHeight", optionHeight) Then Exit Function
    If optionHeight <= 0 Then
        MsgBox "Control '" & headerControlName & "' has invalid attribute 'itemHeight'. Value must be > 0.", vbExclamation
        Exit Function
    End If

    If Not mp_TryGetOptionalBooleanAttr(headerControlName, "itemMatchWidth", matchWidth) Then Exit Function
    optionStyleName = mp_GetOptionalTextAttr(headerControlName, "itemStyle", optionStyleName)

    mp_GetOptionLayout = True
End Function

Private Function mp_GetOptionalTextAttr(ByVal controlName As String, ByVal attrName As String, ByVal defaultValue As String) As String
    Dim valueText As String

    valueText = Trim$(ex_UiXmlProvider.m_GetControlAttribute(controlName, attrName, ThisWorkbook))
    If Len(valueText) = 0 Then
        mp_GetOptionalTextAttr = defaultValue
    Else
        mp_GetOptionalTextAttr = valueText
    End If
End Function

Private Function mp_ApplyOptionStyle(ByVal shp As Shape, ByVal styleName As String) As Boolean
    Dim stylesMap As Object

    If shp Is Nothing Then Exit Function

    styleName = Trim$(styleName)
    If Len(styleName) = 0 Then
        mp_ApplyOptionStyle = True
        Exit Function
    End If

    Set stylesMap = ex_UiXmlProvider.m_ReadButtonStyles(ThisWorkbook)
    If stylesMap Is Nothing Then Exit Function

    mp_ApplyOptionStyle = ex_UiXmlProvider.m_ApplyButtonStyleByName(shp, styleName, stylesMap)
End Function

Private Function mp_TryGetOptionalDoubleAttr(ByVal controlName As String, ByVal attrName As String, ByRef outValue As Double) As Boolean
    Dim valueText As String
    Dim parsedValue As Double

    valueText = Trim$(ex_UiXmlProvider.m_GetControlAttribute(controlName, attrName, ThisWorkbook))
    If Len(valueText) = 0 Then
        mp_TryGetOptionalDoubleAttr = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseDouble(valueText, parsedValue, True) Then
        MsgBox "Invalid numeric value '" & valueText & "' for attribute '" & attrName & "' on control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    outValue = parsedValue
    mp_TryGetOptionalDoubleAttr = True
End Function

Private Function mp_TryGetOptionalBooleanAttr(ByVal controlName As String, ByVal attrName As String, ByRef outValue As Boolean) As Boolean
    Dim valueText As String
    Dim parsedValue As Boolean

    valueText = Trim$(ex_UiXmlProvider.m_GetControlAttribute(controlName, attrName, ThisWorkbook))
    If Len(valueText) = 0 Then
        mp_TryGetOptionalBooleanAttr = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseBoolean(valueText, parsedValue) Then
        MsgBox "Invalid boolean value '" & valueText & "' for attribute '" & attrName & "' on control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    outValue = parsedValue
    mp_TryGetOptionalBooleanAttr = True
End Function

Private Function mp_HasRecords(ByVal records As Variant) As Boolean
    On Error GoTo EH
    If Not IsArray(records) Then Exit Function
    mp_HasRecords = (UBound(records, 1) >= LBound(records, 1))
    Exit Function
EH:
    mp_HasRecords = False
End Function

Private Sub mp_SetOptionShapeMetadata( _
    ByVal shp As Shape, _
    ByVal sourceControlName As String, _
    ByVal keyText As String, _
    ByVal setContextText As String, _
    ByVal selectionChangedMacro As String)
    Dim meta As Object

    Set meta = mp_ReadShapeMetaMap(shp)
    mp_SetMetaValue meta, TAG_KEY, keyText
    mp_SetMetaValue meta, TAG_SET_CONTEXT, setContextText
    mp_SetMetaValue meta, TAG_SOURCE_CONTROL, sourceControlName
    mp_SetMetaValue meta, TAG_SELECTION_CHANGED_MACRO, selectionChangedMacro
    mp_WriteShapeMetaMap shp, meta
End Sub

Private Sub mp_SetShapeTag(ByVal shp As Shape, ByVal tagName As String, ByVal tagValue As String)
    Dim meta As Object

    Set meta = mp_ReadShapeMetaMap(shp)
    mp_SetMetaValue meta, tagName, tagValue
    mp_WriteShapeMetaMap shp, meta
End Sub

Private Function mp_GetShapeTag(ByVal shp As Shape, ByVal tagName As String) As String
    Dim meta As Object

    Set meta = mp_ReadShapeMetaMap(shp)
    tagName = Trim$(tagName)
    If Len(tagName) = 0 Then Exit Function
    If meta.Exists(tagName) Then
        mp_GetShapeTag = CStr(meta(tagName))
    End If
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

Private Sub mp_SetShapeOnAction(ByVal shp As Shape, ByVal macroName As String)
    If shp Is Nothing Then Exit Sub
    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then Exit Sub

    On Error GoTo EH
    shp.OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
    Exit Sub
EH:
    MsgBox "Failed to assign macro '" & macroName & "' to shape '" & shp.Name & "': " & Err.Description, vbExclamation
End Sub

Private Sub mp_HandleDynamicOptionSelect(ByVal ws As Worksheet, ByVal optionPrefix As String)
    Dim callerName As String
    Dim optionShape As Shape
    Dim keyText As String
    Dim captionText As String
    Dim setContextText As String
    Dim selectionChangedMacro As String
    Dim sourceControlName As String

    callerName = vbNullString
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0
    callerName = Trim$(callerName)
    If Len(callerName) = 0 Then Exit Sub

    If StrComp(Left$(callerName, Len(optionPrefix)), optionPrefix, vbTextCompare) <> 0 Then Exit Sub

    Set optionShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, callerName)
    If optionShape Is Nothing Then Exit Sub

    keyText = Trim$(mp_GetShapeTag(optionShape, TAG_KEY))
    captionText = Trim$(optionShape.TextFrame.Characters.Text)
    If Len(captionText) = 0 Then captionText = keyText

    setContextText = Trim$(mp_GetShapeTag(optionShape, TAG_SET_CONTEXT))
    selectionChangedMacro = Trim$(mp_GetShapeTag(optionShape, TAG_SELECTION_CHANGED_MACRO))
    sourceControlName = Trim$(mp_GetShapeTag(optionShape, TAG_SOURCE_CONTROL))
    If Len(sourceControlName) = 0 Then
        MsgBox "Dynamic dropdown option '" & callerName & "' has no source control metadata.", vbExclamation
        Exit Sub
    End If

    If Len(keyText) = 0 Then
        mp_FillOptionDataByCallerIndex ws, sourceControlName, callerName, optionPrefix, keyText, captionText, setContextText
    End If

    If Len(setContextText) > 0 Then
        ex_UiXmlProvider.m_ApplyDropdownSetContext setContextText
    End If

    If Len(selectionChangedMacro) > 0 Then
        If Not mp_RunSelectionChangedMacro(selectionChangedMacro, keyText, captionText, sourceControlName) Then
            Exit Sub
        End If
    End If
End Sub

Private Sub mp_FillOptionDataByCallerIndex( _
    ByVal ws As Worksheet, _
    ByVal sourceControlName As String, _
    ByVal callerName As String, _
    ByVal optionPrefix As String, _
    ByRef keyText As String, _
    ByRef captionText As String, _
    ByRef setContextText As String)

    Dim suffix As String
    Dim itemIndex As Long
    Dim records As Variant
    Dim rowIndex As Long

    suffix = Mid$(callerName, Len(optionPrefix) + 1)
    If Not IsNumeric(suffix) Then Exit Sub

    itemIndex = CLng(suffix)
    If itemIndex <= 0 Then Exit Sub

    sourceControlName = Trim$(sourceControlName)
    If Len(sourceControlName) = 0 Then Exit Sub

    records = ex_UiXmlProvider.m_GetDropdownItemRecordsByControl(sourceControlName, ThisWorkbook, ex_ConfigProfilesManager.m_GetActiveModeKey(ws))
    If mp_HasRecords(records) Then
        rowIndex = LBound(records, 1) + itemIndex - 1
        If rowIndex >= LBound(records, 1) And rowIndex <= UBound(records, 1) Then
            If Len(keyText) = 0 Then keyText = Trim$(CStr(records(rowIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY)))
            If Len(captionText) = 0 Then captionText = Trim$(CStr(records(rowIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION)))
            If Len(setContextText) = 0 Then setContextText = Trim$(CStr(records(rowIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_SET_CONTEXT)))
        End If
    End If
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

    On Error GoTo EH
    Application.Run fullyQualified
    Exit Sub
EH:
    MsgBox "Failed to run macro '" & macroName & "': " & Err.Description, vbExclamation
End Sub

Private Sub mp_SetOptionsVisible(ByVal ws As Worksheet, ByVal optionPrefix As String, ByVal isVisible As Boolean)
    Dim shp As Shape
    Dim suffix As String
    Dim isNumberedOption As Boolean
    Dim visibility As MsoTriState

    visibility = IIf(isVisible, msoTrue, msoFalse)

    For Each shp In ws.Shapes
        If StrComp(Left$(shp.Name, Len(optionPrefix)), optionPrefix, vbTextCompare) = 0 Then
            suffix = Mid$(shp.Name, Len(optionPrefix) + 1)
            isNumberedOption = IsNumeric(suffix)
            On Error Resume Next
            If isVisible Then
                shp.Visible = IIf(isNumberedOption, msoTrue, msoFalse)
            Else
                shp.Visible = msoFalse
            End If
            On Error GoTo 0
        End If
    Next shp
End Sub

Private Sub mp_SyncDropdownContext(ByVal ws As Worksheet)
    Dim modeKey As String
    Dim profileName As String

    modeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey(ws))
    If Len(modeKey) > 0 Then
        ex_UiXmlProvider.m_SetDropdownContextValue "activeMode", modeKey
    End If

    profileName = Trim$(ex_ConfigProfilesManager.m_GetActiveProfileName(ws))
    If Len(profileName) > 0 Then
        ex_UiXmlProvider.m_SetDropdownContextValue "activeProfile", profileName
    End If
End Sub

Private Function mp_TryGetStableZoneColumns(ByVal ws As Worksheet, ByRef stableStartCol As Long, ByRef bufferCol As Long) As Boolean
    Dim stableColText As String

    stableColText = Trim$(ex_UiXmlProvider.m_GetLayoutAttribute("stableZone", "startCol", ThisWorkbook))
    If Len(stableColText) = 0 Then
        MsgBox "UI layout must define /uiDefinition/layout/stableZone@startCol in DevUI.xml.", vbExclamation
        Exit Function
    End If

    If Not mp_TryResolveColumnIndex(ws, stableColText, stableStartCol) Then
        MsgBox "Invalid stable zone startCol in DevUI.xml: '" & stableColText & "'.", vbExclamation
        Exit Function
    End If

    If stableStartCol <= 1 Then
        MsgBox "Layout stable zone startCol must be greater than column A.", vbExclamation
        Exit Function
    End If

    bufferCol = stableStartCol - 1
    mp_TryGetStableZoneColumns = True
End Function

Private Function mp_GetStableZoneMinBufferWidthUnits() As Double
    Dim widthText As String

    widthText = Trim$(ex_UiXmlProvider.m_GetLayoutAttribute("stableZone", "minBufferWidth", ThisWorkbook))
    If Len(widthText) = 0 Then
        MsgBox "UI layout must define /uiDefinition/layout/stableZone@minBufferWidth in DevUI.xml.", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(widthText) Then
        MsgBox "Invalid stable zone minBufferWidth in DevUI.xml: '" & widthText & "'.", vbExclamation
        Exit Function
    End If

    mp_GetStableZoneMinBufferWidthUnits = CDbl(widthText)
    If mp_GetStableZoneMinBufferWidthUnits <= 0 Then
        MsgBox "Invalid stable zone minBufferWidth in DevUI.xml: value must be > 0.", vbExclamation
        mp_GetStableZoneMinBufferWidthUnits = 0
    End If
End Function

Private Function mp_TryResolveColumnIndex(ByVal ws As Worksheet, ByVal valueText As String, ByRef outColumnIndex As Long) As Boolean
    Dim parsed As Long
    Dim refRange As Range

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    If IsNumeric(valueText) Then
        parsed = CLng(valueText)
        If parsed > 0 Then
            outColumnIndex = parsed
            mp_TryResolveColumnIndex = True
            Exit Function
        End If
    End If

    On Error Resume Next
    Set refRange = ws.Range(valueText)
    If refRange Is Nothing Then Set refRange = ws.Range(valueText & "1")
    If refRange Is Nothing Then Set refRange = ws.Columns(valueText & ":" & valueText)
    On Error GoTo 0

    If refRange Is Nothing Then Exit Function

    outColumnIndex = refRange.Column
    mp_TryResolveColumnIndex = (outColumnIndex > 0)
End Function

Private Function mp_SetColumnWidthByPoints(ByVal colRange As Range, ByVal targetPoints As Double) As Boolean
    Dim i As Long
    Dim currentPoints As Double
    Dim currentWidthUnits As Double
    Dim slope As Double
    Dim deltaPoints As Double

    If colRange Is Nothing Then Exit Function
    If targetPoints <= 0 Then Exit Function

    On Error GoTo EH
    currentPoints = colRange.Width
    currentWidthUnits = colRange.ColumnWidth
    If currentPoints <= 0 Or currentWidthUnits <= 0 Then Exit Function

    colRange.ColumnWidth = currentWidthUnits * (targetPoints / currentPoints)

    For i = 1 To 8
        currentPoints = colRange.Width
        deltaPoints = targetPoints - currentPoints
        If Abs(deltaPoints) < 0.1 Then
            mp_SetColumnWidthByPoints = True
            Exit Function
        End If

        currentWidthUnits = colRange.ColumnWidth
        If currentWidthUnits <= 0 Then Exit For
        slope = currentPoints / currentWidthUnits
        If slope <= 0 Then Exit For

        colRange.ColumnWidth = currentWidthUnits + (deltaPoints / slope)
        If colRange.ColumnWidth < 0.1 Then colRange.ColumnWidth = 0.1
    Next i

    mp_SetColumnWidthByPoints = (Abs(targetPoints - colRange.Width) < 0.5)
    Exit Function
EH:
    mp_SetColumnWidthByPoints = False
End Function

Private Function mp_GetColumnWidthPointsForUnits(ByVal colRange As Range, ByVal widthUnits As Double) As Double
    Dim prevUnits As Double

    If colRange Is Nothing Then Exit Function
    If widthUnits <= 0 Then Exit Function

    On Error GoTo EH
    prevUnits = colRange.ColumnWidth
    colRange.ColumnWidth = widthUnits
    mp_GetColumnWidthPointsForUnits = colRange.Width
    colRange.ColumnWidth = prevUnits
    Exit Function
EH:
    On Error Resume Next
    If prevUnits > 0 Then colRange.ColumnWidth = prevUnits
    On Error GoTo 0
    mp_GetColumnWidthPointsForUnits = 0
End Function

Private Function mp_GetExpandedState() As Boolean
    mp_GetExpandedState = ex_Settings.m_GetBoolFlag(STATE_DROPDOWN_EXPANDED_FLAG, False)
End Function

Private Sub mp_SetExpandedState(ByVal isExpanded As Boolean)
    ex_Settings.m_SetBoolFlag STATE_DROPDOWN_EXPANDED_FLAG, isExpanded
End Sub

Private Function mp_GetProfileExpandedState() As Boolean
    mp_GetProfileExpandedState = ex_Settings.m_GetBoolFlag(STATE_PROFILE_DROPDOWN_EXPANDED_FLAG, False)
End Function

Private Sub mp_SetProfileExpandedState(ByVal isExpanded As Boolean)
    ex_Settings.m_SetBoolFlag STATE_PROFILE_DROPDOWN_EXPANDED_FLAG, isExpanded
End Sub

Private Function mp_GetDevSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set mp_GetDevSheet = wb.Worksheets(DEV_SHEET_NAME)
    On Error GoTo 0
    If mp_GetDevSheet Is Nothing Then
        MsgBox "Sheet '" & DEV_SHEET_NAME & "' was not found for custom dropdown.", vbExclamation
    End If
End Function

Private Function mp_EnsureShapesExist(ByVal ws As Worksheet, ByVal wb As Workbook) As Boolean
    If ex_ConfigProfilesManager.m_GetShapeByName(ws, HEADER_SHAPE_NAME) Is Nothing Then
        ex_UILoader.m_LoadUiFromConfig wb
    End If

    mp_EnsureShapesExist = Not (ex_ConfigProfilesManager.m_GetShapeByName(ws, HEADER_SHAPE_NAME) Is Nothing)
End Function

Private Function mp_EnsureProfileShapesExist(ByVal ws As Worksheet, ByVal wb As Workbook) As Boolean
    If ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_HEADER_SHAPE_NAME) Is Nothing Then
        ex_UILoader.m_LoadUiFromConfig wb
    End If

    mp_EnsureProfileShapesExist = Not (ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_HEADER_SHAPE_NAME) Is Nothing)
End Function
