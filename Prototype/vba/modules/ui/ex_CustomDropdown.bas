Attribute VB_Name = "ex_CustomDropdown"
Option Explicit

Private Const DEV_SHEET_NAME As String = "Dev"
Private Const HEADER_SHAPE_NAME As String = "btnCustomMode"
Private Const OPTION_SHAPE_PREFIX As String = "btnCustomModeOption_"
Private Const STATE_DROPDOWN_EXPANDED_FLAG As String = "Settings.CustomModeDropdownExpanded"
Private Const PROFILE_HEADER_SHAPE_NAME As String = "btnCustomProfile"
Private Const PROFILE_OPTION_SHAPE_PREFIX As String = "btnCustomProfileOption_"
Private Const STATE_PROFILE_DROPDOWN_EXPANDED_FLAG As String = "Settings.CustomProfileDropdownExpanded"

Private Const TAG_KIND As String = "dd_kind"
Private Const TAG_KEY As String = "dd_key"
Private Const TAG_TARGET As String = "dd_target"
Private Const TAG_SET_CONTEXT As String = "dd_setContext"
Private Const TAG_ACTION_KEY As String = "dd_actionKey"
Private Const TAG_MACRO As String = "dd_macro"

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

Public Sub m_ToggleDevTestDropdown(Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim isExpanded As Boolean

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    If Not mp_EnsureShapesExist(ws, wb) Then Exit Sub

    mp_SyncDropdownContext ws
    mp_RebuildModeOptions ws

    isExpanded = Not mp_GetExpandedState()
    mp_SetExpandedState isExpanded
    mp_SetOptionsVisible ws, OPTION_SHAPE_PREFIX, isExpanded
End Sub

Public Sub m_SelectDevTestOption(Optional ByVal wb As Workbook)
    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    mp_HandleDynamicOptionSelect ws, "mode", OPTION_SHAPE_PREFIX

    mp_SetExpandedState False
    mp_SetOptionsVisible ws, OPTION_SHAPE_PREFIX, False
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

Public Sub m_ToggleCustomProfileDropdown(Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim isExpanded As Boolean

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    If Not mp_EnsureProfileShapesExist(ws, wb) Then Exit Sub

    mp_SyncDropdownContext ws
    mp_RebuildProfileOptions ws

    isExpanded = Not mp_GetProfileExpandedState()
    mp_SetProfileExpandedState isExpanded
    mp_SetOptionsVisible ws, PROFILE_OPTION_SHAPE_PREFIX, isExpanded
End Sub

Public Sub m_SelectCustomProfileOption(Optional ByVal wb As Workbook)
    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    mp_HandleDynamicOptionSelect ws, "profile", PROFILE_OPTION_SHAPE_PREFIX

    mp_SetProfileExpandedState False
    mp_SetOptionsVisible ws, PROFILE_OPTION_SHAPE_PREFIX, False
End Sub

Public Sub m_StabilizeChooseModeAnchorX(Optional ByVal ws As Worksheet, Optional ByVal targetLeft As Double = -1)
    Dim stableStartCol As Long
    Dim bufferCol As Long
    Dim currentLeft As Double
    Dim delta As Double
    Dim bufferRange As Range
    Dim targetBufferWidth As Double

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
    If targetBufferWidth < 4 Then targetBufferWidth = 4

    If Not mp_SetColumnWidthByPoints(bufferRange, targetBufferWidth) Then
        MsgBox "Failed to stabilize Choose mode anchor: unable to set buffer column width.", vbExclamation
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
    mp_RebuildDynamicOptions ws, HEADER_SHAPE_NAME, OPTION_SHAPE_PREFIX, HEADER_SHAPE_NAME, "mode", "ex_UIActions.m_SelectCustomModeOption_OnClick", "ddMode"
End Sub

Private Sub mp_RebuildProfileOptions(ByVal ws As Worksheet)
    mp_RebuildDynamicOptions ws, PROFILE_HEADER_SHAPE_NAME, PROFILE_OPTION_SHAPE_PREFIX, PROFILE_HEADER_SHAPE_NAME, "profile", "ex_UIActions.m_SelectCustomProfileOption_OnClick", "ddProfile"
End Sub

Private Sub mp_RebuildDynamicOptions( _
    ByVal ws As Worksheet, _
    ByVal headerShapeName As String, _
    ByVal optionPrefix As String, _
    ByVal sourceControlName As String, _
    ByVal optionKind As String, _
    ByVal onClickMacro As String, _
    ByVal fallbackDropdownShapeName As String)

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
    Dim currentTop As Double

    Set headerShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, headerShapeName)
    If headerShape Is Nothing Then
        MsgBox "Header shape '" & headerShapeName & "' was not found for dynamic dropdown.", vbExclamation
        Exit Sub
    End If

    Set templateShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, optionPrefix & "1")
    If templateShape Is Nothing Then
        MsgBox "Template shape '" & optionPrefix & "1' was not found for dynamic dropdown.", vbExclamation
        Exit Sub
    End If

    itemRecords = ex_UiXmlProvider.m_GetDropdownItemRecordsByControl(sourceControlName, ThisWorkbook)
    If Not mp_HasRecords(itemRecords) Then
        itemRecords = mp_BuildRecordsFromDropdownShape(ws, fallbackDropdownShapeName)
    End If
    If Not mp_HasRecords(itemRecords) Then
        MsgBox "No items were resolved for dynamic dropdown control '" & sourceControlName & "'.", vbExclamation
        Exit Sub
    End If

    recordLower = LBound(itemRecords, 1)
    recordUpper = UBound(itemRecords, 1)
    recordCount = recordUpper - recordLower + 1

    If Not mp_GetOptionLayout(optionPrefix, marginLeft, firstGap, nextGap, matchWidth) Then Exit Sub

    currentTop = headerShape.Top + headerShape.Height + firstGap

    For i = recordLower To recordUpper
        Set optionShape = mp_GetOrCreateOptionShape(ws, optionPrefix, i - recordLower + 1, templateShape)
        If optionShape Is Nothing Then Exit Sub

        mp_SetShapeOnAction optionShape, onClickMacro
        mp_SetOptionShapeMetadata optionShape, optionKind, itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY), itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_TARGET), itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_SET_CONTEXT), itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_ACTION_KEY), itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_MACRO)

        optionShape.TextFrame.Characters.Text = CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION))

        optionShape.Left = headerShape.Left + marginLeft
        optionShape.Top = currentTop
        If matchWidth Then optionShape.Width = headerShape.Width
        optionShape.ZOrder msoBringToFront

        currentTop = optionShape.Top + optionShape.Height + nextGap
    Next i

    mp_HideExcessOptionShapes ws, optionPrefix, recordCount
End Sub

Private Function mp_GetOrCreateOptionShape(ByVal ws As Worksheet, ByVal optionPrefix As String, ByVal optionIndex As Long, ByVal templateShape As Shape) As Shape
    Dim shapeName As String
    Dim duplicateRange As Object
    Dim createdShape As Shape

    shapeName = optionPrefix & CStr(optionIndex)
    Set createdShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, shapeName)
    If Not createdShape Is Nothing Then
        Set mp_GetOrCreateOptionShape = createdShape
        Exit Function
    End If

    If templateShape Is Nothing Then Exit Function

    On Error GoTo EH_DUP
    Set duplicateRange = templateShape.Duplicate
    Set createdShape = duplicateRange.Item(1)
    createdShape.Name = shapeName
    createdShape.Visible = msoFalse
    createdShape.Placement = xlFreeFloating

    Set mp_GetOrCreateOptionShape = createdShape
    Exit Function
EH_DUP:
    MsgBox "Failed to create dynamic option shape '" & shapeName & "': " & Err.Description, vbExclamation
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
            End If
        End If
    Next shp
End Sub

Private Function mp_GetOptionLayout(ByVal optionPrefix As String, ByRef marginLeft As Double, ByRef firstGap As Double, ByRef nextGap As Double, ByRef matchWidth As Boolean) As Boolean
    Dim firstName As String
    Dim secondName As String

    firstName = optionPrefix & "1"
    secondName = optionPrefix & "2"

    marginLeft = 0
    firstGap = 2
    nextGap = 2
    matchWidth = True

    If Not mp_TryGetOptionalDoubleAttr(firstName, "marginLeft", marginLeft) Then Exit Function
    If Not mp_TryGetOptionalDoubleAttr(firstName, "marginTop", firstGap) Then Exit Function
    If Not mp_TryGetOptionalDoubleAttr(secondName, "marginTop", nextGap) Then Exit Function

    If Len(Trim$(ex_UiXmlProvider.m_GetControlAttribute(firstName, "matchWidthToRelative", ThisWorkbook))) > 0 Then
        If Not mp_TryGetOptionalBooleanAttr(firstName, "matchWidthToRelative", matchWidth) Then Exit Function
    End If

    mp_GetOptionLayout = True
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

Private Function mp_BuildRecordsFromDropdownShape(ByVal ws As Worksheet, ByVal dropdownShapeName As String) As Variant
    Dim shp As Shape
    Dim cf As Object
    Dim listCount As Long
    Dim records() As Variant
    Dim i As Long
    Dim itemText As String

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, dropdownShapeName)
    If shp Is Nothing Then Exit Function

    On Error Resume Next
    Set cf = shp.ControlFormat
    On Error GoTo 0
    If cf Is Nothing Then Exit Function

    On Error Resume Next
    listCount = CLng(cf.ListCount)
    On Error GoTo 0
    If listCount <= 0 Then Exit Function

    ReDim records(1 To listCount, 1 To ex_UiXmlProvider.DROPDOWN_ITEM_COL_MACRO)
    For i = 1 To listCount
        On Error Resume Next
        itemText = CStr(cf.List(i))
        On Error GoTo 0
        itemText = Trim$(itemText)
        If Len(itemText) = 0 Then itemText = "Item " & CStr(i)

        records(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY) = itemText
        records(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION) = itemText
        records(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_TARGET) = itemText
        records(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_SET_CONTEXT) = vbNullString
        records(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_ACTION_KEY) = vbNullString
        records(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_MACRO) = vbNullString
    Next i

    mp_BuildRecordsFromDropdownShape = records
End Function

Private Function mp_HasRecords(ByVal records As Variant) As Boolean
    On Error GoTo EH
    If Not IsArray(records) Then Exit Function
    mp_HasRecords = (UBound(records, 1) >= LBound(records, 1))
    Exit Function
EH:
    mp_HasRecords = False
End Function

Private Sub mp_SetOptionShapeMetadata(ByVal shp As Shape, ByVal optionKind As String, ByVal keyText As String, ByVal targetText As String, ByVal setContextText As String, ByVal actionKey As String, ByVal macroName As String)
    mp_SetShapeTag shp, TAG_KIND, optionKind
    mp_SetShapeTag shp, TAG_KEY, keyText
    mp_SetShapeTag shp, TAG_TARGET, targetText
    mp_SetShapeTag shp, TAG_SET_CONTEXT, setContextText
    mp_SetShapeTag shp, TAG_ACTION_KEY, actionKey
    mp_SetShapeTag shp, TAG_MACRO, macroName
End Sub

Private Sub mp_SetShapeTag(ByVal shp As Shape, ByVal tagName As String, ByVal tagValue As String)
    On Error Resume Next
    shp.Tags.Delete tagName
    shp.Tags.Add tagName, CStr(tagValue)
    On Error GoTo 0
End Sub

Private Function mp_GetShapeTag(ByVal shp As Shape, ByVal tagName As String) As String
    On Error Resume Next
    mp_GetShapeTag = CStr(shp.Tags(tagName))
    On Error GoTo 0
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

Private Sub mp_HandleDynamicOptionSelect(ByVal ws As Worksheet, ByVal expectedKind As String, ByVal optionPrefix As String)
    Dim callerName As String
    Dim optionShape As Shape
    Dim optionKind As String
    Dim targetText As String
    Dim keyText As String
    Dim setContextText As String
    Dim actionKey As String
    Dim macroName As String

    callerName = vbNullString
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0
    callerName = Trim$(callerName)
    If Len(callerName) = 0 Then Exit Sub

    If StrComp(Left$(callerName, Len(optionPrefix)), optionPrefix, vbTextCompare) <> 0 Then Exit Sub

    Set optionShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, callerName)
    If optionShape Is Nothing Then Exit Sub

    optionKind = LCase$(Trim$(mp_GetShapeTag(optionShape, TAG_KIND)))
    If Len(optionKind) = 0 Then optionKind = LCase$(expectedKind)
    If StrComp(optionKind, LCase$(expectedKind), vbTextCompare) <> 0 Then Exit Sub

    keyText = Trim$(mp_GetShapeTag(optionShape, TAG_KEY))
    targetText = Trim$(mp_GetShapeTag(optionShape, TAG_TARGET))
    If Len(targetText) = 0 Then targetText = keyText

    setContextText = Trim$(mp_GetShapeTag(optionShape, TAG_SET_CONTEXT))
    actionKey = Trim$(mp_GetShapeTag(optionShape, TAG_ACTION_KEY))
    macroName = Trim$(mp_GetShapeTag(optionShape, TAG_MACRO))

    If Len(keyText) = 0 And Len(targetText) = 0 Then
        mp_FillOptionDataByCallerIndex ws, expectedKind, callerName, optionPrefix, keyText, targetText, setContextText, actionKey, macroName
    End If

    If Len(setContextText) > 0 Then
        ex_UiXmlProvider.m_ApplyDropdownSetContext setContextText
    End If

    If StrComp(optionKind, "mode", vbTextCompare) = 0 Then
        If Len(targetText) = 0 Then
            MsgBox "Dynamic mode option '" & callerName & "' has no target mode.", vbExclamation
            Exit Sub
        End If

        If Not mp_SelectDropdownByItemText(ws, "ddMode", targetText) Then
            MsgBox "Mode '" & targetText & "' was not found in hidden dropdown 'ddMode'.", vbExclamation
            Exit Sub
        End If

        ex_ConfigProfilesManager.m_OnModeChanged
        If Len(keyText) > 0 Then
            ex_UiXmlProvider.m_SetDropdownContextValue "activeMode", keyText
        Else
            ex_UiXmlProvider.m_SetDropdownContextValue "activeMode", targetText
        End If

    ElseIf StrComp(optionKind, "profile", vbTextCompare) = 0 Then
        If Len(targetText) = 0 Then
            MsgBox "Dynamic profile option '" & callerName & "' has no target profile.", vbExclamation
            Exit Sub
        End If

        If Not mp_SelectDropdownByItemText(ws, "ddProfile", targetText) Then
            MsgBox "Profile '" & targetText & "' was not found in hidden dropdown 'ddProfile'.", vbExclamation
            Exit Sub
        End If

        ex_ConfigProfilesManager.m_OnProfileChanged
        ex_UiXmlProvider.m_SetDropdownContextValue "activeProfile", ex_ConfigProfilesManager.m_GetActiveProfileName(ws)
    End If

    If Len(macroName) = 0 And Len(actionKey) > 0 Then
        macroName = ex_UiXmlProvider.m_ResolveMacroByActionKey(actionKey, ThisWorkbook)
    End If

    If Len(macroName) > 0 Then
        mp_RunMacroByName macroName
    End If
End Sub

Private Sub mp_FillOptionDataByCallerIndex( _
    ByVal ws As Worksheet, _
    ByVal optionKind As String, _
    ByVal callerName As String, _
    ByVal optionPrefix As String, _
    ByRef keyText As String, _
    ByRef targetText As String, _
    ByRef setContextText As String, _
    ByRef actionKey As String, _
    ByRef macroName As String)

    Dim suffix As String
    Dim itemIndex As Long
    Dim records As Variant
    Dim rowIndex As Long
    Dim sourceControlName As String
    Dim dropdownShapeName As String
    Dim shp As Shape
    Dim cf As Object
    Dim fallbackText As String

    suffix = Mid$(callerName, Len(optionPrefix) + 1)
    If Not IsNumeric(suffix) Then Exit Sub

    itemIndex = CLng(suffix)
    If itemIndex <= 0 Then Exit Sub

    If StrComp(optionKind, "profile", vbTextCompare) = 0 Then
        sourceControlName = PROFILE_HEADER_SHAPE_NAME
        dropdownShapeName = "ddProfile"
    Else
        sourceControlName = HEADER_SHAPE_NAME
        dropdownShapeName = "ddMode"
    End If

    records = ex_UiXmlProvider.m_GetDropdownItemRecordsByControl(sourceControlName, ThisWorkbook)
    If mp_HasRecords(records) Then
        rowIndex = LBound(records, 1) + itemIndex - 1
        If rowIndex >= LBound(records, 1) And rowIndex <= UBound(records, 1) Then
            If Len(keyText) = 0 Then keyText = Trim$(CStr(records(rowIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY)))
            If Len(targetText) = 0 Then targetText = Trim$(CStr(records(rowIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_TARGET)))
            If Len(setContextText) = 0 Then setContextText = Trim$(CStr(records(rowIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_SET_CONTEXT)))
            If Len(actionKey) = 0 Then actionKey = Trim$(CStr(records(rowIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_ACTION_KEY)))
            If Len(macroName) = 0 Then macroName = Trim$(CStr(records(rowIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_MACRO)))
            If Len(targetText) = 0 Then targetText = Trim$(CStr(records(rowIndex, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION)))
        End If
    End If

    If Len(targetText) > 0 Then Exit Sub

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, dropdownShapeName)
    If shp Is Nothing Then Exit Sub

    On Error Resume Next
    Set cf = shp.ControlFormat
    On Error GoTo 0
    If cf Is Nothing Then Exit Sub

    On Error Resume Next
    fallbackText = Trim$(CStr(cf.List(itemIndex)))
    On Error GoTo 0
    If Len(fallbackText) = 0 Then Exit Sub

    If Len(keyText) = 0 Then keyText = fallbackText
    targetText = fallbackText
End Sub

Private Function mp_SelectDropdownByItemText(ByVal ws As Worksheet, ByVal dropdownShapeName As String, ByVal itemText As String) As Boolean
    Dim shp As Shape
    Dim cf As Object
    Dim listCount As Long
    Dim i As Long

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, dropdownShapeName)
    If shp Is Nothing Then Exit Function

    On Error Resume Next
    Set cf = shp.ControlFormat
    On Error GoTo 0
    If cf Is Nothing Then Exit Function

    On Error Resume Next
    listCount = CLng(cf.ListCount)
    On Error GoTo 0
    If listCount <= 0 Then Exit Function

    For i = 1 To listCount
        On Error Resume Next
        If StrComp(CStr(cf.List(i)), itemText, vbTextCompare) = 0 Then
            cf.Value = i
            On Error GoTo 0
            mp_SelectDropdownByItemText = True
            Exit Function
        End If
        On Error GoTo 0
    Next i
End Function

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
    Dim visibility As MsoTriState

    visibility = IIf(isVisible, msoTrue, msoFalse)

    For Each shp In ws.Shapes
        If StrComp(Left$(shp.Name, Len(optionPrefix)), optionPrefix, vbTextCompare) = 0 Then
            shp.Visible = visibility
        End If
    Next shp
End Sub

Private Sub mp_SyncDropdownContext(ByVal ws As Worksheet)
    Dim modeName As String
    Dim modeKey As String
    Dim profileName As String

    modeName = Trim$(ex_ConfigProfilesManager.m_GetActiveModeName(ws))
    If Len(modeName) > 0 Then
        modeKey = Trim$(ex_UiXmlProvider.m_GetDropdownItemKeyByTarget(HEADER_SHAPE_NAME, modeName, ThisWorkbook))
        If Len(modeKey) = 0 Then modeKey = modeName
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

    If ex_ConfigProfilesManager.m_GetShapeByName(ws, HEADER_SHAPE_NAME) Is Nothing Then Exit Function
    If ex_ConfigProfilesManager.m_GetShapeByName(ws, OPTION_SHAPE_PREFIX & "1") Is Nothing Then
        ex_UILoader.m_LoadUiFromConfig wb
    End If

    mp_EnsureShapesExist = Not (ex_ConfigProfilesManager.m_GetShapeByName(ws, OPTION_SHAPE_PREFIX & "1") Is Nothing)
End Function

Private Function mp_EnsureProfileShapesExist(ByVal ws As Worksheet, ByVal wb As Workbook) As Boolean
    If ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_HEADER_SHAPE_NAME) Is Nothing Then
        ex_UILoader.m_LoadUiFromConfig wb
    End If

    If ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_HEADER_SHAPE_NAME) Is Nothing Then Exit Function
    If ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_OPTION_SHAPE_PREFIX & "1") Is Nothing Then
        ex_UILoader.m_LoadUiFromConfig wb
    End If

    mp_EnsureProfileShapesExist = Not (ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_OPTION_SHAPE_PREFIX & "1") Is Nothing)
End Function
