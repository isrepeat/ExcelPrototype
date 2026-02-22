Attribute VB_Name = "ex_CustomDropdown"
Option Explicit

Private Const DEV_SHEET_NAME As String = "Dev"
Private Const HEADER_SHAPE_NAME As String = "btnCustomMode"
Private Const OPTION_SHAPE_PREFIX As String = "btnCustomModeOption_"
Private Const OPTION_COUNT As Long = 2
Private Const STATE_DROPDOWN_EXPANDED_FLAG As String = "Settings.CustomModeDropdownExpanded"
Private Const PROFILE_HEADER_SHAPE_NAME As String = "btnCustomProfile"
Private Const PROFILE_OPTION_SHAPE_PREFIX As String = "btnCustomProfileOption_"
Private Const PROFILE_OPTION_COUNT As Long = 2
Private Const STATE_PROFILE_DROPDOWN_EXPANDED_FLAG As String = "Settings.CustomProfileDropdownExpanded"

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

    mp_SyncProfileOptionCaptions ws
    mp_RepositionOptionShapes ws
    mp_RepositionProfileOptionShapes ws

    isExpanded = mp_GetExpandedState()
    mp_SetOptionsVisible ws, isExpanded

    isProfileExpanded = mp_GetProfileExpandedState()
    mp_SetProfileOptionsVisible ws, isProfileExpanded
End Sub

Public Sub m_ToggleDevTestDropdown(Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim isExpanded As Boolean

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    If Not mp_EnsureShapesExist(ws, wb) Then Exit Sub

    mp_RepositionOptionShapes ws
    isExpanded = Not mp_GetExpandedState()
    mp_SetExpandedState isExpanded
    mp_SetOptionsVisible ws, isExpanded
End Sub

Public Sub m_SelectDevTestOption(Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim callerName As String
    Dim optionIndex As Long

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    callerName = vbNullString
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0

    optionIndex = mp_ParseOptionIndex(callerName)
    If optionIndex < 1 Then Exit Sub
    If optionIndex > OPTION_COUNT Then Exit Sub

    mp_SelectModeByOption ws, optionIndex
    mp_SetExpandedState False
    mp_SetOptionsVisible ws, False
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
    mp_SetOptionsVisible ws, False
End Sub

Public Sub m_HideCustomProfileDropdown(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        Set ws = mp_GetDevSheet(ThisWorkbook)
    End If
    If ws Is Nothing Then Exit Sub

    mp_SetProfileExpandedState False
    mp_SetProfileOptionsVisible ws, False
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

    mp_SyncProfileOptionCaptions ws
    mp_RepositionProfileOptionShapes ws
    isExpanded = Not mp_GetProfileExpandedState()
    mp_SetProfileExpandedState isExpanded
    mp_SetProfileOptionsVisible ws, isExpanded
End Sub

Public Sub m_SelectCustomProfileOption(Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim callerName As String
    Dim optionIndex As Long

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = mp_GetDevSheet(wb)
    If ws Is Nothing Then Exit Sub

    callerName = vbNullString
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0

    optionIndex = mp_ParseOptionIndexByPrefix(callerName, PROFILE_OPTION_SHAPE_PREFIX)
    If optionIndex < 1 Then Exit Sub
    If optionIndex > PROFILE_OPTION_COUNT Then Exit Sub

    mp_SelectProfileByOption ws, optionIndex
    mp_SetProfileExpandedState False
    mp_SetProfileOptionsVisible ws, False
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

Private Function mp_GetHeaderShape(ByVal ws As Worksheet) As Shape
    Dim headerShape As Shape

    Set headerShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, HEADER_SHAPE_NAME)
    Set mp_GetHeaderShape = headerShape
End Function

Private Function mp_EnsureShapesExist(ByVal ws As Worksheet, ByVal wb As Workbook) As Boolean
    Dim i As Long

    If mp_GetHeaderShape(ws) Is Nothing Then
        ex_UILoader.m_LoadUiFromConfig wb
    End If

    If mp_GetHeaderShape(ws) Is Nothing Then Exit Function

    For i = 1 To OPTION_COUNT
        If ex_ConfigProfilesManager.m_GetShapeByName(ws, OPTION_SHAPE_PREFIX & CStr(i)) Is Nothing Then
            ex_UILoader.m_LoadUiFromConfig wb
            Exit For
        End If
    Next i

    For i = 1 To OPTION_COUNT
        If ex_ConfigProfilesManager.m_GetShapeByName(ws, OPTION_SHAPE_PREFIX & CStr(i)) Is Nothing Then Exit Function
    Next i

    mp_EnsureShapesExist = True
End Function

Private Function mp_EnsureProfileShapesExist(ByVal ws As Worksheet, ByVal wb As Workbook) As Boolean
    Dim i As Long

    If ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_HEADER_SHAPE_NAME) Is Nothing Then
        ex_UILoader.m_LoadUiFromConfig wb
    End If

    If ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_HEADER_SHAPE_NAME) Is Nothing Then Exit Function

    For i = 1 To PROFILE_OPTION_COUNT
        If ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_OPTION_SHAPE_PREFIX & CStr(i)) Is Nothing Then
            ex_UILoader.m_LoadUiFromConfig wb
            Exit For
        End If
    Next i

    For i = 1 To PROFILE_OPTION_COUNT
        If ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_OPTION_SHAPE_PREFIX & CStr(i)) Is Nothing Then Exit Function
    Next i

    mp_EnsureProfileShapesExist = True
End Function

Private Sub mp_RepositionOptionShapes(ByVal ws As Worksheet)
    Dim i As Long
    Dim optionControlName As String
    Dim baseControlName As String
    Dim optionShape As Shape
    Dim baseShape As Shape
    Dim marginTop As Double
    Dim marginLeft As Double

    For i = 1 To OPTION_COUNT
        optionControlName = OPTION_SHAPE_PREFIX & CStr(i)
        Set optionShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, optionControlName)
        If optionShape Is Nothing Then GoTo NextOption

        If Not mp_TryGetRequiredTextAttr(optionControlName, "relativeTo", baseControlName) Then GoTo NextOption
        Set baseShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, baseControlName)
        If baseShape Is Nothing Then
            MsgBox "Relative anchor control '" & baseControlName & "' was not found for '" & optionControlName & "'.", vbExclamation
            GoTo NextOption
        End If

        If Not mp_TryGetRequiredDoubleAttr(optionControlName, "marginTop", marginTop) Then GoTo NextOption
        If Not mp_TryGetRequiredDoubleAttr(optionControlName, "marginLeft", marginLeft) Then GoTo NextOption

        optionShape.Left = baseShape.Left + marginLeft
        optionShape.Top = baseShape.Top + baseShape.Height + marginTop
        If mp_ShouldMatchWidthToRelative(optionControlName) Then
            optionShape.Width = baseShape.Width
        End If
        optionShape.ZOrder msoBringToFront
NextOption:
    Next i
End Sub

Private Function mp_ShouldMatchWidthToRelative(ByVal controlName As String) As Boolean
    Dim valueText As String
    Dim valueBool As Boolean

    valueText = Trim$(ex_UiXmlProvider.m_GetControlAttribute(controlName, "matchWidthToRelative", ThisWorkbook))
    If Len(valueText) = 0 Then Exit Function

    If Not ex_XmlCore.m_TryParseBoolean(valueText, valueBool) Then
        MsgBox "Invalid boolean value '" & valueText & "' for attribute 'matchWidthToRelative' on control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    mp_ShouldMatchWidthToRelative = valueBool
End Function

Private Function mp_TryGetRequiredTextAttr(ByVal controlName As String, ByVal attrName As String, ByRef outValue As String) As Boolean
    Dim valueText As String

    valueText = Trim$(ex_UiXmlProvider.m_GetControlAttribute(controlName, attrName, ThisWorkbook))
    If Len(valueText) = 0 Then
        MsgBox "Control '" & controlName & "' must define required attribute '" & attrName & "' in DevUI.xml.", vbExclamation
        Exit Function
    End If

    outValue = valueText
    mp_TryGetRequiredTextAttr = True
End Function

Private Function mp_TryGetRequiredDoubleAttr(ByVal controlName As String, ByVal attrName As String, ByRef outValue As Double) As Boolean
    Dim valueText As String
    Dim parsedValue As Double

    valueText = Trim$(ex_UiXmlProvider.m_GetControlAttribute(controlName, attrName, ThisWorkbook))
    If Len(valueText) = 0 Then
        MsgBox "Control '" & controlName & "' must define required attribute '" & attrName & "' in DevUI.xml.", vbExclamation
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseDouble(valueText, parsedValue, True) Then
        MsgBox "Invalid numeric value '" & valueText & "' for attribute '" & attrName & "' on control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    outValue = parsedValue
    mp_TryGetRequiredDoubleAttr = True
End Function

Private Function mp_ParseOptionIndex(ByVal callerName As String) As Long
    mp_ParseOptionIndex = mp_ParseOptionIndexByPrefix(callerName, OPTION_SHAPE_PREFIX)
End Function

Private Function mp_ParseOptionIndexByPrefix(ByVal callerName As String, ByVal prefix As String) As Long
    Dim rawValue As String

    callerName = Trim$(callerName)
    If Len(callerName) = 0 Then Exit Function

    If StrComp(Left$(callerName, Len(prefix)), prefix, vbTextCompare) <> 0 Then Exit Function

    rawValue = Mid$(callerName, Len(prefix) + 1)
    If Len(rawValue) = 0 Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function

    mp_ParseOptionIndexByPrefix = CLng(rawValue)
End Function

Private Function mp_HasOptionShapes(ByVal ws As Worksheet) As Boolean
    Dim firstOption As Shape

    Set firstOption = ex_ConfigProfilesManager.m_GetShapeByName(ws, OPTION_SHAPE_PREFIX & "1")
    mp_HasOptionShapes = Not firstOption Is Nothing
End Function

Private Sub mp_SelectModeByOption(ByVal ws As Worksheet, ByVal optionIndex As Long)
    Dim modeShape As Shape
    Dim cf As Object
    Dim listCount As Long

    Set modeShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, "ddMode")
    If modeShape Is Nothing Then Exit Sub

    On Error Resume Next
    Set cf = modeShape.ControlFormat
    On Error GoTo 0
    If cf Is Nothing Then Exit Sub

    On Error Resume Next
    listCount = CLng(cf.ListCount)
    On Error GoTo 0
    If listCount < optionIndex Then Exit Sub

    On Error Resume Next
    cf.Value = optionIndex
    On Error GoTo 0

    ex_ConfigProfilesManager.m_OnModeChanged
End Sub

Private Sub mp_SelectProfileByOption(ByVal ws As Worksheet, ByVal optionIndex As Long)
    Dim profileShape As Shape
    Dim cf As Object
    Dim listCount As Long

    Set profileShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, "ddProfile")
    If profileShape Is Nothing Then Exit Sub

    On Error Resume Next
    Set cf = profileShape.ControlFormat
    On Error GoTo 0
    If cf Is Nothing Then Exit Sub

    On Error Resume Next
    listCount = CLng(cf.ListCount)
    On Error GoTo 0
    If listCount < optionIndex Then Exit Sub

    On Error Resume Next
    cf.Value = optionIndex
    On Error GoTo 0

    ex_ConfigProfilesManager.m_OnProfileChanged
End Sub

Private Function mp_AreOptionsVisible(ByVal ws As Worksheet) As Boolean
    Dim firstOption As Shape

    Set firstOption = ex_ConfigProfilesManager.m_GetShapeByName(ws, OPTION_SHAPE_PREFIX & "1")
    If firstOption Is Nothing Then Exit Function

    mp_AreOptionsVisible = (firstOption.Visible = msoTrue)
End Function

Private Sub mp_SetOptionsVisible(ByVal ws As Worksheet, ByVal isVisible As Boolean)
    Dim i As Long
    Dim optionShape As Shape
    Dim visibility As MsoTriState

    visibility = IIf(isVisible, msoTrue, msoFalse)

    For i = 1 To OPTION_COUNT
        Set optionShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, OPTION_SHAPE_PREFIX & CStr(i))
        If Not optionShape Is Nothing Then
            optionShape.Visible = visibility
        End If
    Next i
End Sub

Private Sub mp_SetProfileOptionsVisible(ByVal ws As Worksheet, ByVal isVisible As Boolean)
    Dim i As Long
    Dim optionShape As Shape
    Dim visibility As MsoTriState

    visibility = IIf(isVisible, msoTrue, msoFalse)

    For i = 1 To PROFILE_OPTION_COUNT
        Set optionShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_OPTION_SHAPE_PREFIX & CStr(i))
        If Not optionShape Is Nothing Then
            optionShape.Visible = visibility
        End If
    Next i
End Sub

Private Sub mp_RepositionProfileOptionShapes(ByVal ws As Worksheet)
    Dim i As Long
    Dim optionControlName As String
    Dim baseControlName As String
    Dim optionShape As Shape
    Dim baseShape As Shape
    Dim marginTop As Double
    Dim marginLeft As Double

    For i = 1 To PROFILE_OPTION_COUNT
        optionControlName = PROFILE_OPTION_SHAPE_PREFIX & CStr(i)
        Set optionShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, optionControlName)
        If optionShape Is Nothing Then GoTo NextOption

        If Not mp_TryGetRequiredTextAttr(optionControlName, "relativeTo", baseControlName) Then GoTo NextOption
        Set baseShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, baseControlName)
        If baseShape Is Nothing Then
            MsgBox "Relative anchor control '" & baseControlName & "' was not found for '" & optionControlName & "'.", vbExclamation
            GoTo NextOption
        End If

        If Not mp_TryGetRequiredDoubleAttr(optionControlName, "marginTop", marginTop) Then GoTo NextOption
        If Not mp_TryGetRequiredDoubleAttr(optionControlName, "marginLeft", marginLeft) Then GoTo NextOption

        optionShape.Left = baseShape.Left + marginLeft
        optionShape.Top = baseShape.Top + baseShape.Height + marginTop
        If mp_ShouldMatchWidthToRelative(optionControlName) Then
            optionShape.Width = baseShape.Width
        End If
        optionShape.ZOrder msoBringToFront
NextOption:
    Next i
End Sub

Private Sub mp_SyncProfileOptionCaptions(ByVal ws As Worksheet)
    Dim profileShape As Shape
    Dim cf As Object
    Dim i As Long
    Dim listCount As Long
    Dim itemText As String
    Dim optionShape As Shape

    Set profileShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, "ddProfile")
    If profileShape Is Nothing Then Exit Sub

    On Error Resume Next
    Set cf = profileShape.ControlFormat
    On Error GoTo 0
    If cf Is Nothing Then Exit Sub

    On Error Resume Next
    listCount = CLng(cf.ListCount)
    On Error GoTo 0

    For i = 1 To PROFILE_OPTION_COUNT
        Set optionShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, PROFILE_OPTION_SHAPE_PREFIX & CStr(i))
        If optionShape Is Nothing Then GoTo NextOption

        If i <= listCount Then
            On Error Resume Next
            itemText = CStr(cf.List(i))
            On Error GoTo 0
            If Len(Trim$(itemText)) = 0 Then
                itemText = "Profile " & CStr(i)
            End If
            optionShape.TextFrame.Characters.Text = itemText
            optionShape.Visible = msoTrue
        Else
            optionShape.Visible = msoFalse
        End If
NextOption:
    Next i
End Sub
