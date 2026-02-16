Attribute VB_Name = "ex_ProfileUI"
Option Explicit

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const GLOBAL_BUTTONS_REL_PATH As String = "config\GlobalButtons.xml"
Private Const UI_BLOCK_GROUP_NAME As String = "grpUiBlock"
Private Const MODE_DROPDOWN_SHAPE As String = "ddMode"
Private Const UPDATE_BUTTON_SHAPE As String = "btnUpdateCode"
Private Const CLEAR_BUTTON_SHAPE As String = "btnClear"
Private Const MODE_BUTTON_SHAPE As String = "btnMode"
Private Const PERSONAL_BUTTON_SHAPE As String = "btnPersonalCard"
Private Const COMPARING_BUTTON_SHAPE As String = "btnComparing"

' Initial absolute layout in points.
Private Const UI_DDMODE_LEFT As Double = 758.25
Private Const UI_DDMODE_TOP As Double = 2.25
Private Const UI_DDMODE_WIDTH As Double = 156#
Private Const UI_DDMODE_HEIGHT As Double = 15#
Private Const UI_CLEAR_LEFT As Double = 758.25
Private Const UI_CLEAR_TOP As Double = 30.75
Private Const UI_CLEAR_WIDTH As Double = 156#
Private Const UI_CLEAR_HEIGHT As Double = 56.69
Private Const UI_PERSONAL_LEFT As Double = 758.25
Private Const UI_PERSONAL_TOP As Double = 93.11
Private Const UI_PERSONAL_WIDTH As Double = 155.91
Private Const UI_PERSONAL_HEIGHT As Double = 56.69
Private Const UI_COMPARING_LEFT As Double = 758.25
Private Const UI_COMPARING_TOP As Double = 93.11
Private Const UI_COMPARING_WIDTH As Double = 155.91
Private Const UI_COMPARING_HEIGHT As Double = 56.69
Private Const UI_MODEBTN_LEFT As Double = 912#
Private Const UI_MODEBTN_TOP As Double = 102.99
Private Const UI_MODEBTN_WIDTH As Double = 155.25
Private Const UI_MODEBTN_HEIGHT As Double = 36.3

Public Sub m_ApplyProfileUI(ByVal ws As Worksheet, ByVal profileNode As Object, Optional ByVal profileName As String = vbNullString)
    Dim uiNodes As Object
    Dim node As Object
    Dim shapeName As String
    Dim shp As Shape

    If ws Is Nothing Then
        MsgBox "Failed to apply profile UI: worksheet is not specified.", vbExclamation
        Exit Sub
    End If
    If profileNode Is Nothing Then
        MsgBox "Failed to apply profile UI: profile node is not specified.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    profileNode.OwnerDocument.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"
    On Error GoTo 0

    Set uiNodes = profileNode.selectNodes("p:ui/p:shape")
    If uiNodes Is Nothing Then Exit Sub
    If uiNodes.Length = 0 Then Exit Sub

    For Each node In uiNodes
        shapeName = Trim$(mp_NodeAttrText(node, "name"))
        If Len(shapeName) = 0 Then
            MsgBox "Profile UI contains shape entry without 'name' attribute.", vbExclamation
            Exit Sub
        End If

        Set shp = m_GetShapeByName(ws, shapeName)
        If shp Is Nothing Then
            MsgBox "Profile UI shape '" & shapeName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
            Exit Sub
        End If

        If Not mp_ApplyShapeVisible(node, shp) Then Exit Sub
        If Not mp_ApplyShapePlacement(node, shp, ws) Then Exit Sub
        If Not mp_ApplyShapeGeometry(node, shp) Then Exit Sub
        If Not mp_ApplyShapeColor(node, shp, profileName) Then Exit Sub

        Set shp = Nothing
    Next node
End Sub

' Keeps UI controls detached from cell grid so their coordinates stay absolute.
' Managed block: all btn* except btnUpdateCode + ddMode dropdown.
Public Sub m_EnsureUiControlsAbsolute(Optional ByVal ws As Worksheet)
    Dim shp As Shape

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If
    If ws Is Nothing Then
        MsgBox "Failed to apply absolute UI layout: worksheet is not specified.", vbExclamation
        Exit Sub
    End If

    For Each shp In ws.Shapes
        If mp_IsManagedUiBlockShape(shp.Name) Then
            On Error GoTo EH_PLACEMENT
            shp.Placement = xlFreeFloating
            On Error GoTo 0
        End If
    Next shp

    Exit Sub
EH_PLACEMENT:
    MsgBox "Failed to set absolute placement for shape '" & shp.Name & "': " & Err.Description, vbExclamation
End Sub

Public Sub m_InitUiBlockLayoutAndGroup(Optional ByVal ws As Worksheet)
    Dim shp As Shape
    Dim names As Variant
    Dim groupShape As Shape
    Dim shapeName As Variant

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If
    If ws Is Nothing Then
        MsgBox "Failed to initialize UI block group: worksheet is not specified.", vbExclamation
        Exit Sub
    End If

    m_EnsureUiControlsAbsolute ws
    mp_ApplyInitialUiLayout ws

    On Error GoTo EH_UNGROUP
    mp_UngroupManagedUiShapes ws
    On Error GoTo 0

    names = Array(MODE_DROPDOWN_SHAPE, CLEAR_BUTTON_SHAPE, MODE_BUTTON_SHAPE, PERSONAL_BUTTON_SHAPE, COMPARING_BUTTON_SHAPE)
    For Each shapeName In names
        Set shp = m_GetShapeByName(ws, CStr(shapeName))
        If shp Is Nothing Then
            MsgBox "Failed to initialize UI block group: shape '" & CStr(shapeName) & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
            Exit Sub
        End If
    Next shapeName

    On Error GoTo EH_GROUP
    Set groupShape = ws.Shapes.Range(names).Group
    groupShape.Name = UI_BLOCK_GROUP_NAME
    Exit Sub

EH_UNGROUP:
    MsgBox "Failed to ungroup existing UI block shapes before creating '" & UI_BLOCK_GROUP_NAME & "': " & Err.Description, vbExclamation
    Exit Sub
EH_GROUP:
    MsgBox "Failed to create group '" & UI_BLOCK_GROUP_NAME & "'. Group ddMode + buttons manually if needed: " & Err.Description, vbExclamation
End Sub

Public Sub m_ApplyModeVisibility(ByVal ws As Worksheet, ByVal profileNode As Object)
    Dim globalDoc As Object
    Dim globalNodes As Object
    Dim uiNodes As Object

    If ws Is Nothing Then
        MsgBox "Failed to apply mode visibility: worksheet is not specified.", vbExclamation
        Exit Sub
    End If
    If profileNode Is Nothing Then
        MsgBox "Failed to apply mode visibility: profile node is not specified.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    profileNode.OwnerDocument.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"
    On Error GoTo 0

    ' Guardrail: any shape named as button (btn*) must be explicitly enabled by current profile.
    mp_HideAllButtons ws

    Set globalDoc = mp_LoadGlobalButtonsDom()
    If globalDoc Is Nothing Then Exit Sub

    Set globalNodes = globalDoc.selectNodes("/p:globalButtons/p:shape")
    If globalNodes Is Nothing Then
        MsgBox "Invalid global buttons file format. Expected '/globalButtons/shape'.", vbExclamation
        Exit Sub
    End If
    mp_ApplyFilteredVisibilityFromNodes ws, globalNodes

    Set uiNodes = profileNode.selectNodes("p:ui/p:shape")
    If uiNodes Is Nothing Then Exit Sub
    If uiNodes.Length = 0 Then Exit Sub
    mp_ApplyFilteredVisibilityFromNodes ws, uiNodes
End Sub

Private Function mp_IsShapeVisibleByFilters(ByVal node As Object) As Boolean
    Dim visibleText As String
    Dim isBaseVisible As Boolean

    visibleText = Trim$(mp_NodeAttrText(node, "visible"))
    isBaseVisible = False
    If Len(visibleText) > 0 Then
        If Not mp_TryParseBoolean(visibleText, isBaseVisible) Then
            MsgBox "Invalid boolean value for UI attribute 'visible' in mode filter block: " & visibleText, vbExclamation
            Exit Function
        End If
    End If
    If Not isBaseVisible Then Exit Function
    mp_IsShapeVisibleByFilters = True
End Function

Private Sub mp_ApplyFilteredVisibilityFromNodes(ByVal ws As Worksheet, ByVal nodes As Object)
    Dim node As Object
    Dim shapeName As String
    Dim shp As Shape

    For Each node In nodes
        shapeName = Trim$(mp_NodeAttrText(node, "name"))
        If Len(shapeName) = 0 Then
            MsgBox "UI visibility block contains shape entry without 'name'.", vbExclamation
            Exit Sub
        End If
        If Not mp_IsButtonShapeName(shapeName) Then GoTo NextNode

        Set shp = m_GetShapeByName(ws, shapeName)
        If shp Is Nothing Then
            MsgBox "UI shape '" & shapeName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
            Exit Sub
        End If

        If mp_IsShapeVisibleByFilters(node) Then
            shp.Visible = msoTrue
        End If
        Set shp = Nothing
NextNode:
    Next node
End Sub

Private Function mp_LoadGlobalButtonsDom() As Object
    Dim filePath As String
    Dim doc As Object

    filePath = mp_GetGlobalButtonsFilePath()
    If Len(Dir(filePath)) = 0 Then
        MsgBox "Global buttons config file was not found: " & filePath, vbExclamation
        Exit Function
    End If

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False

    If Not doc.Load(filePath) Then
        MsgBox "Failed to parse global buttons config file: " & filePath, vbExclamation
        Exit Function
    End If

    doc.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"
    Set mp_LoadGlobalButtonsDom = doc
End Function

Private Function mp_GetGlobalButtonsFilePath() As String
    Dim basePath As String

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        basePath = CurDir$
    End If

    mp_GetGlobalButtonsFilePath = basePath & "\" & GLOBAL_BUTTONS_REL_PATH
End Function

Private Sub mp_HideAllButtons(ByVal ws As Worksheet)
    mp_HideAllButtonsInContainer ws.Shapes
End Sub

Private Function mp_IsButtonShapeName(ByVal shapeName As String) As Boolean
    mp_IsButtonShapeName = (LCase$(Left$(Trim$(shapeName), 3)) = "btn")
End Function

Private Function mp_IsManagedUiBlockShape(ByVal shapeName As String) As Boolean
    Dim normalized As String

    normalized = Trim$(shapeName)
    If Len(normalized) = 0 Then Exit Function

    If StrComp(normalized, MODE_DROPDOWN_SHAPE, vbTextCompare) = 0 Then
        mp_IsManagedUiBlockShape = True
        Exit Function
    End If

    If mp_IsButtonShapeName(normalized) Then
        mp_IsManagedUiBlockShape = (StrComp(normalized, UPDATE_BUTTON_SHAPE, vbTextCompare) <> 0)
    End If
End Function

Public Function m_GetShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    If ws Is Nothing Then Exit Function
    Set m_GetShapeByName = mp_FindShapeInContainer(ws.Shapes, shapeName)
End Function

Private Sub mp_HideAllButtonsInContainer(ByVal shapeContainer As Object)
    Dim shp As Shape
    Dim groupItem As Shape

    For Each shp In shapeContainer
        If mp_IsButtonShapeName(shp.Name) Then
            shp.Visible = msoFalse
        End If
        If shp.Type = msoGroup Then
            For Each groupItem In shp.GroupItems
                If mp_IsButtonShapeName(groupItem.Name) Then
                    groupItem.Visible = msoFalse
                End If
            Next groupItem
        End If
    Next shp
End Sub

Private Function mp_FindShapeInContainer(ByVal shapeContainer As Object, ByVal shapeName As String) As Shape
    Dim shp As Shape
    Dim groupItem As Shape
    Dim normalized As String

    normalized = Trim$(shapeName)
    If Len(normalized) = 0 Then Exit Function

    For Each shp In shapeContainer
        If StrComp(shp.Name, normalized, vbTextCompare) = 0 Then
            Set mp_FindShapeInContainer = shp
            Exit Function
        End If
        If shp.Type = msoGroup Then
            For Each groupItem In shp.GroupItems
                If StrComp(groupItem.Name, normalized, vbTextCompare) = 0 Then
                    Set mp_FindShapeInContainer = groupItem
                    Exit Function
                End If
            Next groupItem
        End If
    Next shp
End Function

Private Sub mp_UngroupManagedUiShapes(ByVal ws As Worksheet)
    Dim hasGroupsToUngroup As Boolean
    Dim i As Long
    Dim shp As Shape

    Do
        hasGroupsToUngroup = False

        For i = ws.Shapes.Count To 1 Step -1
            Set shp = ws.Shapes(i)
            If shp.Type = msoGroup Then
                If mp_GroupContainsManagedShapes(shp) Then
                    shp.Ungroup
                    hasGroupsToUngroup = True
                    Exit For
                End If
            End If
        Next i
    Loop While hasGroupsToUngroup
End Sub

Private Function mp_GroupContainsManagedShapes(ByVal groupShape As Shape) As Boolean
    Dim groupItem As Shape

    For Each groupItem In groupShape.GroupItems
        If mp_IsManagedUiBlockShape(groupItem.Name) Then
            mp_GroupContainsManagedShapes = True
            Exit Function
        End If
    Next groupItem
End Function

Private Sub mp_ApplyInitialUiLayout(ByVal ws As Worksheet)
    Dim shp As Shape

    Set shp = m_GetShapeByName(ws, MODE_DROPDOWN_SHAPE)
    If shp Is Nothing Then
        MsgBox "Initial UI layout failed: '" & MODE_DROPDOWN_SHAPE & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    shp.Left = UI_DDMODE_LEFT
    shp.Top = UI_DDMODE_TOP
    shp.Width = UI_DDMODE_WIDTH
    shp.Height = UI_DDMODE_HEIGHT

    Set shp = m_GetShapeByName(ws, CLEAR_BUTTON_SHAPE)
    If shp Is Nothing Then
        MsgBox "Initial UI layout failed: '" & CLEAR_BUTTON_SHAPE & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    shp.Left = UI_CLEAR_LEFT
    shp.Top = UI_CLEAR_TOP
    shp.Width = UI_CLEAR_WIDTH
    shp.Height = UI_CLEAR_HEIGHT

    Set shp = m_GetShapeByName(ws, MODE_BUTTON_SHAPE)
    If shp Is Nothing Then
        MsgBox "Initial UI layout failed: '" & MODE_BUTTON_SHAPE & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    shp.Left = UI_MODEBTN_LEFT
    shp.Top = UI_MODEBTN_TOP
    shp.Width = UI_MODEBTN_WIDTH
    shp.Height = UI_MODEBTN_HEIGHT

    Set shp = m_GetShapeByName(ws, PERSONAL_BUTTON_SHAPE)
    If shp Is Nothing Then
        MsgBox "Initial UI layout failed: '" & PERSONAL_BUTTON_SHAPE & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    shp.Left = UI_PERSONAL_LEFT
    shp.Top = UI_PERSONAL_TOP
    shp.Width = UI_PERSONAL_WIDTH
    shp.Height = UI_PERSONAL_HEIGHT

    Set shp = m_GetShapeByName(ws, COMPARING_BUTTON_SHAPE)
    If shp Is Nothing Then
        MsgBox "Initial UI layout failed: '" & COMPARING_BUTTON_SHAPE & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    shp.Left = UI_COMPARING_LEFT
    shp.Top = UI_COMPARING_TOP
    shp.Width = UI_COMPARING_WIDTH
    shp.Height = UI_COMPARING_HEIGHT
End Sub

Private Function mp_ApplyShapeVisible(ByVal node As Object, ByVal shp As Shape) As Boolean
    Dim valueText As String
    Dim valueBool As Boolean

    valueText = Trim$(mp_NodeAttrText(node, "visible"))
    If Len(valueText) = 0 Then
        shp.Visible = msoFalse
        mp_ApplyShapeVisible = True
        Exit Function
    End If

    If Not mp_TryParseBoolean(valueText, valueBool) Then
        MsgBox "Invalid boolean value for UI attribute 'visible' on shape '" & shp.Name & "': " & valueText, vbExclamation
        Exit Function
    End If

    shp.Visible = IIf(valueBool, msoTrue, msoFalse)
    mp_ApplyShapeVisible = True
End Function

Private Function mp_ApplyShapeGeometry(ByVal node As Object, ByVal shp As Shape) As Boolean
    If Not mp_ApplySingleGeometryAttribute(node, shp, "left") Then Exit Function
    If Not mp_ApplySingleGeometryAttribute(node, shp, "top") Then Exit Function
    If Not mp_ApplySingleGeometryAttribute(node, shp, "width") Then Exit Function
    If Not mp_ApplySingleGeometryAttribute(node, shp, "height") Then Exit Function
    mp_ApplyShapeGeometry = True
End Function

Private Function mp_ApplyShapePlacement(ByVal node As Object, ByVal shp As Shape, ByVal ws As Worksheet) As Boolean
    Dim placementText As String
    Dim placementValue As XlPlacement
    Dim anchorCellText As String
    Dim anchorCell As Range
    Dim dx As Double
    Dim dy As Double

    placementText = Trim$(mp_NodeAttrText(node, "placement"))
    If Len(placementText) > 0 Then
        If Not mp_TryParsePlacement(placementText, placementValue) Then
            MsgBox "Invalid UI placement value on shape '" & shp.Name & "': " & placementText, vbExclamation
            Exit Function
        End If
        shp.Placement = placementValue
    End If

    anchorCellText = Trim$(mp_NodeAttrText(node, "anchorCell"))
    If Len(anchorCellText) = 0 Then
        mp_ApplyShapePlacement = True
        Exit Function
    End If

    On Error GoTo EH_ANCHOR
    Set anchorCell = ws.Range(anchorCellText)
    On Error GoTo 0

    If Not mp_ReadOffset(node, "anchorDx", dx) Then
        MsgBox "Invalid numeric value for UI attribute 'anchorDx' on shape '" & shp.Name & "'.", vbExclamation
        Exit Function
    End If
    If Not mp_ReadOffset(node, "anchorDy", dy) Then
        MsgBox "Invalid numeric value for UI attribute 'anchorDy' on shape '" & shp.Name & "'.", vbExclamation
        Exit Function
    End If

    shp.Left = anchorCell.Left + dx
    shp.Top = anchorCell.Top + dy

    mp_ApplyShapePlacement = True
    Exit Function
EH_ANCHOR:
    MsgBox "Invalid range in UI attribute 'anchorCell' for shape '" & shp.Name & "': " & anchorCellText, vbExclamation
End Function

Private Function mp_ApplySingleGeometryAttribute(ByVal node As Object, ByVal shp As Shape, ByVal attrName As String) As Boolean
    Dim valueText As String
    Dim valueNumber As Double

    valueText = Trim$(mp_NodeAttrText(node, attrName))
    If Len(valueText) = 0 Then
        mp_ApplySingleGeometryAttribute = True
        Exit Function
    End If

    If Not mp_TryParseDouble(valueText, valueNumber) Then
        MsgBox "Invalid numeric value for UI attribute '" & attrName & "' on shape '" & shp.Name & "': " & valueText, vbExclamation
        Exit Function
    End If

    Select Case LCase$(attrName)
        Case "left": shp.Left = valueNumber
        Case "top": shp.Top = valueNumber
        Case "width": shp.Width = valueNumber
        Case "height": shp.Height = valueNumber
    End Select

    mp_ApplySingleGeometryAttribute = True
End Function

Private Function mp_ApplyShapeColor(ByVal node As Object, ByVal shp As Shape, ByVal profileName As String) As Boolean
    Dim valueText As String
    Dim colorValue As Long

    valueText = Trim$(mp_NodeAttrText(node, "backColor"))
    If Len(valueText) = 0 Then
        mp_ApplyShapeColor = True
        Exit Function
    End If

    If Not mp_TryParseColor(valueText, colorValue) Then
        MsgBox "Invalid color value for UI attribute 'backColor' on shape '" & shp.Name & "': " & valueText, vbExclamation
        Exit Function
    End If

    On Error GoTo EH
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = colorValue
    mp_ApplyShapeColor = True
    Exit Function
EH:
    MsgBox "Failed to apply 'backColor' for shape '" & shp.Name & "' in profile '" & profileName & "': " & Err.Description, vbExclamation
End Function

Private Function mp_ReadOffset(ByVal node As Object, ByVal attrName As String, ByRef value As Double) As Boolean
    Dim valueText As String

    valueText = Trim$(mp_NodeAttrText(node, attrName))
    If Len(valueText) = 0 Then
        value = 0#
        mp_ReadOffset = True
        Exit Function
    End If

    mp_ReadOffset = mp_TryParseDouble(valueText, value)
End Function

Private Function mp_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    mp_NodeAttrText = CStr(node.selectSingleNode("@*[local-name()='" & attrName & "']").Text)
    If Err.Number <> 0 Then
        Err.Clear
        mp_NodeAttrText = vbNullString
    End If
    On Error GoTo 0
End Function

Private Function mp_TryParseBoolean(ByVal valueText As String, ByRef result As Boolean) As Boolean
    Select Case LCase$(Trim$(valueText))
        Case "1", "true", "yes"
            result = True
            mp_TryParseBoolean = True
        Case "0", "false", "no"
            result = False
            mp_TryParseBoolean = True
    End Select
End Function

Private Function mp_TryParsePlacement(ByVal valueText As String, ByRef result As XlPlacement) As Boolean
    Select Case LCase$(Trim$(valueText))
        Case "absolute", "free", "freefloating"
            result = xlFreeFloating
            mp_TryParsePlacement = True
        Case "move", "movewithcells"
            result = xlMove
            mp_TryParsePlacement = True
        Case "moveandsize", "move_and_size", "move-size", "moveandresize"
            result = xlMoveAndSize
            mp_TryParsePlacement = True
    End Select
End Function

Private Function mp_TryParseDouble(ByVal valueText As String, ByRef result As Double) As Boolean
    Dim normalized As String
    Dim decSep As String
    Dim altSep As String

    On Error GoTo EH

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function

    decSep = CStr(Application.International(xlDecimalSeparator))
    If decSep = "." Then
        altSep = ","
    Else
        altSep = "."
    End If

    normalized = Replace(normalized, altSep, decSep)
    If Not IsNumeric(normalized) Then Exit Function

    result = CDbl(normalized)
    mp_TryParseDouble = True
    Exit Function
EH:
    mp_TryParseDouble = False
End Function

Private Function mp_TryParseColor(ByVal valueText As String, ByRef colorValue As Long) As Boolean
    Dim hexText As String
    Dim r As Long
    Dim g As Long
    Dim b As Long

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    If Left$(valueText, 1) = "#" And Len(valueText) = 7 Then
        hexText = Mid$(valueText, 2)
        If Not mp_IsHex(hexText) Then Exit Function
        r = CLng("&H" & Mid$(hexText, 1, 2))
        g = CLng("&H" & Mid$(hexText, 3, 2))
        b = CLng("&H" & Mid$(hexText, 5, 2))
        colorValue = RGB(r, g, b)
        mp_TryParseColor = True
        Exit Function
    End If

    If IsNumeric(valueText) Then
        colorValue = CLng(valueText)
        mp_TryParseColor = True
    End If
End Function

Private Function mp_IsHex(ByVal valueText As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(valueText) = 0 Then Exit Function
    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next i
    mp_IsHex = True
End Function
