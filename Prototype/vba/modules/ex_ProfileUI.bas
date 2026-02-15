Attribute VB_Name = "ex_ProfileUI"
Option Explicit

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const GLOBAL_BUTTONS_REL_PATH As String = "config\GlobalButtons.xml"

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

        On Error Resume Next
        Set shp = ws.Shapes(shapeName)
        On Error GoTo 0
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

        On Error Resume Next
        Set shp = ws.Shapes(shapeName)
        On Error GoTo 0
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
    Dim shp As Shape
    For Each shp In ws.Shapes
        If mp_IsButtonShapeName(shp.Name) Then
            shp.Visible = msoFalse
        End If
    Next shp
End Sub

Private Function mp_IsButtonShapeName(ByVal shapeName As String) As Boolean
    mp_IsButtonShapeName = (LCase$(Left$(Trim$(shapeName), 3)) = "btn")
End Function

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
