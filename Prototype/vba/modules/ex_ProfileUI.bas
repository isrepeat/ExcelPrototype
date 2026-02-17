Attribute VB_Name = "ex_ProfileUI"
Option Explicit

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const UI_DEFINITION_REL_PATH As String = "config\UI.xml"
Private Const UI_BLOCK_GROUP_NAME As String = "grpUiBlock"
Private Const PROFILE_DROPDOWN_SHAPE As String = "ddProfile"
Private Const MODE_DROPDOWN_SHAPE As String = "ddMode"
Private Const UPDATE_BUTTON_SHAPE As String = "btnUpdateCode"
Private Const CLEAR_BUTTON_SHAPE As String = "btnClear"
Private Const MODE_BUTTON_SHAPE As String = "btnMode"
Private Const PERSONAL_BUTTON_SHAPE As String = "btnPersonalCard"
Private Const COMPARING_BUTTON_SHAPE As String = "btnComparing"

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

    Set uiNodes = profileNode.selectNodes("p:ui/p:control")
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
            If mp_IsButtonShapeName(shapeName) Then GoTo NextNode
            MsgBox "Profile UI shape '" & shapeName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
            Exit Sub
        End If

        If Not mp_ApplyShapeVisible(node, shp) Then Exit Sub

        Set shp = Nothing
NextNode:
    Next node
End Sub

' Keeps UI controls detached from cell grid so their coordinates stay absolute.
' Managed block: all btn* except btnUpdateCode + ddProfile/ddMode dropdowns.
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
    Dim uiDefDoc As Object
    Dim uiControlNodes As Object
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

    Set uiDefDoc = mp_LoadUiDefinitionDom()
    If uiDefDoc Is Nothing Then Exit Sub

    Set uiControlNodes = uiDefDoc.selectNodes("/p:uiDefinition/p:controls/p:control")
    If uiControlNodes Is Nothing Then
        MsgBox "Invalid UI definition format. Expected '/uiDefinition/controls/control'.", vbExclamation
        Exit Sub
    End If
    mp_ApplyGlobalVisibilityFromUiControls ws, uiControlNodes

    Set uiNodes = profileNode.selectNodes("p:ui/p:control")
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
            GoTo NextNode
        End If

        If mp_IsShapeVisibleByFilters(node) Then
            shp.Visible = msoTrue
        End If
        Set shp = Nothing
NextNode:
    Next node
End Sub

Private Sub mp_ApplyGlobalVisibilityFromUiControls(ByVal ws As Worksheet, ByVal controlNodes As Object)
    Dim node As Object
    Dim shapeName As String
    Dim visibleText As String
    Dim isGlobalVisible As Boolean
    Dim shp As Shape

    For Each node In controlNodes
        shapeName = Trim$(mp_NodeAttrText(node, "name"))
        If Len(shapeName) = 0 Then
            MsgBox "UI control entry contains no 'name' attribute.", vbExclamation
            Exit Sub
        End If
        If Not mp_IsButtonShapeName(shapeName) Then GoTo NextNode

        visibleText = Trim$(mp_NodeAttrText(node, "globalVisible"))
        If Len(visibleText) = 0 Then GoTo NextNode
        If Not mp_TryParseBoolean(visibleText, isGlobalVisible) Then
            MsgBox "Invalid boolean value for UI control attribute 'globalVisible' on '" & shapeName & "': " & visibleText, vbExclamation
            Exit Sub
        End If
        If Not isGlobalVisible Then GoTo NextNode

        Set shp = m_GetShapeByName(ws, shapeName)
        If shp Is Nothing Then GoTo NextNode
        shp.Visible = msoTrue
NextNode:
    Next node
End Sub

Private Function mp_LoadUiDefinitionDom() As Object
    Dim filePath As String
    Dim doc As Object

    filePath = mp_GetUiDefinitionFilePath()
    If Len(Dir(filePath)) = 0 Then
        MsgBox "UI definition config file was not found: " & filePath, vbExclamation
        Exit Function
    End If

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False

    If Not doc.Load(filePath) Then
        MsgBox "Failed to parse UI definition config file: " & filePath, vbExclamation
        Exit Function
    End If

    doc.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"
    Set mp_LoadUiDefinitionDom = doc
End Function

Private Function mp_GetUiDefinitionFilePath() As String
    Dim basePath As String

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        basePath = CurDir$
    End If

    mp_GetUiDefinitionFilePath = basePath & "\" & UI_DEFINITION_REL_PATH
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

    If StrComp(normalized, PROFILE_DROPDOWN_SHAPE, vbTextCompare) = 0 Then
        mp_IsManagedUiBlockShape = True
        Exit Function
    End If

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
