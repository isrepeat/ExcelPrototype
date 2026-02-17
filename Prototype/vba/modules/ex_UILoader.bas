Attribute VB_Name = "ex_UILoader"
Option Explicit

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const UI_CONFIG_REL_PATH As String = "config\UI.xml"
Private Const DEFAULT_SHEET_NAME As String = "Dev"
Private Const UI_BLOCK_GROUP_NAME As String = "grpUiBlock"

Public Sub m_LoadUiFromConfig(Optional ByVal wb As Workbook)
    Dim doc As Object
    Dim controlNodes As Object
    Dim controlNode As Object
    Dim stylesMap As Object
    Dim ws As Worksheet
    Dim controlName As String
    Dim controlType As String
    Dim sheetName As String
    Dim isRequired As Boolean
    Dim createIfMissing As Boolean
    Dim didUngroupUiBlock As Boolean
    Dim regroupSheets As Object
    Dim regroupSheetName As Variant

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    If wb Is Nothing Then
        MsgBox "Failed to load UI from config: workbook is not specified.", vbExclamation
        Exit Sub
    End If

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Sub

    Set controlNodes = doc.selectNodes("/p:uiDefinition/p:controls/p:control")
    If controlNodes Is Nothing Then
        MsgBox "Invalid UI config format. Expected '/uiDefinition/controls/control'.", vbExclamation
        Exit Sub
    End If
    If controlNodes.Length = 0 Then
        MsgBox "UI config is empty: " & mp_GetUiConfigFilePath(wb), vbExclamation
        Exit Sub
    End If

    Set stylesMap = mp_ReadButtonStyles(doc)

    If Not mp_RemoveButtonsMissingInConfig(wb, controlNodes) Then Exit Sub

    Set regroupSheets = CreateObject("Scripting.Dictionary")

    For Each controlNode In controlNodes
        controlName = Trim$(mp_NodeAttrText(controlNode, "name"))
        If Len(controlName) = 0 Then
            MsgBox "UI config contains <control> without 'name' attribute.", vbExclamation
            Exit Sub
        End If

        controlType = LCase$(Trim$(mp_NodeAttrText(controlNode, "type")))
        If Len(controlType) = 0 Then
            controlType = "button"
        End If
        If Not mp_IsSupportedControlType(controlType) Then
            MsgBox "Unsupported UI control type '" & controlType & "' for control '" & controlName & "'.", vbExclamation
            Exit Sub
        End If

        sheetName = Trim$(mp_NodeAttrText(controlNode, "sheet"))
        If Len(sheetName) = 0 Then
            sheetName = DEFAULT_SHEET_NAME
        End If
        Set ws = mp_GetWorksheetByName(wb, sheetName)
        If ws Is Nothing Then
            MsgBox "Sheet '" & sheetName & "' for control '" & controlName & "' was not found in workbook '" & wb.Name & "'.", vbExclamation
            Exit Sub
        End If

        isRequired = mp_NodeAttrBool(controlNode, "required", True)
        createIfMissing = mp_NodeAttrBool(controlNode, "createIfMissing", False)

        If Not mp_EnsureControlShape(ws, controlNode, controlName, controlType, createIfMissing, isRequired) Then
            Exit Sub
        End If
        If ex_ProfileUI.m_GetShapeByName(ws, controlName) Is Nothing Then
            If isRequired Then
                MsgBox "Control '" & controlName & "' is required but was not resolved on sheet '" & ws.Name & "'.", vbExclamation
                Exit Sub
            End If
            GoTo NextControl
        End If

        If Not mp_ApplyControlAttributes(ws, controlNode, controlName, controlType, stylesMap) Then
            Exit Sub
        End If

        If Not mp_AssignShapeMacro(ws, controlNode, controlName, wb, isRequired, didUngroupUiBlock) Then
            Exit Sub
        End If
        If didUngroupUiBlock Then
            If Not regroupSheets.Exists(ws.Name) Then
                regroupSheets.Add ws.Name, True
            End If
        End If
NextControl:
    Next controlNode

    For Each regroupSheetName In regroupSheets.Keys
        Set ws = mp_GetWorksheetByName(wb, CStr(regroupSheetName))
        If Not ws Is Nothing Then
            mp_TryRegroupUiBlock ws
        End If
    Next regroupSheetName
End Sub

Private Function mp_LoadUiDom(ByVal wb As Workbook) As Object
    Dim doc As Object
    Dim filePath As String

    filePath = mp_GetUiConfigFilePath(wb)
    If Len(Dir(filePath)) = 0 Then
        MsgBox "UI config file was not found: " & filePath, vbExclamation
        Exit Function
    End If

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False

    If Not doc.Load(filePath) Then
        MsgBox "Failed to parse UI config file: " & filePath, vbExclamation
        Exit Function
    End If

    doc.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"
    Set mp_LoadUiDom = doc
End Function

Private Function mp_GetUiConfigFilePath(ByVal wb As Workbook) As String
    Dim basePath As String

    basePath = wb.Path
    If Len(basePath) = 0 Then
        basePath = CurDir$
    End If

    mp_GetUiConfigFilePath = basePath & "\" & UI_CONFIG_REL_PATH
End Function

Private Function mp_IsSupportedControlType(ByVal controlType As String) As Boolean
    Select Case LCase$(Trim$(controlType))
        Case "button", "dropdown", "combo"
            mp_IsSupportedControlType = True
    End Select
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set mp_GetWorksheetByName = ws
            Exit Function
        End If
    Next ws
End Function

Private Function mp_EnsureControlShape(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByVal controlType As String, ByVal createIfMissing As Boolean, ByVal isRequired As Boolean) As Boolean
    Dim shp As Shape

    Set shp = ex_ProfileUI.m_GetShapeByName(ws, controlName)
    If Not shp Is Nothing Then
        mp_EnsureControlShape = True
        Exit Function
    End If

    If Not createIfMissing Then
        If isRequired Then
            MsgBox "Control '" & controlName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Else
            mp_EnsureControlShape = True
        End If
        Exit Function
    End If

    Select Case LCase$(controlType)
        Case "button"
            If Not mp_TryCreateMissingButton(ws, controlNode, controlName, shp) Then Exit Function
        Case "dropdown", "combo"
            If Not mp_TryCreateMissingDropdown(ws, controlNode, controlName, shp) Then Exit Function
        Case Else
            If isRequired Then
                MsgBox "Auto-create is not supported for type='" & controlType & "'. Control '" & controlName & "' cannot be created.", vbExclamation
            Else
                mp_EnsureControlShape = True
            End If
            Exit Function
    End Select

    mp_EnsureControlShape = True
End Function

Private Function mp_TryCreateMissingButton(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByRef createdShape As Shape) As Boolean
    Dim captionText As String
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim templateName As String
    Dim templateShape As Shape

    captionText = Trim$(mp_NodeAttrText(controlNode, "caption"))
    If Len(captionText) = 0 Then
        captionText = controlName
    End If

    If Not mp_ReadRequiredControlRect(controlNode, controlName, leftPos, topPos, widthVal, heightVal) Then Exit Function

    On Error GoTo EH_CREATE
    Set createdShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)

    templateName = Trim$(mp_NodeAttrText(controlNode, "template"))
    If Len(templateName) > 0 Then
        Set templateShape = ex_ProfileUI.m_GetShapeByName(ws, templateName)
    End If

    If Not templateShape Is Nothing Then
        On Error Resume Next
        templateShape.PickUp
        createdShape.Apply
        On Error GoTo 0
    End If

    createdShape.Left = leftPos
    createdShape.Top = topPos
    createdShape.Width = widthVal
    createdShape.Height = heightVal
    createdShape.Name = controlName
    createdShape.TextFrame.Characters.Text = captionText
    createdShape.Placement = xlFreeFloating
    mp_TryCreateMissingButton = True
    Exit Function

EH_CREATE:
    MsgBox "Failed to auto-create button control '" & controlName & "': " & Err.Description, vbExclamation
End Function

Private Function mp_TryCreateMissingDropdown(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByRef createdShape As Shape) As Boolean
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim dd As DropDown

    If Not mp_ReadRequiredControlRect(controlNode, controlName, leftPos, topPos, widthVal, heightVal) Then Exit Function

    On Error GoTo EH_CREATE
    Set dd = ws.DropDowns.Add(leftPos, topPos, widthVal, heightVal)
    dd.Name = controlName
    Set createdShape = ex_ProfileUI.m_GetShapeByName(ws, controlName)
    If createdShape Is Nothing Then
        MsgBox "Dropdown control '" & controlName & "' was created but shape lookup failed on sheet '" & ws.Name & "'.", vbExclamation
        Exit Function
    End If
    mp_TryCreateMissingDropdown = True
    Exit Function
EH_CREATE:
    MsgBox "Failed to auto-create dropdown control '" & controlName & "': " & Err.Description, vbExclamation
End Function

Private Function mp_ReadRequiredControlRect(ByVal controlNode As Object, ByVal controlName As String, ByRef leftPos As Double, ByRef topPos As Double, ByRef widthVal As Double, ByRef heightVal As Double) As Boolean
    If Not mp_ReadRequiredNumber(controlNode, "left", leftPos) Then
        MsgBox "Control '" & controlName & "' with createIfMissing='true' must define numeric 'left' in UI.xml.", vbExclamation
        Exit Function
    End If
    If Not mp_ReadRequiredNumber(controlNode, "top", topPos) Then
        MsgBox "Control '" & controlName & "' with createIfMissing='true' must define numeric 'top' in UI.xml.", vbExclamation
        Exit Function
    End If
    If Not mp_ReadRequiredNumber(controlNode, "width", widthVal) Then
        MsgBox "Control '" & controlName & "' with createIfMissing='true' must define numeric 'width' in UI.xml.", vbExclamation
        Exit Function
    End If
    If Not mp_ReadRequiredNumber(controlNode, "height", heightVal) Then
        MsgBox "Control '" & controlName & "' with createIfMissing='true' must define numeric 'height' in UI.xml.", vbExclamation
        Exit Function
    End If

    mp_ReadRequiredControlRect = True
End Function

Private Function mp_ReadRequiredNumber(ByVal node As Object, ByVal attrName As String, ByRef result As Double) As Boolean
    Dim valueText As String

    valueText = Trim$(mp_NodeAttrText(node, attrName))
    If Len(valueText) = 0 Then Exit Function
    mp_ReadRequiredNumber = mp_TryParseDouble(valueText, result)
End Function

Private Function mp_ApplyControlAttributes(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByVal controlType As String, ByVal stylesMap As Object) As Boolean
    Dim shp As Shape
    Dim styleName As String

    Set shp = ex_ProfileUI.m_GetShapeByName(ws, controlName)
    If shp Is Nothing Then
        MsgBox "Control '" & controlName & "' was not found while applying attributes on sheet '" & ws.Name & "'.", vbExclamation
        Exit Function
    End If

    If Not mp_ApplyShapeVisible(controlNode, shp) Then Exit Function
    If Not mp_ApplyShapePlacement(controlNode, shp, ws) Then Exit Function
    If Not mp_ApplyShapeGeometry(controlNode, shp) Then Exit Function

    If StrComp(controlType, "button", vbTextCompare) = 0 Then
        If Not mp_ApplyButtonCaption(controlNode, shp) Then Exit Function

        styleName = Trim$(mp_NodeAttrText(controlNode, "style"))
        If Len(styleName) > 0 Then
            If Not mp_ApplyButtonStyleByName(shp, styleName, stylesMap) Then Exit Function
        End If
    ElseIf StrComp(controlType, "dropdown", vbTextCompare) = 0 Or StrComp(controlType, "combo", vbTextCompare) = 0 Then
        If Not mp_ApplyDropdownItems(controlNode, shp, controlName) Then Exit Function
    End If

    mp_ApplyControlAttributes = True
End Function

Private Function mp_ApplyButtonCaption(ByVal node As Object, ByVal shp As Shape) As Boolean
    Dim captionText As String

    captionText = mp_NodeAttrText(node, "caption")
    If Len(captionText) = 0 Then
        mp_ApplyButtonCaption = True
        Exit Function
    End If

    On Error GoTo EH
    shp.TextFrame.Characters.Text = captionText
    mp_ApplyButtonCaption = True
    Exit Function
EH:
    MsgBox "Failed to apply caption for control '" & shp.Name & "': " & Err.Description, vbExclamation
End Function

Private Function mp_ApplyDropdownItems(ByVal node As Object, ByVal shp As Shape, ByVal controlName As String) As Boolean
    Dim sourceName As String
    Dim items As Variant
    Dim itemsNode As Object
    Dim itemNodes As Object
    Dim itemNode As Object
    Dim cf As Object
    Dim selectedText As String
    Dim selectedIndex As Long
    Dim itemText As String

    On Error Resume Next
    Set cf = shp.ControlFormat
    On Error GoTo 0
    If cf Is Nothing Then
        MsgBox "Control '" & controlName & "' is not a valid dropdown/combo Form Control.", vbExclamation
        Exit Function
    End If

    sourceName = Trim$(mp_NodeAttrText(node, "itemsSource"))
    If Len(sourceName) > 0 Then
        items = m_GetDropdownItemsByName(controlName, ThisWorkbook)
        If Not mp_ArrayHasItems(items) Then
            MsgBox "Control '" & controlName & "' did not resolve any items from source '" & sourceName & "'.", vbExclamation
            Exit Function
        End If

        On Error GoTo EH_CLEAR
        cf.RemoveAllItems
        On Error GoTo 0

        For selectedIndex = LBound(items) To UBound(items)
            itemText = CStr(items(selectedIndex))
            On Error GoTo EH_ADD
            cf.AddItem itemText
            On Error GoTo 0
        Next selectedIndex
        GoTo ApplySelection
    End If

    Set itemsNode = node.selectSingleNode("p:items")
    If itemsNode Is Nothing Then
        mp_ApplyDropdownItems = True
        Exit Function
    End If

    Set itemNodes = node.selectNodes("p:items/p:item")
    If itemNodes Is Nothing Then
        MsgBox "UI config control '" & controlName & "' contains <items> but no valid <item> entries.", vbExclamation
        Exit Function
    End If

    On Error GoTo EH_CLEAR
    cf.RemoveAllItems
    On Error GoTo 0

    For Each itemNode In itemNodes
        itemText = Trim$(mp_NodeAttrText(itemNode, "value"))
        If Len(itemText) = 0 Then
            itemText = Trim$(CStr(itemNode.Text))
        End If
        If Len(itemText) = 0 Then
            MsgBox "UI config control '" & controlName & "' contains empty <item>.", vbExclamation
            Exit Function
        End If

        On Error GoTo EH_ADD
        cf.AddItem itemText
        On Error GoTo 0
    Next itemNode

ApplySelection:
    selectedText = Trim$(mp_NodeAttrText(node, "selectedItem"))
    If Len(selectedText) > 0 Then
        selectedIndex = mp_FindDropdownItemIndex(cf, selectedText)
        If selectedIndex = 0 Then
            MsgBox "UI config selectedItem '" & selectedText & "' for control '" & controlName & "' was not found in its <items> list.", vbExclamation
            Exit Function
        End If
        On Error Resume Next
        cf.Value = selectedIndex
        On Error GoTo 0
    End If

    mp_ApplyDropdownItems = True
    Exit Function
EH_CLEAR:
    MsgBox "Failed to clear dropdown items for control '" & controlName & "': " & Err.Description, vbExclamation
    Exit Function
EH_ADD:
    MsgBox "Failed to add dropdown item '" & itemText & "' for control '" & controlName & "': " & Err.Description, vbExclamation
End Function

Private Function mp_ArrayHasItems(ByVal values As Variant) As Boolean
    On Error GoTo EH
    If IsArray(values) Then
        mp_ArrayHasItems = (UBound(values) >= LBound(values))
    End If
    Exit Function
EH:
    mp_ArrayHasItems = False
End Function

Private Function mp_FindDropdownItemIndex(ByVal cf As Object, ByVal itemText As String) As Long
    Dim i As Long
    Dim itemCount As Long

    On Error Resume Next
    itemCount = CLng(cf.ListCount)
    On Error GoTo 0
    If itemCount <= 0 Then Exit Function

    For i = 1 To itemCount
        On Error Resume Next
        If StrComp(CStr(cf.List(i)), itemText, vbTextCompare) = 0 Then
            On Error GoTo 0
            mp_FindDropdownItemIndex = i
            Exit Function
        End If
        On Error GoTo 0
    Next i
End Function

Public Function m_GetDropdownItemsByName(ByVal controlName As String, Optional ByVal wb As Workbook, Optional ByVal modeName As String = vbNullString) As Variant
    Dim doc As Object
    Dim controlNode As Object
    Dim sourceName As String
    Dim itemNodes As Object
    Dim itemNode As Object
    Dim items() As String
    Dim idx As Long
    Dim itemText As String

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, controlName)
    If controlNode Is Nothing Then Exit Function

    sourceName = Trim$(mp_NodeAttrText(controlNode, "itemsSource"))
    If Len(sourceName) > 0 Then
        m_GetDropdownItemsByName = mp_GetItemsFromSource(doc, sourceName, modeName, wb)
        Exit Function
    End If

    Set itemNodes = controlNode.selectNodes("p:items/p:item")
    If itemNodes Is Nothing Then Exit Function
    If itemNodes.Length = 0 Then Exit Function

    ReDim items(0 To itemNodes.Length - 1)
    idx = 0
    For Each itemNode In itemNodes
        itemText = Trim$(mp_NodeAttrText(itemNode, "value"))
        If Len(itemText) = 0 Then
            itemText = Trim$(CStr(itemNode.Text))
        End If
        If Len(itemText) = 0 Then
            MsgBox "UI config control '" & controlName & "' contains empty <item>.", vbExclamation
            Exit Function
        End If
        items(idx) = itemText
        idx = idx + 1
    Next itemNode

    m_GetDropdownItemsByName = items
End Function

Public Function m_GetProfilesFilePathByMode(Optional ByVal modeName As String = vbNullString, Optional ByVal wb As Workbook, Optional ByVal sourceName As String = "profilesByMode") As String
    Dim doc As Object
    Dim sourceNode As Object
    Dim relPath As String
    Dim basePath As String

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set sourceNode = mp_GetProfilesSourceNode(doc, sourceName)
    If sourceNode Is Nothing Then Exit Function

    relPath = mp_ResolveProfilesSourceRelPath(sourceNode, modeName)
    If Len(relPath) = 0 Then Exit Function

    basePath = wb.Path
    If Len(basePath) = 0 Then
        basePath = CurDir$
    End If

    m_GetProfilesFilePathByMode = basePath & "\" & relPath
End Function

Public Function m_GetModeVariantsByControl(ByVal controlName As String, Optional ByVal wb As Workbook) As Variant
    Dim doc As Object
    Dim controlNode As Object
    Dim variantNodes As Object
    Dim variantNode As Object
    Dim variants() As Variant
    Dim idx As Long
    Dim valueText As String
    Dim captionText As String
    Dim styleText As String
    Dim displayText As String

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, controlName)
    If controlNode Is Nothing Then Exit Function

    Set variantNodes = controlNode.selectNodes("p:modeVariants/p:variant")
    If variantNodes Is Nothing Then
        m_GetModeVariantsByControl = Array()
        Exit Function
    End If
    If variantNodes.Length = 0 Then
        m_GetModeVariantsByControl = Array()
        Exit Function
    End If

    ReDim variants(1 To variantNodes.Length, 1 To 4)
    idx = 1
    For Each variantNode In variantNodes
        valueText = Trim$(mp_NodeAttrText(variantNode, "value"))
        If Len(valueText) = 0 Then
            MsgBox "Control '" & controlName & "' contains mode variant without 'value'.", vbExclamation
            Exit Function
        End If
        If Not IsNumeric(valueText) Then
            MsgBox "Control '" & controlName & "' has non-numeric mode variant value: " & valueText, vbExclamation
            Exit Function
        End If

        captionText = Trim$(mp_NodeAttrText(variantNode, "caption"))
        styleText = Trim$(mp_NodeAttrText(variantNode, "style"))
        displayText = Trim$(mp_NodeAttrText(variantNode, "display"))

        variants(idx, 1) = CLng(valueText)
        variants(idx, 2) = captionText
        variants(idx, 3) = styleText
        variants(idx, 4) = displayText
        idx = idx + 1
    Next variantNode

    m_GetModeVariantsByControl = variants
End Function

Public Function m_GetControlAttribute(ByVal controlName As String, ByVal attrName As String, Optional ByVal wb As Workbook) As String
    Dim doc As Object
    Dim controlNode As Object

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, controlName)
    If controlNode Is Nothing Then Exit Function

    m_GetControlAttribute = Trim$(mp_NodeAttrText(controlNode, attrName))
End Function

Public Function m_ApplyControlStyleByName(ByVal ws As Worksheet, ByVal controlName As String, ByVal styleName As String, Optional ByVal wb As Workbook) As Boolean
    Dim doc As Object
    Dim stylesMap As Object
    Dim shp As Shape

    If ws Is Nothing Then
        MsgBox "Failed to apply control style: worksheet is not specified.", vbExclamation
        Exit Function
    End If

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    If wb Is Nothing Then Exit Function

    Set shp = ex_ProfileUI.m_GetShapeByName(ws, controlName)
    If shp Is Nothing Then
        MsgBox "Failed to apply control style: control '" & controlName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Function
    End If

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set stylesMap = mp_ReadButtonStyles(doc)
    If stylesMap Is Nothing Then Exit Function

    m_ApplyControlStyleByName = mp_ApplyButtonStyleByName(shp, styleName, stylesMap)
End Function

Private Function mp_GetControlNode(ByVal doc As Object, ByVal controlName As String) As Object
    Set mp_GetControlNode = doc.selectSingleNode("/p:uiDefinition/p:controls/p:control[@name=" & mp_XPathLiteral(controlName) & "]")
    If mp_GetControlNode Is Nothing Then
        MsgBox "Control '" & controlName & "' was not found in UI config.", vbExclamation
    End If
End Function

Private Function mp_GetItemsFromSource(ByVal doc As Object, ByVal sourceName As String, ByVal modeName As String, ByVal wb As Workbook) As Variant
    Dim sourceNode As Object
    Dim relPath As String
    Dim filePath As String
    Dim srcDoc As Object
    Dim profileNodes As Object
    Dim names() As String
    Dim i As Long

    Set sourceNode = mp_GetProfilesSourceNode(doc, sourceName)
    If sourceNode Is Nothing Then Exit Function

    relPath = mp_ResolveProfilesSourceRelPath(sourceNode, modeName)
    If Len(relPath) = 0 Then Exit Function

    filePath = mp_CombineBasePath(wb, relPath)
    If Len(Dir(filePath)) = 0 Then
        MsgBox "Profiles source file was not found: " & filePath, vbExclamation
        Exit Function
    End If

    Set srcDoc = CreateObject("MSXML2.DOMDocument.6.0")
    srcDoc.async = False
    srcDoc.validateOnParse = False
    If Not srcDoc.Load(filePath) Then
        MsgBox "Failed to parse profiles source file: " & filePath, vbExclamation
        Exit Function
    End If
    srcDoc.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"

    Set profileNodes = srcDoc.selectNodes("/p:presets/p:profile")
    If profileNodes Is Nothing Then Exit Function
    If profileNodes.Length = 0 Then
        mp_GetItemsFromSource = Array()
        Exit Function
    End If

    ReDim names(0 To profileNodes.Length - 1)
    For i = 0 To profileNodes.Length - 1
        names(i) = CStr(profileNodes.Item(i).getAttribute("name"))
    Next i

    mp_GetItemsFromSource = names
End Function

Private Function mp_GetProfilesSourceNode(ByVal doc As Object, ByVal sourceName As String) As Object
    Set mp_GetProfilesSourceNode = doc.selectSingleNode("/p:uiDefinition/p:dataSources/p:profilesSource[@name=" & mp_XPathLiteral(sourceName) & "]")
    If mp_GetProfilesSourceNode Is Nothing Then
        MsgBox "Profiles source '" & sourceName & "' was not found in UI config.", vbExclamation
    End If
End Function

Private Function mp_ResolveProfilesSourceRelPath(ByVal sourceNode As Object, ByVal modeName As String) As String
    Dim modePersonal As String
    Dim modeComparing As String
    Dim pathPersonal As String
    Dim pathComparing As String
    Dim defaultMode As String

    modePersonal = Trim$(mp_NodeAttrText(sourceNode, "modePersonalCard"))
    modeComparing = Trim$(mp_NodeAttrText(sourceNode, "modeComparing"))
    pathPersonal = Trim$(mp_NodeAttrText(sourceNode, "pathPersonalCard"))
    pathComparing = Trim$(mp_NodeAttrText(sourceNode, "pathComparing"))
    defaultMode = Trim$(mp_NodeAttrText(sourceNode, "defaultMode"))

    If Len(modePersonal) = 0 Or Len(modeComparing) = 0 Then
        MsgBox "Profiles source is missing required mode labels: modePersonalCard/modeComparing.", vbExclamation
        Exit Function
    End If
    If Len(pathPersonal) = 0 Or Len(pathComparing) = 0 Then
        MsgBox "Profiles source is missing required paths: pathPersonalCard/pathComparing.", vbExclamation
        Exit Function
    End If

    modeName = Trim$(modeName)
    If Len(modeName) = 0 Then
        modeName = defaultMode
    End If
    If Len(modeName) = 0 Then
        modeName = modePersonal
    End If

    If StrComp(modeName, modeComparing, vbTextCompare) = 0 Then
        mp_ResolveProfilesSourceRelPath = pathComparing
        Exit Function
    End If
    If StrComp(modeName, modePersonal, vbTextCompare) = 0 Then
        mp_ResolveProfilesSourceRelPath = pathPersonal
        Exit Function
    End If

    MsgBox "Invalid mode '" & modeName & "' for profiles source. Allowed values: '" & modePersonal & "', '" & modeComparing & "'.", vbExclamation
End Function

Private Function mp_CombineBasePath(ByVal wb As Workbook, ByVal relPath As String) As String
    Dim basePath As String

    basePath = wb.Path
    If Len(basePath) = 0 Then
        basePath = CurDir$
    End If

    mp_CombineBasePath = basePath & "\" & relPath
End Function

Private Function mp_XPathLiteral(ByVal value As String) As String
    If InStr(value, "'") = 0 Then
        mp_XPathLiteral = "'" & value & "'"
        Exit Function
    End If

    If InStr(value, """") = 0 Then
        mp_XPathLiteral = """" & value & """"
        Exit Function
    End If

    Dim parts() As String
    Dim i As Long

    parts = Split(value, "'")
    mp_XPathLiteral = "concat("
    For i = 0 To UBound(parts)
        If i > 0 Then
            mp_XPathLiteral = mp_XPathLiteral & ", ""'"" , "
        End If
        mp_XPathLiteral = mp_XPathLiteral & "'" & parts(i) & "'"
    Next i
    mp_XPathLiteral = mp_XPathLiteral & ")"
End Function

Private Function mp_AssignShapeMacro(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByVal wb As Workbook, ByVal isRequired As Boolean, ByRef didUngroupUiBlock As Boolean) As Boolean
    Dim shp As Shape
    Dim macroName As String
    Dim onActionText As String
    Dim assignmentErr As String
    Dim parentGroupName As String

    macroName = Trim$(mp_NodeAttrText(controlNode, "macro"))
    didUngroupUiBlock = False

    If Len(macroName) = 0 Then
        If isRequired Then
            MsgBox "UI config control '" & controlName & "' does not define required attribute 'macro'.", vbExclamation
        Else
            mp_AssignShapeMacro = True
        End If
        Exit Function
    End If

    onActionText = "'" & wb.Name & "'!" & macroName
    Set shp = ex_ProfileUI.m_GetShapeByName(ws, controlName)
    If shp Is Nothing Then
        If isRequired Then
            MsgBox "Control '" & controlName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Else
            mp_AssignShapeMacro = True
        End If
        Exit Function
    End If

    On Error GoTo EH_ASSIGN_FIRST
    shp.OnAction = onActionText
    mp_AssignShapeMacro = True
    Exit Function

EH_ASSIGN_FIRST:
    assignmentErr = Err.Description
    Err.Clear

    If mp_TryUngroupParentShape(shp, parentGroupName) Then
        If StrComp(parentGroupName, UI_BLOCK_GROUP_NAME, vbTextCompare) = 0 Then
            didUngroupUiBlock = True
        End If

        Set shp = ex_ProfileUI.m_GetShapeByName(ws, controlName)
        If shp Is Nothing Then
            MsgBox "Control '" & controlName & "' is unavailable after ungroup on sheet '" & ws.Name & "'.", vbExclamation
            Exit Function
        End If

        On Error GoTo EH_ASSIGN_SECOND
        shp.OnAction = onActionText
        mp_AssignShapeMacro = True
        Exit Function
    End If

    If isRequired Then
        MsgBox "Failed to assign macro '" & macroName & "' to control '" & controlName & "': " & assignmentErr, vbExclamation
    Else
        mp_AssignShapeMacro = True
    End If
    Exit Function

EH_ASSIGN_SECOND:
    If isRequired Then
        MsgBox "Failed to assign macro '" & macroName & "' to control '" & controlName & "' after ungroup retry: " & Err.Description, vbExclamation
    Else
        mp_AssignShapeMacro = True
    End If
End Function

Private Function mp_TryUngroupParentShape(ByVal shp As Shape, ByRef parentGroupName As String) As Boolean
    Dim parentGroup As Shape

    parentGroupName = vbNullString
    On Error Resume Next
    Set parentGroup = shp.ParentGroup
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If parentGroup Is Nothing Then Exit Function

    parentGroupName = parentGroup.Name
    On Error GoTo EH_UNGROUP
    parentGroup.Ungroup
    mp_TryUngroupParentShape = True
    Exit Function

EH_UNGROUP:
    MsgBox "Failed to ungroup parent group for control '" & shp.Name & "': " & Err.Description, vbExclamation
End Function

Private Sub mp_TryRegroupUiBlock(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error GoTo EH_REGROUP
    ex_ProfileUI.m_InitUiBlockLayoutAndGroup ws
    Exit Sub
EH_REGROUP:
    MsgBox "Failed to regroup UI block '" & UI_BLOCK_GROUP_NAME & "' on sheet '" & ws.Name & "': " & Err.Description, vbExclamation
End Sub

Private Function mp_ReadButtonStyles(ByVal doc As Object) As Object
    Dim styleNodes As Object
    Dim styleNode As Object
    Dim styleName As String
    Dim stylesMap As Object
    Dim styleData As Object

    Set stylesMap = CreateObject("Scripting.Dictionary")

    Set styleNodes = doc.selectNodes("/p:uiDefinition/p:styles/p:buttonStyle")
    If styleNodes Is Nothing Then
        Set mp_ReadButtonStyles = stylesMap
        Exit Function
    End If

    For Each styleNode In styleNodes
        styleName = Trim$(mp_NodeAttrText(styleNode, "name"))
        If Len(styleName) = 0 Then
            MsgBox "UI config contains <buttonStyle> without 'name' attribute.", vbExclamation
            Exit Function
        End If

        Set styleData = CreateObject("Scripting.Dictionary")
        mp_SetStyleValue styleData, "backColor", mp_NodeAttrText(styleNode, "backColor")
        mp_SetStyleValue styleData, "textColor", mp_NodeAttrText(styleNode, "textColor")
        mp_SetStyleValue styleData, "borderColor", mp_NodeAttrText(styleNode, "borderColor")
        mp_SetStyleValue styleData, "borderWeight", mp_NodeAttrText(styleNode, "borderWeight")
        mp_SetStyleValue styleData, "fontName", mp_NodeAttrText(styleNode, "fontName")
        mp_SetStyleValue styleData, "fontSize", mp_NodeAttrText(styleNode, "fontSize")
        mp_SetStyleValue styleData, "fontBold", mp_NodeAttrText(styleNode, "fontBold")

        Set stylesMap(styleName) = styleData
    Next styleNode

    Set mp_ReadButtonStyles = stylesMap
End Function

Private Sub mp_SetStyleValue(ByVal styleData As Object, ByVal key As String, ByVal value As String)
    value = Trim$(value)
    If Len(value) = 0 Then Exit Sub
    styleData(key) = value
End Sub

Private Function mp_ApplyButtonStyleByName(ByVal shp As Shape, ByVal styleName As String, ByVal stylesMap As Object) As Boolean
    Dim styleData As Object
    Dim colorValue As Long
    Dim numberValue As Double
    Dim boolValue As Boolean

    If stylesMap Is Nothing Then
        MsgBox "UI style '" & styleName & "' is referenced by '" & shp.Name & "', but styles map is unavailable.", vbExclamation
        Exit Function
    End If
    If Not stylesMap.Exists(styleName) Then
        MsgBox "UI style '" & styleName & "' is referenced by '" & shp.Name & "' but not defined in /uiDefinition/styles.", vbExclamation
        Exit Function
    End If

    Set styleData = stylesMap(styleName)

    On Error GoTo EH
    If styleData.Exists("backColor") Then
        If Not mp_TryParseColor(CStr(styleData("backColor")), colorValue) Then
            MsgBox "Invalid style backColor for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.Fill.Visible = msoTrue
        shp.Fill.ForeColor.RGB = colorValue
    End If

    If styleData.Exists("textColor") Then
        If Not mp_TryParseColor(CStr(styleData("textColor")), colorValue) Then
            MsgBox "Invalid style textColor for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.TextFrame.Characters.Font.Color = colorValue
    End If

    If styleData.Exists("borderColor") Then
        If Not mp_TryParseColor(CStr(styleData("borderColor")), colorValue) Then
            MsgBox "Invalid style borderColor for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = colorValue
    End If

    If styleData.Exists("borderWeight") Then
        If Not mp_TryParseDouble(CStr(styleData("borderWeight")), numberValue) Then
            MsgBox "Invalid style borderWeight for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.Line.Visible = msoTrue
        shp.Line.Weight = numberValue
    End If

    If styleData.Exists("fontName") Then
        shp.TextFrame.Characters.Font.Name = CStr(styleData("fontName"))
    End If

    If styleData.Exists("fontSize") Then
        If Not mp_TryParseDouble(CStr(styleData("fontSize")), numberValue) Then
            MsgBox "Invalid style fontSize for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.TextFrame.Characters.Font.Size = numberValue
    End If

    If styleData.Exists("fontBold") Then
        If Not mp_TryParseBoolean(CStr(styleData("fontBold")), boolValue) Then
            MsgBox "Invalid style fontBold for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.TextFrame.Characters.Font.Bold = boolValue
    End If

    mp_ApplyButtonStyleByName = True
    Exit Function
EH:
    MsgBox "Failed to apply style '" & styleName & "' to control '" & shp.Name & "': " & Err.Description, vbExclamation
End Function

Private Function mp_ApplyShapeVisible(ByVal node As Object, ByVal shp As Shape) As Boolean
    Dim valueText As String
    Dim valueBool As Boolean

    valueText = Trim$(mp_NodeAttrText(node, "visible"))
    If Len(valueText) = 0 Then
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

Private Function mp_ApplyShapeGeometry(ByVal node As Object, ByVal shp As Shape) As Boolean
    If Not mp_ApplySingleGeometryAttribute(node, shp, "left") Then Exit Function
    If Not mp_ApplySingleGeometryAttribute(node, shp, "top") Then Exit Function
    If Not mp_ApplySingleGeometryAttribute(node, shp, "width") Then Exit Function
    If Not mp_ApplySingleGeometryAttribute(node, shp, "height") Then Exit Function
    mp_ApplyShapeGeometry = True
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

Private Function mp_RemoveButtonsMissingInConfig(ByVal wb As Workbook, ByVal controlNodes As Object) As Boolean
    Dim allowedButtons As Object
    Dim ws As Worksheet
    Dim existingButtons As Object
    Dim controlNode As Object
    Dim controlName As String
    Dim controlType As String
    Dim sheetName As String
    Dim buttonName As Variant

    Set allowedButtons = CreateObject("Scripting.Dictionary")

    For Each controlNode In controlNodes
        controlName = Trim$(mp_NodeAttrText(controlNode, "name"))
        controlType = LCase$(Trim$(mp_NodeAttrText(controlNode, "type")))
        If Len(controlType) = 0 Then
            controlType = "button"
        End If
        If StrComp(controlType, "button", vbTextCompare) <> 0 Then GoTo NextNode
        If Not mp_IsButtonShapeName(controlName) Then GoTo NextNode

        sheetName = Trim$(mp_NodeAttrText(controlNode, "sheet"))
        If Len(sheetName) = 0 Then
            sheetName = DEFAULT_SHEET_NAME
        End If
        allowedButtons(mp_SheetButtonKey(sheetName, controlName)) = True
NextNode:
    Next controlNode

    For Each ws In wb.Worksheets
        Set existingButtons = CreateObject("Scripting.Dictionary")
        mp_CollectButtonShapeNamesInContainer ws.Shapes, existingButtons

        For Each buttonName In existingButtons.Keys
            If Not allowedButtons.Exists(mp_SheetButtonKey(ws.Name, CStr(buttonName))) Then
                If Not mp_DeleteShapeByName(ws, CStr(buttonName)) Then Exit Function
            End If
        Next buttonName
    Next ws

    mp_RemoveButtonsMissingInConfig = True
End Function

Private Sub mp_CollectButtonShapeNamesInContainer(ByVal shapeContainer As Object, ByVal names As Object)
    Dim shp As Shape
    Dim groupItem As Shape

    For Each shp In shapeContainer
        If mp_IsButtonShapeName(shp.Name) Then
            names(shp.Name) = True
        End If

        If shp.Type = msoGroup Then
            For Each groupItem In shp.GroupItems
                If mp_IsButtonShapeName(groupItem.Name) Then
                    names(groupItem.Name) = True
                End If
                If groupItem.Type = msoGroup Then
                    mp_CollectButtonShapeNamesInContainer groupItem.GroupItems, names
                End If
            Next groupItem
        End If
    Next shp
End Sub

Private Function mp_DeleteShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    Dim shp As Shape
    Dim parentGroupName As String

    Set shp = ex_ProfileUI.m_GetShapeByName(ws, shapeName)
    If shp Is Nothing Then
        mp_DeleteShapeByName = True
        Exit Function
    End If

    On Error GoTo EH_DELETE
    shp.Delete
    mp_DeleteShapeByName = True
    Exit Function
EH_DELETE:
    Err.Clear

    If mp_TryUngroupParentShape(shp, parentGroupName) Then
        Set shp = ex_ProfileUI.m_GetShapeByName(ws, shapeName)
        If shp Is Nothing Then
            mp_DeleteShapeByName = True
            Exit Function
        End If

        On Error GoTo EH_DELETE_RETRY
        shp.Delete
        mp_DeleteShapeByName = True
        Exit Function
    End If

    MsgBox "Failed to delete UI shape '" & shapeName & "' from sheet '" & ws.Name & "'.", vbExclamation
    Exit Function
EH_DELETE_RETRY:
    MsgBox "Failed to delete UI shape '" & shapeName & "' from sheet '" & ws.Name & "' after ungroup retry: " & Err.Description, vbExclamation
End Function

Private Function mp_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    mp_NodeAttrText = CStr(node.Attributes.getNamedItem(attrName).Text)
    If Err.Number <> 0 Then
        Err.Clear
        mp_NodeAttrText = vbNullString
    End If
    On Error GoTo 0
End Function

Private Function mp_NodeAttrBool(ByVal node As Object, ByVal attrName As String, ByVal defaultValue As Boolean) As Boolean
    Dim valueText As String

    valueText = LCase$(Trim$(mp_NodeAttrText(node, attrName)))
    If Len(valueText) = 0 Then
        mp_NodeAttrBool = defaultValue
        Exit Function
    End If

    Select Case valueText
        Case "true", "1", "yes"
            mp_NodeAttrBool = True
        Case "false", "0", "no"
            mp_NodeAttrBool = False
        Case Else
            mp_NodeAttrBool = defaultValue
    End Select
End Function

Private Function mp_IsButtonShapeName(ByVal shapeName As String) As Boolean
    mp_IsButtonShapeName = (LCase$(Left$(Trim$(shapeName), 3)) = "btn")
End Function

Private Function mp_SheetButtonKey(ByVal sheetName As String, ByVal controlName As String) As String
    mp_SheetButtonKey = LCase$(Trim$(sheetName)) & "|" & LCase$(Trim$(controlName))
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
