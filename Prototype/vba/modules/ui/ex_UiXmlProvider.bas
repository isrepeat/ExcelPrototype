Attribute VB_Name = "ex_UiXmlProvider"
Option Explicit

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const UI_CONFIG_REL_PATH As String = "config\UI.xml"

Public Function m_GetDropdownItemsByName(ByVal controlName As String, Optional ByVal wb As Workbook, Optional ByVal modeName As String = vbNullString) As Variant
    Dim doc As Object
    Dim controlNode As Object
    Dim sourceName As String
    Dim itemNodes As Object
    Dim itemNode As Object
    Dim items() As String
    Dim idx As Long
    Dim itemText As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, controlName)
    If controlNode Is Nothing Then Exit Function

    sourceName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource"))
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
        itemText = Trim$(ex_XmlCore.m_NodeAttrText(itemNode, "value"))
        If Len(itemText) = 0 Then itemText = Trim$(CStr(itemNode.Text))
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

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set sourceNode = mp_GetProfilesSourceNode(doc, sourceName)
    If sourceNode Is Nothing Then Exit Function

    relPath = mp_ResolveProfilesSourceRelPath(sourceNode, modeName)
    If Len(relPath) = 0 Then Exit Function

    m_GetProfilesFilePathByMode = ex_XmlCore.m_CombineBasePath(wb, relPath)
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

    If wb Is Nothing Then Set wb = ThisWorkbook
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
        valueText = Trim$(ex_XmlCore.m_NodeAttrText(variantNode, "value"))
        If Len(valueText) = 0 Then
            MsgBox "Control '" & controlName & "' contains mode variant without 'value'.", vbExclamation
            Exit Function
        End If
        If Not IsNumeric(valueText) Then
            MsgBox "Control '" & controlName & "' has non-numeric mode variant value: " & valueText, vbExclamation
            Exit Function
        End If

        captionText = Trim$(ex_XmlCore.m_NodeAttrText(variantNode, "caption"))
        styleText = Trim$(ex_XmlCore.m_NodeAttrText(variantNode, "style"))
        displayText = Trim$(ex_XmlCore.m_NodeAttrText(variantNode, "display"))

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

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, controlName)
    If controlNode Is Nothing Then Exit Function

    m_GetControlAttribute = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, attrName))
End Function

Public Function m_ReadButtonStyles(Optional ByVal wb As Workbook) As Object
    Dim doc As Object
    Dim styleNodes As Object
    Dim styleNode As Object
    Dim styleName As String
    Dim stylesMap As Object
    Dim styleData As Object

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set stylesMap = CreateObject("Scripting.Dictionary")
    Set styleNodes = doc.selectNodes("/p:uiDefinition/p:styles/p:buttonStyle")
    If styleNodes Is Nothing Then
        Set m_ReadButtonStyles = stylesMap
        Exit Function
    End If

    For Each styleNode In styleNodes
        styleName = Trim$(ex_XmlCore.m_NodeAttrText(styleNode, "name"))
        If Len(styleName) = 0 Then
            MsgBox "UI config contains <buttonStyle> without 'name' attribute.", vbExclamation
            Exit Function
        End If

        Set styleData = CreateObject("Scripting.Dictionary")
        mp_SetStyleValue styleData, "backColor", ex_XmlCore.m_NodeAttrText(styleNode, "backColor")
        mp_SetStyleValue styleData, "textColor", ex_XmlCore.m_NodeAttrText(styleNode, "textColor")
        mp_SetStyleValue styleData, "borderColor", ex_XmlCore.m_NodeAttrText(styleNode, "borderColor")
        mp_SetStyleValue styleData, "borderWeight", ex_XmlCore.m_NodeAttrText(styleNode, "borderWeight")
        mp_SetStyleValue styleData, "fontName", ex_XmlCore.m_NodeAttrText(styleNode, "fontName")
        mp_SetStyleValue styleData, "fontSize", ex_XmlCore.m_NodeAttrText(styleNode, "fontSize")
        mp_SetStyleValue styleData, "fontBold", ex_XmlCore.m_NodeAttrText(styleNode, "fontBold")

        Set stylesMap(styleName) = styleData
    Next styleNode

    Set m_ReadButtonStyles = stylesMap
End Function

Public Function m_ApplyControlStyleByName(ByVal ws As Worksheet, ByVal controlName As String, ByVal styleName As String, Optional ByVal wb As Workbook) As Boolean
    Dim stylesMap As Object
    Dim shp As Shape

    If ws Is Nothing Then
        MsgBox "Failed to apply control style: worksheet is not specified.", vbExclamation
        Exit Function
    End If

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, controlName)
    If shp Is Nothing Then
        MsgBox "Failed to apply control style: control '" & controlName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Function
    End If

    Set stylesMap = m_ReadButtonStyles(wb)
    If stylesMap Is Nothing Then Exit Function

    m_ApplyControlStyleByName = m_ApplyButtonStyleByName(shp, styleName, stylesMap)
End Function

Public Function m_ApplyButtonStyleByName(ByVal shp As Shape, ByVal styleName As String, ByVal stylesMap As Object) As Boolean
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
        If Not ex_XmlCore.m_TryParseColor(CStr(styleData("backColor")), colorValue) Then
            MsgBox "Invalid style backColor for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.Fill.Visible = msoTrue
        shp.Fill.ForeColor.RGB = colorValue
    End If

    If styleData.Exists("textColor") Then
        If Not ex_XmlCore.m_TryParseColor(CStr(styleData("textColor")), colorValue) Then
            MsgBox "Invalid style textColor for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.TextFrame.Characters.Font.Color = colorValue
    End If

    If styleData.Exists("borderColor") Then
        If Not ex_XmlCore.m_TryParseColor(CStr(styleData("borderColor")), colorValue) Then
            MsgBox "Invalid style borderColor for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = colorValue
    End If

    If styleData.Exists("borderWeight") Then
        If Not ex_XmlCore.m_TryParseDouble(CStr(styleData("borderWeight")), numberValue) Then
            MsgBox "Invalid style borderWeight for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.Line.Weight = numberValue
    End If

    If styleData.Exists("fontName") Then
        shp.TextFrame.Characters.Font.Name = CStr(styleData("fontName"))
    End If

    If styleData.Exists("fontSize") Then
        If Not ex_XmlCore.m_TryParseDouble(CStr(styleData("fontSize")), numberValue) Then
            MsgBox "Invalid style fontSize for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.TextFrame.Characters.Font.Size = numberValue
    End If

    If styleData.Exists("fontBold") Then
        If Not ex_XmlCore.m_TryParseBoolean(CStr(styleData("fontBold")), boolValue) Then
            MsgBox "Invalid style fontBold for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.TextFrame.Characters.Font.Bold = boolValue
    End If

    m_ApplyButtonStyleByName = True
    Exit Function
EH:
    MsgBox "Failed to apply style '" & styleName & "' to shape '" & shp.Name & "': " & Err.Description, vbExclamation
End Function

Private Function mp_LoadUiDom(ByVal wb As Workbook) As Object
    Set mp_LoadUiDom = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        UI_CONFIG_REL_PATH, _
        PRESETS_NS, _
        "UI config file was not found: ", _
        "Failed to parse UI config file: ")
End Function

Private Function mp_GetControlNode(ByVal doc As Object, ByVal controlName As String) As Object
    Set mp_GetControlNode = doc.selectSingleNode("/p:uiDefinition/p:controls/p:control[@name=" & ex_XmlCore.m_XPathLiteral(controlName) & "]")
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

    filePath = ex_XmlCore.m_CombineBasePath(wb, relPath)
    Set srcDoc = ex_XmlCore.m_LoadDomByFilePath( _
        filePath, _
        PRESETS_NS, _
        "Profiles source file was not found: ", _
        "Failed to parse profiles source file: ")
    If srcDoc Is Nothing Then Exit Function

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
    Set mp_GetProfilesSourceNode = doc.selectSingleNode("/p:uiDefinition/p:dataSources/p:profilesSource[@name=" & ex_XmlCore.m_XPathLiteral(sourceName) & "]")
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

    modePersonal = Trim$(ex_XmlCore.m_NodeAttrText(sourceNode, "modePersonalCard"))
    modeComparing = Trim$(ex_XmlCore.m_NodeAttrText(sourceNode, "modeComparing"))
    pathPersonal = Trim$(ex_XmlCore.m_NodeAttrText(sourceNode, "pathPersonalCard"))
    pathComparing = Trim$(ex_XmlCore.m_NodeAttrText(sourceNode, "pathComparing"))
    defaultMode = Trim$(ex_XmlCore.m_NodeAttrText(sourceNode, "defaultMode"))

    If Len(modePersonal) = 0 Or Len(modeComparing) = 0 Then
        MsgBox "Profiles source is missing required mode labels: modePersonalCard/modeComparing.", vbExclamation
        Exit Function
    End If
    If Len(pathPersonal) = 0 Or Len(pathComparing) = 0 Then
        MsgBox "Profiles source is missing required paths: pathPersonalCard/pathComparing.", vbExclamation
        Exit Function
    End If

    modeName = Trim$(modeName)
    If Len(modeName) = 0 Then modeName = defaultMode
    If Len(modeName) = 0 Then modeName = modePersonal

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

Private Sub mp_SetStyleValue(ByVal styleData As Object, ByVal key As String, ByVal value As String)
    value = Trim$(value)
    If Len(value) = 0 Then Exit Sub
    styleData(key) = value
End Sub
