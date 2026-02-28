Attribute VB_Name = "ex_UiXmlProvider"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const DEV_UI_CONFIG_REL_PATH As String = "config\DevUI.xml"
Private Const PROFILES_FILE_SUFFIX As String = "Profiles.xml"
Private Const ACTION_MAP_REL_PATH As String = "config\ActionMap.xml"
Private Const DROPDOWN_CONTEXT_PROP_PREFIX As String = "Settings.DropdownContext."
Private Const STATE_ACTIVE_MODE_PROP As String = "Settings.ActiveModeName"

Public Const DROPDOWN_ITEM_COL_KEY As Long = 1
Public Const DROPDOWN_ITEM_COL_CAPTION As Long = 2
Public Const DROPDOWN_ITEM_COL_TARGET As Long = 3
Public Const DROPDOWN_ITEM_COL_SET_CONTEXT As Long = 4
Public Const DROPDOWN_ITEM_COL_ACTION_KEY As Long = 5
Public Const DROPDOWN_ITEM_COL_MACRO As Long = 6

Public Function m_GetDropdownItemsByName(ByVal controlName As String, Optional ByVal wb As Workbook, Optional ByVal modeName As String = vbNullString) As Variant
    Dim doc As Object
    Dim controlNode As Object
    Dim sourceName As String
    Dim itemNodes As Object
    Dim itemNode As Object
    Dim items() As String
    Dim idx As Long
    Dim itemText As String
    Dim itemRecords As Variant

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadDevUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, controlName)
    If controlNode Is Nothing Then Exit Function

    itemRecords = mp_GetDropdownItemRecordsFromControlNode(controlNode, wb)
    If mp_HasDropdownItemRecords(itemRecords) Then
        m_GetDropdownItemsByName = mp_BuildCaptionItemsFromRecords(itemRecords)
        Exit Function
    End If

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

Public Function m_GetDropdownItemRecordsByControl(ByVal controlName As String, Optional ByVal wb As Workbook) As Variant
    Dim doc As Object
    Dim controlNode As Object

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadDevUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, controlName)
    If controlNode Is Nothing Then Exit Function

    m_GetDropdownItemRecordsByControl = mp_GetDropdownItemRecordsFromControlNode(controlNode, wb)
End Function

Public Function m_GetDropdownItemKeyByTarget(ByVal controlName As String, ByVal targetText As String, Optional ByVal wb As Workbook) As String
    Dim records As Variant
    Dim i As Long
    Dim captionText As String
    Dim recordTarget As String

    targetText = Trim$(targetText)
    If Len(targetText) = 0 Then Exit Function

    records = m_GetDropdownItemRecordsByControl(controlName, wb)
    If Not mp_HasDropdownItemRecords(records) Then Exit Function

    For i = LBound(records, 1) To UBound(records, 1)
        captionText = Trim$(CStr(records(i, DROPDOWN_ITEM_COL_CAPTION)))
        recordTarget = Trim$(CStr(records(i, DROPDOWN_ITEM_COL_TARGET)))

        If StrComp(recordTarget, targetText, vbTextCompare) = 0 Or _
           StrComp(captionText, targetText, vbTextCompare) = 0 Or _
           StrComp(CStr(records(i, DROPDOWN_ITEM_COL_KEY)), targetText, vbTextCompare) = 0 Then
            m_GetDropdownItemKeyByTarget = CStr(records(i, DROPDOWN_ITEM_COL_KEY))
            Exit Function
        End If
    Next i
End Function

Public Sub m_SetDropdownContextValue(ByVal contextKey As String, ByVal valueText As String)
    Dim propName As String

    contextKey = Trim$(contextKey)
    If Len(contextKey) = 0 Then
        MsgBox "Dropdown context key cannot be empty.", vbExclamation
        Exit Sub
    End If

    propName = DROPDOWN_CONTEXT_PROP_PREFIX & mp_NormalizeContextKey(contextKey)

    On Error GoTo AddProp
    ThisWorkbook.CustomDocumentProperties(propName).Value = CStr(valueText)
    Exit Sub
AddProp:
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:=propName, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=CStr(valueText)
End Sub

Public Function m_GetDropdownContextValue(ByVal contextKey As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim propName As String

    contextKey = Trim$(contextKey)
    If Len(contextKey) = 0 Then
        m_GetDropdownContextValue = defaultValue
        Exit Function
    End If

    propName = DROPDOWN_CONTEXT_PROP_PREFIX & mp_NormalizeContextKey(contextKey)

    On Error GoTo EH
    m_GetDropdownContextValue = CStr(ThisWorkbook.CustomDocumentProperties(propName).Value)
    Exit Function
EH:
    m_GetDropdownContextValue = defaultValue
End Function

Public Sub m_ApplyDropdownSetContext(ByVal setContextText As String)
    Dim assignments As Variant
    Dim assignment As Variant
    Dim pairText As String
    Dim eqPos As Long
    Dim keyText As String
    Dim valueText As String

    setContextText = Trim$(setContextText)
    If Len(setContextText) = 0 Then Exit Sub

    assignments = Split(setContextText, ";")
    For Each assignment In assignments
        pairText = Trim$(CStr(assignment))
        If Len(pairText) = 0 Then GoTo NextAssignment

        eqPos = InStr(1, pairText, "=", vbTextCompare)
        If eqPos <= 1 Then
            MsgBox "Invalid setContext pair. Expected 'key=value', got: '" & pairText & "'.", vbExclamation
            Exit Sub
        End If

        keyText = Trim$(Left$(pairText, eqPos - 1))
        valueText = Trim$(Mid$(pairText, eqPos + 1))
        If Len(keyText) = 0 Then
            MsgBox "Invalid setContext pair with empty key: '" & pairText & "'.", vbExclamation
            Exit Sub
        End If

        m_SetDropdownContextValue keyText, valueText
NextAssignment:
    Next assignment
End Sub

Public Function m_ResolveMacroByActionKey(ByVal actionKey As String, Optional ByVal wb As Workbook) As String
    Dim filePath As String
    Dim doc As Object
    Dim actionNode As Object
    Dim macroName As String

    actionKey = Trim$(actionKey)
    If Len(actionKey) = 0 Then Exit Function

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    filePath = ex_XmlCore.m_CombineBasePath(wb, ACTION_MAP_REL_PATH)
    If Len(filePath) = 0 Then Exit Function
    If Len(Dir(filePath)) = 0 Then
        MsgBox "Action map file was not found: " & filePath, vbExclamation
        Exit Function
    End If

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    doc.preserveWhiteSpace = False
    If Not doc.Load(filePath) Then
        MsgBox "Failed to parse action map file: " & filePath, vbExclamation
        Exit Function
    End If

    Set actionNode = doc.selectSingleNode("/actionMap/action[@key=" & ex_XmlCore.m_XPathLiteral(actionKey) & "]")
    If actionNode Is Nothing Then
        MsgBox "Action key '" & actionKey & "' was not found in " & ACTION_MAP_REL_PATH & ".", vbExclamation
        Exit Function
    End If

    macroName = mp_GetPlainXmlAttrText(actionNode, "macro")
    If Len(macroName) = 0 Then
        MsgBox "Action key '" & actionKey & "' has empty macro in " & ACTION_MAP_REL_PATH & ".", vbExclamation
        Exit Function
    End If

    m_ResolveMacroByActionKey = macroName
End Function

Public Function m_GetProfilesFilePathByMode(Optional ByVal modeName As String = vbNullString, Optional ByVal wb As Workbook, Optional ByVal sourceName As String = "profilesByMode") As String
    Dim doc As Object
    Dim sourceNode As Object
    Dim relPath As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadDevUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set sourceNode = mp_GetProfilesSourceNode(doc, sourceName)
    If sourceNode Is Nothing Then Exit Function

    relPath = mp_ResolveProfilesSourceRelPath(sourceNode, modeName, PROFILES_FILE_SUFFIX)
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

    Set doc = mp_LoadDevUiDom(wb)
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

    Set doc = mp_LoadDevUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, controlName)
    If controlNode Is Nothing Then Exit Function

    m_GetControlAttribute = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, attrName))
End Function

Public Function m_GetLayoutAttribute(ByVal layoutNodeName As String, ByVal attrName As String, Optional ByVal wb As Workbook) As String
    Dim doc As Object
    Dim layoutNode As Object

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadDevUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set layoutNode = doc.selectSingleNode("/p:uiDefinition/p:layout/p:" & layoutNodeName)
    If layoutNode Is Nothing Then Exit Function

    m_GetLayoutAttribute = Trim$(ex_XmlCore.m_NodeAttrText(layoutNode, attrName))
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

    Set doc = mp_LoadDevUiDom(wb)
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
        shp.Fill.Solid
        shp.Fill.ForeColor.RGB = colorValue
        shp.Fill.Transparency = 0
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

    On Error Resume Next
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    On Error GoTo EH

    m_ApplyButtonStyleByName = True
    Exit Function
EH:
    MsgBox "Failed to apply style '" & styleName & "' to shape '" & shp.Name & "': " & Err.Description, vbExclamation
End Function

Private Function mp_GetDropdownItemRecordsFromControlNode(ByVal controlNode As Object, ByVal wb As Workbook) As Variant
    Dim sourceUri As String

    sourceUri = mp_ResolveControlSourceUri(controlNode, wb)
    If Len(sourceUri) = 0 Then Exit Function

    mp_GetDropdownItemRecordsFromControlNode = mp_GetDropdownItemRecordsBySourceUri(sourceUri, wb)
End Function

Private Function mp_ResolveControlSourceUri(ByVal controlNode As Object, ByVal wb As Workbook) As String
    Dim sourceUri As String
    Dim sourceUriTemplate As String

    sourceUri = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "sourceUri"))
    sourceUriTemplate = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "sourceUriTemplate"))

    If Len(sourceUriTemplate) > 0 Then
        mp_ResolveControlSourceUri = mp_ResolveTemplateValue(sourceUriTemplate)
        Exit Function
    End If

    mp_ResolveControlSourceUri = sourceUri
End Function

Private Function mp_GetDropdownItemRecordsBySourceUri(ByVal sourceUri As String, ByVal wb As Workbook) As Variant
    Dim hashPos As Long
    Dim relPath As String
    Dim listName As String
    Dim filePath As String
    Dim doc As Object
    Dim listNode As Object
    Dim itemNodes As Object
    Dim itemNode As Object
    Dim rowCount As Long
    Dim records() As Variant
    Dim rowIndex As Long
    Dim keyText As String
    Dim captionText As String

    sourceUri = Trim$(sourceUri)
    If Len(sourceUri) = 0 Then Exit Function

    hashPos = InStrRev(sourceUri, "#")
    If hashPos <= 1 Or hashPos >= Len(sourceUri) Then
        MsgBox "Invalid sourceUri format '" & sourceUri & "'. Expected '<relative-path>#<list-name>'.", vbExclamation
        Exit Function
    End If

    relPath = Trim$(Left$(sourceUri, hashPos - 1))
    listName = Trim$(Mid$(sourceUri, hashPos + 1))
    If Len(relPath) = 0 Or Len(listName) = 0 Then
        MsgBox "Invalid sourceUri format '" & sourceUri & "'. Path or list name is empty.", vbExclamation
        Exit Function
    End If

    filePath = ex_XmlCore.m_CombineBasePath(wb, relPath)
    Set doc = mp_LoadPlainDomByFilePath(filePath, "DropDown items file was not found: ", "Failed to parse DropDown items file: ")
    If doc Is Nothing Then Exit Function

    Set listNode = doc.selectSingleNode("/dropdownItems/list[@name=" & ex_XmlCore.m_XPathLiteral(listName) & "]")
    If listNode Is Nothing Then
        MsgBox "Dropdown list '" & listName & "' was not found in file: " & filePath, vbExclamation
        Exit Function
    End If

    Set itemNodes = listNode.selectNodes("item")
    If itemNodes Is Nothing Then Exit Function
    If itemNodes.Length = 0 Then
        mp_GetDropdownItemRecordsBySourceUri = Array()
        Exit Function
    End If

    rowCount = CLng(itemNodes.Length)
    ReDim records(1 To rowCount, 1 To DROPDOWN_ITEM_COL_MACRO)

    rowIndex = 1
    For Each itemNode In itemNodes
        keyText = mp_GetPlainXmlAttrText(itemNode, "key")
        If Len(keyText) = 0 Then
            MsgBox "Dropdown item in list '" & listName & "' has empty required attribute 'key' in file: " & filePath, vbExclamation
            Exit Function
        End If

        captionText = mp_GetPlainXmlAttrText(itemNode, "caption")
        If Len(captionText) = 0 Then captionText = keyText

        records(rowIndex, DROPDOWN_ITEM_COL_KEY) = keyText
        records(rowIndex, DROPDOWN_ITEM_COL_CAPTION) = captionText
        records(rowIndex, DROPDOWN_ITEM_COL_TARGET) = mp_GetPlainXmlAttrText(itemNode, "target")
        records(rowIndex, DROPDOWN_ITEM_COL_SET_CONTEXT) = mp_GetPlainXmlAttrText(itemNode, "setContext")
        records(rowIndex, DROPDOWN_ITEM_COL_ACTION_KEY) = mp_GetPlainXmlAttrText(itemNode, "actionKey")
        records(rowIndex, DROPDOWN_ITEM_COL_MACRO) = mp_GetPlainXmlAttrText(itemNode, "macro")
        rowIndex = rowIndex + 1
    Next itemNode

    mp_GetDropdownItemRecordsBySourceUri = records
End Function

Private Function mp_GetPlainXmlAttrText(ByVal node As Object, ByVal attrName As String) As String
    Dim rawValue As Variant

    On Error GoTo EH
    rawValue = node.getAttribute(attrName)
    If IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    mp_GetPlainXmlAttrText = Trim$(CStr(rawValue))
    Exit Function
EH:
    mp_GetPlainXmlAttrText = vbNullString
End Function

Private Function mp_LoadPlainDomByFilePath(ByVal filePath As String, ByVal notFoundPrefix As String, ByVal parseErrorPrefix As String) As Object
    Dim doc As Object

    filePath = Trim$(filePath)
    If Len(filePath) = 0 Then Exit Function

    If Len(Dir(filePath)) = 0 Then
        MsgBox notFoundPrefix & filePath, vbExclamation
        Exit Function
    End If

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    doc.preserveWhiteSpace = False
    If Not doc.Load(filePath) Then
        MsgBox parseErrorPrefix & filePath, vbExclamation
        Exit Function
    End If

    Set mp_LoadPlainDomByFilePath = doc
End Function

Private Function mp_HasDropdownItemRecords(ByVal itemRecords As Variant) As Boolean
    On Error GoTo EH
    If Not IsArray(itemRecords) Then Exit Function
    mp_HasDropdownItemRecords = (UBound(itemRecords, 1) >= LBound(itemRecords, 1))
    Exit Function
EH:
    mp_HasDropdownItemRecords = False
End Function

Private Function mp_BuildCaptionItemsFromRecords(ByVal itemRecords As Variant) As Variant
    Dim rowIndex As Long
    Dim result() As String
    Dim lowerBound As Long
    Dim upperBound As Long

    If Not mp_HasDropdownItemRecords(itemRecords) Then Exit Function

    lowerBound = LBound(itemRecords, 1)
    upperBound = UBound(itemRecords, 1)
    ReDim result(0 To upperBound - lowerBound)

    For rowIndex = lowerBound To upperBound
        result(rowIndex - lowerBound) = CStr(itemRecords(rowIndex, DROPDOWN_ITEM_COL_CAPTION))
    Next rowIndex

    mp_BuildCaptionItemsFromRecords = result
End Function

Private Function mp_ResolveTemplateValue(ByVal templateText As String) As String
    Dim resultText As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tokenName As String
    Dim tokenValue As String

    resultText = templateText
    startPos = InStr(1, resultText, "{", vbTextCompare)

    Do While startPos > 0
        endPos = InStr(startPos + 1, resultText, "}", vbTextCompare)
        If endPos <= startPos Then
            MsgBox "Invalid sourceUriTemplate '" & templateText & "': missing closing '}'.", vbExclamation
            Exit Function
        End If

        tokenName = Trim$(Mid$(resultText, startPos + 1, endPos - startPos - 1))
        If Len(tokenName) = 0 Then
            MsgBox "Invalid sourceUriTemplate '" & templateText & "': empty placeholder is not allowed.", vbExclamation
            Exit Function
        End If

        tokenValue = mp_GetContextOrFallbackValue(tokenName)
        If Len(tokenValue) = 0 Then
            MsgBox "Unable to resolve sourceUriTemplate placeholder '{" & tokenName & "}'.", vbExclamation
            Exit Function
        End If

        resultText = Left$(resultText, startPos - 1) & tokenValue & Mid$(resultText, endPos + 1)
        startPos = InStr(1, resultText, "{", vbTextCompare)
    Loop

    mp_ResolveTemplateValue = resultText
End Function

Private Function mp_GetContextOrFallbackValue(ByVal contextKey As String) As String
    Dim valueText As String
    Dim mappedModeKey As String

    valueText = m_GetDropdownContextValue(contextKey, vbNullString)
    If Len(valueText) > 0 Then
        mp_GetContextOrFallbackValue = valueText
        Exit Function
    End If

    If StrComp(contextKey, "activeMode", vbTextCompare) = 0 Then
        valueText = mp_GetStatePropertyText(STATE_ACTIVE_MODE_PROP)
        If Len(valueText) = 0 Then Exit Function

        mappedModeKey = m_GetDropdownItemKeyByTarget("btnCustomMode", valueText, ThisWorkbook)
        If Len(mappedModeKey) > 0 Then
            mp_GetContextOrFallbackValue = mappedModeKey
        Else
            mp_GetContextOrFallbackValue = valueText
        End If
        Exit Function
    End If
End Function

Private Function mp_GetStatePropertyText(ByVal propName As String) As String
    On Error GoTo EH
    mp_GetStatePropertyText = CStr(ThisWorkbook.CustomDocumentProperties(propName).Value)
    Exit Function
EH:
    mp_GetStatePropertyText = vbNullString
End Function

Private Function mp_NormalizeContextKey(ByVal contextKey As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim resultText As String

    contextKey = Trim$(contextKey)
    If Len(contextKey) = 0 Then
        mp_NormalizeContextKey = "_"
        Exit Function
    End If

    For i = 1 To Len(contextKey)
        ch = Mid$(contextKey, i, 1)
        code = AscW(ch)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
            resultText = resultText & ch
        Else
            resultText = resultText & "_"
        End If
    Next i

    If Len(resultText) = 0 Then resultText = "_"
    mp_NormalizeContextKey = resultText
End Function

Private Function mp_LoadDevUiDom(ByVal wb As Workbook) As Object
    Set mp_LoadDevUiDom = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        DEV_UI_CONFIG_REL_PATH, _
        PROFILES_NS, _
        "Dev UI config file was not found: ", _
        "Failed to parse Dev UI config file: ")
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

    relPath = mp_ResolveProfilesSourceRelPath(sourceNode, modeName, PROFILES_FILE_SUFFIX)
    If Len(relPath) = 0 Then Exit Function

    filePath = ex_XmlCore.m_CombineBasePath(wb, relPath)
    Set srcDoc = ex_XmlCore.m_LoadDomByFilePath( _
        filePath, _
        PROFILES_NS, _
        "Profiles source file was not found: ", _
        "Failed to parse profiles source file: ")
    If srcDoc Is Nothing Then Exit Function

    Set profileNodes = srcDoc.selectNodes("/p:profiles/p:profile")
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

Private Function mp_ResolveProfilesSourceRelPath(ByVal sourceNode As Object, ByVal modeName As String, Optional ByVal fileSuffix As String = PROFILES_FILE_SUFFIX) As String
    Dim modePersonal As String
    Dim modeComparing As String
    Dim pathPersonal As String
    Dim pathComparing As String
    Dim personalDir As String
    Dim comparingDir As String
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

    personalDir = mp_NormalizeDirectoryPath(pathPersonal)
    comparingDir = mp_NormalizeDirectoryPath(pathComparing)
    If Len(personalDir) = 0 Or Len(comparingDir) = 0 Then
        MsgBox "Profiles source paths are invalid: pathPersonalCard/pathComparing.", vbExclamation
        Exit Function
    End If

    modeName = Trim$(modeName)
    If Len(modeName) = 0 Then modeName = defaultMode
    If Len(modeName) = 0 Then modeName = modePersonal

    If Len(fileSuffix) = 0 Then fileSuffix = PROFILES_FILE_SUFFIX

    If StrComp(modeName, modeComparing, vbTextCompare) = 0 Then
        mp_ResolveProfilesSourceRelPath = mp_BuildPatternBasedFilePath(comparingDir, fileSuffix)
        Exit Function
    End If
    If StrComp(modeName, modePersonal, vbTextCompare) = 0 Then
        mp_ResolveProfilesSourceRelPath = mp_BuildPatternBasedFilePath(personalDir, fileSuffix)
        Exit Function
    End If

    MsgBox "Invalid mode '" & modeName & "' for profiles source. Allowed values: '" & modePersonal & "', '" & modeComparing & "'.", vbExclamation
End Function

Private Function mp_BuildPatternBasedFilePath(ByVal directoryRelPath As String, ByVal fileSuffix As String) As String
    Dim normalizedDir As String
    Dim dirName As String

    normalizedDir = mp_NormalizeDirectoryPath(directoryRelPath)
    If Len(normalizedDir) = 0 Then Exit Function

    dirName = mp_GetLastPathSegment(normalizedDir)
    If Len(dirName) = 0 Then Exit Function

    mp_BuildPatternBasedFilePath = normalizedDir & "\" & dirName & fileSuffix
End Function

Private Function mp_NormalizeDirectoryPath(ByVal value As String) As String
    Dim slashPos As Long
    Dim pathLeaf As String

    value = Trim$(value)
    If Len(value) = 0 Then Exit Function

    value = Replace$(value, "/", "\")
    Do While Right$(value, 1) = "\"
        value = Left$(value, Len(value) - 1)
    Loop

    slashPos = InStrRev(value, "\")
    If slashPos > 0 Then
        pathLeaf = Mid$(value, slashPos + 1)
        If InStr(1, pathLeaf, ".", vbTextCompare) > 0 Then
            value = Left$(value, slashPos - 1)
        End If
    End If

    mp_NormalizeDirectoryPath = value
End Function

Private Function mp_GetLastPathSegment(ByVal pathValue As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(pathValue, "\")
    If slashPos <= 0 Then
        mp_GetLastPathSegment = pathValue
    Else
        mp_GetLastPathSegment = Mid$(pathValue, slashPos + 1)
    End If
End Function

Private Sub mp_SetStyleValue(ByVal styleData As Object, ByVal key As String, ByVal value As String)
    value = Trim$(value)
    If Len(value) = 0 Then Exit Sub
    styleData(key) = value
End Sub
