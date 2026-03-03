Attribute VB_Name = "ex_UiXmlProvider"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const DEV_UI_CONFIG_REL_PATH As String = "config\DevUI.xml"
Private Const ACTION_MAP_REL_PATH As String = "config\ActionMap.xml"
Private Const DROPDOWN_CONTEXT_PROP_PREFIX As String = "Settings.DropdownContext."
Private Const STATE_ACTIVE_MODE_KEY_PROP As String = "Settings.ActiveModeKey"

Public Const DROPDOWN_ITEM_COL_KEY As Long = 1
Public Const DROPDOWN_ITEM_COL_CAPTION As Long = 2
Public Const DROPDOWN_ITEM_COL_TARGET As Long = 3
Public Const DROPDOWN_ITEM_COL_SET_CONTEXT As Long = 4
Public Const DROPDOWN_ITEM_COL_ACTION_KEY As Long = 5
Public Const DROPDOWN_ITEM_COL_MACRO As Long = 6

Private Const MODE_LIST_CONTROL As String = "btnCustomMode"
Private Const PROFILES_FILE_SOURCE_NAME As String = "profilesFileByMode"

Private g_DevUiDomCache As Object
Private g_DevUiDomCacheWbKey As String
Private g_DevUiDomCacheStamp As Date
Private g_SourceUriRecordsCache As Object
Private g_SourceUriRecordsStampCache As Object
Private g_SourceUriDefaultKeyCache As Object
Private g_SourceUriDefaultKeyStampCache As Object

Public Function m_GetDropdownItemsByName(ByVal controlName As String, Optional ByVal wb As Workbook, Optional ByVal modeKey As String = vbNullString) As Variant
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

    itemRecords = mp_GetDropdownItemRecordsFromControlNode(controlNode, wb, modeKey, doc)
    If mp_HasDropdownItemRecords(itemRecords) Then
        m_GetDropdownItemsByName = mp_BuildCaptionItemsFromRecords(itemRecords)
        Exit Function
    End If

    sourceName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource"))
    If Len(sourceName) > 0 Then
        m_GetDropdownItemsByName = Array()
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

Public Function m_GetDropdownItemRecordsByControl(ByVal controlName As String, Optional ByVal wb As Workbook, Optional ByVal modeKey As String = vbNullString) As Variant
    Dim doc As Object
    Dim controlNode As Object

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadDevUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, controlName)
    If controlNode Is Nothing Then Exit Function

    m_GetDropdownItemRecordsByControl = mp_GetDropdownItemRecordsFromControlNode(controlNode, wb, modeKey, doc)
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

Public Function m_GetDropdownItemTargetByKey(ByVal controlName As String, ByVal keyText As String, Optional ByVal wb As Workbook) As String
    Dim records As Variant
    Dim i As Long
    Dim recordKey As String
    Dim recordTarget As String
    Dim captionText As String

    keyText = Trim$(keyText)
    If Len(keyText) = 0 Then Exit Function

    records = m_GetDropdownItemRecordsByControl(controlName, wb)
    If Not mp_HasDropdownItemRecords(records) Then Exit Function

    For i = LBound(records, 1) To UBound(records, 1)
        recordKey = Trim$(CStr(records(i, DROPDOWN_ITEM_COL_KEY)))
        If StrComp(recordKey, keyText, vbTextCompare) <> 0 Then GoTo ContinueRow

        recordTarget = Trim$(CStr(records(i, DROPDOWN_ITEM_COL_TARGET)))
        captionText = Trim$(CStr(records(i, DROPDOWN_ITEM_COL_CAPTION)))
        If Len(recordTarget) > 0 Then
            m_GetDropdownItemTargetByKey = recordTarget
        Else
            m_GetDropdownItemTargetByKey = captionText
        End If
        Exit Function
ContinueRow:
    Next i
End Function

Public Function m_GetDropdownItemCaptionByKey(ByVal controlName As String, ByVal keyText As String, Optional ByVal wb As Workbook) As String
    Dim records As Variant
    Dim i As Long
    Dim recordKey As String

    keyText = Trim$(keyText)
    If Len(keyText) = 0 Then Exit Function

    records = m_GetDropdownItemRecordsByControl(controlName, wb)
    If Not mp_HasDropdownItemRecords(records) Then Exit Function

    For i = LBound(records, 1) To UBound(records, 1)
        recordKey = Trim$(CStr(records(i, DROPDOWN_ITEM_COL_KEY)))
        If StrComp(recordKey, keyText, vbTextCompare) = 0 Then
            m_GetDropdownItemCaptionByKey = Trim$(CStr(records(i, DROPDOWN_ITEM_COL_CAPTION)))
            Exit Function
        End If
    Next i
End Function

Public Function m_GetDefaultModeKey(Optional ByVal wb As Workbook) As String
    Dim doc As Object
    Dim controlNode As Object
    Dim sourceName As String
    Dim sourceUri As String
    Dim defaultModeKey As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadDevUiDom(wb)
    If doc Is Nothing Then Exit Function

    Set controlNode = mp_GetControlNode(doc, MODE_LIST_CONTROL)
    If controlNode Is Nothing Then Exit Function

    sourceUri = mp_ResolveControlSourceUri(controlNode, wb, vbNullString)
    If Len(sourceUri) = 0 Then
        sourceName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource"))
        If Len(sourceName) = 0 Then Exit Function
        sourceUri = mp_ResolveItemsSourceUriByName(doc, sourceName, wb, vbNullString)
    End If
    If Len(sourceUri) = 0 Then Exit Function

    If Not mp_TryGetDefaultDropdownItemKeyBySourceUri(sourceUri, wb, defaultModeKey, MODE_LIST_CONTROL) Then
        Exit Function
    End If

    If Len(defaultModeKey) > 0 Then
        m_GetDefaultModeKey = defaultModeKey
        Exit Function
    End If

    m_GetDefaultModeKey = Trim$(m_GetModeKeyByIndex(1, wb))
End Function

Public Function m_GetModeKeyByIndex(ByVal modeIndex As Long, Optional ByVal wb As Workbook) As String
    Dim modeRecords As Variant
    Dim rowStart As Long
    Dim rowIndex As Long

    If modeIndex <= 0 Then Exit Function
    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    modeRecords = m_GetDropdownItemRecordsByControl(MODE_LIST_CONTROL, wb)
    If Not mp_HasDropdownItemRecords(modeRecords) Then Exit Function

    rowStart = LBound(modeRecords, 1)
    rowIndex = rowStart + modeIndex - 1
    If rowIndex < rowStart Or rowIndex > UBound(modeRecords, 1) Then Exit Function

    m_GetModeKeyByIndex = Trim$(CStr(modeRecords(rowIndex, DROPDOWN_ITEM_COL_KEY)))
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

Public Function m_GetProfilesFilePathByMode(Optional ByVal modeKey As String = vbNullString, Optional ByVal wb As Workbook, Optional ByVal sourceName As String = PROFILES_FILE_SOURCE_NAME) As String
    Dim doc As Object
    Dim sourceUri As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    Set doc = mp_LoadDevUiDom(wb)
    If doc Is Nothing Then Exit Function

    sourceUri = mp_ResolveItemsSourceUriByName(doc, sourceName, wb, modeKey)
    If Len(sourceUri) = 0 Then Exit Function

    If InStr(1, sourceUri, "#", vbTextCompare) > 0 Then
        MsgBox "Items source '" & sourceName & "' must resolve to a file path without '#list': " & sourceUri, vbExclamation
        Exit Function
    End If

    m_GetProfilesFilePathByMode = ex_XmlCore.m_CombineBasePath(wb, sourceUri)
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

Private Function mp_GetDropdownItemRecordsFromControlNode(ByVal controlNode As Object, ByVal wb As Workbook, Optional ByVal modeKey As String = vbNullString, Optional ByVal doc As Object = Nothing) As Variant
    Dim sourceUri As String
    Dim sourceName As String

    sourceUri = mp_ResolveControlSourceUri(controlNode, wb, modeKey)
    If Len(sourceUri) > 0 Then
        mp_GetDropdownItemRecordsFromControlNode = mp_GetDropdownItemRecordsBySourceUri(sourceUri, wb)
        If mp_HasDropdownItemRecords(mp_GetDropdownItemRecordsFromControlNode) Then Exit Function
    End If

    sourceName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource"))
    If Len(sourceName) = 0 Then Exit Function

    If doc Is Nothing Then
        On Error Resume Next
        Set doc = controlNode.OwnerDocument
        On Error GoTo 0
    End If

    If doc Is Nothing Then
        Set doc = mp_LoadDevUiDom(wb)
        If doc Is Nothing Then Exit Function
    End If

    mp_GetDropdownItemRecordsFromControlNode = mp_GetDropdownItemRecordsByItemsSource(doc, sourceName, wb, modeKey)
End Function

Private Function mp_ResolveControlSourceUri(ByVal controlNode As Object, ByVal wb As Workbook, Optional ByVal modeKey As String = vbNullString) As String
    Dim sourceUri As String
    Dim sourceUriTemplate As String

    sourceUri = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "uri"))
    sourceUriTemplate = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "uriTemplate"))
    If Len(sourceUri) = 0 Then sourceUri = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "sourceUri"))
    If Len(sourceUriTemplate) = 0 Then sourceUriTemplate = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "sourceUriTemplate"))

    If Len(sourceUriTemplate) > 0 Then
        mp_ResolveControlSourceUri = mp_ResolveTemplateValue(sourceUriTemplate, modeKey)
        Exit Function
    End If

    mp_ResolveControlSourceUri = sourceUri
End Function

Private Function mp_GetDropdownItemRecordsByItemsSource(ByVal doc As Object, ByVal sourceName As String, ByVal wb As Workbook, Optional ByVal modeKey As String = vbNullString) As Variant
    Dim sourceUri As String

    sourceName = Trim$(sourceName)
    If Len(sourceName) = 0 Then Exit Function
    If doc Is Nothing Then Exit Function

    sourceUri = mp_ResolveItemsSourceUriByName(doc, sourceName, wb, modeKey)
    If Len(sourceUri) = 0 Then Exit Function

    mp_GetDropdownItemRecordsByItemsSource = mp_GetDropdownItemRecordsBySourceUri(sourceUri, wb)
End Function

Private Function mp_ResolveItemSourceUri(ByVal sourceNode As Object, Optional ByVal modeKey As String = vbNullString) As String
    Dim sourceUri As String
    Dim sourceUriTemplate As String

    sourceUri = Trim$(ex_XmlCore.m_NodeAttrText(sourceNode, "uri"))
    sourceUriTemplate = Trim$(ex_XmlCore.m_NodeAttrText(sourceNode, "uriTemplate"))

    If Len(sourceUriTemplate) > 0 Then
        mp_ResolveItemSourceUri = mp_ResolveTemplateValue(sourceUriTemplate, modeKey)
        Exit Function
    End If

    mp_ResolveItemSourceUri = sourceUri
End Function

Private Function mp_GetItemSourceNode(ByVal doc As Object, ByVal sourceName As String) As Object
    Set mp_GetItemSourceNode = doc.selectSingleNode("/p:uiDefinition/p:dataSources/p:itemsSource[@name=" & ex_XmlCore.m_XPathLiteral(sourceName) & "]")
End Function

Private Function mp_ResolveItemsSourceUriByName( _
    ByVal doc As Object, _
    ByVal sourceName As String, _
    ByVal wb As Workbook, _
    Optional ByVal modeKey As String = vbNullString) As String

    Dim sourceNode As Object
    Dim sourceUri As String

    sourceName = Trim$(sourceName)
    If Len(sourceName) = 0 Then Exit Function
    If doc Is Nothing Then Exit Function

    Set sourceNode = mp_GetItemSourceNode(doc, sourceName)
    If sourceNode Is Nothing Then
        MsgBox "Items source '" & sourceName & "' was not found in UI config.", vbExclamation
        Exit Function
    End If

    sourceUri = mp_ResolveItemSourceUri(sourceNode, modeKey)
    If Len(sourceUri) = 0 Then
        MsgBox "Items source '" & sourceName & "' has empty uri/uriTemplate.", vbExclamation
        Exit Function
    End If

    mp_ResolveItemsSourceUriByName = sourceUri
End Function

Private Function mp_GetDropdownItemRecordsBySourceUri(ByVal sourceUri As String, ByVal wb As Workbook) As Variant
    Dim filePath As String
    Dim listName As String
    Dim doc As Object
    Dim listNode As Object
    Dim itemNodes As Object
    Dim itemNode As Object
    Dim rowCount As Long
    Dim records() As Variant
    Dim rowIndex As Long
    Dim keyText As String
    Dim captionText As String
    Dim defaultText As String
    Dim isDefault As Boolean
    Dim defaultCount As Long
    Dim cacheKey As String
    Dim fileStamp As Date

    sourceUri = Trim$(sourceUri)
    If Len(sourceUri) = 0 Then Exit Function
    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    If Not mp_TryResolveSourceUriParts(sourceUri, wb, vbNullString, filePath, listName) Then
        Exit Function
    End If

    mp_EnsureRuntimeCaches
    cacheKey = mp_BuildSourceUriCacheKey(wb, sourceUri)
    If mp_TryGetFileStamp(filePath, fileStamp) Then
        If g_SourceUriRecordsCache.Exists(cacheKey) Then
            If g_SourceUriRecordsStampCache.Exists(cacheKey) Then
                If g_SourceUriRecordsStampCache(cacheKey) = fileStamp Then
                    mp_GetDropdownItemRecordsBySourceUri = g_SourceUriRecordsCache(cacheKey)
                    Exit Function
                End If
            End If
        End If
    End If

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

        defaultText = mp_GetPlainXmlAttrText(itemNode, "default")
        If Len(defaultText) > 0 Then
            If Not ex_XmlCore.m_TryParseBoolean(defaultText, isDefault) Then
                MsgBox "Invalid boolean value '" & defaultText & "' in attribute 'default' for list '" & listName & "' (file: " & filePath & ").", vbExclamation
                Exit Function
            End If
            If isDefault Then defaultCount = defaultCount + 1
        End If

        records(rowIndex, DROPDOWN_ITEM_COL_KEY) = keyText
        records(rowIndex, DROPDOWN_ITEM_COL_CAPTION) = captionText
        records(rowIndex, DROPDOWN_ITEM_COL_TARGET) = mp_GetPlainXmlAttrText(itemNode, "target")
        records(rowIndex, DROPDOWN_ITEM_COL_SET_CONTEXT) = mp_GetPlainXmlAttrText(itemNode, "setContext")
        records(rowIndex, DROPDOWN_ITEM_COL_ACTION_KEY) = mp_GetPlainXmlAttrText(itemNode, "actionKey")
        records(rowIndex, DROPDOWN_ITEM_COL_MACRO) = mp_GetPlainXmlAttrText(itemNode, "macro")
        rowIndex = rowIndex + 1
    Next itemNode

    If defaultCount > 1 Then
        MsgBox "Dropdown list '" & listName & "' in file '" & filePath & "' contains multiple default items. Only one item with default='true' is allowed.", vbExclamation
        Exit Function
    End If

    mp_GetDropdownItemRecordsBySourceUri = records
    If mp_TryGetFileStamp(filePath, fileStamp) Then
        g_SourceUriRecordsCache(cacheKey) = records
        g_SourceUriRecordsStampCache(cacheKey) = fileStamp
    End If
End Function

Private Function mp_TryGetDefaultDropdownItemKeyBySourceUri( _
    ByVal sourceUri As String, _
    ByVal wb As Workbook, _
    ByRef outDefaultKey As String, _
    ByVal controlName As String) As Boolean

    Dim filePath As String
    Dim listName As String
    Dim doc As Object
    Dim listNode As Object
    Dim itemNodes As Object
    Dim itemNode As Object
    Dim keyText As String
    Dim defaultText As String
    Dim isDefault As Boolean
    Dim defaultCount As Long
    Dim cacheKey As String
    Dim fileStamp As Date

    sourceUri = Trim$(sourceUri)
    If Len(sourceUri) = 0 Then
        mp_TryGetDefaultDropdownItemKeyBySourceUri = True
        Exit Function
    End If
    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    If Not mp_TryResolveSourceUriParts(sourceUri, wb, controlName, filePath, listName) Then
        Exit Function
    End If

    mp_EnsureRuntimeCaches
    cacheKey = mp_BuildSourceUriCacheKey(wb, sourceUri)
    If mp_TryGetFileStamp(filePath, fileStamp) Then
        If g_SourceUriDefaultKeyCache.Exists(cacheKey) Then
            If g_SourceUriDefaultKeyStampCache.Exists(cacheKey) Then
                If g_SourceUriDefaultKeyStampCache(cacheKey) = fileStamp Then
                    outDefaultKey = CStr(g_SourceUriDefaultKeyCache(cacheKey))
                    mp_TryGetDefaultDropdownItemKeyBySourceUri = True
                    Exit Function
                End If
            End If
        End If
    End If

    Set doc = mp_LoadPlainDomByFilePath(filePath, "DropDown items file was not found: ", "Failed to parse DropDown items file: ")
    If doc Is Nothing Then Exit Function

    Set listNode = doc.selectSingleNode("/dropdownItems/list[@name=" & ex_XmlCore.m_XPathLiteral(listName) & "]")
    If listNode Is Nothing Then
        MsgBox "Dropdown list '" & listName & "' was not found in file: " & filePath, vbExclamation
        Exit Function
    End If

    Set itemNodes = listNode.selectNodes("item")
    If itemNodes Is Nothing Then
        mp_TryGetDefaultDropdownItemKeyBySourceUri = True
        Exit Function
    End If

    For Each itemNode In itemNodes
        defaultText = mp_GetPlainXmlAttrText(itemNode, "default")
        If Len(defaultText) = 0 Then GoTo NextItem

        If Not ex_XmlCore.m_TryParseBoolean(defaultText, isDefault) Then
            MsgBox "Invalid boolean value '" & defaultText & "' in attribute 'default' for list '" & listName & "' (file: " & filePath & ").", vbExclamation
            Exit Function
        End If
        If Not isDefault Then GoTo NextItem

        keyText = mp_GetPlainXmlAttrText(itemNode, "key")
        If Len(keyText) = 0 Then
            MsgBox "Default item in list '" & listName & "' has empty key (file: " & filePath & ").", vbExclamation
            Exit Function
        End If

        defaultCount = defaultCount + 1
        outDefaultKey = keyText
NextItem:
    Next itemNode

    If defaultCount > 1 Then
        MsgBox "Dropdown list '" & listName & "' in file '" & filePath & "' contains multiple default items. Only one item with default='true' is allowed.", vbExclamation
        Exit Function
    End If

    mp_TryGetDefaultDropdownItemKeyBySourceUri = True
    If mp_TryGetFileStamp(filePath, fileStamp) Then
        g_SourceUriDefaultKeyCache(cacheKey) = outDefaultKey
        g_SourceUriDefaultKeyStampCache(cacheKey) = fileStamp
    End If
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

Private Function mp_ResolveTemplateValue(ByVal templateText As String, Optional ByVal modeKeyOverride As String = vbNullString, Optional ByVal profileOverride As String = vbNullString) As String
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
            MsgBox "Invalid uriTemplate '" & templateText & "': missing closing '}'.", vbExclamation
            Exit Function
        End If

        tokenName = Trim$(Mid$(resultText, startPos + 1, endPos - startPos - 1))
        If Len(tokenName) = 0 Then
            MsgBox "Invalid uriTemplate '" & templateText & "': empty placeholder is not allowed.", vbExclamation
            Exit Function
        End If

        tokenValue = mp_GetContextOrFallbackValue(tokenName, modeKeyOverride, profileOverride)
        If Len(tokenValue) = 0 Then
            MsgBox "Unable to resolve uriTemplate placeholder '{" & tokenName & "}'.", vbExclamation
            Exit Function
        End If

        resultText = Left$(resultText, startPos - 1) & tokenValue & Mid$(resultText, endPos + 1)
        startPos = InStr(1, resultText, "{", vbTextCompare)
    Loop

    mp_ResolveTemplateValue = resultText
End Function

Private Function mp_GetContextOrFallbackValue(ByVal contextKey As String, Optional ByVal modeKeyOverride As String = vbNullString, Optional ByVal profileOverride As String = vbNullString) As String
    Dim valueText As String

    If StrComp(contextKey, "activeMode", vbTextCompare) = 0 Then
        modeKeyOverride = Trim$(modeKeyOverride)
        If Len(modeKeyOverride) > 0 Then
            mp_GetContextOrFallbackValue = modeKeyOverride
            Exit Function
        End If
    End If

    If StrComp(contextKey, "activeProfile", vbTextCompare) = 0 Then
        profileOverride = Trim$(profileOverride)
        If Len(profileOverride) > 0 Then
            mp_GetContextOrFallbackValue = profileOverride
            Exit Function
        End If
    End If

    valueText = m_GetDropdownContextValue(contextKey, vbNullString)
    If Len(valueText) > 0 Then
        mp_GetContextOrFallbackValue = valueText
        Exit Function
    End If

    If StrComp(contextKey, "activeMode", vbTextCompare) = 0 Then
        valueText = Trim$(mp_GetStatePropertyText(STATE_ACTIVE_MODE_KEY_PROP))
        mp_GetContextOrFallbackValue = valueText
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
    Dim wbKey As String
    Dim fileStamp As Date

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    mp_EnsureRuntimeCaches
    wbKey = mp_BuildWorkbookCacheKey(wb)

    If mp_TryGetRelativeFileStamp(wb, DEV_UI_CONFIG_REL_PATH, fileStamp) Then
        If Not g_DevUiDomCache Is Nothing Then
            If StrComp(g_DevUiDomCacheWbKey, wbKey, vbTextCompare) = 0 Then
                If g_DevUiDomCacheStamp = fileStamp Then
                    Set mp_LoadDevUiDom = g_DevUiDomCache
                    Exit Function
                End If
            End If
        End If
    End If

    Set g_DevUiDomCache = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        DEV_UI_CONFIG_REL_PATH, _
        PROFILES_NS, _
        "Dev UI config file was not found: ", _
        "Failed to parse Dev UI config file: ")
    Set mp_LoadDevUiDom = g_DevUiDomCache
    g_DevUiDomCacheWbKey = wbKey
    If mp_TryGetRelativeFileStamp(wb, DEV_UI_CONFIG_REL_PATH, fileStamp) Then
        g_DevUiDomCacheStamp = fileStamp
    Else
        g_DevUiDomCacheStamp = 0
    End If
End Function

Private Sub mp_EnsureRuntimeCaches()
    If g_SourceUriRecordsCache Is Nothing Then
        Set g_SourceUriRecordsCache = CreateObject("Scripting.Dictionary")
        g_SourceUriRecordsCache.CompareMode = 1
    End If
    If g_SourceUriRecordsStampCache Is Nothing Then
        Set g_SourceUriRecordsStampCache = CreateObject("Scripting.Dictionary")
        g_SourceUriRecordsStampCache.CompareMode = 1
    End If
    If g_SourceUriDefaultKeyCache Is Nothing Then
        Set g_SourceUriDefaultKeyCache = CreateObject("Scripting.Dictionary")
        g_SourceUriDefaultKeyCache.CompareMode = 1
    End If
    If g_SourceUriDefaultKeyStampCache Is Nothing Then
        Set g_SourceUriDefaultKeyStampCache = CreateObject("Scripting.Dictionary")
        g_SourceUriDefaultKeyStampCache.CompareMode = 1
    End If
End Sub

Private Function mp_BuildWorkbookCacheKey(ByVal wb As Workbook) As String
    If wb Is Nothing Then
        mp_BuildWorkbookCacheKey = "wb:none"
        Exit Function
    End If

    mp_BuildWorkbookCacheKey = LCase$(Trim$(wb.FullName))
    If Len(mp_BuildWorkbookCacheKey) = 0 Then
        mp_BuildWorkbookCacheKey = "wb:" & LCase$(Trim$(wb.Path)) & "|" & LCase$(Trim$(wb.Name))
    End If
End Function

Private Function mp_BuildSourceUriCacheKey(ByVal wb As Workbook, ByVal sourceUri As String) As String
    mp_BuildSourceUriCacheKey = mp_BuildWorkbookCacheKey(wb) & "|" & LCase$(Trim$(sourceUri))
End Function

Private Function mp_TryGetRelativeFileStamp(ByVal wb As Workbook, ByVal relPath As String, ByRef outStamp As Date) As Boolean
    Dim fullPath As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    fullPath = ex_XmlCore.m_CombineBasePath(wb, relPath)
    If Len(fullPath) = 0 Then Exit Function

    mp_TryGetRelativeFileStamp = mp_TryGetFileStamp(fullPath, outStamp)
End Function

Private Function mp_TryGetFileStamp(ByVal fullPath As String, ByRef outStamp As Date) As Boolean
    On Error GoTo EH
    If Len(Trim$(fullPath)) = 0 Then Exit Function
    If Len(Dir(fullPath)) = 0 Then Exit Function

    outStamp = FileDateTime(fullPath)
    mp_TryGetFileStamp = True
    Exit Function
EH:
    mp_TryGetFileStamp = False
End Function

Private Function mp_TryResolveSourceUriParts( _
    ByVal sourceUri As String, _
    ByVal wb As Workbook, _
    ByVal controlName As String, _
    ByRef outFilePath As String, _
    ByRef outListName As String) As Boolean

    Dim hashPos As Long
    Dim relPath As String
    Dim messagePrefix As String

    sourceUri = Trim$(sourceUri)
    If Len(sourceUri) = 0 Then Exit Function

    If Len(Trim$(controlName)) > 0 Then
        messagePrefix = "Invalid uri/sourceUri format '" & sourceUri & "' for control '" & controlName & "'. "
    Else
        messagePrefix = "Invalid uri/sourceUri format '" & sourceUri & "'. "
    End If

    hashPos = InStrRev(sourceUri, "#")
    If hashPos <= 1 Or hashPos >= Len(sourceUri) Then
        MsgBox messagePrefix & "Expected '<relative-path>#<list-name>'.", vbExclamation
        Exit Function
    End If

    relPath = Trim$(Left$(sourceUri, hashPos - 1))
    outListName = Trim$(Mid$(sourceUri, hashPos + 1))
    If Len(relPath) = 0 Or Len(outListName) = 0 Then
        MsgBox messagePrefix & "Path or list name is empty.", vbExclamation
        Exit Function
    End If

    outFilePath = ex_XmlCore.m_CombineBasePath(wb, relPath)
    If Len(outFilePath) = 0 Then Exit Function

    mp_TryResolveSourceUriParts = True
End Function

Private Function mp_GetControlNode(ByVal doc As Object, ByVal controlName As String) As Object
    Set mp_GetControlNode = doc.selectSingleNode("/p:uiDefinition/p:controls/p:control[@name=" & ex_XmlCore.m_XPathLiteral(controlName) & "]")
    If mp_GetControlNode Is Nothing Then
        MsgBox "Control '" & controlName & "' was not found in UI config.", vbExclamation
    End If
End Function

Private Sub mp_SetStyleValue(ByVal styleData As Object, ByVal key As String, ByVal value As String)
    value = Trim$(value)
    If Len(value) = 0 Then Exit Sub
    styleData(key) = value
End Sub
