Attribute VB_Name = "ex_StylePipelineEngine"
Option Explicit

Private Const UI_NS As String = "urn:excelprototype:profiles"
Private Const SHEET_SCOPE_MIN_COL As Long = 40
Private Const SHEET_SCOPE_MIN_ROW As Long = 100
Private Const SHEET_SCOPE_EXPAND_STEP As Long = 30

Public Function m_ApplyPageStyles(ByVal ws As Worksheet, ByVal wsUiDoc As Object) As Boolean
    Dim stylesByName As Object

    If ws Is Nothing Then
        MsgBox "PrototypeNew: worksheet is not specified for style pass.", vbExclamation
        Exit Function
    End If
    If wsUiDoc Is Nothing Then
        MsgBox "PrototypeNew: page UI document is not specified for style pass.", vbExclamation
        Exit Function
    End If

    Set stylesByName = mp_LoadControlStyles(wsUiDoc)
    If stylesByName Is Nothing Then Exit Function

    If Not mp_ApplyControlStyles(ws, stylesByName) Then Exit Function
    If Not mp_ApplyPipelineStageByName(ws, wsUiDoc, "default", True) Then Exit Function

    m_ApplyPageStyles = True
End Function

Public Function m_ApplyPageStyleStage( _
    ByVal ws As Worksheet, _
    ByVal wsUiDoc As Object, _
    ByVal stageName As String _
) As Boolean
    If ws Is Nothing Then
        MsgBox "PrototypeNew: worksheet is not specified for stage style pass.", vbExclamation
        Exit Function
    End If
    If wsUiDoc Is Nothing Then
        MsgBox "PrototypeNew: page UI document is not specified for stage style pass.", vbExclamation
        Exit Function
    End If

    stageName = Trim$(stageName)
    If Len(stageName) = 0 Then
        MsgBox "PrototypeNew: stage name is required for explicit style stage apply.", vbExclamation
        Exit Function
    End If

    If Not mp_ApplyPipelineStageByName(ws, wsUiDoc, stageName, True) Then Exit Function
    m_ApplyPageStyleStage = True
End Function

Private Function mp_ApplyControlStyles(ByVal ws As Worksheet, ByVal stylesByName As Object) As Boolean
    Dim shp As Shape
    Dim styleName As String

    If ws Is Nothing Then Exit Function
    If stylesByName Is Nothing Then Exit Function

    For Each shp In ws.Shapes
        If Not mp_IsControlShape(shp) Then GoTo ContinueShape

        styleName = LCase$(Trim$(mp_ReadShapeMetaValue(shp, "pn.style")))
        If Len(styleName) = 0 Then GoTo ContinueShape

        If Not stylesByName.Exists(styleName) Then
            MsgBox "PrototypeNew: control style '" & styleName & "' is not declared in <styles>.", vbExclamation
            Exit Function
        End If

        If Not mp_ApplyShapeStyle(shp, stylesByName(styleName), "controlStyle:" & styleName) Then Exit Function

ContinueShape:
    Next shp

    mp_ApplyControlStyles = True
End Function

Private Function mp_ApplyPipelineStageByName( _
    ByVal ws As Worksheet, _
    ByVal wsUiDoc As Object, _
    ByVal stageName As String, _
    ByVal stageMustExist As Boolean _
) As Boolean
    Dim stageNodes As Object
    Dim stageNode As Object
    Dim stageEnabled As Boolean
    Dim stageKey As String
    Dim requestedStageKey As String
    Dim stageFound As Boolean

    requestedStageKey = LCase$(Trim$(stageName))
    If Len(requestedStageKey) = 0 Then
        MsgBox "PrototypeNew: style stage name is required.", vbExclamation
        Exit Function
    End If

    Set stageNodes = wsUiDoc.selectNodes("/p:uiDefinition/p:styles/p:stylePipelineStage")
    If stageNodes Is Nothing Or stageNodes.Length = 0 Then
        If stageMustExist Then
            MsgBox "PrototypeNew: style stage '" & requestedStageKey & "' was not found.", vbExclamation
            Exit Function
        End If

        mp_ApplyPipelineStageByName = True
        Exit Function
    End If

    For Each stageNode In stageNodes
        If stageNode.NodeType <> 1 Then GoTo ContinueStage

        stageKey = LCase$(Trim$(ex_XmlCore.m_NodeAttrText(stageNode, "name")))
        If Len(stageKey) = 0 Then
            MsgBox "PrototypeNew: stylePipelineStage@name is required.", vbExclamation
            Exit Function
        End If
        If StrComp(stageKey, requestedStageKey, vbBinaryCompare) <> 0 Then GoTo ContinueStage

        If stageFound Then
            MsgBox "PrototypeNew: duplicate stylePipelineStage with name '" & stageKey & "'.", vbExclamation
            Exit Function
        End If
        stageFound = True

        If Not mp_TryReadNodeEnabled(stageNode, True, stageEnabled) Then Exit Function
        If Not stageEnabled Then GoTo ContinueStage

        If Not mp_ApplyStageLayers(ws, stageNode) Then Exit Function

ContinueStage:
    Next stageNode

    If Not stageFound Then
        If stageMustExist Then
            MsgBox "PrototypeNew: style stage '" & requestedStageKey & "' was not found.", vbExclamation
            Exit Function
        End If
    End If

    mp_ApplyPipelineStageByName = True
End Function

Private Function mp_ApplyStageLayers(ByVal ws As Worksheet, ByVal stageNode As Object) As Boolean
    Dim layerNode As Object
    Dim layerEnabled As Boolean

    For Each layerNode In stageNode.ChildNodes
        If layerNode.NodeType <> 1 Then GoTo ContinueLayer

        If StrComp(LCase$(CStr(layerNode.baseName)), "layer", vbBinaryCompare) <> 0 Then
            MsgBox "PrototypeNew: stylePipelineStage supports only <layer> children.", vbExclamation
            Exit Function
        End If

        If Not mp_TryReadNodeEnabled(layerNode, True, layerEnabled) Then Exit Function
        If Not layerEnabled Then GoTo ContinueLayer

        If Not mp_ApplyLayerRules(ws, layerNode) Then Exit Function

ContinueLayer:
    Next layerNode

    mp_ApplyStageLayers = True
End Function

Private Function mp_ApplyLayerRules(ByVal ws As Worksheet, ByVal layerNode As Object) As Boolean
    Dim ruleNode As Object
    Dim ruleEnabled As Boolean

    For Each ruleNode In layerNode.ChildNodes
        If ruleNode.NodeType <> 1 Then GoTo ContinueRule

        If StrComp(LCase$(CStr(ruleNode.baseName)), "rule", vbBinaryCompare) <> 0 Then
            MsgBox "PrototypeNew: layer supports only <rule> children.", vbExclamation
            Exit Function
        End If

        If Not mp_TryReadNodeEnabled(ruleNode, True, ruleEnabled) Then Exit Function
        If Not ruleEnabled Then GoTo ContinueRule

        If Not mp_ApplySingleRule(ws, ruleNode) Then Exit Function

ContinueRule:
    Next ruleNode

    mp_ApplyLayerRules = True
End Function

Private Function mp_ApplySingleRule(ByVal ws As Worksheet, ByVal ruleNode As Object) As Boolean
    Dim ruleTarget As String
    Dim selector As Object
    Dim declarations As Object
    Dim scopeRange As Range
    Dim columnScope As Range

    If ws Is Nothing Then Exit Function
    If ruleNode Is Nothing Then Exit Function

    ruleTarget = LCase$(Trim$(ex_XmlCore.m_NodeAttrText(ruleNode, "target")))
    If Len(ruleTarget) = 0 Then
        MsgBox "PrototypeNew: rule target is required.", vbExclamation
        Exit Function
    End If

    If Not mp_RuleTargetIsSupported(ruleTarget) Then
        MsgBox "PrototypeNew: unsupported rule target '" & ruleTarget & "'.", vbExclamation
        Exit Function
    End If

    If Not mp_TryReadRuleSelector(ruleNode, selector) Then Exit Function

    Set declarations = mp_ReadStyleDeclarations(ruleNode)
    If declarations Is Nothing Then Exit Function

    Select Case ruleTarget
        Case "row"
            If Not mp_TryResolveRowTargetScope(ws, selector, scopeRange, columnScope) Then Exit Function

        Case "column"
            If Not mp_TryResolveColumnTargetScope(ws, selector, scopeRange, columnScope) Then Exit Function

        Case "cell"
            If Not mp_TryResolveCellTargetScope(ws, selector, scopeRange, columnScope) Then Exit Function

        Case "range"
            If Not selector.Exists("address") Then
                MsgBox "PrototypeNew: range rule requires selector address.", vbExclamation
                Exit Function
            End If
            If Not mp_TryGetRangeByAddress(ws, CStr(selector("address")), scopeRange) Then Exit Function
            Set columnScope = scopeRange.EntireColumn

        Case "usedrange"
            Set scopeRange = mp_GetUsedScopeRange(ws)
            If scopeRange Is Nothing Then Exit Function
            Set columnScope = scopeRange.EntireColumn

        Case "sheet"
            Set scopeRange = mp_GetExpandedSheetScopeRange(ws)
            If scopeRange Is Nothing Then Exit Function
            Set columnScope = scopeRange.EntireColumn

        Case "controlpart"
            If Not mp_TryResolveControlPartTargetScope(ws, selector, scopeRange, columnScope) Then Exit Function
    End Select

    If scopeRange Is Nothing Then
        mp_ApplySingleRule = True
        Exit Function
    End If

    If Not mp_ApplyRangeDeclarations(scopeRange, columnScope, declarations, ruleTarget) Then Exit Function

    mp_ApplySingleRule = True
End Function

Private Function mp_LoadControlStyles(ByVal wsUiDoc As Object) As Object
    Dim result As Object
    Dim styleNodes As Object
    Dim styleNode As Object
    Dim styleName As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    Set styleNodes = wsUiDoc.selectNodes("/p:uiDefinition/p:styles/p:controlStyle")
    If styleNodes Is Nothing Then
        Set mp_LoadControlStyles = result
        Exit Function
    End If

    For Each styleNode In styleNodes
        styleName = LCase$(Trim$(ex_XmlCore.m_NodeAttrText(styleNode, "name")))
        If Len(styleName) = 0 Then
            MsgBox "PrototypeNew: controlStyle has empty name.", vbExclamation
            Exit Function
        End If

        If result.Exists(styleName) Then
            MsgBox "PrototypeNew: duplicate controlStyle '" & styleName & "'.", vbExclamation
            Exit Function
        End If

        result.Add styleName, mp_ReadStyleDeclarations(styleNode)
    Next styleNode

    Set mp_LoadControlStyles = result
End Function

Private Function mp_ReadStyleDeclarations(ByVal styleNode As Object) As Object
    Dim declarations As Object
    Dim inlineStyles As String

    If styleNode Is Nothing Then Exit Function

    Set declarations = CreateObject("Scripting.Dictionary")
    declarations.CompareMode = 1

    mp_TrySetDeclaration declarations, "backColor", ex_XmlCore.m_NodeAttrText(styleNode, "backColor")
    mp_TrySetDeclaration declarations, "textColor", ex_XmlCore.m_NodeAttrText(styleNode, "textColor")
    mp_TrySetDeclaration declarations, "fontColor", ex_XmlCore.m_NodeAttrText(styleNode, "fontColor")
    mp_TrySetDeclaration declarations, "borderColor", ex_XmlCore.m_NodeAttrText(styleNode, "borderColor")
    mp_TrySetDeclaration declarations, "borderWeight", ex_XmlCore.m_NodeAttrText(styleNode, "borderWeight")
    mp_TrySetDeclaration declarations, "fontName", ex_XmlCore.m_NodeAttrText(styleNode, "fontName")
    mp_TrySetDeclaration declarations, "fontSize", ex_XmlCore.m_NodeAttrText(styleNode, "fontSize")
    mp_TrySetDeclaration declarations, "fontBold", ex_XmlCore.m_NodeAttrText(styleNode, "fontBold")
    mp_TrySetDeclaration declarations, "fontItalic", ex_XmlCore.m_NodeAttrText(styleNode, "fontItalic")
    mp_TrySetDeclaration declarations, "horizontal", ex_XmlCore.m_NodeAttrText(styleNode, "horizontal")
    mp_TrySetDeclaration declarations, "vertical", ex_XmlCore.m_NodeAttrText(styleNode, "vertical")
    mp_TrySetDeclaration declarations, "overflow", ex_XmlCore.m_NodeAttrText(styleNode, "overflow")
    mp_TrySetDeclaration declarations, "width", ex_XmlCore.m_NodeAttrText(styleNode, "width")
    mp_TrySetDeclaration declarations, "rowHeight", ex_XmlCore.m_NodeAttrText(styleNode, "rowHeight")

    inlineStyles = ex_XmlCore.m_NodeAttrText(styleNode, "styles")
    If Len(Trim$(inlineStyles)) > 0 Then
        If Not mp_ParseInlineStyles(inlineStyles, declarations) Then Exit Function
    End If

    Set mp_ReadStyleDeclarations = declarations
End Function

Private Sub mp_TrySetDeclaration(ByVal declarations As Object, ByVal keyName As String, ByVal rawValue As String)
    Dim normalizedKey As String

    rawValue = Trim$(rawValue)
    If Len(rawValue) = 0 Then Exit Sub

    normalizedKey = mp_NormalizeStyleKey(keyName)
    If Not mp_IsSupportedStyleKey(normalizedKey) Then Exit Sub

    declarations(normalizedKey) = rawValue
End Sub

Private Function mp_ParseInlineStyles(ByVal stylesText As String, ByVal declarations As Object) As Boolean
    Dim normalized As String
    Dim parts As Variant
    Dim part As Variant
    Dim pairText As String
    Dim sepPos As Long
    Dim keyName As String
    Dim keyValue As String

    normalized = Replace(stylesText, vbCr, vbNullString)
    normalized = Replace(normalized, vbLf, vbNullString)
    normalized = Replace(normalized, "{", vbNullString)
    normalized = Replace(normalized, "}", vbNullString)

    parts = Split(normalized, ";")
    For Each part In parts
        pairText = Trim$(CStr(part))
        If Len(pairText) = 0 Then GoTo ContinuePart

        sepPos = InStr(1, pairText, ":", vbBinaryCompare)
        If sepPos <= 1 Or sepPos >= Len(pairText) Then
            MsgBox "PrototypeNew: invalid styles declaration segment '" & pairText & "'.", vbExclamation
            Exit Function
        End If

        keyName = mp_NormalizeStyleKey(Trim$(Left$(pairText, sepPos - 1)))
        keyValue = Trim$(Mid$(pairText, sepPos + 1))

        If Not mp_IsSupportedStyleKey(keyName) Then
            MsgBox "PrototypeNew: unsupported style key '" & keyName & "' in styles declaration.", vbExclamation
            Exit Function
        End If
        If Len(keyValue) = 0 Then
            MsgBox "PrototypeNew: empty style value for key '" & keyName & "'.", vbExclamation
            Exit Function
        End If

        declarations(keyName) = keyValue

ContinuePart:
    Next part

    mp_ParseInlineStyles = True
End Function

Private Function mp_NormalizeStyleKey(ByVal keyName As String) As String
    keyName = LCase$(Trim$(keyName))

    Select Case keyName
        Case "textcolor"
            mp_NormalizeStyleKey = "fontcolor"
        Case Else
            mp_NormalizeStyleKey = keyName
    End Select
End Function

Private Function mp_IsSupportedStyleKey(ByVal keyName As String) As Boolean
    Select Case LCase$(Trim$(keyName))
        Case "backcolor", "fontcolor", "bordercolor", "borderweight", "fontname", "fontsize", "fontbold", "fontitalic", "horizontal", "vertical", "overflow", "width", "rowheight"
            mp_IsSupportedStyleKey = True
    End Select
End Function

Private Function mp_TryReadRuleSelector(ByVal ruleNode As Object, ByRef outSelector As Object) As Boolean
    Dim selectorText As String
    Dim selectorParts As Variant
    Dim part As Variant
    Dim pairText As String
    Dim sepPos As Long
    Dim keyName As String
    Dim keyValue As String

    Set outSelector = CreateObject("Scripting.Dictionary")
    outSelector.CompareMode = 1

    selectorText = Trim$(ex_XmlCore.m_NodeAttrText(ruleNode, "selector"))
    If Len(selectorText) = 0 Then
        mp_TryReadRuleSelector = True
        Exit Function
    End If

    selectorParts = Split(selectorText, ";")
    For Each part In selectorParts
        pairText = Trim$(CStr(part))
        If Len(pairText) = 0 Then GoTo ContinueSelectorPart

        sepPos = InStr(1, pairText, "=", vbBinaryCompare)
        If sepPos <= 1 Or sepPos >= Len(pairText) Then
            MsgBox "PrototypeNew: invalid selector segment '" & pairText & "'.", vbExclamation
            Exit Function
        End If

        keyName = LCase$(Trim$(Left$(pairText, sepPos - 1)))
        keyValue = Trim$(Mid$(pairText, sepPos + 1))
        If Len(keyValue) = 0 Then
            MsgBox "PrototypeNew: selector value is empty for key '" & keyName & "'.", vbExclamation
            Exit Function
        End If

        Select Case keyName
            Case "col", "row", "address", "type", "name", "part"
                If outSelector.Exists(keyName) Then
                    MsgBox "PrototypeNew: duplicate selector key '" & keyName & "'.", vbExclamation
                    Exit Function
                End If
                outSelector(keyName) = keyValue

            Case Else
                MsgBox "PrototypeNew: unsupported selector key '" & keyName & "'.", vbExclamation
                Exit Function
        End Select

ContinueSelectorPart:
    Next part

    mp_TryReadRuleSelector = True
End Function

Private Function mp_RuleTargetIsSupported(ByVal targetName As String) As Boolean
    Select Case LCase$(Trim$(targetName))
        Case "row", "column", "cell", "range", "usedrange", "sheet", "controlpart"
            mp_RuleTargetIsSupported = True
    End Select
End Function

Private Function mp_TryResolveControlPartTargetScope( _
    ByVal ws As Worksheet, _
    ByVal selector As Object, _
    ByRef outScope As Range, _
    ByRef outColumnScope As Range _
) As Boolean
    Dim controlType As String
    Dim controlName As String
    Dim partName As String

    If ws Is Nothing Then Exit Function
    If selector Is Nothing Then
        MsgBox "PrototypeNew: controlPart rule requires selector.", vbExclamation
        Exit Function
    End If

    If Not selector.Exists("type") Then
        MsgBox "PrototypeNew: controlPart rule requires selector key 'type'.", vbExclamation
        Exit Function
    End If
    If Not selector.Exists("part") Then
        MsgBox "PrototypeNew: controlPart rule requires selector key 'part'.", vbExclamation
        Exit Function
    End If

    controlType = LCase$(Trim$(CStr(selector("type"))))
    partName = LCase$(Trim$(CStr(selector("part"))))
    If selector.Exists("name") Then
        controlName = LCase$(Trim$(CStr(selector("name"))))
    End If

    If Len(controlType) = 0 Then
        MsgBox "PrototypeNew: controlPart selector 'type' is empty.", vbExclamation
        Exit Function
    End If
    If Len(partName) = 0 Then
        MsgBox "PrototypeNew: controlPart selector 'part' is empty.", vbExclamation
        Exit Function
    End If

    If Not ex_ControlPartsRuntime.m_TryResolveControlPartScope( _
        ws, controlType, controlName, partName, outScope, outColumnScope) Then Exit Function

    mp_TryResolveControlPartTargetScope = True
End Function

Private Function mp_TryResolveRowTargetScope( _
    ByVal ws As Worksheet, _
    ByVal selector As Object, _
    ByRef outScope As Range, _
    ByRef outColumnScope As Range _
) As Boolean
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim colStart As Long
    Dim colEnd As Long

    If ws Is Nothing Then Exit Function
    If selector Is Nothing Then Exit Function

    If selector.Exists("address") Then
        If Not mp_TryGetRangeByAddress(ws, CStr(selector("address")), outScope) Then Exit Function
        Set outColumnScope = outScope.EntireColumn
        mp_TryResolveRowTargetScope = True
        Exit Function
    End If

    rowStart = 1
    rowEnd = mp_GetLastUsedRow(ws)
    colStart = 1
    colEnd = mp_GetLastUsedColumn(ws)

    If selector.Exists("row") Then
        If Not mp_TryResolveSpan(CStr(selector("row")), rowStart, rowEnd) Then
            MsgBox "PrototypeNew: invalid row selector span '" & CStr(selector("row")) & "'.", vbExclamation
            Exit Function
        End If
    End If
    If selector.Exists("col") Then
        If Not mp_TryResolveSpan(CStr(selector("col")), colStart, colEnd) Then
            MsgBox "PrototypeNew: invalid col selector span '" & CStr(selector("col")) & "'.", vbExclamation
            Exit Function
        End If
    End If

    If rowEnd < rowStart Then rowEnd = rowStart
    If colEnd < colStart Then colEnd = colStart

    Set outScope = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    Set outColumnScope = ws.Range(ws.Columns(colStart), ws.Columns(colEnd))
    mp_TryResolveRowTargetScope = True
End Function

Private Function mp_TryResolveColumnTargetScope( _
    ByVal ws As Worksheet, _
    ByVal selector As Object, _
    ByRef outScope As Range, _
    ByRef outColumnScope As Range _
) As Boolean
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim colStart As Long
    Dim colEnd As Long

    If ws Is Nothing Then Exit Function
    If selector Is Nothing Then Exit Function

    If selector.Exists("address") Then
        If Not mp_TryGetRangeByAddress(ws, CStr(selector("address")), outScope) Then Exit Function
        Set outColumnScope = outScope.EntireColumn
        mp_TryResolveColumnTargetScope = True
        Exit Function
    End If

    rowStart = 1
    rowEnd = mp_GetLastUsedRow(ws)
    colStart = 1
    colEnd = mp_GetLastUsedColumn(ws)

    If selector.Exists("row") Then
        If Not mp_TryResolveSpan(CStr(selector("row")), rowStart, rowEnd) Then
            MsgBox "PrototypeNew: invalid row selector span '" & CStr(selector("row")) & "'.", vbExclamation
            Exit Function
        End If
    End If
    If selector.Exists("col") Then
        If Not mp_TryResolveSpan(CStr(selector("col")), colStart, colEnd) Then
            MsgBox "PrototypeNew: invalid col selector span '" & CStr(selector("col")) & "'.", vbExclamation
            Exit Function
        End If
    End If

    If rowEnd < rowStart Then rowEnd = rowStart
    If colEnd < colStart Then colEnd = colStart

    Set outScope = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    Set outColumnScope = ws.Range(ws.Columns(colStart), ws.Columns(colEnd))
    mp_TryResolveColumnTargetScope = True
End Function

Private Function mp_TryResolveCellTargetScope( _
    ByVal ws As Worksheet, _
    ByVal selector As Object, _
    ByRef outScope As Range, _
    ByRef outColumnScope As Range _
) As Boolean
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim colStart As Long
    Dim colEnd As Long

    If ws Is Nothing Then Exit Function
    If selector Is Nothing Then Exit Function

    If selector.Exists("address") Then
        If Not mp_TryGetRangeByAddress(ws, CStr(selector("address")), outScope) Then Exit Function
        Set outColumnScope = outScope.EntireColumn
        mp_TryResolveCellTargetScope = True
        Exit Function
    End If

    If Not selector.Exists("row") Or Not selector.Exists("col") Then
        MsgBox "PrototypeNew: cell rule requires selector row+col or address.", vbExclamation
        Exit Function
    End If

    If Not mp_TryResolveSpan(CStr(selector("row")), rowStart, rowEnd) Then
        MsgBox "PrototypeNew: invalid row selector span '" & CStr(selector("row")) & "'.", vbExclamation
        Exit Function
    End If
    If Not mp_TryResolveSpan(CStr(selector("col")), colStart, colEnd) Then
        MsgBox "PrototypeNew: invalid col selector span '" & CStr(selector("col")) & "'.", vbExclamation
        Exit Function
    End If

    If rowEnd < rowStart Then rowEnd = rowStart
    If colEnd < colStart Then colEnd = colStart

    Set outScope = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    Set outColumnScope = ws.Range(ws.Columns(colStart), ws.Columns(colEnd))
    mp_TryResolveCellTargetScope = True
End Function

Private Function mp_TryResolveSpan(ByVal spanText As String, ByRef outStart As Long, ByRef outEnd As Long) As Boolean
    Dim normalized As String
    Dim parts As Variant

    normalized = Trim$(spanText)
    If Len(normalized) = 0 Then Exit Function

    If InStr(1, normalized, ":", vbBinaryCompare) > 0 Then
        parts = Split(normalized, ":")
        If UBound(parts) <> 1 Then Exit Function
        If Not IsNumeric(Trim$(CStr(parts(0)))) Then Exit Function
        If Not IsNumeric(Trim$(CStr(parts(1)))) Then Exit Function

        outStart = CLng(Trim$(CStr(parts(0))))
        outEnd = CLng(Trim$(CStr(parts(1))))
    Else
        If Not IsNumeric(normalized) Then Exit Function
        outStart = CLng(normalized)
        outEnd = outStart
    End If

    If outStart <= 0 Or outEnd <= 0 Then Exit Function
    If outEnd < outStart Then Exit Function

    mp_TryResolveSpan = True
End Function

Private Function mp_TryGetRangeByAddress(ByVal ws As Worksheet, ByVal addressText As String, ByRef outRange As Range) As Boolean
    If ws Is Nothing Then Exit Function

    addressText = Trim$(addressText)
    If Len(addressText) = 0 Then
        MsgBox "PrototypeNew: selector address is empty.", vbExclamation
        Exit Function
    End If

    On Error GoTo EH_RANGE
    Set outRange = ws.Range(addressText)
    On Error GoTo 0

    mp_TryGetRangeByAddress = True
    Exit Function

EH_RANGE:
    MsgBox "PrototypeNew: invalid selector address '" & addressText & "'.", vbExclamation
End Function

Private Function mp_ApplyRangeDeclarations( _
    ByVal targetRange As Range, _
    ByVal columnScope As Range, _
    ByVal declarations As Object, _
    ByVal contextName As String _
) As Boolean
    Dim colorValue As Long
    Dim sizeValue As Double
    Dim boolValue As Boolean
    Dim hAlign As Long
    Dim vAlign As Long
    Dim borderWeightValue As Variant
    Dim scopeColumns As Range

    If targetRange Is Nothing Then Exit Function
    If declarations Is Nothing Then
        mp_ApplyRangeDeclarations = True
        Exit Function
    End If

    If declarations.Exists("backcolor") Then
        If Not mp_TryParseColor(CStr(declarations("backcolor")), colorValue) Then
            MsgBox "PrototypeNew: invalid backColor in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.Interior.Color = colorValue
    End If

    If declarations.Exists("fontcolor") Then
        If Not mp_TryParseColor(CStr(declarations("fontcolor")), colorValue) Then
            MsgBox "PrototypeNew: invalid fontColor in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.Font.Color = colorValue
    End If

    If declarations.Exists("bordercolor") Then
        If Not mp_TryParseColor(CStr(declarations("bordercolor")), colorValue) Then
            MsgBox "PrototypeNew: invalid borderColor in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.Borders.LineStyle = xlContinuous
        targetRange.Borders.Color = colorValue
    End If

    If declarations.Exists("borderweight") Then
        If Not mp_TryParseCellBorderWeight(CStr(declarations("borderweight")), borderWeightValue) Then
            MsgBox "PrototypeNew: invalid borderWeight in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.Borders.LineStyle = xlContinuous
        targetRange.Borders.Weight = borderWeightValue
    End If

    If declarations.Exists("fontname") Then
        targetRange.Font.Name = CStr(declarations("fontname"))
    End If

    If declarations.Exists("fontsize") Then
        If Not mp_TryParsePositiveDouble(CStr(declarations("fontsize")), sizeValue) Then
            MsgBox "PrototypeNew: invalid fontSize in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.Font.Size = sizeValue
    End If

    If declarations.Exists("fontbold") Then
        If Not mp_TryParseBoolean(CStr(declarations("fontbold")), boolValue) Then
            MsgBox "PrototypeNew: invalid fontBold in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.Font.Bold = boolValue
    End If

    If declarations.Exists("fontitalic") Then
        If Not mp_TryParseBoolean(CStr(declarations("fontitalic")), boolValue) Then
            MsgBox "PrototypeNew: invalid fontItalic in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.Font.Italic = boolValue
    End If

    If declarations.Exists("horizontal") Then
        If Not mp_TryParseHorizontalAlignment(CStr(declarations("horizontal")), hAlign) Then
            MsgBox "PrototypeNew: invalid horizontal alignment in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.HorizontalAlignment = hAlign
    End If

    If declarations.Exists("vertical") Then
        If Not mp_TryParseVerticalAlignment(CStr(declarations("vertical")), vAlign) Then
            MsgBox "PrototypeNew: invalid vertical alignment in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.VerticalAlignment = vAlign
    End If

    If declarations.Exists("overflow") Then
        If Not mp_ApplyOverflow(CStr(declarations("overflow")), targetRange) Then
            MsgBox "PrototypeNew: invalid overflow value in " & contextName & ".", vbExclamation
            Exit Function
        End If
    End If

    If declarations.Exists("width") Then
        If Not mp_TryParsePositiveDouble(CStr(declarations("width")), sizeValue) Then
            MsgBox "PrototypeNew: invalid width in " & contextName & ".", vbExclamation
            Exit Function
        End If

        If columnScope Is Nothing Then
            Set scopeColumns = targetRange.EntireColumn
        Else
            Set scopeColumns = columnScope
        End If
        scopeColumns.ColumnWidth = sizeValue
    End If

    If declarations.Exists("rowheight") Then
        If Not mp_TryParsePositiveDouble(CStr(declarations("rowheight")), sizeValue) Then
            MsgBox "PrototypeNew: invalid rowHeight in " & contextName & ".", vbExclamation
            Exit Function
        End If
        targetRange.EntireRow.RowHeight = sizeValue
    End If

    mp_ApplyRangeDeclarations = True
End Function

Private Function mp_ApplyShapeStyle(ByVal shp As Shape, ByVal declarations As Object, ByVal styleName As String) As Boolean
    Dim colorValue As Long
    Dim sizeValue As Double
    Dim boolValue As Boolean

    If shp Is Nothing Then Exit Function
    If declarations Is Nothing Then
        mp_ApplyShapeStyle = True
        Exit Function
    End If

    If declarations.Exists("backcolor") Then
        If Not mp_TryParseColor(CStr(declarations("backcolor")), colorValue) Then
            MsgBox "PrototypeNew: invalid backColor in style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.Fill.Visible = msoTrue
        shp.Fill.ForeColor.RGB = colorValue
    End If

    If declarations.Exists("fontcolor") Then
        If Not mp_TryParseColor(CStr(declarations("fontcolor")), colorValue) Then
            MsgBox "PrototypeNew: invalid fontColor in style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        On Error Resume Next
        shp.TextFrame.Characters.Font.Color = colorValue
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = colorValue
        On Error GoTo 0
    End If

    If declarations.Exists("bordercolor") Then
        If Not mp_TryParseColor(CStr(declarations("bordercolor")), colorValue) Then
            MsgBox "PrototypeNew: invalid borderColor in style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = colorValue
    End If

    If declarations.Exists("borderweight") Then
        If Not mp_TryParseShapeBorderWeight(CStr(declarations("borderweight")), sizeValue) Then
            MsgBox "PrototypeNew: invalid borderWeight in style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        shp.Line.Weight = sizeValue
    End If

    If declarations.Exists("fontname") Then
        On Error Resume Next
        shp.TextFrame.Characters.Font.Name = CStr(declarations("fontname"))
        shp.TextFrame2.TextRange.Font.Name = CStr(declarations("fontname"))
        On Error GoTo 0
    End If

    If declarations.Exists("fontsize") Then
        If Not mp_TryParsePositiveDouble(CStr(declarations("fontsize")), sizeValue) Then
            MsgBox "PrototypeNew: invalid fontSize in style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        On Error Resume Next
        shp.TextFrame.Characters.Font.Size = sizeValue
        shp.TextFrame2.TextRange.Font.Size = sizeValue
        On Error GoTo 0
    End If

    If declarations.Exists("fontbold") Then
        If Not mp_TryParseBoolean(CStr(declarations("fontbold")), boolValue) Then
            MsgBox "PrototypeNew: invalid fontBold in style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        On Error Resume Next
        shp.TextFrame.Characters.Font.Bold = boolValue
        shp.TextFrame2.TextRange.Font.Bold = boolValue
        On Error GoTo 0
    End If

    If declarations.Exists("fontitalic") Then
        If Not mp_TryParseBoolean(CStr(declarations("fontitalic")), boolValue) Then
            MsgBox "PrototypeNew: invalid fontItalic in style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        On Error Resume Next
        shp.TextFrame.Characters.Font.Italic = boolValue
        shp.TextFrame2.TextRange.Font.Italic = boolValue
        On Error GoTo 0
    End If

    mp_ApplyShapeStyle = True
End Function

Private Function mp_ReadShapeMetaValue(ByVal shp As Shape, ByVal keyName As String) As String
    mp_ReadShapeMetaValue = Trim$(ex_ShapeMetaRuntime.m_GetShapeMetaValue(shp, keyName, vbNullString))
End Function

Private Function mp_IsControlShape(ByVal shp As Shape) As Boolean
    If shp Is Nothing Then Exit Function
    mp_IsControlShape = (Len(Trim$(mp_ReadShapeMetaValue(shp, "pn.control"))) > 0)
End Function

Private Function mp_TryReadNodeEnabled( _
    ByVal node As Object, _
    ByVal defaultValue As Boolean, _
    ByRef outEnabled As Boolean _
) As Boolean
    Dim rawValue As String

    If node Is Nothing Then Exit Function

    rawValue = Trim$(ex_XmlCore.m_NodeAttrText(node, "enabled"))
    If Len(rawValue) = 0 Then
        outEnabled = defaultValue
        mp_TryReadNodeEnabled = True
        Exit Function
    End If

    If Not mp_TryParseBoolean(rawValue, outEnabled) Then
        MsgBox "PrototypeNew: attribute 'enabled' must be boolean.", vbExclamation
        Exit Function
    End If

    mp_TryReadNodeEnabled = True
End Function

Private Function mp_TryParsePositiveDouble(ByVal valueText As String, ByRef outValue As Double) As Boolean
    Dim normalized As String
    Dim decimalSep As String

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    If IsNumeric(valueText) Then
        outValue = CDbl(valueText)
        If outValue <= 0# Then Exit Function
        mp_TryParsePositiveDouble = True
        Exit Function
    End If

    ' Locale-safe parse for values like "0.75" on systems with decimal separator ",".
    decimalSep = Application.DecimalSeparator
    If Len(decimalSep) = 0 Then decimalSep = "."

    normalized = Replace$(valueText, ".", decimalSep)
    normalized = Replace$(normalized, ",", decimalSep)

    If Not IsNumeric(normalized) Then Exit Function
    outValue = CDbl(normalized)
    If outValue <= 0# Then Exit Function
    mp_TryParsePositiveDouble = True
End Function

Private Function mp_TryParseBoolean(ByVal valueText As String, ByRef outValue As Boolean) As Boolean
    valueText = LCase$(Trim$(valueText))

    Select Case valueText
        Case "true", "1", "yes", "on"
            outValue = True
            mp_TryParseBoolean = True
        Case "false", "0", "no", "off"
            outValue = False
            mp_TryParseBoolean = True
    End Select
End Function

Private Function mp_TryParseColor(ByVal valueText As String, ByRef outColor As Long) As Boolean
    Dim r As Long
    Dim g As Long
    Dim b As Long

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    If Left$(valueText, 1) = "#" Then
        If Len(valueText) <> 7 Then Exit Function
        If Not mp_IsHexPair(Mid$(valueText, 2, 2)) Then Exit Function
        If Not mp_IsHexPair(Mid$(valueText, 4, 2)) Then Exit Function
        If Not mp_IsHexPair(Mid$(valueText, 6, 2)) Then Exit Function

        r = CLng("&H" & Mid$(valueText, 2, 2))
        g = CLng("&H" & Mid$(valueText, 4, 2))
        b = CLng("&H" & Mid$(valueText, 6, 2))
        outColor = RGB(r, g, b)
        mp_TryParseColor = True
        Exit Function
    End If

    If IsNumeric(valueText) Then
        outColor = CLng(valueText)
        mp_TryParseColor = True
    End If
End Function

Private Function mp_IsHexPair(ByVal pairText As String) As Boolean
    Dim value As Long

    On Error GoTo EH
    If Len(pairText) <> 2 Then Exit Function
    value = CLng("&H" & pairText)
    If value < 0 Or value > 255 Then Exit Function
    mp_IsHexPair = True
    Exit Function
EH:
    mp_IsHexPair = False
End Function

Private Function mp_TryParseShapeBorderWeight(ByVal valueText As String, ByRef outValue As Double) As Boolean
    valueText = LCase$(Trim$(valueText))

    Select Case valueText
        Case "hairline"
            outValue = 0.25
            mp_TryParseShapeBorderWeight = True
            Exit Function
        Case "thin"
            outValue = 0.75
            mp_TryParseShapeBorderWeight = True
            Exit Function
        Case "medium"
            outValue = 1.5
            mp_TryParseShapeBorderWeight = True
            Exit Function
        Case "thick"
            outValue = 2.25
            mp_TryParseShapeBorderWeight = True
            Exit Function
    End Select

    mp_TryParseShapeBorderWeight = mp_TryParsePositiveDouble(valueText, outValue)
End Function

Private Function mp_TryParseCellBorderWeight(ByVal valueText As String, ByRef outValue As Variant) As Boolean
    Dim numericValue As Double

    valueText = LCase$(Trim$(valueText))

    Select Case valueText
        Case "hairline"
            outValue = xlHairline
            mp_TryParseCellBorderWeight = True
            Exit Function
        Case "thin"
            outValue = xlThin
            mp_TryParseCellBorderWeight = True
            Exit Function
        Case "medium"
            outValue = xlMedium
            mp_TryParseCellBorderWeight = True
            Exit Function
        Case "thick"
            outValue = xlThick
            mp_TryParseCellBorderWeight = True
            Exit Function
    End Select

    If Not mp_TryParsePositiveDouble(valueText, numericValue) Then Exit Function
    outValue = numericValue
    mp_TryParseCellBorderWeight = True
End Function

Private Function mp_TryParseHorizontalAlignment(ByVal valueText As String, ByRef outValue As Long) As Boolean
    valueText = LCase$(Trim$(valueText))

    Select Case valueText
        Case "left"
            outValue = xlHAlignLeft
        Case "center"
            outValue = xlHAlignCenter
        Case "right"
            outValue = xlHAlignRight
        Case "general"
            outValue = xlHAlignGeneral
        Case Else
            Exit Function
    End Select

    mp_TryParseHorizontalAlignment = True
End Function

Private Function mp_TryParseVerticalAlignment(ByVal valueText As String, ByRef outValue As Long) As Boolean
    valueText = LCase$(Trim$(valueText))

    Select Case valueText
        Case "top"
            outValue = xlVAlignTop
        Case "center", "middle"
            outValue = xlVAlignCenter
        Case "bottom"
            outValue = xlVAlignBottom
        Case Else
            Exit Function
    End Select

    mp_TryParseVerticalAlignment = True
End Function

Private Function mp_ApplyOverflow(ByVal valueText As String, ByVal targetRange As Range) As Boolean
    valueText = LCase$(Trim$(valueText))

    Select Case valueText
        Case "wrap"
            targetRange.WrapText = True
        Case "clip"
            targetRange.WrapText = False
        Case Else
            Exit Function
    End Select

    mp_ApplyOverflow = True
End Function

Private Function mp_GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        mp_GetLastUsedRow = 1
    Else
        mp_GetLastUsedRow = lastCell.Row
        If mp_GetLastUsedRow <= 0 Then mp_GetLastUsedRow = 1
    End If
End Function

Private Function mp_GetLastUsedColumn(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        mp_GetLastUsedColumn = 1
    Else
        mp_GetLastUsedColumn = lastCell.Column
        If mp_GetLastUsedColumn <= 0 Then mp_GetLastUsedColumn = 1
    End If
End Function

Private Function mp_GetUsedScopeRange(ByVal ws As Worksheet) As Range
    Dim lastRow As Long
    Dim lastCol As Long

    If ws Is Nothing Then Exit Function

    lastRow = mp_GetLastUsedRow(ws)
    lastCol = mp_GetLastUsedColumn(ws)
    If lastRow < 1 Then lastRow = 1
    If lastCol < 1 Then lastCol = 1

    Set mp_GetUsedScopeRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Function

Private Function mp_GetExpandedSheetScopeRange(ByVal ws As Worksheet) As Range
    Dim usedScope As Range
    Dim usedLastRow As Long
    Dim usedLastCol As Long
    Dim endRow As Long
    Dim endCol As Long

    If ws Is Nothing Then Exit Function

    Set usedScope = mp_GetUsedScopeRange(ws)
    If usedScope Is Nothing Then Exit Function

    usedLastRow = usedScope.Row + usedScope.Rows.Count - 1
    usedLastCol = usedScope.Column + usedScope.Columns.Count - 1

    endCol = SHEET_SCOPE_MIN_COL
    If usedLastCol > SHEET_SCOPE_MIN_COL Then
        endCol = usedLastCol + SHEET_SCOPE_EXPAND_STEP
    End If

    endRow = SHEET_SCOPE_MIN_ROW
    If usedLastRow > SHEET_SCOPE_MIN_ROW Then
        endRow = usedLastRow + SHEET_SCOPE_EXPAND_STEP
    End If

    If endCol < 1 Then endCol = 1
    If endRow < 1 Then endRow = 1
    If endCol > ws.Columns.Count Then endCol = ws.Columns.Count
    If endRow > ws.Rows.Count Then endRow = ws.Rows.Count

    Set mp_GetExpandedSheetScopeRange = ws.Range(ws.Cells(1, 1), ws.Cells(endRow, endCol))
End Function
