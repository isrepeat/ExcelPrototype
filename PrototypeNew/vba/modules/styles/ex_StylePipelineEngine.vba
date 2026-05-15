Attribute VB_Name = "ex_StylePipelineEngine"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const UI_NS As String = "urn:excelprototype:profiles"
Private Const SHEET_SCOPE_MIN_COL As Long = 40
Private Const SHEET_SCOPE_MIN_ROW As Long = 100
Private Const SHEET_SCOPE_EXPAND_STEP As Long = 30
' Формат записи: Array(sheetName, rowStart, colStart, rowEnd, colEnd, tag, name, tagDepth)
Private m_LayoutBounds As Collection

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_StylePipelineEngine.fn_Module_Dispose"
#End If
    Set m_LayoutBounds = Nothing
End Sub
' //
' // API
' //
Public Function fn_ApplyPageStyles(ByVal ws As Worksheet, ByVal wsUiDoc As Object) As Boolean
    Dim stylesByName As Object

    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for style pass."
#End If
        Exit Function
    End If
    If wsUiDoc Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page UI document is not specified for style pass."
#End If
        Exit Function
    End If

    Set stylesByName = private_LoadControlStyles(wsUiDoc)
    If stylesByName Is Nothing Then Exit Function

    If Not private_ApplyControlStyles(ws, stylesByName) Then Exit Function
    If Not private_ApplyPipelineStageByName(ws, wsUiDoc, "default", True) Then Exit Function

    fn_ApplyPageStyles = True
End Function


Public Function fn_ApplyPageStyleStage( _
    ByVal ws As Worksheet, _
    ByVal wsUiDoc As Object, _
    ByVal stageName As String _
) As Boolean
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for stage style pass."
#End If
        Exit Function
    End If
    If wsUiDoc Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page UI document is not specified for stage style pass."
#End If
        Exit Function
    End If

    stageName = VBA.Trim$(stageName)
    If VBA.Len(stageName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: stage name is required for explicit style stage apply."
#End If
        Exit Function
    End If

    If Not private_ApplyPipelineStageByName(ws, wsUiDoc, stageName, True) Then Exit Function
    fn_ApplyPageStyleStage = True
End Function


Public Sub fn_ResetLayoutBounds()
    Set m_LayoutBounds = Nothing
End Sub


Public Sub fn_RegisterLayoutBound( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal tagName As String = "", _
    Optional ByVal nodeName As String = "", _
    Optional ByVal tagDepth As Long = -1 _
)
    If ws Is Nothing Then Exit Sub
    If rowStart <= 0 Or colStart <= 0 Then Exit Sub
    If rowEnd < rowStart Or colEnd < colStart Then Exit Sub

    If m_LayoutBounds Is Nothing Then
        Set m_LayoutBounds = New Collection
    End If

    m_LayoutBounds.Add Array( _
        ws.Name, _
        CLng(rowStart), _
        CLng(colStart), _
        CLng(rowEnd), _
        CLng(colEnd), _
        VBA.LCase$(VBA.Trim$(tagName)), _
        VBA.LCase$(VBA.Trim$(nodeName)), _
        CLng(tagDepth))
End Sub

' //
' // Internal
' //

Private Function private_ApplyControlStyles(ByVal ws As Worksheet, ByVal stylesByName As Object) As Boolean
    Dim shp As Shape
    Dim styleName As String

    If ws Is Nothing Then Exit Function
    If stylesByName Is Nothing Then Exit Function

    For Each shp In ws.Shapes
        If Not private_IsControlShape(shp) Then GoTo ContinueShape

        styleName = VBA.LCase$(VBA.Trim$(private_ReadShapeMetaValue(shp, "pn.style")))
        If VBA.Len(styleName) = 0 Then GoTo ContinueShape

        If Not stylesByName.Exists(styleName) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: control style '" & styleName & "' is not declared in <styles>."
#End If
            Exit Function
        End If

        If Not private_ApplyShapeStyle(shp, stylesByName(styleName), "controlStyle:" & styleName) Then Exit Function

ContinueShape:
    Next shp

    private_ApplyControlStyles = True
End Function


Private Function private_ApplyPipelineStageByName( _
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

    requestedStageKey = VBA.LCase$(VBA.Trim$(stageName))
    If VBA.Len(requestedStageKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: style stage name is required."
#End If
        Exit Function
    End If

    Set stageNodes = wsUiDoc.selectNodes("/p:page/p:styles/p:stylePipelineStage | /p:uiDefinition/p:styles/p:stylePipelineStage")
    If stageNodes Is Nothing Or stageNodes.Length = 0 Then
        If stageMustExist Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: style stage '" & requestedStageKey & "' was not found."
#End If
            Exit Function
        End If

        private_ApplyPipelineStageByName = True
        Exit Function
    End If

    For Each stageNode In stageNodes
        If stageNode.NodeType <> 1 Then GoTo ContinueStage

        stageKey = VBA.LCase$(VBA.Trim$(ex_XmlCore.fn_NodeAttrText(stageNode, "name")))
        If VBA.Len(stageKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: stylePipelineStage@name is required."
#End If
            Exit Function
        End If
        If VBA.StrComp(stageKey, requestedStageKey, VBA.vbBinaryCompare) <> 0 Then GoTo ContinueStage

        If stageFound Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: duplicate stylePipelineStage with name '" & stageKey & "'."
#End If
            Exit Function
        End If
        stageFound = True

        If Not private_TryReadNodeEnabled(stageNode, True, stageEnabled) Then Exit Function
        If Not stageEnabled Then GoTo ContinueStage

        If Not private_ApplyStageLayers(ws, stageNode) Then Exit Function

ContinueStage:
    Next stageNode

    If Not stageFound Then
        If stageMustExist Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: style stage '" & requestedStageKey & "' was not found."
#End If
            Exit Function
        End If
    End If

    private_ApplyPipelineStageByName = True
End Function


Private Function private_ApplyStageLayers(ByVal ws As Worksheet, ByVal stageNode As Object) As Boolean
    Dim layerNode As Object

    For Each layerNode In stageNode.ChildNodes
        If layerNode.NodeType <> 1 Then GoTo ContinueLayer

        If VBA.StrComp(VBA.LCase$(VBA.CStr(layerNode.baseName)), "layer", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: stylePipelineStage supports only <layer> children."
#End If
            Exit Function
        End If

        If Not private_ApplyLayerRules(ws, layerNode) Then Exit Function

ContinueLayer:
    Next layerNode

    private_ApplyStageLayers = True
End Function


Private Function private_ApplyLayerRules(ByVal ws As Worksheet, ByVal layerNode As Object) As Boolean
    Dim ruleNode As Object
    Dim ruleEnabled As Boolean

    For Each ruleNode In layerNode.ChildNodes
        If ruleNode.NodeType <> 1 Then GoTo ContinueRule

        If VBA.StrComp(VBA.LCase$(VBA.CStr(ruleNode.baseName)), "rule", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: layer supports only <rule> children."
#End If
            Exit Function
        End If

        If Not private_TryReadNodeEnabled(ruleNode, True, ruleEnabled) Then Exit Function
        If Not ruleEnabled Then GoTo ContinueRule

        If Not private_ApplySingleRule(ws, ruleNode) Then Exit Function

ContinueRule:
    Next ruleNode

    private_ApplyLayerRules = True
End Function


Private Function private_ApplySingleRule(ByVal ws As Worksheet, ByVal ruleNode As Object) As Boolean
    Dim ruleTarget As String
    Dim selector As Object
    Dim declarations As Object
    Dim scopeRange As Range
    Dim columnScope As Range

    If ws Is Nothing Then Exit Function
    If ruleNode Is Nothing Then Exit Function

    ruleTarget = VBA.LCase$(VBA.Trim$(ex_XmlCore.fn_NodeAttrText(ruleNode, "target")))
    If VBA.Len(ruleTarget) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: rule target is required."
#End If
        Exit Function
    End If

    If Not private_RuleTargetIsSupported(ruleTarget) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: unsupported rule target '" & ruleTarget & "'."
#End If
        Exit Function
    End If

    If Not private_TryReadRuleSelector(ruleNode, selector) Then Exit Function

    Set declarations = private_ReadStyleDeclarations(ruleNode)
    If declarations Is Nothing Then Exit Function

    Select Case ruleTarget
        Case "layoutbound"
            If Not private_ApplyLayoutBoundRule(ws, selector, declarations, "layoutBound rule") Then Exit Function
            private_ApplySingleRule = True
            Exit Function

        Case "row"
            If Not private_TryResolveRowTargetScope(ws, selector, scopeRange, columnScope) Then Exit Function

        Case "column"
            If Not private_TryResolveColumnTargetScope(ws, selector, scopeRange, columnScope) Then Exit Function

        Case "cell"
            If Not private_TryResolveCellTargetScope(ws, selector, scopeRange, columnScope) Then Exit Function

        Case "range"
            If Not selector.Exists("address") Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "PrototypeNew: range rule requires selector address."
#End If
                Exit Function
            End If
            If Not private_TryGetRangeByAddress(ws, VBA.CStr(selector("address")), scopeRange) Then Exit Function
            Set columnScope = scopeRange.EntireColumn

        Case "usedrange"
            Set scopeRange = private_GetUsedScopeRange(ws)
            If scopeRange Is Nothing Then Exit Function
            Set columnScope = scopeRange.EntireColumn

        Case "sheet"
            Set scopeRange = private_GetExpandedSheetScopeRange(ws)
            If scopeRange Is Nothing Then Exit Function
            Set columnScope = scopeRange.EntireColumn

        Case "controlpart"
            If Not private_TryResolveControlPartTargetScope(ws, selector, scopeRange, columnScope) Then Exit Function
    End Select

    If scopeRange Is Nothing Then
        private_ApplySingleRule = True
        Exit Function
    End If

    If Not private_ApplyRangeDeclarations(scopeRange, columnScope, declarations, ruleTarget) Then Exit Function

    private_ApplySingleRule = True
End Function


Private Function private_ApplyLayoutBoundRule( _
    ByVal ws As Worksheet, _
    ByVal selector As Object, _
    ByVal declarations As Object, _
    ByVal contextName As String _
) As Boolean
    Dim entry As Variant
    Dim hasBorderColor As Boolean
    Dim borderColor As Long
    Dim hasBorderWeight As Boolean
    Dim borderWeight As Variant
    Dim hasBorderLineStyle As Boolean
    Dim borderLineStyle As Variant
    Dim hasAnyBorderStyle As Boolean

    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for layoutBound style rule."
#End If
        Exit Function
    End If

    If declarations Is Nothing Then
        private_ApplyLayoutBoundRule = True
        Exit Function
    End If

    If Not private_TryReadLayoutBoundBorderDeclarations( _
        declarations, contextName, _
        hasAnyBorderStyle, _
        hasBorderColor, borderColor, _
        hasBorderWeight, borderWeight, _
        hasBorderLineStyle, borderLineStyle) Then Exit Function

    If Not hasAnyBorderStyle Then
        private_ApplyLayoutBoundRule = True
        Exit Function
    End If

    If m_LayoutBounds Is Nothing Then
        private_ApplyLayoutBoundRule = True
        Exit Function
    End If

    For Each entry In m_LayoutBounds
        If VBA.StrComp(VBA.CStr(entry(0)), ws.Name, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        If Not private_LayoutBoundEntryMatchesSelector(entry, selector) Then GoTo ContinueEntry

        If Not private_PaintLayoutBoundFrame( _
            ws:=ws, _
            rowStart:=CLng(entry(1)), _
            colStart:=CLng(entry(2)), _
            rowEnd:=CLng(entry(3)), _
            colEnd:=CLng(entry(4)), _
            tagName:=VBA.CStr(entry(5)), _
            hasBorderColor:=hasBorderColor, _
            borderColor:=borderColor, _
            hasBorderWeight:=hasBorderWeight, _
            borderWeight:=borderWeight, _
            hasBorderLineStyle:=hasBorderLineStyle, _
            borderLineStyle:=borderLineStyle) Then Exit Function

ContinueEntry:
    Next entry

    private_ApplyLayoutBoundRule = True
End Function


Private Function private_LayoutBoundEntryMatchesSelector(ByRef entry As Variant, ByVal selector As Object) As Boolean
    Dim selectorTag As String
    Dim selectorName As String
    Dim selectorTagDepth As String
    Dim depthMin As Long
    Dim depthMax As Long
    Dim nodeDepth As Long

    If selector Is Nothing Then
        private_LayoutBoundEntryMatchesSelector = True
        Exit Function
    End If

    If selector.Exists("tag") Then
        selectorTag = VBA.LCase$(VBA.Trim$(VBA.CStr(selector("tag"))))
        If VBA.Len(selectorTag) = 0 Then Exit Function
        If VBA.StrComp(selectorTag, VBA.CStr(entry(5)), VBA.vbBinaryCompare) <> 0 Then Exit Function
    End If

    If selector.Exists("name") Then
        selectorName = VBA.LCase$(VBA.Trim$(VBA.CStr(selector("name"))))
        If VBA.Len(selectorName) = 0 Then Exit Function
        If VBA.StrComp(selectorName, VBA.CStr(entry(6)), VBA.vbBinaryCompare) <> 0 Then Exit Function
    End If

    If selector.Exists("tagdepth") Then
        selectorTagDepth = VBA.Trim$(VBA.CStr(selector("tagdepth")))
        If Not private_TryParseLayoutBoundDepthSpan(selectorTagDepth, depthMin, depthMax) Then Exit Function

        nodeDepth = CLng(entry(7))
        If nodeDepth < depthMin Or nodeDepth > depthMax Then Exit Function
    End If

    private_LayoutBoundEntryMatchesSelector = True
End Function


Private Function private_TryParseLayoutBoundDepthSpan( _
    ByVal spanText As String, _
    ByRef outStart As Long, _
    ByRef outEnd As Long _
) As Boolean
    Dim normalized As String
    Dim parts As Variant

    normalized = VBA.Trim$(spanText)
    If VBA.Len(normalized) = 0 Then Exit Function

    If VBA.InStr(1, normalized, ":", VBA.vbBinaryCompare) > 0 Then
        parts = VBA.Split(normalized, ":")
        If UBound(parts) <> 1 Then Exit Function
        If Not VBA.IsNumeric(VBA.Trim$(VBA.CStr(parts(0)))) Then Exit Function
        If Not VBA.IsNumeric(VBA.Trim$(VBA.CStr(parts(1)))) Then Exit Function

        outStart = VBA.CLng(VBA.Trim$(VBA.CStr(parts(0))))
        outEnd = VBA.CLng(VBA.Trim$(VBA.CStr(parts(1))))
    Else
        If Not VBA.IsNumeric(normalized) Then Exit Function
        outStart = VBA.CLng(normalized)
        outEnd = outStart
    End If

    If outStart < 0 Or outEnd < 0 Then Exit Function
    If outEnd < outStart Then Exit Function

    private_TryParseLayoutBoundDepthSpan = True
End Function


Private Function private_PaintLayoutBoundFrame( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    ByVal tagName As String, _
    ByVal hasBorderColor As Boolean, _
    ByVal borderColor As Long, _
    ByVal hasBorderWeight As Boolean, _
    ByVal borderWeight As Variant, _
    ByVal hasBorderLineStyle As Boolean, _
    ByVal borderLineStyle As Variant _
) As Boolean
    Dim targetRange As Range
    Dim resolvedLineStyle As Variant

    If ws Is Nothing Then Exit Function
    If rowStart <= 0 Or colStart <= 0 Then Exit Function
    If rowEnd < rowStart Or colEnd < colStart Then Exit Function

    On Error GoTo EH_FRAME
    Set targetRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    If targetRange Is Nothing Then Exit Function

    resolvedLineStyle = xlContinuous
    If VBA.StrComp(tagName, "stackpanel", VBA.vbBinaryCompare) = 0 Then
        resolvedLineStyle = xlDash
    End If
    If hasBorderLineStyle Then
        resolvedLineStyle = borderLineStyle
    End If

    With targetRange.Borders(xlEdgeLeft)
        .LineStyle = resolvedLineStyle
        If hasBorderWeight Then .Weight = borderWeight
        If hasBorderColor Then .Color = borderColor
    End With
    With targetRange.Borders(xlEdgeTop)
        .LineStyle = resolvedLineStyle
        If hasBorderWeight Then .Weight = borderWeight
        If hasBorderColor Then .Color = borderColor
    End With
    With targetRange.Borders(xlEdgeRight)
        .LineStyle = resolvedLineStyle
        If hasBorderWeight Then .Weight = borderWeight
        If hasBorderColor Then .Color = borderColor
    End With
    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = resolvedLineStyle
        If hasBorderWeight Then .Weight = borderWeight
        If hasBorderColor Then .Color = borderColor
    End With

    private_PaintLayoutBoundFrame = True
    Exit Function

EH_FRAME:
    On Error GoTo 0
End Function


Private Function private_TryReadLayoutBoundBorderDeclarations( _
    ByVal declarations As Object, _
    ByVal contextName As String, _
    ByRef outHasAnyBorderStyle As Boolean, _
    ByRef outHasBorderColor As Boolean, _
    ByRef outBorderColor As Long, _
    ByRef outHasBorderWeight As Boolean, _
    ByRef outBorderWeight As Variant, _
    ByRef outHasBorderLineStyle As Boolean, _
    ByRef outBorderLineStyle As Variant _
) As Boolean
    Dim valueText As String

    outHasAnyBorderStyle = False

    If declarations.Exists("bordercolor") Then
        valueText = VBA.CStr(declarations("bordercolor"))
        If Not ex_HelpersCSS.fn_TryParseColor(valueText, outBorderColor) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid borderColor in " & contextName & "."
#End If
            Exit Function
        End If
        outHasBorderColor = True
        outHasAnyBorderStyle = True
    End If

    If declarations.Exists("borderweight") Then
        valueText = VBA.CStr(declarations("borderweight"))
        If Not ex_HelpersCSS.fn_TryParseCellBorderWeight(valueText, outBorderWeight) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid borderWeight in " & contextName & "."
#End If
            Exit Function
        End If
        outHasBorderWeight = True
        outHasAnyBorderStyle = True
    End If

    If declarations.Exists("borderlinestyle") Then
        valueText = VBA.CStr(declarations("borderlinestyle"))
        If Not private_TryParseLayoutBoundBorderLineStyle(valueText, outBorderLineStyle) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid borderLineStyle in " & contextName & "."
#End If
            Exit Function
        End If
        outHasBorderLineStyle = True
        outHasAnyBorderStyle = True
    End If

    private_TryReadLayoutBoundBorderDeclarations = True
End Function


Private Function private_TryParseLayoutBoundBorderLineStyle(ByVal valueText As String, ByRef outLineStyle As Variant) As Boolean
    valueText = VBA.LCase$(VBA.Trim$(valueText))

    Select Case valueText
        Case "solid", "continuous"
            outLineStyle = xlContinuous
        Case "dash", "dashed"
            outLineStyle = xlDash
        Case "dot", "dotted"
            outLineStyle = xlDot
        Case "dashdot"
            outLineStyle = xlDashDot
        Case "dashdotdot"
            outLineStyle = xlDashDotDot
        Case "double"
            outLineStyle = xlDouble
        Case "none"
            outLineStyle = xlLineStyleNone
        Case Else
            Exit Function
    End Select

    private_TryParseLayoutBoundBorderLineStyle = True
End Function


Private Function private_LoadControlStyles(ByVal wsUiDoc As Object) As Object
    Dim result As Object
    Dim styleNodes As Object
    Dim styleNode As Object
    Dim styleName As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    Set styleNodes = wsUiDoc.selectNodes("/p:page/p:styles/p:controlStyle | /p:uiDefinition/p:styles/p:controlStyle")
    If styleNodes Is Nothing Then
        Set private_LoadControlStyles = result
        Exit Function
    End If

    For Each styleNode In styleNodes
        styleName = VBA.LCase$(VBA.Trim$(ex_XmlCore.fn_NodeAttrText(styleNode, "name")))
        If VBA.Len(styleName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: controlStyle has empty name."
#End If
            Exit Function
        End If

        If result.Exists(styleName) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: duplicate controlStyle '" & styleName & "'."
#End If
            Exit Function
        End If

        result.Add styleName, private_ReadStyleDeclarations(styleNode)
    Next styleNode

    Set private_LoadControlStyles = result
End Function


Private Function private_ReadStyleDeclarations(ByVal styleNode As Object) As Object
    Dim declarations As Object
    Dim inlineStyles As String

    If styleNode Is Nothing Then Exit Function

    Set declarations = CreateObject("Scripting.Dictionary")
    declarations.CompareMode = 1

    private_TrySetDeclaration declarations, "backColor", ex_XmlCore.fn_NodeAttrText(styleNode, "backColor")
    private_TrySetDeclaration declarations, "textColor", ex_XmlCore.fn_NodeAttrText(styleNode, "textColor")
    private_TrySetDeclaration declarations, "fontColor", ex_XmlCore.fn_NodeAttrText(styleNode, "fontColor")
    private_TrySetDeclaration declarations, "borderColor", ex_XmlCore.fn_NodeAttrText(styleNode, "borderColor")
    private_TrySetDeclaration declarations, "borderWeight", ex_XmlCore.fn_NodeAttrText(styleNode, "borderWeight")
    private_TrySetDeclaration declarations, "borderLineStyle", ex_XmlCore.fn_NodeAttrText(styleNode, "borderLineStyle")
    private_TrySetDeclaration declarations, "fontName", ex_XmlCore.fn_NodeAttrText(styleNode, "fontName")
    private_TrySetDeclaration declarations, "fontSize", ex_XmlCore.fn_NodeAttrText(styleNode, "fontSize")
    private_TrySetDeclaration declarations, "fontBold", ex_XmlCore.fn_NodeAttrText(styleNode, "fontBold")
    private_TrySetDeclaration declarations, "fontItalic", ex_XmlCore.fn_NodeAttrText(styleNode, "fontItalic")
    private_TrySetDeclaration declarations, "horizontal", ex_XmlCore.fn_NodeAttrText(styleNode, "horizontal")
    private_TrySetDeclaration declarations, "vertical", ex_XmlCore.fn_NodeAttrText(styleNode, "vertical")
    private_TrySetDeclaration declarations, "overflow", ex_XmlCore.fn_NodeAttrText(styleNode, "overflow")
    private_TrySetDeclaration declarations, "width", ex_XmlCore.fn_NodeAttrText(styleNode, "width")
    private_TrySetDeclaration declarations, "rowHeight", ex_XmlCore.fn_NodeAttrText(styleNode, "rowHeight")

    inlineStyles = ex_XmlCore.fn_NodeAttrText(styleNode, "styles")
    If VBA.Len(VBA.Trim$(inlineStyles)) > 0 Then
        If Not private_ParseInlineStyles(inlineStyles, declarations) Then Exit Function
    End If

    Set private_ReadStyleDeclarations = declarations
End Function


Private Sub private_TrySetDeclaration(ByVal declarations As Object, ByVal keyName As String, ByVal rawValue As String)
    Dim normalizedKey As String

    rawValue = VBA.Trim$(rawValue)
    If VBA.Len(rawValue) = 0 Then Exit Sub

    normalizedKey = private_NormalizeStyleKey(keyName)
    If Not private_IsSupportedStyleKey(normalizedKey) Then Exit Sub

    declarations(normalizedKey) = rawValue
End Sub


Private Function private_ParseInlineStyles(ByVal stylesText As String, ByVal declarations As Object) As Boolean
    Dim normalized As String
    Dim parts As Variant
    Dim part As Variant
    Dim pairText As String
    Dim sepPos As Long
    Dim keyName As String
    Dim keyValue As String

    normalized = VBA.Replace(stylesText, VBA.vbCr, VBA.vbNullString)
    normalized = VBA.Replace(normalized, VBA.vbLf, VBA.vbNullString)
    normalized = VBA.Replace(normalized, "{", VBA.vbNullString)
    normalized = VBA.Replace(normalized, "}", VBA.vbNullString)

    parts = VBA.Split(normalized, ";")
    For Each part In parts
        pairText = VBA.Trim$(VBA.CStr(part))
        If VBA.Len(pairText) = 0 Then GoTo ContinuePart

        sepPos = VBA.InStr(1, pairText, ":", VBA.vbBinaryCompare)
        If sepPos <= 1 Or sepPos >= VBA.Len(pairText) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid styles declaration segment '" & pairText & "'."
#End If
            Exit Function
        End If

        keyName = private_NormalizeStyleKey(VBA.Trim$(VBA.Left$(pairText, sepPos - 1)))
        keyValue = VBA.Trim$(VBA.Mid$(pairText, sepPos + 1))

        If Not private_IsSupportedStyleKey(keyName) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: unsupported style key '" & keyName & "' in styles declaration."
#End If
            Exit Function
        End If
        If VBA.Len(keyValue) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: empty style value for key '" & keyName & "'."
#End If
            Exit Function
        End If

        declarations(keyName) = keyValue

ContinuePart:
    Next part

    private_ParseInlineStyles = True
End Function


Private Function private_NormalizeStyleKey(ByVal keyName As String) As String
    keyName = VBA.LCase$(VBA.Trim$(keyName))

    Select Case keyName
        Case "textcolor"
            private_NormalizeStyleKey = "fontcolor"
        Case Else
            private_NormalizeStyleKey = keyName
    End Select
End Function


Private Function private_IsSupportedStyleKey(ByVal keyName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(keyName))
        Case "backcolor", "fontcolor", "bordercolor", "borderweight", "borderlinestyle", "fontname", "fontsize", "fontbold", "fontitalic", "horizontal", "vertical", "overflow", "width", "rowheight"
            private_IsSupportedStyleKey = True
    End Select
End Function


Private Function private_TryReadRuleSelector(ByVal ruleNode As Object, ByRef outSelector As Object) As Boolean
    Dim selectorText As String
    Dim selectorParts As Variant
    Dim part As Variant
    Dim pairText As String
    Dim sepPos As Long
    Dim keyName As String
    Dim keyValue As String

    Set outSelector = CreateObject("Scripting.Dictionary")
    outSelector.CompareMode = 1

    selectorText = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(ruleNode, "selector"))
    If VBA.Len(selectorText) = 0 Then
        private_TryReadRuleSelector = True
        Exit Function
    End If

    selectorParts = VBA.Split(selectorText, ";")
    For Each part In selectorParts
        pairText = VBA.Trim$(VBA.CStr(part))
        If VBA.Len(pairText) = 0 Then GoTo ContinueSelectorPart

        sepPos = VBA.InStr(1, pairText, "=", VBA.vbBinaryCompare)
        If sepPos <= 1 Or sepPos >= VBA.Len(pairText) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid selector segment '" & pairText & "'."
#End If
            Exit Function
        End If

        keyName = VBA.LCase$(VBA.Trim$(VBA.Left$(pairText, sepPos - 1)))
        keyValue = VBA.Trim$(VBA.Mid$(pairText, sepPos + 1))
        If VBA.Len(keyValue) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: selector value is empty for key '" & keyName & "'."
#End If
            Exit Function
        End If

        Select Case keyName
            Case "col", "row", "address", "type", "name", "part", "tag", "tagdepth"
                If outSelector.Exists(keyName) Then
#If LOGGING_DEBUG_ENABLED Then
                    ex_Core.fn_Diagnostic_LogError "PrototypeNew: duplicate selector key '" & keyName & "'."
#End If
                    Exit Function
                End If
                outSelector(keyName) = keyValue

            Case Else
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "PrototypeNew: unsupported selector key '" & keyName & "'."
#End If
                Exit Function
        End Select

ContinueSelectorPart:
    Next part

    private_TryReadRuleSelector = True
End Function


Private Function private_RuleTargetIsSupported(ByVal targetName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(targetName))
        Case "row", "column", "cell", "range", "usedrange", "sheet", "controlpart", "layoutbound"
            private_RuleTargetIsSupported = True
    End Select
End Function


Private Function private_TryResolveControlPartTargetScope( _
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
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: controlPart rule requires selector."
#End If
        Exit Function
    End If

    If Not selector.Exists("type") Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: controlPart rule requires selector key 'type'."
#End If
        Exit Function
    End If
    If Not selector.Exists("part") Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: controlPart rule requires selector key 'part'."
#End If
        Exit Function
    End If

    controlType = VBA.LCase$(VBA.Trim$(VBA.CStr(selector("type"))))
    partName = VBA.LCase$(VBA.Trim$(VBA.CStr(selector("part"))))
    If selector.Exists("name") Then
        controlName = VBA.LCase$(VBA.Trim$(VBA.CStr(selector("name"))))
    End If

    If VBA.Len(controlType) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: controlPart selector 'type' is empty."
#End If
        Exit Function
    End If
    If VBA.Len(partName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: controlPart selector 'part' is empty."
#End If
        Exit Function
    End If

    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, controlType, controlName, partName, outScope, outColumnScope) Then Exit Function

    private_TryResolveControlPartTargetScope = True
End Function


Private Function private_TryResolveRowTargetScope( _
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
        If Not private_TryGetRangeByAddress(ws, VBA.CStr(selector("address")), outScope) Then Exit Function
        Set outColumnScope = outScope.EntireColumn
        private_TryResolveRowTargetScope = True
        Exit Function
    End If

    rowStart = 1
    rowEnd = private_GetLastUsedRow(ws)
    colStart = 1
    colEnd = private_GetLastUsedColumn(ws)

    If selector.Exists("row") Then
        If Not private_TryResolveSpan(VBA.CStr(selector("row")), rowStart, rowEnd) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid row selector span '" & VBA.CStr(selector("row")) & "'."
#End If
            Exit Function
        End If
    End If
    If selector.Exists("col") Then
        If Not private_TryResolveSpan(VBA.CStr(selector("col")), colStart, colEnd) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid col selector span '" & VBA.CStr(selector("col")) & "'."
#End If
            Exit Function
        End If
    End If

    If rowEnd < rowStart Then rowEnd = rowStart
    If colEnd < colStart Then colEnd = colStart

    Set outScope = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    Set outColumnScope = ws.Range(ws.Columns(colStart), ws.Columns(colEnd))
    private_TryResolveRowTargetScope = True
End Function


Private Function private_TryResolveColumnTargetScope( _
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
        If Not private_TryGetRangeByAddress(ws, VBA.CStr(selector("address")), outScope) Then Exit Function
        Set outColumnScope = outScope.EntireColumn
        private_TryResolveColumnTargetScope = True
        Exit Function
    End If

    rowStart = 1
    rowEnd = private_GetLastUsedRow(ws)
    colStart = 1
    colEnd = private_GetLastUsedColumn(ws)

    If selector.Exists("row") Then
        If Not private_TryResolveSpan(VBA.CStr(selector("row")), rowStart, rowEnd) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid row selector span '" & VBA.CStr(selector("row")) & "'."
#End If
            Exit Function
        End If
    End If
    If selector.Exists("col") Then
        If Not private_TryResolveSpan(VBA.CStr(selector("col")), colStart, colEnd) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid col selector span '" & VBA.CStr(selector("col")) & "'."
#End If
            Exit Function
        End If
    End If

    If rowEnd < rowStart Then rowEnd = rowStart
    If colEnd < colStart Then colEnd = colStart

    Set outScope = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    Set outColumnScope = ws.Range(ws.Columns(colStart), ws.Columns(colEnd))
    private_TryResolveColumnTargetScope = True
End Function


Private Function private_TryResolveCellTargetScope( _
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
        If Not private_TryGetRangeByAddress(ws, VBA.CStr(selector("address")), outScope) Then Exit Function
        Set outColumnScope = outScope.EntireColumn
        private_TryResolveCellTargetScope = True
        Exit Function
    End If

    If Not selector.Exists("row") Or Not selector.Exists("col") Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: cell rule requires selector row+col or address."
#End If
        Exit Function
    End If

    If Not private_TryResolveSpan(VBA.CStr(selector("row")), rowStart, rowEnd) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid row selector span '" & VBA.CStr(selector("row")) & "'."
#End If
        Exit Function
    End If
    If Not private_TryResolveSpan(VBA.CStr(selector("col")), colStart, colEnd) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid col selector span '" & VBA.CStr(selector("col")) & "'."
#End If
        Exit Function
    End If

    If rowEnd < rowStart Then rowEnd = rowStart
    If colEnd < colStart Then colEnd = colStart

    Set outScope = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    Set outColumnScope = ws.Range(ws.Columns(colStart), ws.Columns(colEnd))
    private_TryResolveCellTargetScope = True
End Function


Private Function private_TryResolveSpan(ByVal spanText As String, ByRef outStart As Long, ByRef outEnd As Long) As Boolean
    Dim normalized As String
    Dim parts As Variant

    normalized = VBA.Trim$(spanText)
    If VBA.Len(normalized) = 0 Then Exit Function

    If VBA.InStr(1, normalized, ":", VBA.vbBinaryCompare) > 0 Then
        parts = VBA.Split(normalized, ":")
        If UBound(parts) <> 1 Then Exit Function
        If Not VBA.IsNumeric(VBA.Trim$(VBA.CStr(parts(0)))) Then Exit Function
        If Not VBA.IsNumeric(VBA.Trim$(VBA.CStr(parts(1)))) Then Exit Function

        outStart = VBA.CLng(VBA.Trim$(VBA.CStr(parts(0))))
        outEnd = VBA.CLng(VBA.Trim$(VBA.CStr(parts(1))))
    Else
        If Not VBA.IsNumeric(normalized) Then Exit Function
        outStart = VBA.CLng(normalized)
        outEnd = outStart
    End If

    If outStart <= 0 Or outEnd <= 0 Then Exit Function
    If outEnd < outStart Then Exit Function

    private_TryResolveSpan = True
End Function


Private Function private_TryGetRangeByAddress(ByVal ws As Worksheet, ByVal addressText As String, ByRef outRange As Range) As Boolean
    If ws Is Nothing Then Exit Function

    addressText = VBA.Trim$(addressText)
    If VBA.Len(addressText) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: selector address is empty."
#End If
        Exit Function
    End If

    On Error GoTo EH_RANGE
    Set outRange = ws.Range(addressText)
    On Error GoTo 0

    private_TryGetRangeByAddress = True
    Exit Function

EH_RANGE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid selector address '" & addressText & "'."
#End If
End Function


Private Function private_ApplyRangeDeclarations( _
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
        private_ApplyRangeDeclarations = True
        Exit Function
    End If

    If declarations.Exists("backcolor") Then
        If Not ex_HelpersCSS.fn_TryParseColor(VBA.CStr(declarations("backcolor")), colorValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid backColor in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.Interior.Color = colorValue
    End If

    If declarations.Exists("fontcolor") Then
        If Not ex_HelpersCSS.fn_TryParseColor(VBA.CStr(declarations("fontcolor")), colorValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid fontColor in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.Font.Color = colorValue
    End If

    If declarations.Exists("bordercolor") Then
        If Not ex_HelpersCSS.fn_TryParseColor(VBA.CStr(declarations("bordercolor")), colorValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid borderColor in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.Borders.LineStyle = xlContinuous
        targetRange.Borders.Color = colorValue
    End If

    If declarations.Exists("borderweight") Then
        If Not ex_HelpersCSS.fn_TryParseCellBorderWeight(VBA.CStr(declarations("borderweight")), borderWeightValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid borderWeight in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.Borders.LineStyle = xlContinuous
        targetRange.Borders.Weight = borderWeightValue
    End If

    If declarations.Exists("fontname") Then
        targetRange.Font.Name = VBA.CStr(declarations("fontname"))
    End If

    If declarations.Exists("fontsize") Then
        If Not ex_HelpersCSS.fn_TryParsePositiveDouble(VBA.CStr(declarations("fontsize")), sizeValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid fontSize in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.Font.Size = sizeValue
    End If

    If declarations.Exists("fontbold") Then
        If Not private_TryParseBoolean(VBA.CStr(declarations("fontbold")), boolValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid fontBold in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.Font.Bold = boolValue
    End If

    If declarations.Exists("fontitalic") Then
        If Not private_TryParseBoolean(VBA.CStr(declarations("fontitalic")), boolValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid fontItalic in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.Font.Italic = boolValue
    End If

    If declarations.Exists("horizontal") Then
        If Not private_TryParseHorizontalAlignment(VBA.CStr(declarations("horizontal")), hAlign) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid horizontal alignment in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.HorizontalAlignment = hAlign
    End If

    If declarations.Exists("vertical") Then
        If Not private_TryParseVerticalAlignment(VBA.CStr(declarations("vertical")), vAlign) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid vertical alignment in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.VerticalAlignment = vAlign
    End If

    If declarations.Exists("overflow") Then
        If Not private_ApplyOverflow(VBA.CStr(declarations("overflow")), targetRange) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid overflow value in " & contextName & "."
#End If
            Exit Function
        End If
    End If

    If declarations.Exists("width") Then
        If Not ex_HelpersCSS.fn_TryParsePositiveDouble(VBA.CStr(declarations("width")), sizeValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid width in " & contextName & "."
#End If
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
        If Not ex_HelpersCSS.fn_TryParsePositiveDouble(VBA.CStr(declarations("rowheight")), sizeValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid rowHeight in " & contextName & "."
#End If
            Exit Function
        End If
        targetRange.EntireRow.RowHeight = sizeValue
    End If

    private_ApplyRangeDeclarations = True
End Function


Private Function private_ApplyShapeStyle(ByVal shp As Shape, ByVal declarations As Object, ByVal styleName As String) As Boolean
    Dim colorValue As Long
    Dim sizeValue As Double
    Dim boolValue As Boolean

    If shp Is Nothing Then Exit Function
    If declarations Is Nothing Then
        private_ApplyShapeStyle = True
        Exit Function
    End If

    If declarations.Exists("backcolor") Then
        If Not ex_HelpersCSS.fn_TryParseColor(VBA.CStr(declarations("backcolor")), colorValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid backColor in style '" & styleName & "'."
#End If
            Exit Function
        End If
        shp.Fill.Visible = msoTrue
        shp.Fill.ForeColor.RGB = colorValue
    End If

    If declarations.Exists("fontcolor") Then
        If Not ex_HelpersCSS.fn_TryParseColor(VBA.CStr(declarations("fontcolor")), colorValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid fontColor in style '" & styleName & "'."
#End If
            Exit Function
        End If
        On Error Resume Next
        shp.TextFrame.Characters.Font.Color = colorValue
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = colorValue
        On Error GoTo 0
    End If

    If declarations.Exists("bordercolor") Then
        If Not ex_HelpersCSS.fn_TryParseColor(VBA.CStr(declarations("bordercolor")), colorValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid borderColor in style '" & styleName & "'."
#End If
            Exit Function
        End If
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = colorValue
    End If

    If declarations.Exists("borderweight") Then
        If Not ex_HelpersCSS.fn_TryParseShapeBorderWeight(VBA.CStr(declarations("borderweight")), sizeValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid borderWeight in style '" & styleName & "'."
#End If
            Exit Function
        End If
        shp.Line.Weight = sizeValue
    End If

    If declarations.Exists("fontname") Then
        On Error Resume Next
        shp.TextFrame.Characters.Font.Name = VBA.CStr(declarations("fontname"))
        shp.TextFrame2.TextRange.Font.Name = VBA.CStr(declarations("fontname"))
        On Error GoTo 0
    End If

    If declarations.Exists("fontsize") Then
        If Not ex_HelpersCSS.fn_TryParsePositiveDouble(VBA.CStr(declarations("fontsize")), sizeValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid fontSize in style '" & styleName & "'."
#End If
            Exit Function
        End If
        On Error Resume Next
        shp.TextFrame.Characters.Font.Size = sizeValue
        shp.TextFrame2.TextRange.Font.Size = sizeValue
        On Error GoTo 0
    End If

    If declarations.Exists("fontbold") Then
        If Not private_TryParseBoolean(VBA.CStr(declarations("fontbold")), boolValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid fontBold in style '" & styleName & "'."
#End If
            Exit Function
        End If
        On Error Resume Next
        shp.TextFrame.Characters.Font.Bold = boolValue
        shp.TextFrame2.TextRange.Font.Bold = boolValue
        On Error GoTo 0
    End If

    If declarations.Exists("fontitalic") Then
        If Not private_TryParseBoolean(VBA.CStr(declarations("fontitalic")), boolValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: invalid fontItalic in style '" & styleName & "'."
#End If
            Exit Function
        End If
        On Error Resume Next
        shp.TextFrame.Characters.Font.Italic = boolValue
        shp.TextFrame2.TextRange.Font.Italic = boolValue
        On Error GoTo 0
    End If

    private_ApplyShapeStyle = True
End Function


Private Function private_ReadShapeMetaValue(ByVal shp As Shape, ByVal keyName As String) As String
    private_ReadShapeMetaValue = VBA.Trim$(ex_ShapeMetaRuntime.fn_GetShapeMetaValue(shp, keyName, VBA.vbNullString))
End Function


Private Function private_IsControlShape(ByVal shp As Shape) As Boolean
    If shp Is Nothing Then Exit Function
    private_IsControlShape = (VBA.Len(VBA.Trim$(private_ReadShapeMetaValue(shp, "pn.control"))) > 0)
End Function


Private Function private_TryReadNodeEnabled( _
    ByVal node As Object, _
    ByVal defaultValue As Boolean, _
    ByRef outEnabled As Boolean _
) As Boolean
    Dim rawValue As String

    If node Is Nothing Then Exit Function

    rawValue = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(node, "enabled"))
    If VBA.Len(rawValue) = 0 Then
        outEnabled = defaultValue
        private_TryReadNodeEnabled = True
        Exit Function
    End If

    If Not private_TryParseBoolean(rawValue, outEnabled) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: attribute 'enabled' must be boolean."
#End If
        Exit Function
    End If

    private_TryReadNodeEnabled = True
End Function


Private Function private_TryParseBoolean(ByVal valueText As String, ByRef outValue As Boolean) As Boolean
    valueText = VBA.LCase$(VBA.Trim$(valueText))

    Select Case valueText
        Case "true", "1", "yes", "on"
            outValue = True
            private_TryParseBoolean = True
        Case "false", "0", "no", "off"
            outValue = False
            private_TryParseBoolean = True
    End Select
End Function


Private Function private_TryParseHorizontalAlignment(ByVal valueText As String, ByRef outValue As Long) As Boolean
    valueText = VBA.LCase$(VBA.Trim$(valueText))

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

    private_TryParseHorizontalAlignment = True
End Function


Private Function private_TryParseVerticalAlignment(ByVal valueText As String, ByRef outValue As Long) As Boolean
    valueText = VBA.LCase$(VBA.Trim$(valueText))

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

    private_TryParseVerticalAlignment = True
End Function


Private Function private_ApplyOverflow(ByVal valueText As String, ByVal targetRange As Range) As Boolean
    valueText = VBA.LCase$(VBA.Trim$(valueText))

    Select Case valueText
        Case "wrap"
            targetRange.WrapText = True
        Case "clip"
            targetRange.WrapText = False
        Case Else
            Exit Function
    End Select

    private_ApplyOverflow = True
End Function


Private Function private_GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        private_GetLastUsedRow = 1
    Else
        private_GetLastUsedRow = lastCell.Row
        If private_GetLastUsedRow <= 0 Then private_GetLastUsedRow = 1
    End If
End Function


Private Function private_GetLastUsedColumn(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        private_GetLastUsedColumn = 1
    Else
        private_GetLastUsedColumn = lastCell.Column
        If private_GetLastUsedColumn <= 0 Then private_GetLastUsedColumn = 1
    End If
End Function


Private Function private_GetUsedScopeRange(ByVal ws As Worksheet) As Range
    Dim lastRow As Long
    Dim lastCol As Long

    If ws Is Nothing Then Exit Function

    lastRow = private_GetLastUsedRow(ws)
    lastCol = private_GetLastUsedColumn(ws)
    If lastRow < 1 Then lastRow = 1
    If lastCol < 1 Then lastCol = 1

    Set private_GetUsedScopeRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Function


Private Function private_GetExpandedSheetScopeRange(ByVal ws As Worksheet) As Range
    Dim usedScope As Range
    Dim usedLastRow As Long
    Dim usedLastCol As Long
    Dim controlsLastRow As Long
    Dim controlsLastCol As Long
    Dim hasControlsBounds As Boolean
    Dim endRow As Long
    Dim endCol As Long

    If ws Is Nothing Then Exit Function

    Set usedScope = private_GetUsedScopeRange(ws)
    If usedScope Is Nothing Then Exit Function

    usedLastRow = usedScope.Row + usedScope.Rows.Count - 1
    usedLastCol = usedScope.Column + usedScope.Columns.Count - 1

    ' Основной режим: размер sheet-scope привязан к layout/grid (по фактическим bounds контролов) + запас.
    hasControlsBounds = ex_ControlRefreshRuntime.fn_TryGetSheetMaxControlBounds(ws.Name, controlsLastRow, controlsLastCol)
    If hasControlsBounds Then
        endRow = controlsLastRow + SHEET_SCOPE_EXPAND_STEP
        endCol = controlsLastCol + SHEET_SCOPE_EXPAND_STEP
    Else
        ' Fallback для страниц без зарегистрированных контролов.
        endCol = SHEET_SCOPE_MIN_COL
        If usedLastCol > SHEET_SCOPE_MIN_COL Then
            endCol = usedLastCol + SHEET_SCOPE_EXPAND_STEP
        End If

        endRow = SHEET_SCOPE_MIN_ROW
        If usedLastRow > SHEET_SCOPE_MIN_ROW Then
            endRow = usedLastRow + SHEET_SCOPE_EXPAND_STEP
        End If
    End If

    If endCol < 1 Then endCol = 1
    If endRow < 1 Then endRow = 1
    If endCol > ws.Columns.Count Then endCol = ws.Columns.Count
    If endRow > ws.Rows.Count Then endRow = ws.Rows.Count

    Set private_GetExpandedSheetScopeRange = ws.Range(ws.Cells(1, 1), ws.Cells(endRow, endCol))
End Function
