Attribute VB_Name = "ex_StylePipelineEngine"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const STYLE_PIPELINE_REL_PATH As String = "config\StylePipeline.xml"
Private Const INLINE_LAYER_ID As String = "profileInline"
Private Const INLINE_LAYER_PRIORITY As Long = 100
Private Const SHEET_SCOPE_MIN_COL As Long = 40      ' AN
Private Const SHEET_SCOPE_MIN_ROW As Long = 100
Private Const SHEET_SCOPE_EXPAND_STEP As Long = 30

' Supported style properties (declarations):
' width, minWidth, maxWidth, autoFitColumns
' overflow, autoHeight, rowHeight, mergeColumns
' fontName, fontSize, fontBold
' backColor, fontColor
' borderColor, borderWeight
' horizontal, vertical
Private Const STYLE_PROP_WIDTH As String = "width"
Private Const STYLE_PROP_MIN_WIDTH As String = "minwidth"
Private Const STYLE_PROP_MAX_WIDTH As String = "maxwidth"
Private Const STYLE_PROP_AUTO_FIT_COLUMNS As String = "autofitcolumns"
Private Const STYLE_PROP_OVERFLOW As String = "overflow"
Private Const STYLE_PROP_AUTO_HEIGHT As String = "autoheight"
Private Const STYLE_PROP_ROW_HEIGHT As String = "rowheight"
Private Const STYLE_PROP_MERGE_COLUMNS As String = "mergecolumns"
Private Const STYLE_PROP_FONT_NAME As String = "fontname"
Private Const STYLE_PROP_FONT_SIZE As String = "fontsize"
Private Const STYLE_PROP_FONT_BOLD As String = "fontbold"
Private Const STYLE_PROP_BACK_COLOR As String = "backcolor"
Private Const STYLE_PROP_FONT_COLOR As String = "fontcolor"
Private Const STYLE_PROP_BORDER_COLOR As String = "bordercolor"
Private Const STYLE_PROP_BORDER_WEIGHT As String = "borderweight"
Private Const STYLE_PROP_HORIZONTAL As String = "horizontal"
Private Const STYLE_PROP_VERTICAL As String = "vertical"

Private g_IsRuntimeCacheInitialized As Boolean
Private g_LayersCache As Object
Private g_SelectorCache As Object

Private g_StylePipelineDocCache As Object
Private g_StylePipelineDocWbKey As String
Private g_StylePipelineDocStamp As Date

Public Function m_CreatePipeline() As Collection
    Set m_CreatePipeline = New Collection
End Function

Public Sub m_ResetRuntimeCaches()
    g_IsRuntimeCacheInitialized = False
    Set g_LayersCache = Nothing
    Set g_SelectorCache = Nothing
    Set g_StylePipelineDocCache = Nothing
    g_StylePipelineDocWbKey = vbNullString
    g_StylePipelineDocStamp = 0
End Sub

Public Sub m_AddLayer(ByVal pipeline As Collection, ByVal layer As obj_StyleLayer)
    If pipeline Is Nothing Then Exit Sub
    If layer Is Nothing Then Exit Sub
    pipeline.Add layer
End Sub

Private Sub mp_ApplyRowKindRule( _
    ByVal ws As Worksheet, _
    ByVal selector As Object, _
    ByVal declarations As Object, _
    ByVal rowKindRanges As Object, _
    ByVal ruleId As String, _
    Optional ByVal autoHeightState As Object = Nothing, _
    Optional ByVal autoHeightOnly As Boolean = False _
)
    Dim kindName As String
    Dim kindRanges As Collection
    Dim rowEntry As Variant
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim colStart As Long
    Dim colEnd As Long
    Dim scopeRange As Range

    kindName = LCase$(Trim$(CStr(selector("kind"))))
    If Len(kindName) = 0 Then
        Err.Raise vbObjectError + 1732, "ex_StylePipelineEngine", _
            "Row rule '" & ruleId & "' has empty selector kind."
    End If

    If rowKindRanges Is Nothing Then Exit Sub
    If Not rowKindRanges.Exists(kindName) Then Exit Sub

    Set kindRanges = rowKindRanges(kindName)
    If kindRanges Is Nothing Then Exit Sub
    If kindRanges.Count = 0 Then Exit Sub

    colStart = 1
    colEnd = mp_GetLastUsedColumn(ws)
    If colEnd < colStart Then colEnd = colStart
    If selector.Exists("col") Then
        If Not mp_TryResolveColumnSpan(CStr(selector("col")), colStart, colEnd) Then
            Err.Raise vbObjectError + 1723, "ex_StylePipelineEngine", _
                "Invalid selector col span for rule '" & ruleId & "': " & CStr(selector("col"))
        End If
    End If

    For Each rowEntry In kindRanges
        If Not mp_TryResolveRowEntry(rowEntry, rowStart, rowEnd) Then GoTo ContinueRow
        If rowStart <= 0 Then GoTo ContinueRow
        If rowEnd < rowStart Then rowEnd = rowStart

        Set scopeRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
        mp_ApplyDeclarations scopeRange, Nothing, declarations, Nothing, autoHeightState, autoHeightOnly
ContinueRow:
    Next rowEntry
End Sub

Public Sub m_InsertLayer(ByVal pipeline As Collection, ByVal atIndex As Long, ByVal layer As obj_StyleLayer)
    If pipeline Is Nothing Then Exit Sub
    If layer Is Nothing Then Exit Sub

    If pipeline.Count = 0 Then
        pipeline.Add layer
    ElseIf atIndex <= 1 Then
        pipeline.Add layer, Before:=1
    ElseIf atIndex > pipeline.Count Then
        pipeline.Add layer
    Else
        pipeline.Add layer, Before:=atIndex
    End If
End Sub

Public Function m_RemoveLayer(ByVal pipeline As Collection, ByVal layerId As String) As Boolean
    Dim i As Long
    Dim candidate As obj_StyleLayer
    Dim normalizedId As String

    If pipeline Is Nothing Then Exit Function
    normalizedId = Trim$(layerId)
    If Len(normalizedId) = 0 Then Exit Function

    For i = pipeline.Count To 1 Step -1
        Set candidate = pipeline(i)
        If Not candidate Is Nothing Then
            If StrComp(candidate.LayerId, normalizedId, vbTextCompare) = 0 Then
                pipeline.Remove i
                m_RemoveLayer = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Function m_FindLayerIndex(ByVal pipeline As Collection, ByVal layerId As String) As Long
    Dim i As Long
    Dim candidate As obj_StyleLayer
    Dim normalizedId As String

    If pipeline Is Nothing Then Exit Function
    normalizedId = Trim$(layerId)
    If Len(normalizedId) = 0 Then Exit Function

    For i = 1 To pipeline.Count
        Set candidate = pipeline(i)
        If Not candidate Is Nothing Then
            If StrComp(candidate.LayerId, normalizedId, vbTextCompare) = 0 Then
                m_FindLayerIndex = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Function m_GetLayer(ByVal pipeline As Collection, ByVal layerId As String) As obj_StyleLayer
    Dim idx As Long

    idx = m_FindLayerIndex(pipeline, layerId)
    If idx <= 0 Then Exit Function
    Set m_GetLayer = pipeline(idx)
End Function

Public Function m_GetSortedLayers(ByVal pipeline As Collection) As Collection
    Dim result As Collection
    Dim i As Long
    Dim j As Long
    Dim shouldSwap As Boolean
    Dim ordered() As obj_StyleLayer
    Dim tmp As obj_StyleLayer

    Set result = New Collection
    If pipeline Is Nothing Then
        Set m_GetSortedLayers = result
        Exit Function
    End If
    If pipeline.Count = 0 Then
        Set m_GetSortedLayers = result
        Exit Function
    End If

    ReDim ordered(1 To pipeline.Count)
    For i = 1 To pipeline.Count
        Set ordered(i) = pipeline(i)
    Next i

    For i = 1 To UBound(ordered) - 1
        For j = i + 1 To UBound(ordered)
            shouldSwap = False
            If ordered(i) Is Nothing And Not ordered(j) Is Nothing Then
                shouldSwap = True
            ElseIf Not ordered(i) Is Nothing And Not ordered(j) Is Nothing Then
                If ordered(j).Priority < ordered(i).Priority Then
                    shouldSwap = True
                End If
            End If

            If shouldSwap Then
                Set tmp = ordered(i)
                Set ordered(i) = ordered(j)
                Set ordered(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To UBound(ordered)
        If Not ordered(i) Is Nothing Then result.Add ordered(i)
    Next i

    Set m_GetSortedLayers = result
End Function

Public Function m_LoadSheetPipelineLayers( _
    ByVal pageName As String, _
    ByVal wb As Workbook _
) As Collection
    Dim doc As Object
    Dim result As Collection
    Dim seenLayerIds As Object
    Dim formatLayers As Collection
    Dim layerObj As obj_StyleLayer
    Dim cacheKey As String

    Set result = New Collection
    Set seenLayerIds = mp_CreateStringDictionary()
    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        Set m_LoadSheetPipelineLayers = result
        Exit Function
    End If

    Set doc = mp_GetStylePipelineDomCached(wb)
    If doc Is Nothing Then
        Set m_LoadSheetPipelineLayers = result
        Exit Function
    End If

    mp_EnsureRuntimeCaches
    cacheKey = mp_BuildLayersCacheKey(wb, pageName)
    If g_LayersCache.Exists(cacheKey) Then
        Set m_LoadSheetPipelineLayers = g_LayersCache(cacheKey)
        Exit Function
    End If

    Set formatLayers = mp_LoadLayersFromSheetPipelineXml(doc, pageName)
    If Not formatLayers Is Nothing Then
        For Each layerObj In formatLayers
            If Not layerObj Is Nothing Then
                If Not seenLayerIds.Exists(LCase$(layerObj.LayerId)) Then
                    seenLayerIds(LCase$(layerObj.LayerId)) = True
                    result.Add layerObj
                End If
            End If
        Next layerObj
    End If

    Set g_LayersCache(cacheKey) = result
    Set m_LoadSheetPipelineLayers = result
End Function

Public Function m_BuildColumnStylesPipeline( _
    ByVal resultFieldRanges As Collection, _
    ByVal cfgStyles As Object, _
    ByVal activeModeName As String, _
    ByVal wb As Workbook, _
    Optional ByVal pageName As String = vbNullString _
) As Collection
    Dim pipeline As Collection
    Dim inlineLayer As obj_StyleLayer
    Dim xmlLayers As Collection
    Dim xmlLayer As obj_StyleLayer

    Set pipeline = m_CreatePipeline()

    Set inlineLayer = mp_BuildInlineLayer(resultFieldRanges, cfgStyles)
    If Not inlineLayer Is Nothing Then m_AddLayer pipeline, inlineLayer

    Set xmlLayers = m_LoadSheetPipelineLayers(pageName, wb)
    If Not xmlLayers Is Nothing Then
        For Each xmlLayer In xmlLayers
            m_AddLayer pipeline, xmlLayer
        Next xmlLayer
    End If

    Set m_BuildColumnStylesPipeline = pipeline
End Function

Public Function m_ValidateColumnStylesPipeline( _
    ByVal resultFieldRanges As Collection, _
    ByVal cfgStyles As Object, _
    ByVal activeModeName As String, _
    ByRef outErrorText As String, _
    ByVal wb As Workbook, _
    Optional ByVal pageName As String = vbNullString _
) As Boolean
    Dim pipeline As Collection
    Dim sortedLayers As Collection
    Dim layerObj As obj_StyleLayer
    Dim ruleObj As obj_StyleRule
    Dim declarations As Object
    Dim stepName As String

    On Error GoTo EH
    outErrorText = vbNullString

    stepName = "build-pipeline"
    Set pipeline = m_BuildColumnStylesPipeline(resultFieldRanges, cfgStyles, activeModeName, wb, pageName)

    stepName = "sort-layers"
    Set sortedLayers = m_GetSortedLayers(pipeline)

    stepName = "validate-layers"
    For Each layerObj In sortedLayers
        If layerObj Is Nothing Then GoTo ContinueLayer
        If Not layerObj.IsEnabled Then GoTo ContinueLayer

        stepName = "validate-rules"
        For Each ruleObj In layerObj.Rules
            If ruleObj Is Nothing Then GoTo ContinueRule
            stepName = "validate-rule-target"
            If Not mp_IsSupportedTarget(ruleObj.Target) Then
                outErrorText = "Unsupported style rule target '" & ruleObj.Target & "' in layer '" & layerObj.LayerId & "'."
                Exit Function
            End If
            stepName = "validate-rule-declarations"
            Set declarations = ruleObj.Declarations
            If Not mp_ValidateDeclarations(declarations, outErrorText) Then Exit Function
ContinueRule:
        Next ruleObj
ContinueLayer:
    Next layerObj

    m_ValidateColumnStylesPipeline = True
    Exit Function
EH:
    outErrorText = "Style pipeline validation failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Public Sub m_ApplyColumnStylesPipeline( _
    ByVal ws As Worksheet, _
    ByVal resultFieldRanges As Collection, _
    ByVal pipeline As Collection, _
    ByVal activeModeName As String, _
    Optional ByVal rowKindRanges As Object = Nothing, _
    Optional ByVal autoHeightOnly As Boolean = False _
)
    Dim sortedLayers As Collection
    Dim layerObj As obj_StyleLayer
    Dim ruleObj As obj_StyleRule
    Dim targetsByMapKey As Object
    Dim autoHeightState As Object

    If ws Is Nothing Then Exit Sub
    If pipeline Is Nothing Then Exit Sub
    If pipeline.Count = 0 Then Exit Sub

    Set autoHeightState = mp_CreateStringDictionary()

    If Not resultFieldRanges Is Nothing Then
        If resultFieldRanges.Count > 0 Then
            Set targetsByMapKey = mp_BuildResultTargetsByMapKey(resultFieldRanges)
        End If
    End If

    Set sortedLayers = m_GetSortedLayers(pipeline)
    For Each layerObj In sortedLayers
        If layerObj Is Nothing Then GoTo ContinueLayer
        If Not layerObj.IsEnabled Then GoTo ContinueLayer

        For Each ruleObj In layerObj.Rules
            If ruleObj Is Nothing Then GoTo ContinueRule
            mp_ApplyRule ws, resultFieldRanges, ruleObj, activeModeName, rowKindRanges, targetsByMapKey, autoHeightState, autoHeightOnly
ContinueRule:
        Next ruleObj
ContinueLayer:
    Next layerObj

    mp_ApplyDeferredAutoHeight ws, autoHeightState
End Sub

Private Function mp_LoadLayersFromSheetPipelineXml( _
    ByVal doc As Object, _
    ByVal pageName As String _
) As Collection
    Dim result As Collection
    Dim sheetPipelines As Object
    Dim sheetPipelineNode As Object
    Dim layerNodes As Object
    Dim layerNode As Object
    Dim layerObj As obj_StyleLayer
    Dim layerId As String
    Dim layerSource As String
    Dim layerEnabled As Boolean
    Dim layerPriority As Long
    Dim ok As Boolean
    Dim pageKey As String

    Set result = New Collection
    If doc Is Nothing Then
        Set mp_LoadLayersFromSheetPipelineXml = result
        Exit Function
    End If

    Set sheetPipelines = doc.selectNodes("/p:stylePipeline/p:sheetPipeline")
    If sheetPipelines Is Nothing Then
        Set mp_LoadLayersFromSheetPipelineXml = result
        Exit Function
    End If

    For Each sheetPipelineNode In sheetPipelines
        pageKey = Trim$(mp_NodeAttrText(sheetPipelineNode, "page"))
        If Len(pageKey) = 0 Then
            Err.Raise vbObjectError + 1736, "ex_StylePipelineEngine", _
                "sheetPipeline@page is required."
        End If
        If Len(Trim$(pageName)) > 0 Then
            If StrComp(pageKey, Trim$(pageName), vbTextCompare) <> 0 Then GoTo ContinueSheetPipeline
        End If

        Set layerNodes = sheetPipelineNode.selectNodes("p:layer")
        If layerNodes Is Nothing Then GoTo ContinueSheetPipeline

        For Each layerNode In layerNodes
            layerId = Trim$(mp_NodeAttrText(layerNode, "id"))
            If Len(layerId) = 0 Then
                Err.Raise vbObjectError + 1710, "ex_StylePipelineEngine", "sheetPipeline/layer@" & "id is required."
            End If
            If Not ex_XmlCore.m_TryParseLong(mp_NodeAttrText(layerNode, "priority"), layerPriority) Then
                Err.Raise vbObjectError + 1711, "ex_StylePipelineEngine", "Invalid style layer priority for '" & layerId & "'."
            End If
            layerSource = Trim$(mp_NodeAttrText(layerNode, "source"))
            If Len(layerSource) = 0 Then layerSource = "sheetPipeline"

            ok = mp_TryParseBoolean(mp_NodeAttrText(layerNode, "enabled"), layerEnabled)
            If Not ok Then layerEnabled = True

            Set layerObj = New obj_StyleLayer
            layerObj.Initialize layerId, layerPriority, layerSource, layerEnabled
            mp_ParseLayerRules layerNode, layerObj

            result.Add layerObj
ContinueLayer:
        Next layerNode
ContinueSheetPipeline:
    Next sheetPipelineNode

    Set mp_LoadLayersFromSheetPipelineXml = result
End Function

Private Sub mp_ParseLayerRules(ByVal layerNode As Object, ByVal layerObj As obj_StyleLayer)
    Dim ruleNodes As Object
    Dim ruleNode As Object
    Dim ruleObj As obj_StyleRule
    Dim ruleId As String
    Dim targetName As String
    Dim selectorText As String
    Dim stylesText As String
    Dim declarations As Object
    Dim hasDecl As Boolean
    Dim parseError As String
    Dim ruleIndex As Long

    If layerNode Is Nothing Then Exit Sub
    If layerObj Is Nothing Then Exit Sub

    Set ruleNodes = layerNode.selectNodes("p:rule")
    If ruleNodes Is Nothing Then Exit Sub

    ruleIndex = 0
    For Each ruleNode In ruleNodes
        ruleIndex = ruleIndex + 1
        ruleId = layerObj.LayerId & ".rule" & CStr(ruleIndex)

        targetName = LCase$(Trim$(mp_NodeAttrText(ruleNode, "target")))
        If Len(targetName) = 0 Then
            Err.Raise vbObjectError + 1712, "ex_StylePipelineEngine", "Missing target in style rule '" & ruleId & "'."
        End If

        selectorText = Trim$(mp_NodeAttrText(ruleNode, "selector"))
        stylesText = Trim$(mp_NodeAttrText(ruleNode, "styles"))
        If Len(stylesText) = 0 Then stylesText = Trim$(CStr(ruleNode.Text))

        If Len(stylesText) = 0 And StrComp(targetName, "pipelinestep", vbTextCompare) = 0 Then
            Set declarations = mp_CreateStringDictionary()
        Else
            parseError = vbNullString
            hasDecl = False
            If Not ex_ConfigStylesParser.m_TryParseStyleDeclarations(stylesText, declarations, hasDecl, parseError) Then
                Err.Raise vbObjectError + 1713, "ex_StylePipelineEngine", _
                    "Invalid style rule declarations for '" & ruleId & "': " & parseError
            End If
            If Not hasDecl Then
                Err.Raise vbObjectError + 1714, "ex_StylePipelineEngine", _
                    "Style rule '" & ruleId & "' must define declarations."
            End If
        End If

        Set ruleObj = New obj_StyleRule
        ruleObj.Initialize ruleId, targetName, selectorText, declarations
        layerObj.AddRule ruleObj
    Next ruleNode
End Sub

Private Function mp_BuildInlineLayer( _
    ByVal resultFieldRanges As Collection, _
    ByVal cfgStyles As Object _
) As obj_StyleLayer
    Dim layerObj As obj_StyleLayer
    Dim target As Object
    Dim mapKey As String
    Dim styleText As String
    Dim declarations As Object
    Dim hasDecl As Boolean
    Dim parseError As String
    Dim dedupe As Object
    Dim ruleObj As obj_StyleRule
    Dim ruleId As String
    Dim ruleIndex As Long

    If resultFieldRanges Is Nothing Then Exit Function
    If cfgStyles Is Nothing Then Exit Function
    If resultFieldRanges.Count = 0 Then Exit Function

    Set dedupe = CreateObject("Scripting.Dictionary")
    dedupe.CompareMode = 1

    Set layerObj = New obj_StyleLayer
    layerObj.Initialize INLINE_LAYER_ID, INLINE_LAYER_PRIORITY, "profileInline", True

    For Each target In resultFieldRanges
        If target Is Nothing Then GoTo ContinueTarget
        mapKey = Trim$(CStr(target("MapKey")))
        If Len(mapKey) = 0 Then GoTo ContinueTarget
        If dedupe.Exists(mapKey) Then GoTo ContinueTarget
        dedupe(mapKey) = True

        If Not cfgStyles.Exists(mapKey) Then GoTo ContinueTarget
        styleText = Trim$(CStr(cfgStyles(mapKey)))
        If Len(styleText) = 0 Then GoTo ContinueTarget

        parseError = vbNullString
        hasDecl = False
        If Not ex_ConfigStylesParser.m_TryParseStyleDeclarations(styleText, declarations, hasDecl, parseError) Then
            Err.Raise vbObjectError + 1715, "ex_StylePipelineEngine", _
                "Invalid inline styles for key '" & mapKey & "': " & parseError
        End If
        If Not hasDecl Then GoTo ContinueTarget

        ruleIndex = ruleIndex + 1
        ruleId = INLINE_LAYER_ID & ".rule" & CStr(ruleIndex)
        Set ruleObj = New obj_StyleRule
        ruleObj.Initialize ruleId, "column", "mapKey=" & mapKey, declarations
        layerObj.AddRule ruleObj

ContinueTarget:
    Next target

    If layerObj.RuleCount = 0 Then Exit Function
    Set mp_BuildInlineLayer = layerObj
End Function

Private Sub mp_ApplyRule( _
    ByVal ws As Worksheet, _
    ByVal resultFieldRanges As Collection, _
    ByVal ruleObj As obj_StyleRule, _
    ByVal activeModeName As String, _
    Optional ByVal rowKindRanges As Object = Nothing, _
    Optional ByVal targetsByMapKey As Object = Nothing, _
    Optional ByVal autoHeightState As Object = Nothing, _
    Optional ByVal autoHeightOnly As Boolean = False _
)
    Dim selector As Object
    Dim targetName As String
    Dim scopeRange As Range
    Dim columnScope As Range
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim colStart As Long
    Dim colEnd As Long
    Dim addressText As String
    Dim hasSpan As Boolean
    Dim target As Object
    Dim targetGroup As Collection
    Dim mapCol As Long
    Dim mapRowStart As Long
    Dim mapRowEnd As Long
    Dim exactMapKey As String

    If ws Is Nothing Then Exit Sub
    If ruleObj Is Nothing Then Exit Sub

    Set selector = mp_GetParsedSelectorCached(ruleObj.Selector)
    If selector.Exists("mode") Then
        If StrComp(LCase$(Trim$(CStr(selector("mode")))), LCase$(Trim$(activeModeName)), vbTextCompare) <> 0 Then Exit Sub
    End If

    targetName = LCase$(Trim$(ruleObj.Target))

    Select Case targetName
        Case "column"
            If selector.Exists("mapkey") And Not targetsByMapKey Is Nothing Then
                If mp_IsExactSelectorMapKey(CStr(selector("mapkey"))) Then
                    exactMapKey = LCase$(Trim$(CStr(selector("mapkey"))))
                    If targetsByMapKey.Exists(exactMapKey) Then
                        Set targetGroup = targetsByMapKey(exactMapKey)
                        For Each target In targetGroup
                            If target Is Nothing Then GoTo ContinueExactMapTarget
                            If Not mp_ResultTargetMatchesSelector(target, selector) Then GoTo ContinueExactMapTarget

                            mapCol = CLng(target("ColumnIndex"))
                            mapRowStart = CLng(target("RowStart"))
                            mapRowEnd = CLng(target("RowEnd"))
                            If mapCol <= 0 Then GoTo ContinueExactMapTarget
                            If mapRowStart <= 0 Then mapRowStart = 1
                            If mapRowEnd < mapRowStart Then mapRowEnd = mapRowStart

                            Set scopeRange = ws.Range(ws.Cells(mapRowStart, mapCol), ws.Cells(mapRowEnd, mapCol))
                            Set columnScope = ws.Columns(mapCol)
                            mp_ApplyDeclarations scopeRange, columnScope, ruleObj.Declarations, Nothing, autoHeightState, autoHeightOnly
ContinueExactMapTarget:
                        Next target
                    End If
                    Exit Sub
                End If
            End If

            If mp_SelectorHasMapFilters(selector) Or (Not selector.Exists("col") And Not selector.Exists("address")) Then
                If resultFieldRanges Is Nothing Then Exit Sub
                For Each target In resultFieldRanges
                    If target Is Nothing Then GoTo ContinueMapTarget
                    If Not mp_ResultTargetMatchesSelector(target, selector) Then GoTo ContinueMapTarget

                    mapCol = CLng(target("ColumnIndex"))
                    mapRowStart = CLng(target("RowStart"))
                    mapRowEnd = CLng(target("RowEnd"))
                    If mapCol <= 0 Then GoTo ContinueMapTarget
                    If mapRowStart <= 0 Then mapRowStart = 1
                    If mapRowEnd < mapRowStart Then mapRowEnd = mapRowStart

                    Set scopeRange = ws.Range(ws.Cells(mapRowStart, mapCol), ws.Cells(mapRowEnd, mapCol))
                    Set columnScope = ws.Columns(mapCol)
                    mp_ApplyDeclarations scopeRange, columnScope, ruleObj.Declarations, Nothing, autoHeightState, autoHeightOnly
ContinueMapTarget:
                Next target
                Exit Sub
            End If

            If selector.Exists("address") Then
                addressText = CStr(selector("address"))
                Set scopeRange = ws.Range(addressText)
                Set columnScope = scopeRange.EntireColumn
                mp_ApplyDeclarations scopeRange, columnScope, ruleObj.Declarations, Nothing, autoHeightState, autoHeightOnly
                Exit Sub
            End If

            If Not mp_TryResolveColumnSpan(CStr(selector("col")), colStart, colEnd) Then
                Err.Raise vbObjectError + 1720, "ex_StylePipelineEngine", _
                    "Invalid selector col span for rule '" & ruleObj.RuleId & "': " & CStr(selector("col"))
            End If
            If selector.Exists("row") Then
                If Not mp_TryResolveRowSpan(CStr(selector("row")), rowStart, rowEnd) Then
                    Err.Raise vbObjectError + 1721, "ex_StylePipelineEngine", _
                        "Invalid selector row span for rule '" & ruleObj.RuleId & "': " & CStr(selector("row"))
                End If
            Else
                rowStart = 1
                rowEnd = mp_GetLastUsedRow(ws)
                If rowEnd < rowStart Then rowEnd = rowStart
            End If

            Set scopeRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
            Set columnScope = ws.Range(ws.Columns(colStart), ws.Columns(colEnd))
            mp_ApplyDeclarations scopeRange, columnScope, ruleObj.Declarations, Nothing, autoHeightState, autoHeightOnly

        Case "row"
            If selector.Exists("kind") Then
                mp_ApplyRowKindRule ws, selector, ruleObj.Declarations, rowKindRanges, ruleObj.RuleId, autoHeightState, autoHeightOnly
                Exit Sub
            End If

            If selector.Exists("address") Then
                Set scopeRange = ws.Range(CStr(selector("address")))
                mp_ApplyDeclarations scopeRange, Nothing, ruleObj.Declarations, Nothing, autoHeightState, autoHeightOnly
                Exit Sub
            End If

            rowStart = 1
            rowEnd = mp_GetLastUsedRow(ws)
            If selector.Exists("row") Then
                If Not mp_TryResolveRowSpan(CStr(selector("row")), rowStart, rowEnd) Then
                    Err.Raise vbObjectError + 1722, "ex_StylePipelineEngine", _
                        "Invalid selector row span for rule '" & ruleObj.RuleId & "': " & CStr(selector("row"))
                End If
            End If
            colStart = 1
            colEnd = mp_GetLastUsedColumn(ws)
            If selector.Exists("col") Then
                If Not mp_TryResolveColumnSpan(CStr(selector("col")), colStart, colEnd) Then
                    Err.Raise vbObjectError + 1723, "ex_StylePipelineEngine", _
                        "Invalid selector col span for rule '" & ruleObj.RuleId & "': " & CStr(selector("col"))
                End If
            End If
            If colEnd < colStart Then colEnd = colStart
            Set scopeRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
            mp_ApplyDeclarations scopeRange, Nothing, ruleObj.Declarations, Nothing, autoHeightState, autoHeightOnly

        Case "range"
            If Not selector.Exists("address") Then
                Err.Raise vbObjectError + 1724, "ex_StylePipelineEngine", "Range rule '" & ruleObj.RuleId & "' requires selector address."
            End If
            Set scopeRange = ws.Range(CStr(selector("address")))
            mp_ApplyDeclarations scopeRange, scopeRange.EntireColumn, ruleObj.Declarations, Nothing, autoHeightState, autoHeightOnly

        Case "cell"
            If selector.Exists("address") Then
                Set scopeRange = ws.Range(CStr(selector("address")))
            Else
                If Not selector.Exists("row") Or Not selector.Exists("col") Then
                    Err.Raise vbObjectError + 1725, "ex_StylePipelineEngine", "Cell rule '" & ruleObj.RuleId & "' requires selector row+col or address."
                End If

                hasSpan = mp_TryResolveRowSpan(CStr(selector("row")), rowStart, rowEnd)
                If Not hasSpan Then
                    Err.Raise vbObjectError + 1726, "ex_StylePipelineEngine", _
                        "Invalid selector row span for rule '" & ruleObj.RuleId & "': " & CStr(selector("row"))
                End If
                hasSpan = mp_TryResolveColumnSpan(CStr(selector("col")), colStart, colEnd)
                If Not hasSpan Then
                    Err.Raise vbObjectError + 1727, "ex_StylePipelineEngine", _
                        "Invalid selector col span for rule '" & ruleObj.RuleId & "': " & CStr(selector("col"))
                End If
                Set scopeRange = ws.Cells(rowStart, colStart)
            End If
            mp_ApplyDeclarations scopeRange, scopeRange.EntireColumn, ruleObj.Declarations, Nothing, autoHeightState, autoHeightOnly

        Case "sheet"
            Set scopeRange = mp_GetExpandedSheetScopeRange(ws)
            If scopeRange Is Nothing Then Exit Sub
            mp_ApplyDeclarations scopeRange, scopeRange.EntireColumn, ruleObj.Declarations, scopeRange, autoHeightState, autoHeightOnly

        Case "usedrange"
            Set scopeRange = ws.UsedRange
            If scopeRange Is Nothing Then Exit Sub
            mp_ApplyDeclarations scopeRange, scopeRange.EntireColumn, ruleObj.Declarations, scopeRange, autoHeightState, autoHeightOnly

        Case Else
            Err.Raise vbObjectError + 1728, "ex_StylePipelineEngine", _
                "Unsupported style target '" & targetName & "' in rule '" & ruleObj.RuleId & "'."
    End Select
End Sub

Private Function mp_IsExactSelectorMapKey(ByVal mapKeyText As String) As Boolean
    Dim normalized As String

    normalized = Trim$(mapKeyText)
    If Len(normalized) = 0 Then Exit Function
    If InStr(1, normalized, "*", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, normalized, "?", vbBinaryCompare) > 0 Then Exit Function

    mp_IsExactSelectorMapKey = True
End Function

Private Function mp_BuildResultTargetsByMapKey(ByVal resultFieldRanges As Collection) As Object
    Dim result As Object
    Dim target As Object
    Dim normalizedMapKey As String
    Dim targetGroup As Collection

    Set result = mp_CreateStringDictionary()
    If resultFieldRanges Is Nothing Then
        Set mp_BuildResultTargetsByMapKey = result
        Exit Function
    End If

    For Each target In resultFieldRanges
        If target Is Nothing Then GoTo ContinueTarget
        normalizedMapKey = LCase$(Trim$(CStr(target("MapKey"))))
        If Len(normalizedMapKey) = 0 Then GoTo ContinueTarget

        If result.Exists(normalizedMapKey) Then
            Set targetGroup = result(normalizedMapKey)
        Else
            Set targetGroup = New Collection
            Set result(normalizedMapKey) = targetGroup
        End If
        targetGroup.Add target
ContinueTarget:
    Next target

    Set mp_BuildResultTargetsByMapKey = result
End Function

Private Sub mp_ApplyDeclarations( _
    ByVal scopeRange As Range, _
    ByVal columnRange As Range, _
    ByVal declarations As Object, _
    Optional ByVal rowRange As Range, _
    Optional ByVal autoHeightState As Object = Nothing, _
    Optional ByVal autoHeightOnly As Boolean = False _
)
    Dim widthValue As Double
    Dim minWidthValue As Double
    Dim maxWidthValue As Double
    Dim hasMinWidth As Boolean
    Dim hasMaxWidth As Boolean
    Dim autoFitColumnsEnabled As Boolean
    Dim overflowValue As String
    Dim autoHeightEnabled As Boolean
    Dim rowHeightValue As Double
    Dim mergeColumnsValue As Long
    Dim fontSizeValue As Double
    Dim fontBoldValue As Boolean
    Dim colorValue As Long
    Dim hasBorderColor As Boolean
    Dim hasBorderWeight As Boolean
    Dim borderColorValue As Long
    Dim borderWeightValue As Long
    Dim horizontalValue As String
    Dim verticalValue As String
    Dim operationRowRange As Range

    If scopeRange Is Nothing Then Exit Sub
    If declarations Is Nothing Then Exit Sub
    If rowRange Is Nothing Then
        Set operationRowRange = scopeRange
    Else
        Set operationRowRange = rowRange
    End If

    If autoHeightOnly Then
        If declarations.Exists(STYLE_PROP_AUTO_HEIGHT) Then
            If Not mp_TryParseBoolean(CStr(declarations(STYLE_PROP_AUTO_HEIGHT)), autoHeightEnabled) Then
                Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid autoHeight declaration: " & CStr(declarations(STYLE_PROP_AUTO_HEIGHT))
            End If
            If autoHeightState Is Nothing Then
                If autoHeightEnabled Then operationRowRange.EntireRow.AutoFit
            Else
                mp_RecordDeferredAutoHeight operationRowRange, autoHeightEnabled, autoHeightState
            End If
        End If
        Exit Sub
    End If

    If declarations.Exists(STYLE_PROP_WIDTH) Then
        If Not mp_TryParseWidth(CStr(declarations(STYLE_PROP_WIDTH)), widthValue) Then
            Err.Raise vbObjectError + 1729, "ex_StylePipelineEngine", "Invalid width declaration: " & CStr(declarations(STYLE_PROP_WIDTH))
        End If
        If Not columnRange Is Nothing Then
            columnRange.ColumnWidth = widthValue
        Else
            scopeRange.EntireColumn.ColumnWidth = widthValue
        End If
    End If

    If declarations.Exists(STYLE_PROP_MIN_WIDTH) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_MIN_WIDTH)), minWidthValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid minWidth declaration: " & CStr(declarations(STYLE_PROP_MIN_WIDTH))
        End If
        hasMinWidth = True
    End If

    If declarations.Exists(STYLE_PROP_MAX_WIDTH) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_MAX_WIDTH)), maxWidthValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid maxWidth declaration: " & CStr(declarations(STYLE_PROP_MAX_WIDTH))
        End If
        hasMaxWidth = True
    End If

    If hasMinWidth Or hasMaxWidth Then
        If hasMinWidth And hasMaxWidth Then
            If maxWidthValue < minWidthValue Then
                Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid width clamp declaration: maxWidth < minWidth."
            End If
        End If
        If Not columnRange Is Nothing Then
            mp_ClampColumnRangeWidths columnRange, hasMinWidth, minWidthValue, hasMaxWidth, maxWidthValue
        Else
            mp_ClampColumnRangeWidths scopeRange.EntireColumn, hasMinWidth, minWidthValue, hasMaxWidth, maxWidthValue
        End If
    End If

    If declarations.Exists(STYLE_PROP_AUTO_FIT_COLUMNS) Then
        If Not mp_TryParseBoolean(CStr(declarations(STYLE_PROP_AUTO_FIT_COLUMNS)), autoFitColumnsEnabled) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid autoFitColumns declaration: " & CStr(declarations(STYLE_PROP_AUTO_FIT_COLUMNS))
        End If
        If autoFitColumnsEnabled Then
            If Not columnRange Is Nothing Then
                columnRange.EntireColumn.AutoFit
            Else
                scopeRange.EntireColumn.AutoFit
            End If
        End If
    End If

    If declarations.Exists(STYLE_PROP_OVERFLOW) Then
        overflowValue = LCase$(Trim$(CStr(declarations(STYLE_PROP_OVERFLOW))))
        Select Case overflowValue
            Case "wrap"
                scopeRange.WrapText = True
                scopeRange.ShrinkToFit = False
            Case "shrink"
                scopeRange.WrapText = False
                scopeRange.ShrinkToFit = True
            Case "clip"
                scopeRange.WrapText = False
                scopeRange.ShrinkToFit = False
            Case Else
                Err.Raise vbObjectError + 1730, "ex_StylePipelineEngine", "Unsupported overflow declaration: " & overflowValue
        End Select
    End If

    If declarations.Exists(STYLE_PROP_AUTO_HEIGHT) Then
        If Not mp_TryParseBoolean(CStr(declarations(STYLE_PROP_AUTO_HEIGHT)), autoHeightEnabled) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid autoHeight declaration: " & CStr(declarations(STYLE_PROP_AUTO_HEIGHT))
        End If
        If autoHeightState Is Nothing Then
            If autoHeightEnabled Then operationRowRange.EntireRow.AutoFit
        Else
            mp_RecordDeferredAutoHeight operationRowRange, autoHeightEnabled, autoHeightState
        End If
    End If

    If declarations.Exists(STYLE_PROP_ROW_HEIGHT) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_ROW_HEIGHT)), rowHeightValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid rowHeight declaration: " & CStr(declarations(STYLE_PROP_ROW_HEIGHT))
        End If
        operationRowRange.RowHeight = rowHeightValue
    End If

    If declarations.Exists(STYLE_PROP_MERGE_COLUMNS) Then
        If Not ex_XmlCore.m_TryParseLong(CStr(declarations(STYLE_PROP_MERGE_COLUMNS)), mergeColumnsValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid mergeColumns declaration: " & CStr(declarations(STYLE_PROP_MERGE_COLUMNS))
        End If
        If mergeColumnsValue < 1 Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid mergeColumns declaration: " & CStr(declarations(STYLE_PROP_MERGE_COLUMNS))
        End If
        mp_MergeRowsInScope operationRowRange, mergeColumnsValue
    End If

    If declarations.Exists(STYLE_PROP_FONT_NAME) Then
        scopeRange.Font.Name = CStr(declarations(STYLE_PROP_FONT_NAME))
    End If

    If declarations.Exists(STYLE_PROP_FONT_SIZE) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_FONT_SIZE)), fontSizeValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid fontSize declaration: " & CStr(declarations(STYLE_PROP_FONT_SIZE))
        End If
        scopeRange.Font.Size = fontSizeValue
    End If

    If declarations.Exists(STYLE_PROP_FONT_BOLD) Then
        If Not mp_TryParseBoolean(CStr(declarations(STYLE_PROP_FONT_BOLD)), fontBoldValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid fontBold declaration: " & CStr(declarations(STYLE_PROP_FONT_BOLD))
        End If
        scopeRange.Font.Bold = fontBoldValue
    End If

    If declarations.Exists(STYLE_PROP_BACK_COLOR) Then
        If Not ex_XmlCore.m_TryParseColor(CStr(declarations(STYLE_PROP_BACK_COLOR)), colorValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid backColor declaration: " & CStr(declarations(STYLE_PROP_BACK_COLOR))
        End If
        scopeRange.Interior.Color = colorValue
    End If

    If declarations.Exists(STYLE_PROP_FONT_COLOR) Then
        If Not ex_XmlCore.m_TryParseColor(CStr(declarations(STYLE_PROP_FONT_COLOR)), colorValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid fontColor declaration: " & CStr(declarations(STYLE_PROP_FONT_COLOR))
        End If
        scopeRange.Font.Color = colorValue
    End If

    If declarations.Exists(STYLE_PROP_BORDER_COLOR) Then
        If Not ex_XmlCore.m_TryParseColor(CStr(declarations(STYLE_PROP_BORDER_COLOR)), borderColorValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid borderColor declaration: " & CStr(declarations(STYLE_PROP_BORDER_COLOR))
        End If
        hasBorderColor = True
    End If

    If declarations.Exists(STYLE_PROP_BORDER_WEIGHT) Then
        If Not mp_TryParseBorderWeight(CStr(declarations(STYLE_PROP_BORDER_WEIGHT)), borderWeightValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid borderWeight declaration: " & CStr(declarations(STYLE_PROP_BORDER_WEIGHT))
        End If
        hasBorderWeight = True
    End If

    If hasBorderColor Or hasBorderWeight Then
        With scopeRange
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            If hasBorderColor Then .Borders.Color = borderColorValue
            If hasBorderWeight Then .Borders.Weight = borderWeightValue
        End With
    End If

    If declarations.Exists(STYLE_PROP_HORIZONTAL) Then
        horizontalValue = LCase$(Trim$(CStr(declarations(STYLE_PROP_HORIZONTAL))))
        scopeRange.HorizontalAlignment = mp_ParseHorizontalAlignment(horizontalValue)
    End If

    If declarations.Exists(STYLE_PROP_VERTICAL) Then
        verticalValue = LCase$(Trim$(CStr(declarations(STYLE_PROP_VERTICAL))))
        scopeRange.VerticalAlignment = mp_ParseVerticalAlignment(verticalValue)
    End If
End Sub

Private Sub mp_RecordDeferredAutoHeight( _
    ByVal targetRange As Range, _
    ByVal enabled As Boolean, _
    ByVal autoHeightState As Object _
)
    Dim rowArea As Range
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim rowIndex As Long

    If targetRange Is Nothing Then Exit Sub
    If autoHeightState Is Nothing Then Exit Sub

    For Each rowArea In targetRange.EntireRow.Areas
        rowStart = rowArea.Row
        rowEnd = rowStart + rowArea.Rows.Count - 1
        For rowIndex = rowStart To rowEnd
            autoHeightState(CStr(rowIndex)) = enabled
        Next rowIndex
    Next rowArea
End Sub

Private Sub mp_ApplyDeferredAutoHeight(ByVal ws As Worksheet, ByVal autoHeightState As Object)
    Dim enabledRows() As Long
    Dim enabledCount As Long
    Dim runStart As Long
    Dim runEnd As Long
    Dim rowIndex As Long
    Dim i As Long

    If ws Is Nothing Then Exit Sub
    If autoHeightState Is Nothing Then Exit Sub
    If autoHeightState.Count = 0 Then Exit Sub

    mp_CollectEnabledAutoHeightRows autoHeightState, enabledRows, enabledCount
    If enabledCount <= 0 Then Exit Sub

    mp_QuickSortLong enabledRows, 1, enabledCount

    runStart = enabledRows(1)
    runEnd = runStart
    For i = 2 To enabledCount
        rowIndex = enabledRows(i)
        If rowIndex = runEnd + 1 Then
            runEnd = rowIndex
        Else
            mp_ApplyAutoHeightToRowSpan ws, runStart, runEnd
            runStart = rowIndex
            runEnd = rowIndex
        End If
    Next i

    mp_ApplyAutoHeightToRowSpan ws, runStart, runEnd
End Sub

Private Sub mp_ApplyAutoHeightToRowSpan(ByVal ws As Worksheet, ByVal rowStart As Long, ByVal rowEnd As Long)
    Dim prevHeights() As Double
    Dim currentHeight As Double
    Dim rowIndex As Long
    Dim itemIndex As Long

    If ws Is Nothing Then Exit Sub
    If rowStart <= 0 Then Exit Sub
    If rowEnd < rowStart Then Exit Sub
    If rowStart > ws.Rows.Count Then Exit Sub
    If rowEnd > ws.Rows.Count Then rowEnd = ws.Rows.Count

    ReDim prevHeights(1 To rowEnd - rowStart + 1)
    itemIndex = 0
    For rowIndex = rowStart To rowEnd
        itemIndex = itemIndex + 1
        prevHeights(itemIndex) = ws.Rows(rowIndex).RowHeight
    Next rowIndex

    On Error Resume Next
    ws.Rows(CStr(rowStart) & ":" & CStr(rowEnd)).AutoFit
    On Error GoTo 0

    itemIndex = 0
    For rowIndex = rowStart To rowEnd
        itemIndex = itemIndex + 1
        currentHeight = ws.Rows(rowIndex).RowHeight
        If currentHeight < prevHeights(itemIndex) Then
            ws.Rows(rowIndex).RowHeight = prevHeights(itemIndex)
        End If
    Next rowIndex
End Sub

Private Sub mp_CollectEnabledAutoHeightRows( _
    ByVal autoHeightState As Object, _
    ByRef outRows() As Long, _
    ByRef outCount As Long _
)
    Dim keyValue As Variant
    Dim enabled As Boolean
    Dim rowIndex As Long

    If autoHeightState Is Nothing Then Exit Sub
    If autoHeightState.Count = 0 Then Exit Sub

    ReDim outRows(1 To autoHeightState.Count)
    For Each keyValue In autoHeightState.Keys
        enabled = False
        On Error Resume Next
        enabled = CBool(autoHeightState(CStr(keyValue)))
        On Error GoTo 0
        If Not enabled Then GoTo ContinueKey

        rowIndex = CLng(keyValue)
        If rowIndex <= 0 Then GoTo ContinueKey
        outCount = outCount + 1
        outRows(outCount) = rowIndex
ContinueKey:
    Next keyValue
End Sub

Private Sub mp_QuickSortLong(ByRef values() As Long, ByVal low As Long, ByVal high As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Long
    Dim tmp As Long

    If low >= high Then Exit Sub

    i = low
    j = high
    pivot = values((low + high) \ 2)

    Do While i <= j
        Do While values(i) < pivot
            i = i + 1
        Loop
        Do While values(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            tmp = values(i)
            values(i) = values(j)
            values(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop

    If low < j Then mp_QuickSortLong values, low, j
    If i < high Then mp_QuickSortLong values, i, high
End Sub

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

Private Sub mp_ClampColumnRangeWidths( _
    ByVal targetColumns As Range, _
    ByVal hasMinWidth As Boolean, _
    ByVal minWidthValue As Double, _
    ByVal hasMaxWidth As Boolean, _
    ByVal maxWidthValue As Double _
)
    Dim areaObj As Range
    Dim colObj As Range
    Dim currentWidth As Double

    If targetColumns Is Nothing Then Exit Sub

    For Each areaObj In targetColumns.Areas
        For Each colObj In areaObj.Columns
            currentWidth = colObj.ColumnWidth
            If hasMinWidth Then
                If currentWidth < minWidthValue Then
                    colObj.ColumnWidth = minWidthValue
                    currentWidth = minWidthValue
                End If
            End If
            If hasMaxWidth Then
                If currentWidth > maxWidthValue Then
                    colObj.ColumnWidth = maxWidthValue
                End If
            End If
        Next colObj
    Next areaObj
End Sub

Private Sub mp_MergeRowsInScope(ByVal scopeRange As Range, ByVal mergeColumns As Long)
    Dim ws As Worksheet
    Dim rowObj As Range
    Dim firstCol As Long
    Dim lastCol As Long
    Dim rowColCount As Long
    Dim mergeLastCol As Long
    Dim mergeRange As Range

    If scopeRange Is Nothing Then Exit Sub
    If mergeColumns < 1 Then Exit Sub

    Set ws = scopeRange.Worksheet
    firstCol = scopeRange.Column
    rowColCount = scopeRange.Columns.Count
    If rowColCount < 1 Then Exit Sub

    mergeLastCol = firstCol + mergeColumns - 1
    lastCol = firstCol + rowColCount - 1
    If mergeLastCol > lastCol Then mergeLastCol = lastCol

    For Each rowObj In scopeRange.Rows
        Set mergeRange = ws.Range(ws.Cells(rowObj.Row, firstCol), ws.Cells(rowObj.Row, mergeLastCol))
        mergeRange.UnMerge
        If mergeRange.Columns.Count > 1 Then mergeRange.Merge
    Next rowObj
End Sub

Private Function mp_IsSupportedTarget(ByVal targetName As String) As Boolean
    Select Case LCase$(Trim$(targetName))
        Case "sheet", "usedrange", "range", "row", "column", "cell"
            mp_IsSupportedTarget = True
    End Select
End Function

Private Function mp_ValidateDeclarations(ByVal declarations As Object, ByRef outErrorText As String) As Boolean
    Dim widthValue As Double
    Dim minWidthValue As Double
    Dim maxWidthValue As Double
    Dim hasMinWidth As Boolean
    Dim hasMaxWidth As Boolean
    Dim overflowValue As String
    Dim boolValue As Boolean
    Dim doubleValue As Double
    Dim longValue As Long
    Dim colorValue As Long
    Dim borderWeightValue As Long

    If declarations Is Nothing Then
        outErrorText = "Style declarations object is not initialized."
        Exit Function
    End If

    If declarations.Exists(STYLE_PROP_WIDTH) Then
        If Not mp_TryParseWidth(CStr(declarations(STYLE_PROP_WIDTH)), widthValue) Then
            outErrorText = "Invalid width declaration: " & CStr(declarations(STYLE_PROP_WIDTH))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_MIN_WIDTH) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_MIN_WIDTH)), minWidthValue) Then
            outErrorText = "Invalid minWidth declaration: " & CStr(declarations(STYLE_PROP_MIN_WIDTH))
            Exit Function
        End If
        hasMinWidth = True
    End If

    If declarations.Exists(STYLE_PROP_MAX_WIDTH) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_MAX_WIDTH)), maxWidthValue) Then
            outErrorText = "Invalid maxWidth declaration: " & CStr(declarations(STYLE_PROP_MAX_WIDTH))
            Exit Function
        End If
        hasMaxWidth = True
    End If

    If hasMinWidth And hasMaxWidth Then
        If maxWidthValue < minWidthValue Then
            outErrorText = "Invalid width clamp declaration: maxWidth < minWidth."
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_AUTO_FIT_COLUMNS) Then
        If Not mp_TryParseBoolean(CStr(declarations(STYLE_PROP_AUTO_FIT_COLUMNS)), boolValue) Then
            outErrorText = "Invalid autoFitColumns declaration: " & CStr(declarations(STYLE_PROP_AUTO_FIT_COLUMNS))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_OVERFLOW) Then
        overflowValue = LCase$(Trim$(CStr(declarations(STYLE_PROP_OVERFLOW))))
        Select Case overflowValue
            Case "wrap", "clip", "shrink"
            Case Else
                outErrorText = "Invalid overflow declaration: " & overflowValue
                Exit Function
        End Select
    End If

    If declarations.Exists(STYLE_PROP_AUTO_HEIGHT) Then
        If Not mp_TryParseBoolean(CStr(declarations(STYLE_PROP_AUTO_HEIGHT)), boolValue) Then
            outErrorText = "Invalid autoHeight declaration: " & CStr(declarations(STYLE_PROP_AUTO_HEIGHT))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_ROW_HEIGHT) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_ROW_HEIGHT)), doubleValue) Then
            outErrorText = "Invalid rowHeight declaration: " & CStr(declarations(STYLE_PROP_ROW_HEIGHT))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_MERGE_COLUMNS) Then
        If Not ex_XmlCore.m_TryParseLong(CStr(declarations(STYLE_PROP_MERGE_COLUMNS)), longValue) Then
            outErrorText = "Invalid mergeColumns declaration: " & CStr(declarations(STYLE_PROP_MERGE_COLUMNS))
            Exit Function
        End If
        If longValue < 1 Then
            outErrorText = "Invalid mergeColumns declaration: " & CStr(declarations(STYLE_PROP_MERGE_COLUMNS))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_FONT_NAME) Then
        If Len(Trim$(CStr(declarations(STYLE_PROP_FONT_NAME)))) = 0 Then
            outErrorText = "Invalid fontName declaration: value is empty."
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_FONT_SIZE) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_FONT_SIZE)), doubleValue) Then
            outErrorText = "Invalid fontSize declaration: " & CStr(declarations(STYLE_PROP_FONT_SIZE))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_FONT_BOLD) Then
        If Not mp_TryParseBoolean(CStr(declarations(STYLE_PROP_FONT_BOLD)), boolValue) Then
            outErrorText = "Invalid fontBold declaration: " & CStr(declarations(STYLE_PROP_FONT_BOLD))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_BACK_COLOR) Then
        If Not ex_XmlCore.m_TryParseColor(CStr(declarations(STYLE_PROP_BACK_COLOR)), colorValue) Then
            outErrorText = "Invalid backColor declaration: " & CStr(declarations(STYLE_PROP_BACK_COLOR))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_FONT_COLOR) Then
        If Not ex_XmlCore.m_TryParseColor(CStr(declarations(STYLE_PROP_FONT_COLOR)), colorValue) Then
            outErrorText = "Invalid fontColor declaration: " & CStr(declarations(STYLE_PROP_FONT_COLOR))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_BORDER_COLOR) Then
        If Not ex_XmlCore.m_TryParseColor(CStr(declarations(STYLE_PROP_BORDER_COLOR)), colorValue) Then
            outErrorText = "Invalid borderColor declaration: " & CStr(declarations(STYLE_PROP_BORDER_COLOR))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_BORDER_WEIGHT) Then
        If Not mp_TryParseBorderWeight(CStr(declarations(STYLE_PROP_BORDER_WEIGHT)), borderWeightValue) Then
            outErrorText = "Invalid borderWeight declaration: " & CStr(declarations(STYLE_PROP_BORDER_WEIGHT))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_HORIZONTAL) Then
        If Not mp_IsSupportedHorizontalAlignment(CStr(declarations(STYLE_PROP_HORIZONTAL))) Then
            outErrorText = "Invalid horizontal declaration: " & CStr(declarations(STYLE_PROP_HORIZONTAL))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_VERTICAL) Then
        If Not mp_IsSupportedVerticalAlignment(CStr(declarations(STYLE_PROP_VERTICAL))) Then
            outErrorText = "Invalid vertical declaration: " & CStr(declarations(STYLE_PROP_VERTICAL))
            Exit Function
        End If
    End If

    mp_ValidateDeclarations = True
End Function

Private Function mp_SelectorHasMapFilters(ByVal selector As Object) As Boolean
    If selector Is Nothing Then Exit Function
    mp_SelectorHasMapFilters = selector.Exists("mapkey") _
        Or selector.Exists("source") _
        Or selector.Exists("table") _
        Or selector.Exists("field")
End Function

Private Function mp_ResultTargetMatchesSelector(ByVal target As Object, ByVal selector As Object) As Boolean
    Dim mapKey As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim fieldAlias As String
    Dim targetCol As Long
    Dim targetRowStart As Long
    Dim targetRowEnd As Long
    Dim selColStart As Long
    Dim selColEnd As Long
    Dim selRowStart As Long
    Dim selRowEnd As Long

    If target Is Nothing Then Exit Function
    If selector Is Nothing Then
        mp_ResultTargetMatchesSelector = True
        Exit Function
    End If

    mapKey = Trim$(CStr(target("MapKey")))
    If selector.Exists("mapkey") Then
        If Not mp_TextMatchesPattern(mapKey, CStr(selector("mapkey"))) Then Exit Function
    End If

    If selector.Exists("source") Or selector.Exists("table") Or selector.Exists("field") Then
        If Not mp_TryParseMapKeyParts(mapKey, sourceAlias, tableAlias, fieldAlias) Then Exit Function
        If selector.Exists("source") Then
            If Not mp_TextMatchesPattern(sourceAlias, CStr(selector("source"))) Then Exit Function
        End If
        If selector.Exists("table") Then
            If Not mp_TextMatchesPattern(tableAlias, CStr(selector("table"))) Then Exit Function
        End If
        If selector.Exists("field") Then
            If Not mp_TextMatchesPattern(fieldAlias, CStr(selector("field"))) Then Exit Function
        End If
    End If

    targetCol = CLng(target("ColumnIndex"))
    targetRowStart = CLng(target("RowStart"))
    targetRowEnd = CLng(target("RowEnd"))
    If targetRowEnd < targetRowStart Then targetRowEnd = targetRowStart

    If selector.Exists("col") Then
        If Not mp_TryResolveColumnSpan(CStr(selector("col")), selColStart, selColEnd) Then Exit Function
        If targetCol < selColStart Or targetCol > selColEnd Then Exit Function
    End If

    If selector.Exists("row") Then
        If Not mp_TryResolveRowSpan(CStr(selector("row")), selRowStart, selRowEnd) Then Exit Function
        If targetRowEnd < selRowStart Or targetRowStart > selRowEnd Then Exit Function
    End If

    mp_ResultTargetMatchesSelector = True
End Function

Private Function mp_ParseSelector(ByVal selectorText As String) As Object
    Dim dict As Object
    Dim normalized As String
    Dim tokens As Variant
    Dim token As String
    Dim separatorPos As Long
    Dim keyText As String
    Dim valueText As String
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    normalized = Trim$(selectorText)
    If Len(normalized) = 0 Then
        Set mp_ParseSelector = dict
        Exit Function
    End If

    normalized = Replace$(normalized, vbCrLf, ";")
    normalized = Replace$(normalized, vbCr, ";")
    normalized = Replace$(normalized, vbLf, ";")
    tokens = Split(normalized, ";")

    For i = LBound(tokens) To UBound(tokens)
        token = Trim$(CStr(tokens(i)))
        If Len(token) = 0 Then GoTo ContinueToken

        separatorPos = InStr(1, token, "=", vbBinaryCompare)
        If separatorPos <= 0 Then
            separatorPos = InStr(1, token, ":", vbBinaryCompare)
        End If
        If separatorPos <= 1 Then GoTo ContinueToken

        keyText = LCase$(Trim$(Left$(token, separatorPos - 1)))
        valueText = Trim$(Mid$(token, separatorPos + 1))
        If Len(keyText) = 0 Then GoTo ContinueToken
        If Len(valueText) = 0 Then GoTo ContinueToken
        dict(keyText) = valueText
ContinueToken:
    Next i

    Set mp_ParseSelector = dict
End Function

Private Function mp_TryParseMapKeyParts( _
    ByVal mapKey As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outFieldAlias As String _
) As Boolean
    Dim rx As Object
    Dim matches As Object

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "^\s*([^.]+)\.Sheet\[(.+)\]\.Map\[(.+)\]\s*$"

    If Not rx.Test(mapKey) Then Exit Function
    Set matches = rx.Execute(mapKey)
    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    outSourceAlias = Trim$(CStr(matches(0).SubMatches(0)))
    outTableAlias = Trim$(CStr(matches(0).SubMatches(1)))
    outFieldAlias = Trim$(CStr(matches(0).SubMatches(2)))

    mp_TryParseMapKeyParts = (Len(outSourceAlias) > 0 And Len(outTableAlias) > 0 And Len(outFieldAlias) > 0)
End Function

Private Function mp_TextMatchesPattern(ByVal valueText As String, ByVal patternText As String) As Boolean
    Dim normalizedValue As String
    Dim normalizedPattern As String

    normalizedValue = LCase$(Trim$(valueText))
    normalizedPattern = LCase$(Trim$(patternText))

    If Len(normalizedPattern) = 0 Then
        mp_TextMatchesPattern = (Len(normalizedValue) = 0)
        Exit Function
    End If

    If InStr(1, normalizedPattern, "*", vbBinaryCompare) > 0 Or InStr(1, normalizedPattern, "?", vbBinaryCompare) > 0 Then
        mp_TextMatchesPattern = (normalizedValue Like normalizedPattern)
    Else
        mp_TextMatchesPattern = (StrComp(normalizedValue, normalizedPattern, vbTextCompare) = 0)
    End If
End Function

Private Function mp_TryResolveRowSpan(ByVal spanText As String, ByRef outStart As Long, ByRef outEnd As Long) As Boolean
    Dim normalized As String
    Dim parts As Variant
    Dim tmpStart As Long
    Dim tmpEnd As Long

    normalized = Trim$(spanText)
    If Len(normalized) = 0 Then Exit Function
    normalized = Replace$(normalized, "-", ":")

    If InStr(1, normalized, ":", vbBinaryCompare) > 0 Then
        parts = Split(normalized, ":")
        If UBound(parts) <> 1 Then Exit Function
        If Not ex_XmlCore.m_TryParseLong(Trim$(CStr(parts(0))), tmpStart) Then Exit Function
        If Not ex_XmlCore.m_TryParseLong(Trim$(CStr(parts(1))), tmpEnd) Then Exit Function
    Else
        If Not ex_XmlCore.m_TryParseLong(normalized, tmpStart) Then Exit Function
        tmpEnd = tmpStart
    End If

    If tmpStart <= 0 Or tmpEnd <= 0 Then Exit Function
    If tmpEnd < tmpStart Then
        outStart = tmpEnd
        outEnd = tmpStart
    Else
        outStart = tmpStart
        outEnd = tmpEnd
    End If

    mp_TryResolveRowSpan = True
End Function

Private Function mp_TryResolveRowEntry( _
    ByVal rowEntry As Variant, _
    ByRef outStart As Long, _
    ByRef outEnd As Long _
) As Boolean
    Dim obj As Object

    If IsObject(rowEntry) Then
        Set obj = rowEntry
        If obj Is Nothing Then Exit Function
        If obj.Exists("RowStart") Then outStart = CLng(obj("RowStart"))
        If obj.Exists("RowEnd") Then outEnd = CLng(obj("RowEnd"))
        If outStart <= 0 Then Exit Function
        If outEnd <= 0 Then outEnd = outStart
        If outEnd < outStart Then outEnd = outStart
        mp_TryResolveRowEntry = True
        Exit Function
    End If

    If IsNumeric(rowEntry) Then
        outStart = CLng(rowEntry)
        outEnd = outStart
        mp_TryResolveRowEntry = (outStart > 0)
        Exit Function
    End If
End Function

Private Function mp_TryResolveColumnSpan(ByVal spanText As String, ByRef outStart As Long, ByRef outEnd As Long) As Boolean
    Dim normalized As String
    Dim parts As Variant
    Dim tmpStart As Long
    Dim tmpEnd As Long

    normalized = Trim$(spanText)
    If Len(normalized) = 0 Then Exit Function
    normalized = Replace$(normalized, "-", ":")

    If InStr(1, normalized, ":", vbBinaryCompare) > 0 Then
        parts = Split(normalized, ":")
        If UBound(parts) <> 1 Then Exit Function
        If Not mp_TryResolveColumnToken(Trim$(CStr(parts(0))), tmpStart) Then Exit Function
        If Not mp_TryResolveColumnToken(Trim$(CStr(parts(1))), tmpEnd) Then Exit Function
    Else
        If Not mp_TryResolveColumnToken(normalized, tmpStart) Then Exit Function
        tmpEnd = tmpStart
    End If

    If tmpStart <= 0 Or tmpEnd <= 0 Then Exit Function
    If tmpEnd < tmpStart Then
        outStart = tmpEnd
        outEnd = tmpStart
    Else
        outStart = tmpStart
        outEnd = tmpEnd
    End If

    mp_TryResolveColumnSpan = True
End Function

Private Function mp_TryResolveColumnToken(ByVal tokenText As String, ByRef outColumn As Long) As Boolean
    Dim normalized As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim colNumber As Long

    normalized = UCase$(Trim$(tokenText))
    If Len(normalized) = 0 Then Exit Function

    If IsNumeric(normalized) Then
        outColumn = CLng(normalized)
        mp_TryResolveColumnToken = (outColumn > 0)
        Exit Function
    End If

    colNumber = 0
    For i = 1 To Len(normalized)
        ch = Mid$(normalized, i, 1)
        code = AscW(ch)
        If code < 65 Or code > 90 Then Exit Function
        colNumber = colNumber * 26 + (code - 64)
    Next i

    If colNumber <= 0 Then Exit Function
    outColumn = colNumber
    mp_TryResolveColumnToken = True
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

Private Function mp_CreateStringDictionary() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    Set mp_CreateStringDictionary = result
End Function

Private Sub mp_EnsureRuntimeCaches()
    If g_IsRuntimeCacheInitialized Then Exit Sub

    Set g_LayersCache = mp_CreateStringDictionary()
    Set g_SelectorCache = mp_CreateStringDictionary()
    g_IsRuntimeCacheInitialized = True
End Sub

Private Function mp_BuildLayersCacheKey( _
    ByVal wb As Workbook, _
    ByVal pageName As String _
) As String
    mp_BuildLayersCacheKey = mp_BuildWorkbookCacheKey(wb) & "|" & LCase$(Trim$(pageName))
End Function

Private Function mp_BuildWorkbookCacheKey(ByVal wb As Workbook) As String
    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        mp_BuildWorkbookCacheKey = "wb:none"
        Exit Function
    End If

    mp_BuildWorkbookCacheKey = LCase$(Trim$(wb.FullName))
    If Len(mp_BuildWorkbookCacheKey) = 0 Then
        mp_BuildWorkbookCacheKey = "wb:" & LCase$(Trim$(wb.Path)) & "|" & LCase$(Trim$(wb.Name))
    End If
End Function

Private Function mp_GetStylePipelineDomCached(ByVal wb As Workbook) As Object
    Dim wbKey As String
    Dim fileStamp As Date
    Dim hasFileStamp As Boolean
    Dim shouldResetPipelineCaches As Boolean

    mp_EnsureRuntimeCaches

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    wbKey = mp_BuildWorkbookCacheKey(wb)
    hasFileStamp = mp_TryGetConfigFileStamp(wb, STYLE_PIPELINE_REL_PATH, fileStamp)

    If Not g_StylePipelineDocCache Is Nothing Then
        If StrComp(g_StylePipelineDocWbKey, wbKey, vbTextCompare) = 0 Then
            If (Not hasFileStamp) Or g_StylePipelineDocStamp = fileStamp Then
                Set mp_GetStylePipelineDomCached = g_StylePipelineDocCache
                Exit Function
            End If
            ' Reset pipeline caches only when StylePipeline file timestamp really changed.
            If hasFileStamp Then
                shouldResetPipelineCaches = (g_StylePipelineDocStamp <> fileStamp)
            End If
        End If
    End If

    Set g_StylePipelineDocCache = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        STYLE_PIPELINE_REL_PATH, _
        PROFILES_NS, _
        "StylePipeline config file was not found: ", _
        "Failed to parse StylePipeline config file: " _
    )
    g_StylePipelineDocWbKey = wbKey
    If hasFileStamp Then
        g_StylePipelineDocStamp = fileStamp
    Else
        g_StylePipelineDocStamp = 0
    End If

    If shouldResetPipelineCaches Then
        If Not g_LayersCache Is Nothing Then g_LayersCache.RemoveAll
    End If

    Set mp_GetStylePipelineDomCached = g_StylePipelineDocCache
End Function

Private Function mp_TryGetConfigFileStamp( _
    ByVal wb As Workbook, _
    ByVal relativePath As String, _
    ByRef outStamp As Date _
) As Boolean
    Dim normalizedRelPath As String
    Dim fullPath As String

    If wb Is Nothing Then Exit Function
    If Len(Trim$(wb.Path)) = 0 Then Exit Function

    normalizedRelPath = Replace$(Trim$(relativePath), "/", "\")
    If Len(normalizedRelPath) = 0 Then Exit Function
    fullPath = wb.Path & "\" & normalizedRelPath

    On Error Resume Next
    outStamp = FileDateTime(fullPath)
    If Err.Number = 0 Then
        mp_TryGetConfigFileStamp = True
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function mp_GetParsedSelectorCached(ByVal selectorText As String) As Object
    Dim cacheKey As String
    Dim parsedSelector As Object

    mp_EnsureRuntimeCaches

    cacheKey = Trim$(selectorText)
    If g_SelectorCache.Exists(cacheKey) Then
        Set mp_GetParsedSelectorCached = g_SelectorCache(cacheKey)
        Exit Function
    End If

    Set parsedSelector = mp_ParseSelector(selectorText)
    Set g_SelectorCache(cacheKey) = parsedSelector
    Set mp_GetParsedSelectorCached = parsedSelector
End Function

Private Function mp_TryParseWidth(ByVal valueText As String, ByRef outWidth As Double) As Boolean
    Dim normalized As String

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function

    If Len(normalized) >= 2 Then
        If LCase$(Right$(normalized, 2)) = "px" Then
            normalized = Trim$(Left$(normalized, Len(normalized) - 2))
        End If
    End If

    If Not IsNumeric(normalized) Then Exit Function

    outWidth = CDbl(normalized)
    If outWidth <= 0 Then Exit Function

    mp_TryParseWidth = True
End Function

Private Function mp_TryParseBoolean(ByVal valueText As String, ByRef outValue As Boolean) As Boolean
    Dim normalized As String

    normalized = LCase$(Trim$(valueText))
    If Len(normalized) = 0 Then Exit Function

    Select Case normalized
        Case "true", "1", "yes", "on"
            outValue = True
            mp_TryParseBoolean = True
        Case "false", "0", "no", "off"
            outValue = False
            mp_TryParseBoolean = True
    End Select
End Function

Private Function mp_TryParsePositiveDouble(ByVal valueText As String, ByRef outValue As Double) As Boolean
    Dim normalized As String

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseDouble(normalized, outValue) Then Exit Function
    If outValue <= 0 Then Exit Function
    mp_TryParsePositiveDouble = True
End Function

Private Function mp_TryParseBorderWeight(ByVal valueText As String, ByRef outValue As Long) As Boolean
    Select Case LCase$(Trim$(valueText))
        Case "hairline": outValue = xlHairline
        Case "thin": outValue = xlThin
        Case "medium": outValue = xlMedium
        Case "thick": outValue = xlThick
        Case Else: Exit Function
    End Select

    mp_TryParseBorderWeight = True
End Function

Private Function mp_IsSupportedHorizontalAlignment(ByVal valueText As String) As Boolean
    Select Case LCase$(Trim$(valueText))
        Case "left", "center", "right", "fill", "justify", "distributed", "general"
            mp_IsSupportedHorizontalAlignment = True
    End Select
End Function

Private Function mp_IsSupportedVerticalAlignment(ByVal valueText As String) As Boolean
    Select Case LCase$(Trim$(valueText))
        Case "top", "center", "bottom", "justify", "distributed"
            mp_IsSupportedVerticalAlignment = True
    End Select
End Function

Private Function mp_ParseHorizontalAlignment(ByVal valueText As String) As XlHAlign
    Select Case LCase$(Trim$(valueText))
        Case "left": mp_ParseHorizontalAlignment = xlLeft
        Case "center": mp_ParseHorizontalAlignment = xlCenter
        Case "right": mp_ParseHorizontalAlignment = xlRight
        Case "fill": mp_ParseHorizontalAlignment = xlFill
        Case "justify": mp_ParseHorizontalAlignment = xlJustify
        Case "distributed": mp_ParseHorizontalAlignment = xlDistributed
        Case Else: mp_ParseHorizontalAlignment = xlGeneral
    End Select
End Function

Private Function mp_ParseVerticalAlignment(ByVal valueText As String) As XlVAlign
    Select Case LCase$(Trim$(valueText))
        Case "top": mp_ParseVerticalAlignment = xlTop
        Case "center": mp_ParseVerticalAlignment = xlCenter
        Case "bottom": mp_ParseVerticalAlignment = xlBottom
        Case "justify": mp_ParseVerticalAlignment = xlJustify
        Case "distributed": mp_ParseVerticalAlignment = xlDistributed
        Case Else: mp_ParseVerticalAlignment = xlCenter
    End Select
End Function

Private Function mp_GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then
        mp_GetLastUsedRow = 1
    Else
        mp_GetLastUsedRow = lastCell.Row
    End If
End Function

Private Function mp_GetLastUsedColumn(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then
        mp_GetLastUsedColumn = 1
    Else
        mp_GetLastUsedColumn = lastCell.Column
    End If
End Function
