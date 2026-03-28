Attribute VB_Name = "ex_StylePipelineEngine"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const STYLE_PIPELINE_REL_PATH As String = "config\StylePipeline.xml"
Private Const DEFAULT_STAGE_NAME As String = "default"
Private Const SHEET_SCOPE_MIN_COL As Long = 40      ' AN
Private Const SHEET_SCOPE_MIN_ROW As Long = 100
Private Const SHEET_SCOPE_EXPAND_STEP As Long = 30

' Supported style properties (declarations):
' width, minWidth, maxWidth, autoFitColumns
' overflow, autoHeight, customAutoHeight-margin-top, customAutoHeight-margin-bottom, rowHeight, minRowHeight, mergeColumns
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
Private Const STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_TOP As String = "customautoheight-margin-top"
Private Const STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_BOTTOM As String = "customautoheight-margin-bottom"
Private Const STYLE_PROP_ROW_HEIGHT As String = "rowheight"
Private Const STYLE_PROP_MIN_ROW_HEIGHT As String = "minrowheight"
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
Private g_StylePipelineTrackedFiles As Object

Public Function m_CreatePipeline() As Collection
    Set m_CreatePipeline = New Collection
End Function

Public Sub m_ResetRuntimeCaches()
    g_IsRuntimeCacheInitialized = False
    Set g_LayersCache = Nothing
    Set g_SelectorCache = Nothing
    Set g_StylePipelineDocCache = Nothing
    Set g_StylePipelineTrackedFiles = Nothing
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
    ByVal wb As Workbook, _
    Optional ByVal stageName As String = vbNullString _
) As Collection
    Dim doc As Object
    Dim result As Collection
    Dim seenLayerIds As Object
    Dim formatLayers As Collection
    Dim layerObj As obj_StyleLayer
    Dim cacheKey As String
    Dim normalizedStageName As String

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

    normalizedStageName = mp_NormalizeStageName(stageName)

    mp_EnsureRuntimeCaches
    cacheKey = mp_BuildLayersCacheKey(wb, pageName, normalizedStageName)
    If g_LayersCache.Exists(cacheKey) Then
        Set m_LoadSheetPipelineLayers = g_LayersCache(cacheKey)
        Exit Function
    End If

    Set formatLayers = mp_LoadLayersFromSheetPipelineXml(doc, pageName, normalizedStageName)
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
    Dim xmlLayers As Collection
    Dim xmlLayer As obj_StyleLayer

    Set pipeline = m_CreatePipeline()

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
    Dim targetIndexes As Object
    Dim targetsByMapKey As Object
    Dim targetsByKind As Object
    Dim autoHeightState As Object

    If ws Is Nothing Then Exit Sub
    If pipeline Is Nothing Then Exit Sub
    If pipeline.Count = 0 Then Exit Sub

    Set autoHeightState = mp_CreateStringDictionary()

    If Not resultFieldRanges Is Nothing Then
        If resultFieldRanges.Count > 0 Then
            Set targetIndexes = mp_BuildResultTargetIndexes(resultFieldRanges)
            If Not targetIndexes Is Nothing Then
                If targetIndexes.Exists("ByMapKey") Then Set targetsByMapKey = targetIndexes("ByMapKey")
                If targetIndexes.Exists("ByKind") Then Set targetsByKind = targetIndexes("ByKind")
            End If
        End If
    End If

    Set sortedLayers = m_GetSortedLayers(pipeline)
    For Each layerObj In sortedLayers
        If layerObj Is Nothing Then GoTo ContinueLayer
        If Not layerObj.IsEnabled Then GoTo ContinueLayer

        For Each ruleObj In layerObj.Rules
            If ruleObj Is Nothing Then GoTo ContinueRule
            mp_ApplyRule ws, resultFieldRanges, ruleObj, activeModeName, rowKindRanges, targetsByMapKey, targetsByKind, autoHeightState, autoHeightOnly
ContinueRule:
        Next ruleObj
ContinueLayer:
    Next layerObj

    mp_ApplyDeferredAutoHeight ws, autoHeightState
End Sub

Private Function mp_LoadLayersFromSheetPipelineXml( _
    ByVal doc As Object, _
    ByVal pageName As String, _
    ByVal stageName As String _
) As Collection
    Dim result As Collection
    Dim sheetPipelines As Object
    Dim sheetPipelineNode As Object
    Dim layerNodes As Object
    Dim stageNodes As Object
    Dim stageNode As Object
    Dim layerNode As Object
    Dim layerObj As obj_StyleLayer
    Dim layerId As String
    Dim layerSource As String
    Dim layerEnabled As Boolean
    Dim layerPriority As Long
    Dim ok As Boolean
    Dim pageKey As String
    Dim stageKey As String
    Dim hasDefaultStage As Boolean
    Dim hasRequestedStage As Boolean
    Dim hasMatchingSheetPipeline As Boolean
    Dim normalizedRequestedStage As String

    Set result = New Collection
    If doc Is Nothing Then
        Set mp_LoadLayersFromSheetPipelineXml = result
        Exit Function
    End If

    Set sheetPipelines = doc.selectNodes("/*[local-name()='stylePipeline']/*[local-name()='sheetPipeline']")
    If sheetPipelines Is Nothing Then
        Set mp_LoadLayersFromSheetPipelineXml = result
        Exit Function
    End If

    normalizedRequestedStage = mp_NormalizeStageName(stageName)

    For Each sheetPipelineNode In sheetPipelines
        pageKey = Trim$(mp_NodeAttrText(sheetPipelineNode, "page"))
        If Len(pageKey) = 0 Then
            Err.Raise vbObjectError + 1736, "ex_StylePipelineEngine", _
                "sheetPipeline@page is required."
        End If
        If Len(Trim$(pageName)) > 0 Then
            If StrComp(pageKey, Trim$(pageName), vbTextCompare) <> 0 Then GoTo ContinueSheetPipeline
        End If
        hasMatchingSheetPipeline = True

        Set layerNodes = sheetPipelineNode.selectNodes("*[local-name()='layer']")
        If Not layerNodes Is Nothing Then
            If layerNodes.Length > 0 Then
                Err.Raise vbObjectError + 1738, "ex_StylePipelineEngine", "sheetPipeline direct layers are not supported. Move layers under stage name='default'."
            End If
        End If

        Set stageNodes = sheetPipelineNode.selectNodes("*[local-name()='stage']")
        If stageNodes Is Nothing Or stageNodes.Length = 0 Then
            Err.Raise vbObjectError + 1739, "ex_StylePipelineEngine", "sheetPipeline must contain at least one stage and mandatory stage name='default'."
        End If
        hasDefaultStage = False
        hasRequestedStage = False

        For Each stageNode In stageNodes
            stageKey = Trim$(mp_NodeAttrText(stageNode, "name"))
            If Len(stageKey) = 0 Then
                Err.Raise vbObjectError + 1737, "ex_StylePipelineEngine", "sheetPipeline/stage@name is required."
            End If
            If StrComp(stageKey, DEFAULT_STAGE_NAME, vbTextCompare) = 0 Then
                hasDefaultStage = True
            End If
            If StrComp(stageKey, normalizedRequestedStage, vbTextCompare) <> 0 Then GoTo ContinueStage
            hasRequestedStage = True

            Set layerNodes = stageNode.selectNodes("*[local-name()='layer']")
            If layerNodes Is Nothing Then GoTo ContinueStage

            For Each layerNode In layerNodes
                layerId = Trim$(mp_NodeAttrText(layerNode, "name"))
                If Len(layerId) = 0 Then
                    Err.Raise vbObjectError + 1710, "ex_StylePipelineEngine", "sheetPipeline/stage/layer@name is required."
                End If
                If Not ex_XmlCore.m_TryParseLong(mp_NodeAttrText(layerNode, "priority"), layerPriority) Then
                    Err.Raise vbObjectError + 1711, "ex_StylePipelineEngine", "Invalid style layer priority for '" & layerId & "'."
                End If
                layerSource = Trim$(mp_NodeAttrText(layerNode, "source"))
                If Len(layerSource) = 0 Then layerSource = "sheetPipeline.stage[" & stageKey & "]"

                ok = mp_TryParseBoolean(mp_NodeAttrText(layerNode, "enabled"), layerEnabled)
                If Not ok Then layerEnabled = True

                Set layerObj = New obj_StyleLayer
                layerObj.Initialize layerId, layerPriority, layerSource, layerEnabled, stageKey
                mp_ParseLayerRules layerNode, layerObj

                result.Add layerObj
            Next layerNode
ContinueStage:
        Next stageNode
        If Not hasDefaultStage Then
            Err.Raise vbObjectError + 1739, "ex_StylePipelineEngine", "sheetPipeline must contain mandatory stage name='default'."
        End If
        If Not hasRequestedStage Then
            Err.Raise vbObjectError + 1740, "ex_StylePipelineEngine", "Stage '" & normalizedRequestedStage & "' was not found in sheetPipeline page '" & pageKey & "'."
        End If
ContinueSheetPipeline:
    Next sheetPipelineNode

    If Len(Trim$(pageName)) > 0 Then
        If Not hasMatchingSheetPipeline Then
            Err.Raise vbObjectError + 1736, "ex_StylePipelineEngine", "sheetPipeline@page '" & Trim$(pageName) & "' was not found."
        End If
    End If

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

    Set ruleNodes = layerNode.selectNodes("*[local-name()='rule']")
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

Private Sub mp_ApplyRule( _
    ByVal ws As Worksheet, _
    ByVal resultFieldRanges As Collection, _
    ByVal ruleObj As obj_StyleRule, _
    ByVal activeModeName As String, _
    Optional ByVal rowKindRanges As Object = Nothing, _
    Optional ByVal targetsByMapKey As Object = Nothing, _
    Optional ByVal targetsByKind As Object = Nothing, _
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
    Dim candidateTargets As Collection

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
                Set candidateTargets = mp_GetColumnRuleCandidateTargets(resultFieldRanges, selector, targetsByMapKey, targetsByKind)
                If candidateTargets Is Nothing Then Exit Sub

                For Each target In candidateTargets
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

        Case "usedrange"
            Set scopeRange = ws.UsedRange
            If scopeRange Is Nothing Then Exit Sub
            mp_ApplyDeclarations scopeRange, scopeRange.EntireColumn, ruleObj.Declarations, scopeRange, autoHeightState, autoHeightOnly

        Case "range"
            If Not selector.Exists("address") Then
                Err.Raise vbObjectError + 1724, "ex_StylePipelineEngine", "Range rule '" & ruleObj.RuleId & "' requires selector address."
            End If
            Set scopeRange = ws.Range(CStr(selector("address")))
            mp_ApplyDeclarations scopeRange, scopeRange.EntireColumn, ruleObj.Declarations, Nothing, autoHeightState, autoHeightOnly

        Case "sheet"
            Set scopeRange = mp_GetExpandedSheetScopeRange(ws)
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

Private Function mp_BuildResultTargetIndexes(ByVal resultFieldRanges As Collection) As Object
    Dim result As Object

    Set result = mp_CreateStringDictionary()
    Set result("ByMapKey") = mp_BuildResultTargetsByMapKey(resultFieldRanges)
    Set result("ByKind") = mp_BuildResultTargetsByKind(resultFieldRanges)

    Set mp_BuildResultTargetIndexes = result
End Function

Private Function mp_BuildResultTargetsByMapKey(ByVal resultFieldRanges As Collection) As Object
    Dim result As Object
    Dim target As Object
    Dim normalizedMapKey As String

    Set result = mp_CreateStringDictionary()
    If resultFieldRanges Is Nothing Then
        Set mp_BuildResultTargetsByMapKey = result
        Exit Function
    End If

    For Each target In resultFieldRanges
        If target Is Nothing Then GoTo ContinueTarget
        normalizedMapKey = LCase$(Trim$(CStr(target("MapKey"))))
        If Len(normalizedMapKey) = 0 Then GoTo ContinueTarget

        mp_AddTargetToIndexGroup result, normalizedMapKey, target
ContinueTarget:
    Next target

    Set mp_BuildResultTargetsByMapKey = result
End Function

Private Function mp_BuildResultTargetsByKind(ByVal resultFieldRanges As Collection) As Object
    Dim result As Object
    Dim target As Object
    Dim kindText As String
    Dim tokens As Variant
    Dim i As Long
    Dim tokenText As String

    Set result = mp_CreateStringDictionary()
    If resultFieldRanges Is Nothing Then
        Set mp_BuildResultTargetsByKind = result
        Exit Function
    End If

    For Each target In resultFieldRanges
        If target Is Nothing Then GoTo ContinueTarget
        If Not target.Exists("Kind") Then GoTo ContinueTarget

        kindText = LCase$(Trim$(CStr(target("Kind"))))
        If Len(kindText) = 0 Then GoTo ContinueTarget
        tokens = Split(kindText, "|")

        For i = LBound(tokens) To UBound(tokens)
            tokenText = LCase$(Trim$(CStr(tokens(i))))
            If Len(tokenText) = 0 Then GoTo ContinueToken
            mp_AddTargetToIndexGroup result, tokenText, target
ContinueToken:
        Next i
ContinueTarget:
    Next target

    Set mp_BuildResultTargetsByKind = result
End Function

Private Sub mp_AddTargetToIndexGroup(ByVal indexByToken As Object, ByVal tokenText As String, ByVal target As Object)
    Dim targetGroup As Collection

    If indexByToken Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    tokenText = LCase$(Trim$(tokenText))
    If Len(tokenText) = 0 Then Exit Sub

    If indexByToken.Exists(tokenText) Then
        Set targetGroup = indexByToken(tokenText)
    Else
        Set targetGroup = New Collection
        Set indexByToken(tokenText) = targetGroup
    End If
    targetGroup.Add target
End Sub

Private Function mp_GetColumnRuleCandidateTargets( _
    ByVal resultFieldRanges As Collection, _
    ByVal selector As Object, _
    ByVal targetsByMapKey As Object, _
    ByVal targetsByKind As Object _
) As Collection
    Dim exactMapKey As String
    Dim emptyResult As Collection
    Dim kindIndexed As Collection

    If resultFieldRanges Is Nothing Then Exit Function

    If Not selector Is Nothing Then
        If selector.Exists("mapkey") And Not targetsByMapKey Is Nothing Then
            If mp_IsExactSelectorMapKey(CStr(selector("mapkey"))) Then
                exactMapKey = LCase$(Trim$(CStr(selector("mapkey"))))
                If targetsByMapKey.Exists(exactMapKey) Then
                    Set mp_GetColumnRuleCandidateTargets = targetsByMapKey(exactMapKey)
                Else
                    Set emptyResult = New Collection
                    Set mp_GetColumnRuleCandidateTargets = emptyResult
                End If
                Exit Function
            End If
        End If

        If selector.Exists("kind") And Not targetsByKind Is Nothing Then
            Set kindIndexed = mp_TryGetIndexedCandidatesByKind(CStr(selector("kind")), targetsByKind)
            If Not kindIndexed Is Nothing Then
                Set mp_GetColumnRuleCandidateTargets = kindIndexed
                Exit Function
            End If
        End If
    End If

    Set mp_GetColumnRuleCandidateTargets = resultFieldRanges
End Function

Private Function mp_TryGetIndexedCandidatesByKind(ByVal selectorKindText As String, ByVal targetsByKind As Object) As Collection
    Dim kindTokens As Variant
    Dim i As Long
    Dim tokenText As String
    Dim bestToken As String
    Dim bestCount As Long
    Dim groupCount As Long
    Dim emptyResult As Collection

    If targetsByKind Is Nothing Then Exit Function
    If Not mp_TryParseExactSelectorKindTokens(selectorKindText, kindTokens) Then Exit Function
    If IsArray(kindTokens) = False Then Exit Function

    bestCount = -1
    For i = LBound(kindTokens) To UBound(kindTokens)
        tokenText = LCase$(Trim$(CStr(kindTokens(i))))
        If Len(tokenText) = 0 Then GoTo ContinueToken
        If Not targetsByKind.Exists(tokenText) Then
            Set emptyResult = New Collection
            Set mp_TryGetIndexedCandidatesByKind = emptyResult
            Exit Function
        End If

        groupCount = targetsByKind(tokenText).Count
        If bestCount < 0 Or groupCount < bestCount Then
            bestCount = groupCount
            bestToken = tokenText
        End If
ContinueToken:
    Next i

    If Len(bestToken) = 0 Then Exit Function
    Set mp_TryGetIndexedCandidatesByKind = targetsByKind(bestToken)
End Function

Private Function mp_TryParseExactSelectorKindTokens(ByVal selectorKindText As String, ByRef outTokens As Variant) As Boolean
    Dim tokens As Variant
    Dim i As Long
    Dim tokenText As String
    Dim dedupe As Object

    selectorKindText = LCase$(Trim$(selectorKindText))
    If Len(selectorKindText) = 0 Then Exit Function

    tokens = Split(selectorKindText, "|")
    Set dedupe = mp_CreateStringDictionary()

    For i = LBound(tokens) To UBound(tokens)
        tokenText = LCase$(Trim$(CStr(tokens(i))))
        If Len(tokenText) = 0 Then GoTo ContinueToken
        If mp_HasWildcardChars(tokenText) Then Exit Function
        dedupe(tokenText) = True
ContinueToken:
    Next i

    If dedupe.Count = 0 Then Exit Function
    outTokens = dedupe.Keys
    mp_TryParseExactSelectorKindTokens = True
End Function

Private Function mp_HasWildcardChars(ByVal textIn As String) As Boolean
    If InStr(1, textIn, "*", vbBinaryCompare) > 0 Then
        mp_HasWildcardChars = True
        Exit Function
    End If
    If InStr(1, textIn, "?", vbBinaryCompare) > 0 Then
        mp_HasWildcardChars = True
    End If
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
    Dim minRowHeightValue As Double
    Dim effectiveMinRowHeight As Double
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
    Dim hasRowHeight As Boolean
    Dim hasMinRowHeight As Boolean

    If scopeRange Is Nothing Then Exit Sub
    If declarations Is Nothing Then Exit Sub
    If rowRange Is Nothing Then
        Set operationRowRange = scopeRange
    Else
        Set operationRowRange = rowRange
    End If

    If declarations.Exists(STYLE_PROP_ROW_HEIGHT) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_ROW_HEIGHT)), rowHeightValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid rowHeight declaration: " & CStr(declarations(STYLE_PROP_ROW_HEIGHT))
        End If
        hasRowHeight = True
    End If

    If declarations.Exists(STYLE_PROP_MIN_ROW_HEIGHT) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_MIN_ROW_HEIGHT)), minRowHeightValue) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid minRowHeight declaration: " & CStr(declarations(STYLE_PROP_MIN_ROW_HEIGHT))
        End If
        hasMinRowHeight = True
    End If

    If autoHeightOnly Then
        If hasRowHeight Then
            Exit Sub
        End If
        If declarations.Exists(STYLE_PROP_AUTO_HEIGHT) Then
            If Not mp_TryParseBoolean(CStr(declarations(STYLE_PROP_AUTO_HEIGHT)), autoHeightEnabled) Then
                Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid autoHeight declaration: " & CStr(declarations(STYLE_PROP_AUTO_HEIGHT))
            End If
            If autoHeightState Is Nothing Then
                If autoHeightEnabled Then
                    operationRowRange.EntireRow.AutoFit
                    If hasMinRowHeight Then
                        mp_EnforceMinRowHeight operationRowRange, minRowHeightValue
                    End If
                End If
            Else
                effectiveMinRowHeight = 0
                If autoHeightEnabled And hasMinRowHeight Then
                    effectiveMinRowHeight = minRowHeightValue
                End If
                mp_RecordDeferredAutoHeight operationRowRange, autoHeightEnabled, autoHeightState, effectiveMinRowHeight
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

    If hasRowHeight Then
        operationRowRange.RowHeight = rowHeightValue
    ElseIf declarations.Exists(STYLE_PROP_AUTO_HEIGHT) Then
        If Not mp_TryParseBoolean(CStr(declarations(STYLE_PROP_AUTO_HEIGHT)), autoHeightEnabled) Then
            Err.Raise vbObjectError + 1731, "ex_StylePipelineEngine", "Invalid autoHeight declaration: " & CStr(declarations(STYLE_PROP_AUTO_HEIGHT))
        End If
        If autoHeightState Is Nothing Then
            If autoHeightEnabled Then
                operationRowRange.EntireRow.AutoFit
                If hasMinRowHeight Then
                    mp_EnforceMinRowHeight operationRowRange, minRowHeightValue
                End If
            End If
        Else
            effectiveMinRowHeight = 0
            If autoHeightEnabled And hasMinRowHeight Then
                effectiveMinRowHeight = minRowHeightValue
            End If
            mp_RecordDeferredAutoHeight operationRowRange, autoHeightEnabled, autoHeightState, effectiveMinRowHeight
        End If
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
    ByVal autoHeightState As Object, _
    Optional ByVal minRowHeight As Double = 0 _
)
    Dim rowArea As Range
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim rowIndex As Long
    Dim rowState As Object

    If targetRange Is Nothing Then Exit Sub
    If autoHeightState Is Nothing Then Exit Sub

    For Each rowArea In targetRange.EntireRow.Areas
        rowStart = rowArea.Row
        rowEnd = rowStart + rowArea.Rows.Count - 1
        For rowIndex = rowStart To rowEnd
            Set rowState = mp_GetOrCreateDeferredAutoHeightRowState(autoHeightState, CStr(rowIndex))
            If rowState Is Nothing Then GoTo ContinueRow

            rowState("enabled") = enabled
            If enabled Then
                If minRowHeight > CDbl(rowState("minrowheight")) Then
                    rowState("minrowheight") = minRowHeight
                End If
            Else
                rowState("minrowheight") = 0#
            End If
ContinueRow:
        Next rowIndex
    Next rowArea
End Sub

Private Sub mp_ApplyDeferredAutoHeight(ByVal ws As Worksheet, ByVal autoHeightState As Object)
    Dim enabledRows() As Long
    Dim enabledCount As Long
    Dim minHeightsByRow As Object
    Dim runStart As Long
    Dim runEnd As Long
    Dim rowIndex As Long
    Dim i As Long

    If ws Is Nothing Then Exit Sub
    If autoHeightState Is Nothing Then Exit Sub
    If autoHeightState.Count = 0 Then Exit Sub

    Set minHeightsByRow = mp_CreateStringDictionary()
    mp_CollectEnabledAutoHeightRows autoHeightState, enabledRows, enabledCount, minHeightsByRow
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
    mp_ApplyMinRowHeights ws, minHeightsByRow
End Sub

Private Sub mp_ApplyMinRowHeights(ByVal ws As Worksheet, ByVal minHeightsByRow As Object)
    Dim rowKey As Variant
    Dim rowIndex As Long
    Dim minHeight As Double

    If ws Is Nothing Then Exit Sub
    If minHeightsByRow Is Nothing Then Exit Sub
    If minHeightsByRow.Count = 0 Then Exit Sub

    For Each rowKey In minHeightsByRow.Keys
        rowIndex = CLng(rowKey)
        If rowIndex <= 0 Then GoTo ContinueRow
        If rowIndex > ws.Rows.Count Then GoTo ContinueRow

        minHeight = CDbl(minHeightsByRow(CStr(rowKey)))
        If minHeight <= 0 Then GoTo ContinueRow
        If ws.Rows(rowIndex).RowHeight < minHeight Then
            ws.Rows(rowIndex).RowHeight = minHeight
        End If
ContinueRow:
    Next rowKey
End Sub

Private Sub mp_EnforceMinRowHeight(ByVal targetRange As Range, ByVal minRowHeight As Double)
    Dim rowArea As Range
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim rowIndex As Long
    Dim ws As Worksheet

    If targetRange Is Nothing Then Exit Sub
    If minRowHeight <= 0 Then Exit Sub

    Set ws = targetRange.Worksheet
    For Each rowArea In targetRange.EntireRow.Areas
        rowStart = rowArea.Row
        rowEnd = rowStart + rowArea.Rows.Count - 1
        For rowIndex = rowStart To rowEnd
            If ws.Rows(rowIndex).RowHeight < minRowHeight Then
                ws.Rows(rowIndex).RowHeight = minRowHeight
            End If
        Next rowIndex
    Next rowArea
End Sub

Private Function mp_GetOrCreateDeferredAutoHeightRowState( _
    ByVal autoHeightState As Object, _
    ByVal rowKey As String _
) As Object
    Dim rowState As Object
    Dim rawState As Variant
    Dim enabled As Boolean

    If autoHeightState Is Nothing Then Exit Function
    If Len(rowKey) = 0 Then Exit Function

    If autoHeightState.Exists(rowKey) Then
        If IsObject(autoHeightState(rowKey)) Then
            Set rowState = autoHeightState(rowKey)
        Else
            rawState = autoHeightState(rowKey)
            On Error Resume Next
            enabled = CBool(rawState)
            On Error GoTo 0
            Set rowState = mp_CreateDeferredAutoHeightRowState(enabled, 0)
        End If
    Else
        Set rowState = mp_CreateDeferredAutoHeightRowState(False, 0)
    End If

    Set autoHeightState(rowKey) = rowState
    Set mp_GetOrCreateDeferredAutoHeightRowState = rowState
End Function

Private Function mp_CreateDeferredAutoHeightRowState( _
    ByVal enabled As Boolean, _
    ByVal minRowHeight As Double _
) As Object
    Dim rowState As Object

    Set rowState = mp_CreateStringDictionary()
    rowState("enabled") = enabled
    rowState("minrowheight") = minRowHeight

    Set mp_CreateDeferredAutoHeightRowState = rowState
End Function

Private Sub mp_ApplyAutoHeightToRowSpan(ByVal ws As Worksheet, ByVal rowStart As Long, ByVal rowEnd As Long)
    If ws Is Nothing Then Exit Sub
    If rowStart <= 0 Then Exit Sub
    If rowEnd < rowStart Then Exit Sub
    If rowStart > ws.Rows.Count Then Exit Sub
    If rowEnd > ws.Rows.Count Then rowEnd = ws.Rows.Count

    On Error Resume Next
    ws.Rows(CStr(rowStart) & ":" & CStr(rowEnd)).AutoFit
    On Error GoTo 0
End Sub

Private Sub mp_CollectEnabledAutoHeightRows( _
    ByVal autoHeightState As Object, _
    ByRef outRows() As Long, _
    ByRef outCount As Long, _
    Optional ByVal outMinHeightsByRow As Object = Nothing _
)
    Dim keyValue As Variant
    Dim stateValue As Variant
    Dim stateObj As Object
    Dim enabled As Boolean
    Dim rowIndex As Long
    Dim minRowHeight As Double

    If autoHeightState Is Nothing Then Exit Sub
    If autoHeightState.Count = 0 Then Exit Sub

    ReDim outRows(1 To autoHeightState.Count)
    For Each keyValue In autoHeightState.Keys
        enabled = False
        minRowHeight = 0

        If IsObject(autoHeightState(CStr(keyValue))) Then
            Set stateObj = autoHeightState(CStr(keyValue))
            If Not stateObj Is Nothing Then
                On Error Resume Next
                If stateObj.Exists("enabled") Then enabled = CBool(stateObj("enabled"))
                If stateObj.Exists("minrowheight") Then minRowHeight = CDbl(stateObj("minrowheight"))
                On Error GoTo 0
            End If
        Else
            stateValue = autoHeightState(CStr(keyValue))
            On Error Resume Next
            enabled = CBool(stateValue)
            On Error GoTo 0
        End If

        If Not enabled Then GoTo ContinueKey

        rowIndex = CLng(keyValue)
        If rowIndex <= 0 Then GoTo ContinueKey
        outCount = outCount + 1
        outRows(outCount) = rowIndex

        If Not outMinHeightsByRow Is Nothing Then
            If minRowHeight > 0 Then
                If outMinHeightsByRow.Exists(CStr(rowIndex)) Then
                    If CDbl(outMinHeightsByRow(CStr(rowIndex))) < minRowHeight Then
                        outMinHeightsByRow(CStr(rowIndex)) = minRowHeight
                    End If
                Else
                    outMinHeightsByRow(CStr(rowIndex)) = minRowHeight
                End If
            End If
        End If
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

    If declarations.Exists(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_TOP) Then
        If Not mp_TryParseNonNegativeDouble(CStr(declarations(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_TOP)), doubleValue) Then
            outErrorText = "Invalid customAutoHeight-margin-top declaration: " & CStr(declarations(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_TOP))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_BOTTOM) Then
        If Not mp_TryParseNonNegativeDouble(CStr(declarations(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_BOTTOM)), doubleValue) Then
            outErrorText = "Invalid customAutoHeight-margin-bottom declaration: " & CStr(declarations(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_BOTTOM))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_ROW_HEIGHT) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_ROW_HEIGHT)), doubleValue) Then
            outErrorText = "Invalid rowHeight declaration: " & CStr(declarations(STYLE_PROP_ROW_HEIGHT))
            Exit Function
        End If
    End If

    If declarations.Exists(STYLE_PROP_MIN_ROW_HEIGHT) Then
        If Not mp_TryParsePositiveDouble(CStr(declarations(STYLE_PROP_MIN_ROW_HEIGHT)), doubleValue) Then
            outErrorText = "Invalid minRowHeight declaration: " & CStr(declarations(STYLE_PROP_MIN_ROW_HEIGHT))
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
        Or selector.Exists("field") _
        Or selector.Exists("kind")
End Function

Private Function mp_ResultTargetMatchesSelector(ByVal target As Object, ByVal selector As Object) As Boolean
    Dim mapKey As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim fieldAlias As String
    Dim targetKind As String
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
    If selector.Exists("kind") Then
        If target.Exists("Kind") Then
            targetKind = Trim$(CStr(target("Kind")))
        Else
            targetKind = vbNullString
        End If
        If Not mp_KindTagsMatchSelector(targetKind, CStr(selector("kind"))) Then Exit Function
    End If

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

Private Function mp_KindTagsMatchSelector(ByVal kindTags As String, ByVal selectorKindPattern As String) As Boolean
    Dim valueTokens As Variant
    Dim selectorTokens As Variant
    Dim valueIndex As Long
    Dim selectorIndex As Long
    Dim valueToken As String
    Dim selectorToken As String
    Dim hasSelectorToken As Boolean
    Dim matchedSelectorToken As Boolean
    Dim selectorTokenCount As Long

    selectorKindPattern = Trim$(selectorKindPattern)
    If Len(selectorKindPattern) = 0 Then
        mp_KindTagsMatchSelector = (Len(Trim$(kindTags)) = 0)
        Exit Function
    End If

    kindTags = Trim$(kindTags)
    If Len(kindTags) = 0 Then Exit Function

    valueTokens = Split(kindTags, "|")
    selectorTokens = Split(selectorKindPattern, "|")

    For selectorIndex = LBound(selectorTokens) To UBound(selectorTokens)
        selectorToken = Trim$(CStr(selectorTokens(selectorIndex)))
        If Len(selectorToken) = 0 Then GoTo ContinueSelector
        hasSelectorToken = True
        selectorTokenCount = selectorTokenCount + 1
        matchedSelectorToken = False

        For valueIndex = LBound(valueTokens) To UBound(valueTokens)
            valueToken = Trim$(CStr(valueTokens(valueIndex)))
            If Len(valueToken) = 0 Then GoTo ContinueValue
            If mp_TextMatchesPattern(valueToken, selectorToken) Then
                matchedSelectorToken = True
                Exit For
            End If
ContinueValue:
        Next valueIndex

        If Not matchedSelectorToken Then Exit Function
ContinueSelector:
    Next selectorIndex

    If Not hasSelectorToken Then
        mp_KindTagsMatchSelector = (Len(kindTags) = 0)
    ElseIf selectorTokenCount > 0 Then
        mp_KindTagsMatchSelector = True
    End If
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
    ByVal pageName As String, _
    ByVal stageName As String _
) As String
    mp_BuildLayersCacheKey = mp_BuildWorkbookCacheKey(wb) & "|" & LCase$(Trim$(pageName)) & "|" & LCase$(Trim$(stageName))
End Function

Private Function mp_NormalizeStageName(ByVal stageName As String) As String
    stageName = Trim$(stageName)
    If Len(stageName) = 0 Then
        mp_NormalizeStageName = DEFAULT_STAGE_NAME
    Else
        mp_NormalizeStageName = stageName
    End If
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
    Dim shouldResetPipelineCaches As Boolean
    Dim loadedFiles As Object
    Dim resolvedDoc As Object

    mp_EnsureRuntimeCaches

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    wbKey = mp_BuildWorkbookCacheKey(wb)

    If Not g_StylePipelineDocCache Is Nothing Then
        If StrComp(g_StylePipelineDocWbKey, wbKey, vbTextCompare) = 0 Then
            If mp_AreTrackedStylePipelineFilesUnchanged(g_StylePipelineTrackedFiles) Then
                Set mp_GetStylePipelineDomCached = g_StylePipelineDocCache
                Exit Function
            End If
            shouldResetPipelineCaches = True
        Else
            shouldResetPipelineCaches = True
        End If
    End If

    Set resolvedDoc = mp_LoadStylePipelineDomWithIncludes(wb, loadedFiles)
    If resolvedDoc Is Nothing Then Exit Function

    Set g_StylePipelineDocCache = resolvedDoc
    g_StylePipelineDocWbKey = wbKey
    Set g_StylePipelineTrackedFiles = loadedFiles
    g_StylePipelineDocStamp = 0

    If shouldResetPipelineCaches Then
        If Not g_LayersCache Is Nothing Then g_LayersCache.RemoveAll
    End If

    Set mp_GetStylePipelineDomCached = g_StylePipelineDocCache
End Function

Private Function mp_LoadStylePipelineDomWithIncludes( _
    ByVal wb As Workbook, _
    ByRef outTrackedFiles As Object _
) As Object
    Dim rootPath As String
    Dim doc As Object
    Dim resolvingFiles As Object

    rootPath = ex_XmlCore.m_CombineBasePath(wb, STYLE_PIPELINE_REL_PATH)
    rootPath = mp_NormalizeFilePath(rootPath)
    If Len(rootPath) = 0 Then Exit Function

    Set doc = ex_XmlCore.m_LoadDomByFilePath( _
        rootPath, _
        PROFILES_NS, _
        "StylePipeline config file was not found: ", _
        "Failed to parse StylePipeline config file: " _
    )
    If doc Is Nothing Then Exit Function

    Set outTrackedFiles = mp_CreateStringDictionary()
    If Not mp_TryTrackFileStamp(rootPath, outTrackedFiles) Then Exit Function

    Set resolvingFiles = mp_CreateStringDictionary()
    If Not mp_ExpandStylePipelineIncludes(doc, rootPath, outTrackedFiles, resolvingFiles) Then Exit Function

    Set mp_LoadStylePipelineDomWithIncludes = doc
End Function

Private Function mp_ExpandStylePipelineIncludes( _
    ByVal doc As Object, _
    ByVal ownerFilePath As String, _
    ByVal trackedFiles As Object, _
    ByVal resolvingFiles As Object _
) As Boolean
    Dim includeNodes As Collection
    Dim includeNode As Object
    Dim includeDoc As Object
    Dim includeChildren As Object
    Dim includeChild As Object
    Dim importedNode As Object
    Dim includePath As String
    Dim includeFullPath As String
    Dim ownerKey As String

    If doc Is Nothing Then Exit Function
    If trackedFiles Is Nothing Then Exit Function
    If resolvingFiles Is Nothing Then Exit Function

    ownerKey = LCase$(mp_NormalizeFilePath(ownerFilePath))
    If Len(ownerKey) = 0 Then Exit Function
    If resolvingFiles.Exists(ownerKey) Then
        MsgBox "StylePipeline include recursion detected: " & ownerFilePath, vbExclamation
        Exit Function
    End If
    resolvingFiles(ownerKey) = True

    Set includeNodes = mp_CollectRootIncludeNodes(doc)
    For Each includeNode In includeNodes
        includePath = Trim$(ex_XmlCore.m_NodeAttrText(includeNode, "path"))
        If Len(includePath) = 0 Then
            MsgBox "StylePipeline include node must contain non-empty attribute 'path' in file: " & ownerFilePath, vbExclamation
            GoTo CleanupFalse
        End If

        includeFullPath = mp_ResolveIncludeFilePath(ownerFilePath, includePath)
        If Len(includeFullPath) = 0 Then
            MsgBox "StylePipeline include path could not be resolved: " & includePath & " (owner: " & ownerFilePath & ")", vbExclamation
            GoTo CleanupFalse
        End If

        Set includeDoc = ex_XmlCore.m_LoadDomByFilePath( _
            includeFullPath, _
            PROFILES_NS, _
            "StylePipeline include file was not found: ", _
            "Failed to parse StylePipeline include file: " _
        )
        If includeDoc Is Nothing Then GoTo CleanupFalse
        If Not mp_TryTrackFileStamp(includeFullPath, trackedFiles) Then GoTo CleanupFalse

        If Not mp_ExpandStylePipelineIncludes(includeDoc, includeFullPath, trackedFiles, resolvingFiles) Then GoTo CleanupFalse

        Set includeChildren = includeDoc.selectNodes("/*[local-name()='stylePipeline']/*[local-name()!='include']")
        If Not includeChildren Is Nothing Then
            For Each includeChild In includeChildren
                Set importedNode = doc.importNode(includeChild, True)
                includeNode.parentNode.insertBefore importedNode, includeNode
            Next includeChild
        End If

        includeNode.parentNode.RemoveChild includeNode
    Next includeNode

    resolvingFiles.Remove ownerKey
    mp_ExpandStylePipelineIncludes = True
    Exit Function

CleanupFalse:
    On Error Resume Next
    If resolvingFiles.Exists(ownerKey) Then resolvingFiles.Remove ownerKey
    On Error GoTo 0
End Function

Private Function mp_CollectRootIncludeNodes(ByVal doc As Object) As Collection
    Dim result As Collection
    Dim nodes As Object
    Dim node As Object

    Set result = New Collection
    If doc Is Nothing Then
        Set mp_CollectRootIncludeNodes = result
        Exit Function
    End If

    Set nodes = doc.selectNodes("/*[local-name()='stylePipeline']/*[local-name()='include']")
    If nodes Is Nothing Then
        Set mp_CollectRootIncludeNodes = result
        Exit Function
    End If

    For Each node In nodes
        If Not node Is Nothing Then result.Add node
    Next node

    Set mp_CollectRootIncludeNodes = result
End Function

Private Function mp_ResolveIncludeFilePath(ByVal ownerFilePath As String, ByVal includePath As String) As String
    Dim normalizedIncludePath As String
    Dim ownerDir As String
    Dim combinedPath As String

    normalizedIncludePath = mp_NormalizeFilePath(includePath)
    If Len(normalizedIncludePath) = 0 Then Exit Function

    If mp_IsAbsolutePath(normalizedIncludePath) Then
        mp_ResolveIncludeFilePath = normalizedIncludePath
        Exit Function
    End If

    ownerDir = mp_GetParentDirectory(ownerFilePath)
    If Len(ownerDir) = 0 Then Exit Function

    combinedPath = ownerDir & "\" & normalizedIncludePath
    mp_ResolveIncludeFilePath = mp_NormalizeFilePath(combinedPath)
End Function

Private Function mp_GetParentDirectory(ByVal filePath As String) As String
    Dim slashPos As Long
    Dim normalized As String

    normalized = mp_NormalizeFilePath(filePath)
    If Len(normalized) = 0 Then Exit Function

    slashPos = InStrRev(normalized, "\", -1, vbBinaryCompare)
    If slashPos <= 0 Then Exit Function
    If slashPos = 1 Then
        mp_GetParentDirectory = "\"
    Else
        mp_GetParentDirectory = Left$(normalized, slashPos - 1)
    End If
End Function

Private Function mp_IsAbsolutePath(ByVal filePath As String) As Boolean
    Dim normalized As String

    normalized = mp_NormalizeFilePath(filePath)
    If Len(normalized) = 0 Then Exit Function

    If Left$(normalized, 2) = "\\" Then
        mp_IsAbsolutePath = True
        Exit Function
    End If

    If Len(normalized) >= 3 Then
        If Mid$(normalized, 2, 1) = ":" And Mid$(normalized, 3, 1) = "\" Then
            mp_IsAbsolutePath = True
        End If
    End If
End Function

Private Function mp_NormalizeFilePath(ByVal filePath As String) As String
    Dim normalized As String

    normalized = Trim$(filePath)
    normalized = Replace$(normalized, "/", "\")
    Do While InStr(1, normalized, "\\", vbBinaryCompare) > 0 And Left$(normalized, 2) <> "\\"
        normalized = Replace$(normalized, "\\", "\")
    Loop
    mp_NormalizeFilePath = normalized
End Function

Private Function mp_TryTrackFileStamp(ByVal filePath As String, ByVal trackedFiles As Object) As Boolean
    Dim stamp As Date
    Dim fileSize As Long
    Dim key As String

    If trackedFiles Is Nothing Then Exit Function
    key = LCase$(mp_NormalizeFilePath(filePath))
    If Len(key) = 0 Then Exit Function

    On Error Resume Next
    stamp = FileDateTime(filePath)
    fileSize = FileLen(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        MsgBox "StylePipeline file timestamp read failed: " & filePath, vbExclamation
        Exit Function
    End If
    On Error GoTo 0

    trackedFiles(key) = CStr(CDbl(stamp)) & "|" & CStr(fileSize)
    mp_TryTrackFileStamp = True
End Function

Private Function mp_AreTrackedStylePipelineFilesUnchanged(ByVal trackedFiles As Object) As Boolean
    Dim key As Variant
    Dim stamp As Date
    Dim fileSize As Long
    Dim currentToken As String

    If trackedFiles Is Nothing Then Exit Function
    If trackedFiles.Count = 0 Then Exit Function

    For Each key In trackedFiles.Keys
        On Error Resume Next
        stamp = FileDateTime(CStr(key))
        fileSize = FileLen(CStr(key))
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0

        currentToken = CStr(CDbl(stamp)) & "|" & CStr(fileSize)
        If CStr(trackedFiles(CStr(key))) <> currentToken Then Exit Function
    Next key

    mp_AreTrackedStylePipelineFilesUnchanged = True
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

Private Function mp_TryParseNonNegativeDouble(ByVal valueText As String, ByRef outValue As Double) As Boolean
    Dim normalized As String

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseDouble(normalized, outValue) Then Exit Function
    If outValue < 0 Then Exit Function
    mp_TryParseNonNegativeDouble = True
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
