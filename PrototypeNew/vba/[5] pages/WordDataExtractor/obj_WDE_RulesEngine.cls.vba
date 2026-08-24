VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_WDE_RulesEngine"
Option Explicit

Private Const RULES_NS As String = "urn:excelprototype:word-data-extractor:v4"
Private m_Doc As Object
Private m_RulesRelPath As String
Private m_IsDisposed As Boolean
Private m_ScopeBoundaryPositions As Collection
Private m_DatasetScopeRanges As Object
Private m_ClaimedRanges As Collection
Private m_StructureInputContexts As Object
Private m_RegexFragments As Object
Private m_ExpandedRegexFragments As Object

Public Function Initialize(ByVal rulesRelPath As String) As Boolean
    m_RulesRelPath = VBA.Trim$(rulesRelPath)
    If VBA.Len(m_RulesRelPath) = 0 Then
        private_ShowError "Путь к файлу rules пуст."
        Exit Function
    End If
    Set m_Doc = ex_XmlCore.fn_LoadDomByRelativePath(ThisWorkbook, m_RulesRelPath, _
        "Не найден файл rules WordDataExtractor: ", _
        "Не удалось разобрать файл rules WordDataExtractor: ", RULES_NS)
    If m_Doc Is Nothing Then
        private_ShowError "Не удалось загрузить файл rules: " & m_RulesRelPath
        Exit Function
    End If
    If Not private_LoadRegexFragments() Then Exit Function
    If Not private_ValidateRegexFragmentUsage() Then Exit Function
    m_IsDisposed = False
    Initialize = private_ValidateRules()
End Function

Public Sub Dispose()
    m_IsDisposed = True
    Set m_Doc = Nothing
    Set m_ScopeBoundaryPositions = Nothing
    Set m_DatasetScopeRanges = Nothing
    Set m_ClaimedRanges = Nothing
    Set m_StructureInputContexts = Nothing
    Set m_RegexFragments = Nothing
    Set m_ExpandedRegexFragments = Nothing
    m_RulesRelPath = VBA.vbNullString
End Sub

Public Function ExtractTables(ByVal pipelineId As String, ByVal documentText As String, ByRef outTables As Collection) As Boolean
    Dim pipelineNode As Object
    Dim structureNode As Object
    Dim rootContexts As Collection
    Dim rootContext As Object

    Set outTables = Nothing
    If m_IsDisposed Or m_Doc Is Nothing Then Exit Function
    Set pipelineNode = m_Doc.selectSingleNode("/p:wordDataExtractor/p:pipelines/p:pipeline[@id=" & ex_XmlCore.fn_XPathLiteral(pipelineId) & "]")
    If pipelineNode Is Nothing Then
        private_ShowError "Pipeline не найден: " & pipelineId
        Exit Function
    End If
    Set outTables = New Collection
    Set m_ClaimedRanges = New Collection
    Set m_StructureInputContexts = VBA.CreateObject("Scripting.Dictionary")
    m_StructureInputContexts.CompareMode = 1
    Set rootContexts = New Collection
    Set rootContext = VBA.CreateObject("Scripting.Dictionary")
    rootContext.CompareMode = 1
    rootContext("$text") = documentText
    rootContexts.Add rootContext
    Set structureNode = pipelineNode.selectSingleNode("p:structure")
    If structureNode Is Nothing Then
        private_ShowError "Pipeline должен содержать structure."
        Exit Function
    End If
    If Not private_ExecutePatternChildren(pipelineNode, structureNode, _
        rootContexts, documentText, outTables) Then Exit Function
    Set m_ScopeBoundaryPositions = Nothing
    Set m_DatasetScopeRanges = Nothing
    Set m_ClaimedRanges = Nothing
    Set m_StructureInputContexts = Nothing
    ExtractTables = True
End Function

Private Function private_ExecutePatternChildren( _
    ByVal pipelineNode As Object, _
    ByVal parentNode As Object, _
    ByVal parentContexts As Collection, _
    ByVal documentText As String, _
    ByVal tables As Collection _
) As Boolean
    Dim childNodes As Object
    Dim childNode As Object
    Dim childContexts As Collection
    Dim nodeName As String

    Set childNodes = parentNode.selectNodes( _
        "p:context | p:pattern | p:requiredPattern")
    For Each childNode In childNodes
        nodeName = VBA.LCase$(childNode.nodeName)
        Select Case nodeName
            Case "context"
                If Not private_ExpandStructurePath(pipelineNode, _
                    VBA.CStr(ex_XmlCore.fn_NodeAttrText(childNode, "path")), _
                    parentContexts, childContexts, False) Then Exit Function
                If Not private_ExecutePatternChildren(pipelineNode, _
                    childNode, childContexts, documentText, tables) Then _
                    Exit Function
            Case "pattern", "requiredpattern"
                If Not private_ExecutePattern(pipelineNode, childNode, _
                    parentContexts, documentText, tables, _
                    nodeName = "requiredpattern") Then Exit Function
        End Select
    Next childNode
    private_ExecutePatternChildren = True
End Function

Private Function private_ExecutePattern( _
    ByVal pipelineNode As Object, _
    ByVal patternNode As Object, _
    ByVal parentContexts As Collection, _
    ByVal documentText As String, _
    ByVal tables As Collection, _
    ByVal patternRequired As Boolean _
) As Boolean
    Dim segments As Collection
    Dim executionContexts As Collection
    Dim activeContexts As Collection
    Dim finalReferenceNode As Object
    Dim datasetNode As Object
    Dim datasetId As String
    Dim finalRuleId As String
    Dim finalQuantifier As String
    Dim finalSegment As Object
    Dim tableCountBefore As Long
    Dim rowCount As Long
    Dim patternActivated As Boolean

    If Not private_ParseStructurePath(VBA.CStr(patternNode.Text), _
        segments) Then Exit Function
    If segments.Count = 0 Then
        private_ShowError "Pattern не содержит ссылок на rules."
        Exit Function
    End If

    Set finalSegment = segments.Item(segments.Count)
    finalRuleId = VBA.CStr(finalSegment("id"))
    finalQuantifier = VBA.CStr(finalSegment("quantifier"))
    datasetId = finalRuleId
    Set datasetNode = pipelineNode.selectSingleNode( _
        "p:rules/p:rule[@id=" & ex_XmlCore.fn_XPathLiteral(datasetId) & "]")
    If datasetNode Is Nothing Then
        private_ShowError "Pattern ссылается на неизвестное правило: " & _
            datasetId
        Exit Function
    End If
    If VBA.LCase$(private_AttrOrDefault(datasetNode, _
        "kind", "extract")) <> "extract" Then
        private_ShowError "Последняя ссылка pattern должна указывать на " & _
            "extract-rule: " & datasetId
        Exit Function
    End If

    Set executionContexts = private_CopyContexts(parentContexts)
    If segments.Count > 1 Then
        If Not private_ExpandStructureSegments(pipelineNode, segments, _
            1, segments.Count - 1, executionContexts, activeContexts, _
            True) Then Exit Function
        Set executionContexts = activeContexts
        patternActivated = (executionContexts.Count > 0)
    Else
        If Not private_FilterActiveDatasetContexts(datasetNode, _
            executionContexts, activeContexts) Then Exit Function
        Set executionContexts = activeContexts
        patternActivated = (executionContexts.Count > 0)
    End If

    If executionContexts.Count = 0 Then
        If patternRequired Then
            private_ShowError "Не найден обязательный pattern: " & datasetId
            Exit Function
        End If
    End If

    Set finalReferenceNode = patternNode
    If m_StructureInputContexts.Exists(datasetId) Then
        private_ShowError "Pattern уже выполняется: " & datasetId
        Exit Function
    End If
    m_StructureInputContexts.Add datasetId, executionContexts
    tableCountBefore = tables.Count
    If Not private_ExtractDataset(datasetNode, finalReferenceNode, _
        documentText, tables) Then
        m_StructureInputContexts.Remove datasetId
        Exit Function
    End If
    m_StructureInputContexts.Remove datasetId

    If tables.Count > tableCountBefore Then
        rowCount = tables.Item(tables.Count).RowCount
    End If
    If (patternRequired Or _
        (patternActivated And finalQuantifier = "+")) And rowCount = 0 Then
        private_ShowError "Активный pattern не создал обязательных строк: " & _
            datasetId
        Exit Function
    End If
    If finalQuantifier = "?" And rowCount > 1 Then
        private_ShowError "Pattern с квантификатором ? создал больше одной " & _
            "строки: " & datasetId
        Exit Function
    End If
    private_ExecutePattern = True
End Function

Private Function private_ExpandStructurePath( _
    ByVal pipelineNode As Object, _
    ByVal pathText As String, _
    ByVal parentContexts As Collection, _
    ByRef outContexts As Collection, _
    ByVal firstSegmentActivatesPattern As Boolean _
) As Boolean
    Dim segments As Collection

    If Not private_ParseStructurePath(pathText, segments) Then Exit Function
    If segments.Count = 0 Then
        private_ShowError "Context path не содержит ссылок на rules."
        Exit Function
    End If
    private_ExpandStructurePath = private_ExpandStructureSegments( _
        pipelineNode, segments, 1, segments.Count, parentContexts, _
        outContexts, firstSegmentActivatesPattern)
End Function

Private Function private_ExpandStructureSegments( _
    ByVal pipelineNode As Object, _
    ByVal segments As Collection, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long, _
    ByVal parentContexts As Collection, _
    ByRef outContexts As Collection, _
    ByVal firstSegmentActivatesPattern As Boolean _
) As Boolean
    Dim currentContexts As Collection
    Dim nextContexts As Collection
    Dim segment As Object
    Dim segmentIndex As Long

    Set currentContexts = private_CopyContexts(parentContexts)
    For segmentIndex = firstIndex To lastIndex
        Set segment = segments.Item(segmentIndex)
        If Not private_ExpandStructureSegment(pipelineNode, _
            VBA.CStr(segment("id")), VBA.CStr(segment("quantifier")), _
            currentContexts, nextContexts, _
            firstSegmentActivatesPattern And segmentIndex = firstIndex) Then _
            Exit Function
        Set currentContexts = nextContexts
        If currentContexts.Count = 0 Then Exit For
    Next segmentIndex
    Set outContexts = currentContexts
    private_ExpandStructureSegments = True
End Function

Private Function private_ExpandStructureSegment( _
    ByVal pipelineNode As Object, _
    ByVal ruleId As String, _
    ByVal quantifier As String, _
    ByVal parentContexts As Collection, _
    ByRef outContexts As Collection, _
    ByVal absenceDeactivatesPattern As Boolean _
) As Boolean
    Dim ruleNode As Object
    Dim locatorNode As Object
    Dim parentContext As Variant
    Dim singleContexts As Collection
    Dim foundContexts As Collection
    Dim foundContext As Variant
    Dim foundCount As Long

    Set outContexts = New Collection
    Set ruleNode = pipelineNode.selectSingleNode( _
        "p:rules/p:rule[@id=" & ex_XmlCore.fn_XPathLiteral(ruleId) & "]")
    If ruleNode Is Nothing Then
        private_ShowError "Structure ссылается на неизвестное правило: " & _
            ruleId
        Exit Function
    End If
    Set locatorNode = private_RuleLocatorNode(ruleNode)
    If locatorNode Is Nothing Then
        private_ShowError "Структурное правило не содержит match: " & ruleId
        Exit Function
    End If

    For Each parentContext In parentContexts
        Set singleContexts = New Collection
        singleContexts.Add parentContext
        If Not private_LocateChildContexts(pipelineNode, ruleNode, _
            locatorNode, singleContexts, foundContexts) Then Exit Function
        foundCount = foundContexts.Count

        If foundCount = 0 Then
            If absenceDeactivatesPattern Then
                ' Первый сегмент pattern только активирует ветку.
            ElseIf quantifier = "*" Or quantifier = "?" Then
                outContexts.Add parentContext
            Else
                private_ShowError "Не найден обязательный структурный " & _
                    "уровень: " & ruleId
                Exit Function
            End If
        Else
            If (quantifier = VBA.vbNullString Or quantifier = "?") And _
                foundCount > 1 Then
                private_ShowError "Структурный уровень создал больше одного " & _
                    "контекста: " & ruleId
                Exit Function
            End If
            For Each foundContext In foundContexts
                outContexts.Add foundContext
            Next foundContext
        End If
    Next parentContext
    private_ExpandStructureSegment = True
End Function

Private Function private_LocateChildContexts( _
    ByVal pipelineNode As Object, _
    ByVal ruleNode As Object, _
    ByVal locatorNode As Object, _
    ByVal parentContexts As Collection, _
    ByRef outContexts As Collection _
) As Boolean
    Dim parentContext As Variant
    Dim matches As Object
    Dim matchObj As Object
    Dim nextMatchObj As Object
    Dim rx As Object
    Dim parentText As String
    Dim contextText As String
    Dim startPosition As Long
    Dim endPosition As Long
    Dim i As Long

    Set outContexts = New Collection
    Set rx = private_CreateRegex(VBA.CStr(locatorNode.Text), _
        private_BoolAttr(locatorNode, "ignoreCase", True), _
        private_BoolAttr(locatorNode, "multiline", False))
    If rx Is Nothing Then Exit Function
    For Each parentContext In parentContexts
        parentText = VBA.CStr(parentContext("$text"))
        Set matches = rx.Execute(parentText)
        For i = 0 To matches.Count - 1
            Set matchObj = matches.Item(i)
            startPosition = VBA.CLng(matchObj.FirstIndex) + 1
            If i < matches.Count - 1 Then
                Set nextMatchObj = matches.Item(i + 1)
                endPosition = VBA.CLng(nextMatchObj.FirstIndex) + 1
            Else
                endPosition = VBA.Len(parentText) + 1
            End If
            ' Именованная секция заканчивается на ближайшем следующем
            ' поддерживаемом заголовке, даже если он находится на другом
            ' уровне structure и имеет другую нумерацию в исходном документе.
            If Not ruleNode.selectSingleNode("p:scope/p:match") Is Nothing Then
                endPosition = private_FindNextScopeHeaderPosition( _
                    pipelineNode, parentText, startPosition, endPosition)
            End If
            contextText = VBA.Mid$(parentText, startPosition, _
                endPosition - startPosition)
            If Not private_AddSectionContext(ruleNode, parentContext, _
                contextText, matchObj, outContexts) Then Exit Function
        Next i
    Next parentContext
    private_LocateChildContexts = True
End Function

Private Function private_FindNextScopeHeaderPosition( _
    ByVal pipelineNode As Object, _
    ByVal contextText As String, _
    ByVal startPosition As Long, _
    ByVal defaultEnd As Long _
) As Long
    Dim scopeMatchNodes As Object
    Dim scopeMatchNode As Object
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim candidatePosition As Long

    private_FindNextScopeHeaderPosition = defaultEnd
    Set scopeMatchNodes = pipelineNode.selectNodes( _
        "p:rules/p:rule/p:scope/p:match")
    For Each scopeMatchNode In scopeMatchNodes
        Set rx = private_CreateRegex(VBA.CStr(scopeMatchNode.Text), _
            private_BoolAttr(scopeMatchNode, "ignoreCase", True), _
            private_BoolAttr(scopeMatchNode, "multiline", False))
        If rx Is Nothing Then Exit Function
        Set matches = rx.Execute(contextText)
        For Each matchObj In matches
            candidatePosition = VBA.CLng(matchObj.FirstIndex) + 1
            If candidatePosition > startPosition And _
                candidatePosition < private_FindNextScopeHeaderPosition Then
                private_FindNextScopeHeaderPosition = candidatePosition
            End If
        Next matchObj
    Next scopeMatchNode
End Function

Private Function private_FilterActiveDatasetContexts( _
    ByVal datasetNode As Object, _
    ByVal parentContexts As Collection, _
    ByRef outContexts As Collection _
) As Boolean
    Dim locatorNode As Object
    Dim rx As Object
    Dim parentContext As Variant

    Set outContexts = New Collection
    Set locatorNode = private_RuleLocatorNode(datasetNode)
    If locatorNode Is Nothing Then
        private_ShowError "Extract-rule не содержит match: " & _
            private_AttrOrDefault(datasetNode, "id", "?")
        Exit Function
    End If
    Set rx = private_CreateRegex(VBA.CStr(locatorNode.Text), _
        private_BoolAttr(locatorNode, "ignoreCase", True), _
        private_BoolAttr(locatorNode, "multiline", False))
    If rx Is Nothing Then Exit Function
    For Each parentContext In parentContexts
        If rx.Execute(VBA.CStr(parentContext("$text"))).Count > 0 Then
            outContexts.Add parentContext
        End If
    Next parentContext
    private_FilterActiveDatasetContexts = True
End Function

Private Function private_RuleLocatorNode(ByVal ruleNode As Object) As Object
    Set private_RuleLocatorNode = ruleNode.selectSingleNode( _
        "p:scope/p:match")
    If private_RuleLocatorNode Is Nothing Then
        Set private_RuleLocatorNode = ruleNode.selectSingleNode("p:match")
    End If
End Function

Private Function private_ParseStructurePath( _
    ByVal pathText As String, _
    ByRef outSegments As Collection _
) As Boolean
    Dim rawSegments As Variant
    Dim rawSegment As Variant
    Dim rx As Object
    Dim matches As Object
    Dim segment As Object

    Set outSegments = New Collection
    pathText = VBA.Replace(pathText, VBA.vbCr, " ")
    pathText = VBA.Replace(pathText, VBA.vbLf, " ")
    rawSegments = VBA.Split(pathText, "/")
    Set rx = private_CreateRegex( _
        "^\s*\[([A-Za-z0-9_-]+)\]\s*([?*+]?)\s*$", _
        False, False)
    If rx Is Nothing Then Exit Function
    For Each rawSegment In rawSegments
        If VBA.Len(VBA.Trim$(VBA.CStr(rawSegment))) > 0 Then
            Set matches = rx.Execute(VBA.CStr(rawSegment))
            If matches.Count <> 1 Then
                private_ShowError "Некорректный сегмент structure path: " & _
                    VBA.Trim$(VBA.CStr(rawSegment))
                Exit Function
            End If
            Set segment = VBA.CreateObject("Scripting.Dictionary")
            segment.CompareMode = 1
            segment("id") = VBA.CStr(matches.Item(0).SubMatches(0))
            segment("quantifier") = _
                VBA.CStr(matches.Item(0).SubMatches(1))
            outSegments.Add segment
        End If
    Next rawSegment
    private_ParseStructurePath = True
End Function

Private Function private_CopyContexts( _
    ByVal sourceContexts As Collection _
) As Collection
    Dim result As Collection
    Dim contextItem As Variant

    Set result = New Collection
    For Each contextItem In sourceContexts
        result.Add contextItem
    Next contextItem
    Set private_CopyContexts = result
End Function

Private Function private_BuildStructuredRanges( _
    ByVal pipelineNode As Object, _
    ByVal documentText As String _
) As Boolean
    Dim referenceNodes As Object
    Dim referenceNode As Object
    Dim ruleNode As Object
    Dim scopeNode As Object
    Dim locatorNode As Object
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim rangeItem As Object
    Dim ranges As Collection
    Dim processedRuleIds As Object
    Dim ruleId As String
    Dim startPosition As Long
    Dim endPosition As Long

    Set m_ScopeBoundaryPositions = New Collection
    Set m_DatasetScopeRanges = VBA.CreateObject("Scripting.Dictionary")
    m_DatasetScopeRanges.CompareMode = 1
    Set m_ClaimedRanges = New Collection

    Set referenceNodes = pipelineNode.selectNodes( _
        "p:structure//*[@source and not(" & _
        "ancestor::p:content or ancestor::p:requiredContent or " & _
        "ancestor::p:items or ancestor::p:requiredItems)]")
    If referenceNodes.Length = 0 Then
        private_ShowError "Pipeline должен содержать обязательный блок structure."
        Exit Function
    End If

    ' Любое найденное правило из structure является потенциальной границей.
    ' Для именованной секции используется scope/match, для семантического
    ' standalone-пункта — его основной match.
    Set processedRuleIds = VBA.CreateObject("Scripting.Dictionary")
    processedRuleIds.CompareMode = 1
    For Each referenceNode In referenceNodes
        ruleId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
            referenceNode, "source"))
        If processedRuleIds.Exists(ruleId) Then GoTo NextBoundaryReference
        processedRuleIds.Add ruleId, True
        Set ruleNode = pipelineNode.selectSingleNode( _
            "p:rules/p:rule[@id=" & _
            ex_XmlCore.fn_XPathLiteral(ruleId) & "]")
        If ruleNode Is Nothing Then
            private_ShowError "Structure ссылается на неизвестное правило: " & _
                ruleId
            Exit Function
        End If
        Set locatorNode = ruleNode.selectSingleNode("p:scope/p:match")
        If locatorNode Is Nothing Then
            Set locatorNode = ruleNode.selectSingleNode("p:match")
        End If
        If Not locatorNode Is Nothing Then
            Set rx = private_CreateRegex(VBA.CStr(locatorNode.Text), _
                private_BoolAttr(locatorNode, "ignoreCase", True), _
                private_BoolAttr(locatorNode, "multiline", False))
            If rx Is Nothing Then Exit Function
            Set matches = rx.Execute(documentText)
            For Each matchObj In matches
                startPosition = VBA.CLng(matchObj.FirstIndex) + 1
                private_AddScopeBoundaryPosition startPosition
            Next matchObj
        End If
NextBoundaryReference:
    Next referenceNode

    Set processedRuleIds = VBA.CreateObject("Scripting.Dictionary")
    processedRuleIds.CompareMode = 1
    For Each referenceNode In referenceNodes
        ruleId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
            referenceNode, "source"))
        If processedRuleIds.Exists(ruleId) Then GoTo NextRangeReference
        processedRuleIds.Add ruleId, True
        Set ruleNode = pipelineNode.selectSingleNode( _
            "p:rules/p:rule[@id=" & _
            ex_XmlCore.fn_XPathLiteral(ruleId) & "]")
        Set scopeNode = ruleNode.selectSingleNode("p:scope/p:match")
        If Not scopeNode Is Nothing And _
            VBA.StrComp(VBA.LCase$(private_AttrOrDefault( _
            ruleNode, "kind", "extract")), "boundary", _
            VBA.vbBinaryCompare) <> 0 Then
            Set ranges = New Collection
            Set rx = private_CreateRegex(VBA.CStr(scopeNode.Text), _
                private_BoolAttr(scopeNode, "ignoreCase", True), _
                private_BoolAttr(scopeNode, "multiline", False))
            If rx Is Nothing Then Exit Function
            Set matches = rx.Execute(documentText)
            For Each matchObj In matches
                startPosition = VBA.CLng(matchObj.FirstIndex) + 1
                endPosition = private_FindNextRuleBoundary( _
                    startPosition, VBA.Len(documentText) + 1)
                Set rangeItem = VBA.CreateObject("Scripting.Dictionary")
                rangeItem.CompareMode = 1
                rangeItem("start") = startPosition
                rangeItem("end") = endPosition
                ranges.Add rangeItem
                private_AddClaimedRange startPosition, endPosition
            Next matchObj
            m_DatasetScopeRanges.Add ruleId, ranges
        End If
NextRangeReference:
    Next referenceNode
    private_BuildStructuredRanges = True
End Function

Private Function private_ExecuteTemplateChildren( _
    ByVal pipelineNode As Object, _
    ByVal templateNode As Object, _
    ByVal parentContexts As Collection, _
    ByVal documentText As String, _
    ByVal tables As Collection, _
    ByVal contextsAreRoot As Boolean _
) As Boolean
    Dim childNodes As Object
    Dim childNode As Object
    Dim nodeName As String

    Set childNodes = templateNode.selectNodes( _
        "p:template | p:content | p:requiredContent | " & _
        "p:items | p:requiredItems | p:boundary | p:if")
    For Each childNode In childNodes
        nodeName = VBA.LCase$(childNode.nodeName)
        Select Case nodeName
            Case "template"
                If Not private_ExecuteTemplateChildren(pipelineNode, _
                    childNode, parentContexts, documentText, tables, _
                    contextsAreRoot) Then Exit Function
            Case "content", "requiredcontent", "items", "requireditems", _
                "boundary"
                If Not private_ExecuteSourceNode(pipelineNode, childNode, _
                    parentContexts, documentText, tables, _
                    contextsAreRoot) Then Exit Function
            Case "if"
                If Not private_ExecuteIf(pipelineNode, childNode, _
                    parentContexts, documentText, tables, _
                    contextsAreRoot) Then Exit Function
        End Select
    Next childNode
    private_ExecuteTemplateChildren = True
End Function

Private Function private_ExecuteSourceNode( _
    ByVal pipelineNode As Object, _
    ByVal sourceNode As Object, _
    ByVal parentContexts As Collection, _
    ByVal documentText As String, _
    ByVal tables As Collection, _
    ByVal contextsAreRoot As Boolean _
) As Boolean
    Dim ruleNode As Object
    Dim childTemplateNode As Object
    Dim sourceContexts As Collection
    Dim sourceId As String
    Dim ruleKind As String

    sourceId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(sourceNode, "source"))
    Set ruleNode = pipelineNode.selectSingleNode( _
        "p:rules/p:rule[@id=" & ex_XmlCore.fn_XPathLiteral(sourceId) & "]")
    If ruleNode Is Nothing Then
        private_ShowError "Structure ссылается на неизвестное правило: " & _
            sourceId
        Exit Function
    End If
    ruleKind = VBA.LCase$(private_AttrOrDefault(ruleNode, "kind", "extract"))
    Select Case ruleKind
        Case "boundary"
            ' Boundary участвует в индексе диапазонов, но не создаёт результат.
        Case "section"
            If Not private_BuildSectionContexts(sourceNode, ruleNode, _
                parentContexts, documentText, contextsAreRoot, _
                sourceContexts) Then Exit Function
            If VBA.StrComp(sourceNode.nodeName, "items", _
                VBA.vbTextCompare) = 0 Or _
                VBA.StrComp(sourceNode.nodeName, "requiredItems", _
                VBA.vbTextCompare) = 0 Then
                Set childTemplateNode = sourceNode.selectSingleNode( _
                    "p:itemTemplate")
            Else
                Set childTemplateNode = sourceNode.selectSingleNode( _
                    "p:template")
            End If
            If Not childTemplateNode Is Nothing Then
                If Not private_ExecuteTemplateChildren(pipelineNode, _
                    childTemplateNode, sourceContexts, documentText, _
                    tables, False) Then Exit Function
            End If
        Case "extract"
            If contextsAreRoot Then
                If Not private_ExtractDataset(ruleNode, sourceNode, _
                    documentText, tables) Then Exit Function
            Else
                If m_StructureInputContexts.Exists(sourceId) Then
                    private_ShowError "Источник structure уже выполняется: " & _
                        sourceId
                    Exit Function
                End If
                m_StructureInputContexts.Add sourceId, parentContexts
                If Not private_ExtractDataset(ruleNode, sourceNode, _
                    documentText, tables) Then
                    m_StructureInputContexts.Remove sourceId
                    Exit Function
                End If
                m_StructureInputContexts.Remove sourceId
            End If
        Case Else
            private_ShowError "Неизвестный kind правила: " & sourceId
            Exit Function
    End Select
    private_ExecuteSourceNode = True
End Function

Private Function private_ExecuteIf( _
    ByVal pipelineNode As Object, _
    ByVal ifNode As Object, _
    ByVal parentContexts As Collection, _
    ByVal documentText As String, _
    ByVal tables As Collection, _
    ByVal contextsAreRoot As Boolean _
) As Boolean
    Dim parentContext As Variant
    Dim singleContext As Collection
    Dim thenNode As Object
    Dim elseNode As Object
    Dim sourceId As String
    Set thenNode = ifNode.selectSingleNode("p:then")
    Set elseNode = ifNode.selectSingleNode("p:else")
    sourceId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
        ifNode, "sourceExists"))
    For Each parentContext In parentContexts
        Set singleContext = New Collection
        singleContext.Add parentContext
        If private_SourceExists(pipelineNode, sourceId, _
            VBA.CStr(parentContext("$text"))) Then
            If Not private_ExecuteTemplateChildren(pipelineNode, _
                thenNode, singleContext, documentText, tables, _
                contextsAreRoot) Then Exit Function
        ElseIf Not elseNode Is Nothing Then
            If Not private_ExecuteTemplateChildren(pipelineNode, _
                elseNode, singleContext, documentText, tables, _
                contextsAreRoot) Then Exit Function
        End If
    Next parentContext
    private_ExecuteIf = True
End Function

Private Function private_SourceExists( _
    ByVal pipelineNode As Object, _
    ByVal sourceId As String, _
    ByVal contextText As String _
) As Boolean
    Dim ruleNode As Object
    Dim locatorNode As Object
    Dim rx As Object

    Set ruleNode = pipelineNode.selectSingleNode( _
        "p:rules/p:rule[@id=" & ex_XmlCore.fn_XPathLiteral(sourceId) & "]")
    If ruleNode Is Nothing Then
        private_ShowError "Условие sourceExists ссылается на неизвестное " & _
            "правило: " & sourceId
        Exit Function
    End If
    Set locatorNode = ruleNode.selectSingleNode("p:scope/p:match")
    If locatorNode Is Nothing Then
        Set locatorNode = ruleNode.selectSingleNode("p:match")
    End If
    If locatorNode Is Nothing Then
        private_ShowError "Правило sourceExists не содержит match: " & sourceId
        Exit Function
    End If
    Set rx = private_CreateRegex(VBA.CStr(locatorNode.Text), _
        private_BoolAttr(locatorNode, "ignoreCase", True), _
        private_BoolAttr(locatorNode, "multiline", False))
    If rx Is Nothing Then Exit Function
    private_SourceExists = (rx.Execute(contextText).Count > 0)
End Function

Private Function private_BuildSectionContexts( _
    ByVal sectionNode As Object, _
    ByVal ruleNode As Object, _
    ByVal parentContexts As Collection, _
    ByVal documentText As String, _
    ByVal isRootSection As Boolean, _
    ByRef outContexts As Collection _
) As Boolean
    Dim locatorNode As Object
    Dim rx As Object
    Dim parentContext As Variant
    Dim matches As Object
    Dim matchObj As Object
    Dim nextMatchObj As Object
    Dim rangeItem As Variant
    Dim ranges As Collection
    Dim contextText As String
    Dim parentText As String
    Dim nodeId As String
    Dim i As Long
    Dim startPosition As Long
    Dim endPosition As Long
    Dim matchCount As Long

    Set outContexts = New Collection
    nodeId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(sectionNode, "source"))
    Set locatorNode = ruleNode.selectSingleNode("p:scope/p:match")
    If locatorNode Is Nothing Then
        Set locatorNode = ruleNode.selectSingleNode("p:match")
    End If
    If locatorNode Is Nothing Then
        private_ShowError "Section-rule не содержит match: " & nodeId
        Exit Function
    End If
    Set rx = private_CreateRegex(VBA.CStr(locatorNode.Text), _
        private_BoolAttr(locatorNode, "ignoreCase", True), _
        private_BoolAttr(locatorNode, "multiline", False))
    If rx Is Nothing Then Exit Function

    For Each parentContext In parentContexts
        matchCount = 0
        If isRootSection And m_DatasetScopeRanges.Exists(nodeId) Then
            Set ranges = m_DatasetScopeRanges(nodeId)
            For Each rangeItem In ranges
                contextText = VBA.Mid$(documentText, _
                    VBA.CLng(rangeItem("start")), _
                    VBA.CLng(rangeItem("end")) - _
                    VBA.CLng(rangeItem("start")))
                Set matches = rx.Execute(contextText)
                If matches.Count > 0 Then
                    Set matchObj = matches.Item(0)
                    If Not private_AddSectionContext(ruleNode, _
                        parentContext, contextText, matchObj, _
                        outContexts) Then Exit Function
                    matchCount = matchCount + 1
                End If
            Next rangeItem
        Else
            parentText = VBA.CStr(parentContext("$text"))
            Set matches = rx.Execute(parentText)
            For i = 0 To matches.Count - 1
                Set matchObj = matches.Item(i)
                startPosition = VBA.CLng(matchObj.FirstIndex) + 1
                If i < matches.Count - 1 Then
                    Set nextMatchObj = matches.Item(i + 1)
                    endPosition = VBA.CLng(nextMatchObj.FirstIndex) + 1
                Else
                    endPosition = VBA.Len(parentText) + 1
                End If
                contextText = VBA.Mid$(parentText, startPosition, _
                    endPosition - startPosition)
                If Not private_AddSectionContext(ruleNode, _
                    parentContext, contextText, matchObj, _
                    outContexts) Then Exit Function
                matchCount = matchCount + 1
            Next i
        End If

        If matchCount = 0 Then
            If private_MinOccurs(sectionNode) > 0 Then
                private_ShowError "Не найдена обязательная section: " & nodeId
                Exit Function
            End If
        ElseIf Not private_ValidateMaxOccurs( _
            sectionNode, matchCount) Then
            private_ShowError "Источник content создал больше одного " & _
                "контекста: " & nodeId
            Exit Function
        End If
    Next parentContext
    private_BuildSectionContexts = True
End Function

Private Function private_AddSectionContext( _
    ByVal ruleNode As Object, _
    ByVal parentContext As Object, _
    ByVal contextText As String, _
    ByVal matchObj As Object, _
    ByVal outContexts As Collection _
) As Boolean
    Dim values As Object
    Dim fieldNodes As Object
    Dim fieldNode As Object
    Dim groupIndex As Long

    Set values = private_CloneContextValues(parentContext)
    values("$text") = contextText
    values("$match") = VBA.CStr(matchObj.Value)
    For groupIndex = 0 To matchObj.SubMatches.Count - 1
        values("$" & VBA.CStr(groupIndex + 1)) = _
            VBA.CStr(matchObj.SubMatches(groupIndex))
    Next groupIndex
    Set fieldNodes = ruleNode.selectNodes("p:fields/p:field")
    For Each fieldNode In fieldNodes
        If Not private_ExtractField(fieldNode, _
            VBA.CStr(matchObj.Value), values) Then Exit Function
    Next fieldNode
    outContexts.Add values
    private_AddSectionContext = True
End Function

Private Function private_MinOccurs(ByVal structureNode As Object) As Long
    Select Case VBA.LCase$(structureNode.nodeName)
        Case "requiredcontent", "requireditems"
            private_MinOccurs = 1
        Case Else
            private_MinOccurs = 0
    End Select
End Function

Private Function private_ValidateMaxOccurs( _
    ByVal structureNode As Object, _
    ByVal occurrenceCount As Long _
) As Boolean
    Select Case VBA.LCase$(structureNode.nodeName)
        Case "content", "requiredcontent", "boundary"
            private_ValidateMaxOccurs = (occurrenceCount <= 1)
        Case Else
            private_ValidateMaxOccurs = True
    End Select
End Function

Private Function private_ValidateOccurrenceSpec( _
    ByVal structureNode As Object _
) As Boolean
    If VBA.Len(VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
        structureNode, "source"))) = 0 Then
        private_ShowError "Исполняемый узел structure должен иметь source."
        Exit Function
    End If
    private_ValidateOccurrenceSpec = True
End Function

Private Function private_FindNextRuleBoundary( _
    ByVal startPosition As Long, _
    ByVal defaultEnd As Long _
) As Long
    Dim boundaryPosition As Variant

    private_FindNextRuleBoundary = defaultEnd
    For Each boundaryPosition In m_ScopeBoundaryPositions
        If VBA.CLng(boundaryPosition) > startPosition Then
            private_FindNextRuleBoundary = VBA.CLng(boundaryPosition)
            Exit Function
        End If
    Next boundaryPosition
End Function

Private Sub private_AddClaimedRange( _
    ByVal startPosition As Long, _
    ByVal endPosition As Long _
)
    Dim rangeItem As Object

    If endPosition <= startPosition Then Exit Sub
    Set rangeItem = VBA.CreateObject("Scripting.Dictionary")
    rangeItem.CompareMode = 1
    rangeItem("start") = startPosition
    rangeItem("end") = endPosition
    m_ClaimedRanges.Add rangeItem
End Sub

Private Sub private_AddScopeBoundaryPosition(ByVal position As Long)
    Dim i As Long
    Dim currentPosition As Long

    For i = 1 To m_ScopeBoundaryPositions.Count
        currentPosition = VBA.CLng(m_ScopeBoundaryPositions.Item(i))
        If currentPosition = position Then Exit Sub
        If currentPosition > position Then
            m_ScopeBoundaryPositions.Add position, Before:=i
            Exit Sub
        End If
    Next i
    m_ScopeBoundaryPositions.Add position
End Sub

Private Function private_ExtractDataset( _
    ByVal datasetNode As Object, _
    ByVal structureReferenceNode As Object, _
    ByVal sourceText As String, _
    ByVal tables As Collection _
) As Boolean
    Dim rx As Object, matches As Object, matchObj As Object
    Dim tableObj As obj_TableDynamic, rowObj As obj_Row
    Dim fieldValues As Object, columnNodes As Object, columnNode As Object
    Dim fieldNodes As Object, fieldNode As Object
    Dim patternNode As Object, patternText As String, datasetId As String
    Dim matchIndex As Long, datasetMatchIndex As Long, valueText As String
    Dim columnIndex As Long, cellTag As String
    Dim scopeTexts As Collection, scopeText As Variant
    Dim contextItems As Collection, contextItem As Variant
    Dim executionItems As Collection, expandedItem As Variant
    Dim nestedItems As Collection, nestedContextItem As Variant
    Dim contextKey As Variant
    Dim localNodes As Object, commonNodes As Object
    Dim datasetScopeNode As Object
    Dim tableNode As Object
    Dim claimMatches As Boolean
    Dim structureBranchActive As Boolean

    structureBranchActive = True
    datasetId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(datasetNode, "id"))
    Set datasetScopeNode = datasetNode.selectSingleNode("p:scope")
    claimMatches = (datasetScopeNode Is Nothing) And _
        VBA.StrComp(VBA.LCase$(VBA.Trim$( _
        ex_XmlCore.fn_NodeAttrText(datasetNode, "source"))), _
        "unclaimed", VBA.vbBinaryCompare) = 0
    Set patternNode = datasetNode.selectSingleNode("p:match")
    patternText = VBA.CStr(patternNode.Text)
    Set rx = private_CreateRegex(patternText, private_BoolAttr(patternNode, "ignoreCase", True), private_BoolAttr(patternNode, "multiline", True))
    If rx Is Nothing Then Exit Function
    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = private_AttrOrDefault(datasetNode, "caption", datasetId)
    tableObj.SourceAlias = datasetId
    tableObj.SourceAliasTemplate = "WordDataExtractor"
    Set tableNode = datasetNode.selectSingleNode("p:table")
    ' Сначала сохраняем специфичные колонки dataset, затем добавляем общие.
    ' XPath union вернул бы узлы в порядке XML-документа и поставил бы
    ' commonColumns перед ПІБ/датой события, поскольку они объявлены выше.
    Set columnNodes = New Collection
    Set localNodes = datasetNode.selectNodes("p:table/p:column")
    For Each columnNode In localNodes
        private_AddColumnNodeOrdered columnNodes, datasetNode, columnNode
    Next columnNode
    Set commonNodes = datasetNode.selectNodes( _
        "../../p:commonColumns/p:column")
    For Each columnNode In commonNodes
        If private_ShouldIncludeCommonColumn(datasetNode, columnNode) Then
            private_AddColumnNodeOrdered columnNodes, datasetNode, columnNode
        End If
    Next columnNode
    For Each columnNode In columnNodes
        If Not private_AddColumn(tableObj, columnNode) Then Exit Function
    Next columnNode

    ' Общие поля извлекаются из того же $match после полей dataset. Благодаря
    ' этому продовольственные даты и TVO не приходится копировать во все rules.
    Set fieldNodes = New Collection
    Set localNodes = datasetNode.selectNodes("p:fields/p:field")
    For Each fieldNode In localNodes
        fieldNodes.Add fieldNode
    Next fieldNode
    Set commonNodes = datasetNode.selectNodes( _
        "../../p:commonFields/p:field")
    For Each fieldNode In commonNodes
        fieldNodes.Add fieldNode
    Next fieldNode
    Set executionItems = New Collection
    If Not m_StructureInputContexts Is Nothing And _
        m_StructureInputContexts.Exists(datasetId) Then
        Set contextItems = m_StructureInputContexts(datasetId)
        structureBranchActive = (contextItems.Count > 0)
        For Each contextItem In contextItems
            If datasetNode.selectNodes( _
                "p:scope/p:context | p:context").Length > 0 Then
                ' Родительская структурная ветка может содержать несколько
                ' разных событий. Перед вложенными context обязательно
                ' отфильтровываем её собственным scope текущего dataset.
                If Not private_CollectScopedNestedContextItems(datasetNode, _
                    contextItem, nestedItems) Then Exit Function
                For Each nestedContextItem In nestedItems
                    executionItems.Add nestedContextItem
                Next nestedContextItem
            Else
                executionItems.Add contextItem
            End If
        Next contextItem
    Else
        If Not private_CollectScopeTexts(datasetNode, _
            structureReferenceNode, sourceText, scopeTexts) Then Exit Function
        For Each scopeText In scopeTexts
            If Not private_CollectContextItems( _
                datasetNode, VBA.CStr(scopeText), contextItems) Then _
                Exit Function
            For Each expandedItem In contextItems
                executionItems.Add expandedItem
            Next expandedItem
        Next scopeText
    End If
    For Each contextItem In executionItems
            ' Порядок относится к текущему context-блоку: общая групповая
            ' «Підстава» и список людей всегда индексируются внутри него.
            datasetMatchIndex = 0
            Set matches = rx.Execute(VBA.CStr(contextItem("$text")))
            For Each matchObj In matches
                datasetMatchIndex = datasetMatchIndex + 1
                Set fieldValues = VBA.CreateObject("Scripting.Dictionary")
                fieldValues.CompareMode = 1
                For Each contextKey In contextItem.Keys
                    If VBA.CStr(contextKey) <> "$text" Then
                        fieldValues(VBA.CStr(contextKey)) = contextItem(contextKey)
                    End If
                Next contextKey
                fieldValues("$match") = VBA.CStr(matchObj.Value)
                fieldValues("$document") = sourceText
                fieldValues("$matchIndex") = datasetMatchIndex
                For matchIndex = 0 To matchObj.SubMatches.Count - 1
                    fieldValues("$" & VBA.CStr(matchIndex + 1)) = VBA.CStr(matchObj.SubMatches(matchIndex))
                Next matchIndex
                For Each fieldNode In fieldNodes
                    If Not private_ExtractField(fieldNode, VBA.CStr(matchObj.Value), fieldValues) Then Exit Function
                Next fieldNode

                Set rowObj = New obj_Row
                columnIndex = 0
                For Each columnNode In columnNodes
                    columnIndex = columnIndex + 1
                    valueText = private_ResolveValue(ex_XmlCore.fn_NodeAttrText(columnNode, "value"), fieldValues)
                    valueText = private_ApplyTransforms( _
                        valueText, columnNode, fieldValues)
                    valueText = private_ApplyColumnRules( _
                        valueText, datasetNode, columnNode, fieldValues)
                    rowObj.PushCellRaw valueText
                    cellTag = VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
                        columnNode, "cellTag"))
                    If VBA.Len(cellTag) > 0 Then
                        If Not rowObj.AddCellTag( _
                            columnIndex, cellTag) Then Exit Function
                    End If
                Next columnNode
                If Not tableObj.PushRow(rowObj) Then Exit Function
                If claimMatches Then
                    private_AddClaimedRange _
                        VBA.CLng(matchObj.FirstIndex) + 1, _
                        VBA.CLng(matchObj.FirstIndex) + _
                        VBA.Len(VBA.CStr(matchObj.Value)) + 1
                End If
            Next matchObj
    Next contextItem
    If structureBranchActive And tableObj.RowCount < private_MinOccurs( _
        structureReferenceNode) Then
        private_ShowError "Обязательное правило structure не создало строк: " & _
            datasetId
        Exit Function
    End If
    If Not private_ValidateMaxOccurs(structureReferenceNode, _
        tableObj.RowCount) Then
        private_ShowError "Источник content создал больше одного результата: " & _
            datasetId
        Exit Function
    End If
    If tableObj.RowCount = 0 And _
        private_BoolAttr(tableNode, "emitEmptyRow", False) Then
        Set fieldValues = VBA.CreateObject("Scripting.Dictionary")
        fieldValues.CompareMode = 1
        fieldValues("$match") = VBA.vbNullString
        fieldValues("$document") = sourceText
        For Each fieldNode In fieldNodes
            If Not private_ExtractField(fieldNode, VBA.vbNullString, _
                fieldValues) Then Exit Function
        Next fieldNode
        Set rowObj = New obj_Row
        columnIndex = 0
        For Each columnNode In columnNodes
            columnIndex = columnIndex + 1
            valueText = private_ResolveValue( _
                ex_XmlCore.fn_NodeAttrText(columnNode, "value"), _
                fieldValues)
            valueText = private_ApplyTransforms( _
                valueText, columnNode, fieldValues)
            valueText = private_ApplyColumnRules( _
                valueText, datasetNode, columnNode, fieldValues)
            rowObj.PushCellRaw valueText
            cellTag = VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
                columnNode, "cellTag"))
            If VBA.Len(cellTag) > 0 Then
                If Not rowObj.AddCellTag( _
                    columnIndex, cellTag) Then Exit Function
            End If
        Next columnNode
        If Not tableObj.PushRow(rowObj) Then Exit Function
    End If
    ' Пустые таблицы также возвращаются вызывающему коду: UI самостоятельно
    ' решает, показывать все datasets или только содержащие строки.
    tables.Add tableObj
    private_ExtractDataset = True
End Function

Private Function private_CollectScopedNestedContextItems( _
    ByVal datasetNode As Object, _
    ByVal parentContext As Object, _
    ByRef outItems As Collection _
) As Boolean
    Dim scopeMatchNode As Object
    Dim scopeRx As Object
    Dim scopeMatches As Object
    Dim scopeMatch As Object
    Dim nextScopeMatch As Object
    Dim scopedParent As Object
    Dim nestedItems As Collection
    Dim nestedItem As Variant
    Dim parentText As String
    Dim startPosition As Long
    Dim endPosition As Long
    Dim i As Long

    Set outItems = New Collection
    Set scopeMatchNode = datasetNode.selectSingleNode("p:scope/p:match")
    If scopeMatchNode Is Nothing Then
        private_CollectScopedNestedContextItems = _
            private_CollectNestedContextItems( _
                datasetNode, parentContext, outItems)
        Exit Function
    End If

    parentText = VBA.CStr(parentContext("$text"))
    Set scopeRx = private_CreateRegex(VBA.CStr(scopeMatchNode.Text), _
        private_BoolAttr(scopeMatchNode, "ignoreCase", True), _
        private_BoolAttr(scopeMatchNode, "multiline", False))
    If scopeRx Is Nothing Then Exit Function
    Set scopeMatches = scopeRx.Execute(parentText)

    ' Отсутствие scope в конкретном родительском пункте означает, что этот
    ' dataset к нему не относится. Это не отсутствие обязательной секции во
    ' всём документе — кратность проверяется структурным маршрутом отдельно.
    For i = 0 To scopeMatches.Count - 1
        Set scopeMatch = scopeMatches.Item(i)
        startPosition = VBA.CLng(scopeMatch.FirstIndex) + 1
        If i < scopeMatches.Count - 1 Then
            Set nextScopeMatch = scopeMatches.Item(i + 1)
            endPosition = VBA.CLng(nextScopeMatch.FirstIndex) + 1
        Else
            endPosition = VBA.Len(parentText) + 1
        End If

        Set scopedParent = private_CloneContextValues(parentContext)
        scopedParent("$text") = VBA.Mid$( _
            parentText, startPosition, endPosition - startPosition)
        If Not private_CollectNestedContextItems( _
            datasetNode, scopedParent, nestedItems) Then Exit Function
        For Each nestedItem In nestedItems
            outItems.Add nestedItem
        Next nestedItem
    Next i

    private_CollectScopedNestedContextItems = True
End Function

Private Function private_CollectNestedContextItems( _
    ByVal datasetNode As Object, _
    ByVal parentContext As Object, _
    ByRef outItems As Collection _
) As Boolean
    Dim contextNodes As Object
    Dim contextNode As Object
    Dim inheritedValues As Object

    Set outItems = New Collection
    Set contextNodes = datasetNode.selectNodes( _
        "p:scope/p:context | p:context")
    Set inheritedValues = private_CloneContextValues(parentContext)
    If inheritedValues.Exists("$text") Then inheritedValues.Remove "$text"
    For Each contextNode In contextNodes
        If Not private_ExpandContextNode(contextNode, _
            VBA.CStr(parentContext("$text")), inheritedValues, outItems, _
            private_AttrOrDefault(datasetNode, "id", "?")) Then Exit Function
    Next contextNode
    private_CollectNestedContextItems = True
End Function

Private Function private_ApplyColumnRules( _
    ByVal valueText As String, _
    ByVal datasetNode As Object, _
    ByVal columnNode As Object, _
    ByVal values As Object _
) As String
    Dim ruleNodes As Object
    Dim ruleNode As Object
    Dim columnId As String
    Dim datasetId As String
    Dim ruleColumnId As String
    Dim excludedDatasets As String

    private_ApplyColumnRules = valueText
    columnId = VBA.LCase$(VBA.Trim$( _
        ex_XmlCore.fn_NodeAttrText(columnNode, "id")))
    datasetId = VBA.LCase$(VBA.Trim$( _
        ex_XmlCore.fn_NodeAttrText(datasetNode, "id")))
    Set ruleNodes = datasetNode.selectNodes( _
        "../../p:columnRules/p:rule")

    For Each ruleNode In ruleNodes
        ruleColumnId = VBA.LCase$(VBA.Trim$( _
            ex_XmlCore.fn_NodeAttrText(ruleNode, "columnId")))
        If VBA.StrComp(ruleColumnId, columnId, VBA.vbBinaryCompare) = 0 Then
            excludedDatasets = VBA.LCase$(VBA.Replace( _
                ex_XmlCore.fn_NodeAttrText(ruleNode, "excludeDatasets"), _
                " ", VBA.vbNullString))
            If VBA.InStr(1, ";" & excludedDatasets & ";", _
                ";" & datasetId & ";", VBA.vbBinaryCompare) = 0 Then
                valueText = private_ApplyTransforms( _
                    valueText, ruleNode, values)
            End If
        End If
    Next ruleNode
    private_ApplyColumnRules = valueText
End Function

Private Sub private_AddColumnNodeOrdered( _
    ByVal targetNodes As Collection, _
    ByVal datasetNode As Object, _
    ByVal columnNode As Object _
)
    Dim newOrder As Long
    Dim currentOrder As Long
    Dim i As Long

    newOrder = private_GetColumnOrder(datasetNode, columnNode)
    For i = 1 To targetNodes.Count
        currentOrder = private_GetColumnOrder( _
            datasetNode, targetNodes.Item(i))
        If newOrder < currentOrder Then
            targetNodes.Add columnNode, Before:=i
            Exit Sub
        End If
    Next i
    targetNodes.Add columnNode
End Sub

Private Function private_GetColumnOrder( _
    ByVal datasetNode As Object, _
    ByVal columnNode As Object _
) As Long
    Dim aliasText As String
    Dim datasetFlow As String

    aliasText = VBA.LCase$(VBA.Trim$( _
        ex_XmlCore.fn_NodeAttrText(columnNode, "id")))
    datasetFlow = VBA.LCase$(private_AttrOrDefault( _
        datasetNode, "flow", "other"))

    ' Основные кадровые колонки занимают стабильные позиции во всех таблицах.
    ' Неизвестные alias получают порядок 100 и остаются в хвосте в порядке DSL.
    ' Підстава обычно замыкает строку. Супровідний документ выводится после
    ' неё, когда такая специализированная колонка объявлена в dataset.
    Select Case aliasText
        Case "rank": private_GetColumnOrder = 5
        Case "fio": private_GetColumnOrder = 10
        Case "ipn": private_GetColumnOrder = 20
        Case "eventdate"
            If datasetFlow = "arrival" Then _
                private_GetColumnOrder = 70 Else private_GetColumnOrder = 30
        Case "exclusiondate": private_GetColumnOrder = 40
        Case "foodremovaldate": private_GetColumnOrder = 40
        Case "durationdays": private_GetColumnOrder = 50
        Case "eventto": private_GetColumnOrder = 60
        Case "foodenrollmentdate": private_GetColumnOrder = 80
        Case "arrivaldurationdays": private_GetColumnOrder = 81
        Case "position": private_GetColumnOrder = 82
        Case "assignmentorigin": private_GetColumnOrder = 83
        Case "returndate", "enrollmentdate": private_GetColumnOrder = 70
        Case "basis": private_GetColumnOrder = 1000
        Case "accompanyingdocument": private_GetColumnOrder = 1010
        Case Else: private_GetColumnOrder = 100
    End Select
End Function

Private Function private_ShouldIncludeCommonColumn( _
    ByVal datasetNode As Object, _
    ByVal columnNode As Object _
) As Boolean
    Dim datasetFlow As String
    Dim columnFlows As String
    Dim allowedColumnIds As String
    Dim columnId As String

    columnId = VBA.LCase$(VBA.Trim$( _
        ex_XmlCore.fn_NodeAttrText(columnNode, "id")))

    ' Звание является общей кадровой колонкой даже для составных datasets,
    ' отключивших остальные commonColumns. В служебные таблицы без ПІБ оно
    ' при этом не добавляется.
    If columnId = "rank" Then
        private_ShouldIncludeCommonColumn = Not _
            datasetNode.selectSingleNode("p:table/p:column[@id='fio']") Is Nothing
        Exit Function
    End If

    ' Некоторые составные события имеют собственную схему и не должны
    ' получать даже продовольственные commonColumns.
    If Not private_BoolAttr(datasetNode, _
        "includeCommonColumns", True) Then Exit Function

    ' Точечный allowlist нужен секциям, которые относятся к arrival/other,
    ' но используют лишь часть общих колонок. Например, первичное прибытие
    ' ставит на продовольствие, однако не возвращает человека к обязанностям.
    allowedColumnIds = VBA.LCase$(VBA.Replace(VBA.Replace( _
        ex_XmlCore.fn_NodeAttrText(datasetNode, "commonColumnIds"), _
        " ", VBA.vbNullString), ",", ";"))
    If VBA.Len(allowedColumnIds) > 0 Then
        If VBA.InStr(1, ";" & allowedColumnIds & ";", _
            ";" & columnId & ";", VBA.vbTextCompare) = 0 Then Exit Function
    End If

    ' flow управляет только представлением: commonFields вычисляются всегда,
    ' но arrival не показывает пустое "З продовольчого", а departure —
    ' пустое "На продовольче". Значение other сохраняет обе стороны события.
    ' Колонки TVO намеренно разрешены только для одного направления каждая:
    ' в нейтральных событиях (СЗЧ, смена статуса и т. п.) их быть не должно.
    datasetFlow = VBA.LCase$(private_AttrOrDefault( _
        datasetNode, "flow", "other"))
    columnFlows = VBA.LCase$(VBA.Replace( _
        ex_XmlCore.fn_NodeAttrText(columnNode, "flows"), " ", _
        VBA.vbNullString))
    If VBA.Len(columnFlows) = 0 Then
        private_ShouldIncludeCommonColumn = True
    Else
        private_ShouldIncludeCommonColumn = _
            (VBA.InStr(1, ";" & columnFlows & ";", _
            ";" & datasetFlow & ";", VBA.vbTextCompare) > 0)
    End If
End Function

Private Function private_CollectContextItems( _
    ByVal datasetNode As Object, _
    ByVal scopeText As String, _
    ByRef outItems As Collection _
) As Boolean
    Dim contextNodes As Object, contextNode As Object
    Dim inheritedValues As Object, item As Object

    Set outItems = New Collection
    Set contextNodes = datasetNode.selectNodes( _
        "p:scope/p:context | p:context")
    If contextNodes.Length = 0 Then
        Set item = VBA.CreateObject("Scripting.Dictionary")
        item.CompareMode = 1
        item("$text") = scopeText
        outItems.Add item
        private_CollectContextItems = True
        Exit Function
    End If

    ' Каждый корневой context разворачивается рекурсивно. В leaf-словаре
    ' остаются значения всех родителей, поэтому dataset/match видит, например,
    ' одновременно $dateBlock.1 и $subsection.1 при любой глубине вложенности.
    Set inheritedValues = VBA.CreateObject("Scripting.Dictionary")
    inheritedValues.CompareMode = 1
    For Each contextNode In contextNodes
        If Not private_ExpandContextNode( _
            contextNode, scopeText, inheritedValues, outItems, _
            private_AttrOrDefault(datasetNode, "id", "?")) Then Exit Function
    Next contextNode

    If outItems.Count = 0 Then
        private_ShowError "Не найдены context-блоки для dataset '" & _
            private_AttrOrDefault(datasetNode, "id", "?") & "'."
        Exit Function
    End If
    private_CollectContextItems = True
End Function

Private Function private_ExpandContextNode( _
    ByVal contextNode As Object, _
    ByVal parentText As String, _
    ByVal inheritedValues As Object, _
    ByVal outItems As Collection, _
    ByVal datasetId As String _
) As Boolean
    Dim contextId As String, contextKey As String
    Dim contextMatchNode As Object, childNodes As Object, childNode As Object
    Dim rx As Object, matches As Object, matchObj As Object
    Dim item As Object
    Dim matchIndex As Long, groupIndex As Long
    Dim contextText As String
    Dim itemsBefore As Long
    Dim isOptional As Boolean

    contextId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(contextNode, "id"))
    If VBA.Len(contextId) = 0 Then
        private_ShowError "У каждого context должен быть непустой id."
        Exit Function
    End If
    isOptional = private_BoolAttr(contextNode, "optional", False)
    Set contextMatchNode = contextNode.selectSingleNode("p:match")
    If contextMatchNode Is Nothing Then
        private_ShowError "Context '" & contextId & "' не содержит match."
        Exit Function
    End If

    Set rx = private_CreateRegex(VBA.CStr(contextMatchNode.Text), _
        private_BoolAttr(contextMatchNode, "ignoreCase", True), _
        private_BoolAttr(contextMatchNode, "multiline", False))
    If rx Is Nothing Then Exit Function
    Set matches = rx.Execute(parentText)
    groupIndex = VBA.CLng(VBA.Val( _
        private_AttrOrDefault(contextMatchNode, "group", "0")))
    Set childNodes = contextNode.selectNodes("p:context")
    itemsBefore = outItems.Count

    For Each matchObj In matches
        contextText = VBA.vbNullString
        If groupIndex = 0 Then
            contextText = VBA.CStr(matchObj.Value)
        ElseIf groupIndex <= matchObj.SubMatches.Count Then
            contextText = VBA.CStr(matchObj.SubMatches(groupIndex - 1))
        End If
        If VBA.Len(contextText) = 0 Then GoTo NextMatch

        ' Клонирование обязательно: соседние match одного context не должны
        ' перезаписывать у уже созданных строк значения родительских групп.
        Set item = private_CloneContextValues(inheritedValues)
        contextKey = "$" & contextId
        item(contextKey) = VBA.CStr(matchObj.Value)
        For matchIndex = 0 To matchObj.SubMatches.Count - 1
            item(contextKey & "." & VBA.CStr(matchIndex + 1)) = _
                VBA.CStr(matchObj.SubMatches(matchIndex))
        Next matchIndex

        If childNodes.Length = 0 Then
            item("$text") = contextText
            outItems.Add item
        Else
            For Each childNode In childNodes
                If Not private_ExpandContextNode( _
                    childNode, contextText, item, outItems, datasetId) Then Exit Function
            Next childNode
        End If
NextMatch:
    Next matchObj

    If outItems.Count = itemsBefore Then
        If isOptional Then
            private_ExpandContextNode = True
            Exit Function
        End If
        private_LogContextNotFound _
            datasetId, contextId, contextMatchNode, parentText
        private_ShowError "Не найдены блоки context '" & contextId & _
            "' для dataset '" & datasetId & _
            "'. Подробности и начало родительского текста записаны в лог."
        Exit Function
    End If
    private_ExpandContextNode = True
End Function

Private Function private_CloneContextValues( _
    ByVal sourceValues As Object _
) As Object
    Dim result As Object
    Dim key As Variant

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    If Not sourceValues Is Nothing Then
        For Each key In sourceValues.Keys
            result(VBA.CStr(key)) = sourceValues(key)
        Next key
    End If
    Set private_CloneContextValues = result
End Function

Private Sub private_LogContextNotFound( _
    ByVal datasetId As String, _
    ByVal contextId As String, _
    ByVal contextMatchNode As Object, _
    ByVal parentText As String _
)
    Dim previewText As String
    Dim patternText As String

    patternText = VBA.CStr(contextMatchNode.Text)
    previewText = VBA.Left$(parentText, 1500)
    previewText = VBA.Replace(previewText, VBA.vbCr, "<CR>")
    previewText = VBA.Replace(previewText, VBA.vbLf, "<LF>")
    previewText = VBA.Replace(previewText, "'", "''")

    ex_Core.fn_Diagnostic_LogError _
        "WordDataExtractor: context-not-found dataset='" & datasetId & _
        "' context='" & contextId & _
        "' parentLength=" & VBA.CStr(VBA.Len(parentText)) & _
        " pattern='" & VBA.Replace(patternText, "'", "''") & _
        "' preview='" & previewText & "'"
End Sub

Private Function private_CollectScopeTexts( _
    ByVal datasetNode As Object, _
    ByVal structureReferenceNode As Object, _
    ByVal documentText As String, _
    ByRef outScopeTexts As Collection _
) As Boolean
    Dim scopeMatchNode As Object
    Dim ranges As Collection
    Dim rangeItem As Variant
    Dim datasetId As String
    Dim sourceMode As String
    Dim startPosition As Long
    Dim endPosition As Long

    Set outScopeTexts = New Collection
    Set scopeMatchNode = datasetNode.selectSingleNode("p:scope/p:match")
    If scopeMatchNode Is Nothing Then
        sourceMode = VBA.LCase$(VBA.Trim$( _
            ex_XmlCore.fn_NodeAttrText(datasetNode, "source")))
        Select Case sourceMode
            Case "document"
                outScopeTexts.Add documentText
            Case "unclaimed"
                outScopeTexts.Add private_BuildUnclaimedText(documentText)
            Case Else
                private_ShowError "Dataset без scope должен явно задавать " & _
                    "source='document' или source='unclaimed': " & _
                    private_AttrOrDefault(datasetNode, "id", "?")
                Exit Function
        End Select
        private_CollectScopeTexts = True
        Exit Function
    End If

    datasetId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(datasetNode, "id"))
    If m_DatasetScopeRanges.Exists(datasetId) Then
        Set ranges = m_DatasetScopeRanges(datasetId)
        For Each rangeItem In ranges
            startPosition = VBA.CLng(rangeItem("start"))
            endPosition = VBA.CLng(rangeItem("end"))
            outScopeTexts.Add VBA.Mid$(documentText, startPosition, _
                endPosition - startPosition)
        Next rangeItem
    End If

    If outScopeTexts.Count = 0 Then
        ' Optional применяется только к отсутствующей секции целиком. Если
        ' заголовок найден, но вложенный context повреждён, это по-прежнему
        ' явная ошибка, а не молчаливо пропущенные данные.
        If private_MinOccurs(structureReferenceNode) = 0 Then
            private_CollectScopeTexts = True
            Exit Function
        End If
        private_LogScopeNotFound datasetNode, scopeMatchNode, documentText
        private_ShowError "Не найдена область scope для dataset '" & _
            private_AttrOrDefault(datasetNode, "id", "?") & _
            "'. Подробности и начало прочитанного текста записаны в лог."
        Exit Function
    End If
    private_CollectScopeTexts = True
End Function

Private Function private_BuildUnclaimedText( _
    ByVal documentText As String _
) As String
    Dim resultText As String
    Dim rangeItem As Variant
    Dim position As Long
    Dim startPosition As Long
    Dim endPosition As Long
    Dim currentChar As String

    resultText = documentText
    For Each rangeItem In m_ClaimedRanges
        startPosition = VBA.CLng(rangeItem("start"))
        endPosition = VBA.CLng(rangeItem("end")) - 1
        For position = startPosition To endPosition
            currentChar = VBA.Mid$(resultText, position, 1)
            If currentChar <> VBA.vbCr And currentChar <> VBA.vbLf Then
                VBA.Mid$(resultText, position, 1) = " "
            End If
        Next position
    Next rangeItem
    private_BuildUnclaimedText = resultText
End Function

Private Function private_ExtractField(ByVal fieldNode As Object, ByVal recordText As String, ByVal values As Object) As Boolean
    Dim fieldId As String, fromExpr As String, patternText As String, valueText As String
    Dim rx As Object, matches As Object, matchObj As Object, groupIndex As Long
    Dim existingValue As String

    fieldId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(fieldNode, "id"))
    If values.Exists("$" & fieldId) Then
        existingValue = VBA.CStr(values("$" & fieldId))
    End If
    fromExpr = private_AttrOrDefault(fieldNode, "from", "$match")
    valueText = private_ResolveValue(fromExpr, values)
    patternText = ex_XmlCore.fn_NodeAttrText(fieldNode, "regex")
    If VBA.Len(patternText) > 0 Then
        Set rx = private_CreateRegex(patternText, private_BoolAttr(fieldNode, "ignoreCase", True), private_BoolAttr(fieldNode, "multiline", True))
        If rx Is Nothing Then Exit Function
        Set matches = rx.Execute(valueText)
        ' Групповой dataset может извлечь поле из context раньше commonField.
        ' preserveExisting применяется только при отсутствии нового match:
        ' найденное персональное значение по-прежнему имеет приоритет.
        If matches.Count = 0 And _
            private_BoolAttr(fieldNode, "preserveExisting", False) And _
            VBA.Len(existingValue) > 0 Then
            valueText = existingValue
        Else
            valueText = private_AttrOrDefault( _
                fieldNode, "default", VBA.vbNullString)
            If matches.Count > 0 Then
                Set matchObj = matches.Item(0)
                groupIndex = VBA.CLng(VBA.Val( _
                    private_AttrOrDefault(fieldNode, "group", "1")))
                If groupIndex = 0 Then
                    valueText = VBA.CStr(matchObj.Value)
                ElseIf groupIndex <= matchObj.SubMatches.Count Then
                    valueText = VBA.CStr( _
                        matchObj.SubMatches(groupIndex - 1))
                End If
            End If
        End If
    End If
    values("$" & fieldId) = private_ApplyTransforms( _
        valueText, fieldNode, values)
    private_ExtractField = True
End Function

Private Function private_AddColumn(ByVal tableObj As obj_TableDynamic, ByVal columnNode As Object) As Boolean
    Dim colObj As obj_Column, aliasText As String
    aliasText = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(columnNode, "id"))
    Set colObj = New obj_Column
    colObj.Name = private_AttrOrDefault(columnNode, "caption", aliasText)
    colObj.Position = tableObj.ColumnCount + 1
    colObj.AddAlias aliasText
    colObj.FormatKind = VBA.LCase$(VBA.Trim$( _
        ex_XmlCore.fn_NodeAttrText(columnNode, "format")))
    If VBA.Len(colObj.FormatKind) = 0 Then
        ' Основные date-alias едины для всех datasets. Пометка FormatKind не
        ' форматирует значение сама: TableList превратит колонку в part=datelike,
        ' а итоговый NumberFormat задаст style pipeline страницы.
        If private_IsDateColumnAlias(aliasText) Then colObj.FormatKind = "date"
    End If
    private_AddColumn = tableObj.PushColumn(colObj)
End Function

Private Function private_IsDateColumnAlias(ByVal aliasText As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(aliasText))
        Case "eventdate", "eventto", "exclusiondate", _
             "enrollmentdate", "returndate", _
             "foodenrollmentdate", "foodremovaldate", _
             "returnfrom", "assignmentfrom"
            private_IsDateColumnAlias = True
    End Select
End Function

Private Function private_ApplyTransforms( _
    ByVal valueText As String, _
    ByVal ownerNode As Object, _
    ByVal values As Object _
) As String
    Dim nodes As Object, node As Object, transformType As String, rx As Object
    Dim parsedDate As Date
    Dim indexExpression As String
    Dim itemIndex As Long
    Dim selectedValue As String
    Set nodes = ownerNode.selectNodes("p:transform")
    For Each node In nodes
        transformType = VBA.LCase$(VBA.Trim$(ex_XmlCore.fn_NodeAttrText(node, "type")))
        Select Case transformType
            Case "trim": valueText = VBA.Trim$(valueText)
            Case "upper": valueText = VBA.UCase$(valueText)
            Case "lower": valueText = VBA.LCase$(valueText)
            Case "replace"
                valueText = VBA.Replace(valueText, _
                    ex_XmlCore.fn_NodeAttrText(node, "find"), _
                    private_GetTransformReplacement(node))
            Case "regexreplace"
                Set rx = private_CreateRegex(ex_XmlCore.fn_NodeAttrText(node, "pattern"), private_BoolAttr(node, "ignoreCase", True), True)
                If rx Is Nothing Then Exit Function
                rx.Global = True
                valueText = rx.Replace(valueText, private_GetTransformReplacement(node))
            Case "dateformat"
                If private_TryParseDateValue(valueText, parsedDate) Then
                    valueText = VBA.Format$(parsedDate, _
                        private_AttrOrDefault(node, "format", "dd.mm.yyyy"))
                End If
            Case "dateadddays"
                If private_TryParseDateValue(valueText, parsedDate) Then
                    parsedDate = VBA.DateAdd("d", _
                        VBA.CLng(VBA.Val(private_AttrOrDefault(node, "days", "0"))), _
                        parsedDate)
                    valueText = VBA.Format$(parsedDate, _
                        private_AttrOrDefault(node, "format", "dd.mm.yyyy"))
                End If
            Case "join"
                ' Составные ячейки описываются в DSL без специальных правил
                ' конкретного dataset. separatorToken позволяет безопасно
                ' задать tab/newline, которые неудобно хранить в XML-атрибуте.
                valueText = private_JoinResolvedValues( _
                    ex_XmlCore.fn_NodeAttrText(node, "values"), _
                    ex_XmlCore.fn_NodeAttrText(node, "separatorToken"), _
                    values)
            Case "daterangestartformat"
                If private_TryParseDateRangeStart(valueText, parsedDate) Then
                    valueText = VBA.Format$(parsedDate, _
                        private_AttrOrDefault(node, "format", "dd.mm.yyyy"))
                End If
            Case "orderedtripcredential"
                indexExpression = private_AttrOrDefault( _
                    node, "indexFrom", "$matchIndex")
                itemIndex = VBA.CLng(VBA.Val( _
                    private_ResolveValue(indexExpression, values)))
                If private_TrySelectTripCredential( _
                    valueText, itemIndex, selectedValue) Then
                    valueText = selectedValue
                End If
            Case "orderedrangereplace"
                indexExpression = private_AttrOrDefault( _
                    node, "indexFrom", "$matchIndex")
                itemIndex = VBA.CLng(VBA.Val( _
                    private_ResolveValue(indexExpression, values)))
                If private_TryApplyOrderedRangeReplace( _
                    valueText, itemIndex, node, selectedValue) Then
                    valueText = selectedValue
                End If
            Case "boolean"
                If VBA.Len(VBA.Trim$(valueText)) > 0 Then
                    valueText = private_AttrOrDefault(node, "trueValue", "true")
                Else
                    valueText = private_AttrOrDefault(node, "falseValue", "false")
                End If
            Case Else
                private_ShowError "Неподдерживаемый тип transform: " & transformType
                Exit Function
        End Select
    Next node
    private_ApplyTransforms = valueText
End Function

Private Function private_JoinResolvedValues( _
    ByVal expressionsText As String, _
    ByVal separatorToken As String, _
    ByVal values As Object _
) As String
    Dim expressions As Variant
    Dim expressionItem As Variant
    Dim separatorText As String
    Dim resultText As String
    Dim isFirstValue As Boolean

    Select Case VBA.LCase$(VBA.Trim$(separatorToken))
        Case "tab": separatorText = VBA.vbTab
        Case "newline": separatorText = VBA.vbCrLf
        Case "space": separatorText = " "
        Case Else: separatorText = separatorToken
    End Select

    expressions = VBA.Split(expressionsText, ";")
    isFirstValue = True
    For Each expressionItem In expressions
        If Not isFirstValue Then resultText = resultText & separatorText
        resultText = resultText & private_ResolveValue( _
            VBA.Trim$(VBA.CStr(expressionItem)), values)
        isFirstValue = False
    Next expressionItem
    private_JoinResolvedValues = resultText
End Function

Private Function private_TryApplyOrderedRangeReplace( _
    ByVal sourceText As String, _
    ByVal oneBasedIndex As Long, _
    ByVal transformNode As Object, _
    ByRef outValue As String _
) As Boolean
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim patternText As String
    Dim replacementTemplate As String
    Dim startGroup As Long
    Dim endGroup As Long
    Dim startNumber As Long
    Dim endNumber As Long
    Dim rangeCount As Long
    Dim selectedNumber As Long
    Dim replacementText As String

    outValue = sourceText
    If oneBasedIndex <= 0 Then Exit Function

    patternText = ex_XmlCore.fn_NodeAttrText(transformNode, "pattern")
    replacementTemplate = ex_XmlCore.fn_NodeAttrText( _
        transformNode, "replacement")
    ' startGroup/endGroup — однобазные номера круглых групп захвата pattern:
    ' их значения задают левую и правую числовые границы диапазона.
    ' {item} в replacement — вычисленный элемент для текущего matchIndex,
    ' тогда как {groupN} ниже переносит исходное значение N-й группы regex.
    startGroup = VBA.CLng(VBA.Val(private_AttrOrDefault( _
        transformNode, "startGroup", "1")))
    endGroup = VBA.CLng(VBA.Val(private_AttrOrDefault( _
        transformNode, "endGroup", "2")))
    Set rx = private_CreateRegex(patternText, _
        private_BoolAttr(transformNode, "ignoreCase", True), True)
    If rx Is Nothing Then Exit Function
    rx.Global = True
    Set matches = rx.Execute(sourceText)

    ' Отсутствие диапазона означает, что transform к этой записи неприменим.
    ' Это позволяет тем же rule обрабатывать обычные одиночные основания.
    If matches.Count = 0 Then Exit Function

    For Each matchObj In matches
        If startGroup <= 0 Or endGroup <= 0 Or _
           startGroup > matchObj.SubMatches.Count Or _
           endGroup > matchObj.SubMatches.Count Then
            private_ShowError "orderedRangeReplace ссылается на " & _
                "несуществующую группу regex."
            Exit Function
        End If
        If Not VBA.IsNumeric(matchObj.SubMatches(startGroup - 1)) Or _
           Not VBA.IsNumeric(matchObj.SubMatches(endGroup - 1)) Then
            private_ShowError "Границы orderedRangeReplace должны быть числами."
            Exit Function
        End If

        startNumber = VBA.CLng(matchObj.SubMatches(startGroup - 1))
        endNumber = VBA.CLng(matchObj.SubMatches(endGroup - 1))
        If endNumber < startNumber Then
            private_ShowError "Правая граница orderedRangeReplace меньше левой."
            Exit Function
        End If
        rangeCount = endNumber - startNumber + 1

        If oneBasedIndex <= rangeCount Then
            selectedNumber = startNumber + oneBasedIndex - 1
            replacementText = private_ExpandRangeReplacement( _
                replacementTemplate, selectedNumber, matchObj)
            ' FirstIndex имеет нулевую базу. Позиционная замена сохраняет весь
            ' окружающий текст и не привязывает движок к его предметному смыслу.
            outValue = VBA.Left$(sourceText, matchObj.FirstIndex) & _
                replacementText & _
                VBA.Mid$(sourceText, _
                    matchObj.FirstIndex + matchObj.Length + 1)
            private_TryApplyOrderedRangeReplace = True
            Exit Function
        End If
        oneBasedIndex = oneBasedIndex - rangeCount
    Next matchObj

    private_ShowError "Порядковый номер записи выходит за пределы " & _
        "диапазонов orderedRangeReplace."
End Function

Private Function private_ExpandRangeReplacement( _
    ByVal replacementTemplate As String, _
    ByVal selectedNumber As Long, _
    ByVal matchObj As Object _
) As String
    Dim groupIndex As Long
    Dim resultText As String

    resultText = VBA.Replace(replacementTemplate, _
        "{item}", VBA.CStr(selectedNumber))
    ' Именованные группы VBScript.RegExp не поддерживает, поэтому шаблон XML
    ' обращается к захватам как {group1}, {group2} и т. д.
    For groupIndex = matchObj.SubMatches.Count To 1 Step -1
        resultText = VBA.Replace(resultText, _
            "{group" & VBA.CStr(groupIndex) & "}", _
            VBA.CStr(matchObj.SubMatches(groupIndex - 1)))
    Next groupIndex
    private_ExpandRangeReplacement = resultText
End Function

Private Function private_TrySelectTripCredential( _
    ByVal basisText As String, _
    ByVal oneBasedIndex As Long, _
    ByRef outValue As String _
) As Boolean
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim startNumber As String
    Dim endNumber As String
    Dim credentialDate As String
    Dim credentialDateFormatted As String
    Dim startPrefix As String
    Dim endPrefix As String
    Dim startSequence As Long
    Dim endSequence As Long
    Dim rangeCount As Long
    Dim parsedDate As Date

    outValue = VBA.vbNullString
    If oneBasedIndex <= 0 Then Exit Function

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.MultiLine = True
    ' В документах встречаются оба варианта маркера диапазона:
    ' `№№ 47/117-47/118` и `№ № 47/117-47/118`.
    rx.Pattern = "посвідчення\s+про\s+відрядження\s+№\s*№?\s*" & _
        "([0-9]+(?:/[0-9]+)*)(?:\s*[-–—]\s*" & _
        "([0-9]+(?:/[0-9]+)*))?(?:\s+від\s+" & _
        "(\d{1,2}(?:[.\-/]\d{1,2}[.\-/]\d{2,4}|" & _
        "\s+[а-яіїєґ]+\s+\d{4}\s+року)))?"
    Set matches = rx.Execute(basisText)

    For Each matchObj In matches
        startNumber = VBA.CStr(matchObj.SubMatches(0))
        endNumber = VBA.CStr(matchObj.SubMatches(1))
        credentialDate = VBA.CStr(matchObj.SubMatches(2))
        credentialDateFormatted = VBA.Trim$(credentialDate)
        If VBA.Len(credentialDateFormatted) > 0 Then
            If private_TryParseDateValue(credentialDateFormatted, parsedDate) Then
                credentialDateFormatted = VBA.Format$(parsedDate, "dd.mm.yyyy")
            End If
        End If
        rangeCount = 1

        If VBA.Len(endNumber) > 0 Then
            If private_TrySplitCredentialNumber( _
                startNumber, startPrefix, startSequence) And _
               private_TrySplitCredentialNumber( _
                endNumber, endPrefix, endSequence) Then
                ' В сокращённой записи `47/117-118` правая граница наследует
                ' префикс слева; обратные и смешанные диапазоны не разворачиваем.
                If VBA.Len(endPrefix) = 0 Then endPrefix = startPrefix
                If VBA.StrComp(startPrefix, endPrefix, _
                    VBA.vbBinaryCompare) = 0 And endSequence >= startSequence Then
                    rangeCount = endSequence - startSequence + 1
                End If
            End If
        End If

        If oneBasedIndex <= rangeCount Then
            If rangeCount = 1 Then
                outValue = "посвідчення про відрядження № " & startNumber
            Else
                outValue = "посвідчення про відрядження № " & _
                    startPrefix & VBA.CStr(startSequence + oneBasedIndex - 1)
            End If
            If VBA.Len(credentialDateFormatted) > 0 Then
                outValue = outValue & " від " & credentialDateFormatted
            End If
            outValue = outValue & "."
            private_TrySelectTripCredential = True
            Exit Function
        End If
        oneBasedIndex = oneBasedIndex - rangeCount
    Next matchObj
End Function

Private Function private_TrySplitCredentialNumber( _
    ByVal numberText As String, _
    ByRef outPrefix As String, _
    ByRef outSequence As Long _
) As Boolean
    Dim slashPosition As Long
    Dim sequenceText As String

    outPrefix = VBA.vbNullString
    outSequence = 0
    numberText = VBA.Trim$(numberText)
    If VBA.Len(numberText) = 0 Then Exit Function

    slashPosition = VBA.InStrRev(numberText, "/")
    If slashPosition > 0 Then
        outPrefix = VBA.Left$(numberText, slashPosition)
        sequenceText = VBA.Mid$(numberText, slashPosition + 1)
    Else
        sequenceText = numberText
    End If
    If Not VBA.IsNumeric(sequenceText) Then Exit Function
    outSequence = VBA.CLng(sequenceText)
    private_TrySplitCredentialNumber = True
End Function

Private Function private_TryParseDateRangeStart( _
    ByVal valueText As String, _
    ByRef outDate As Date _
) As Boolean
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim startDateText As String
    Dim startMonthText As String
    Dim startYearText As String

    ' В приказах общие части диапазона не дублируются: встречаются как
    ' «з 11 липня по 09 серпня 2026 року», так и «з 11 по 30 липня 2026 року».
    ' Отсутствующие месяц/год наследуем справа; явно заданные имеют приоритет.
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "^\s*(\d{1,2})(?:\s+([а-яіїєґ]+))?" & _
        "(?:\s+(\d{4})\s+року)?\s+по\s+\d{1,2}\s+" & _
        "([а-яіїєґ]+)\s+(\d{4})\s+року\s*$"
    Set matches = rx.Execute(valueText)
    If matches.Count = 0 Then Exit Function

    Set matchObj = matches.Item(0)
    startMonthText = VBA.CStr(matchObj.SubMatches(1))
    If VBA.Len(startMonthText) = 0 Then
        startMonthText = VBA.CStr(matchObj.SubMatches(3))
    End If
    startDateText = VBA.CStr(matchObj.SubMatches(0)) & _
        " " & startMonthText
    startYearText = VBA.CStr(matchObj.SubMatches(2))
    If VBA.Len(startYearText) = 0 Then
        startYearText = VBA.CStr(matchObj.SubMatches(4))
    End If
    private_TryParseDateRangeStart = private_TryParseDateValue( _
        startDateText & " " & startYearText & " року", outDate)
End Function

Private Function private_TryParseDateValue( _
    ByVal valueText As String, _
    ByRef outDate As Date _
) As Boolean
    Dim rx As Object, matches As Object, matchObj As Object
    Dim monthText As String
    Dim monthNo As Long, dayNo As Long, yearNo As Long

    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    ' Числовой формат разбираем вручную до IsDate: иначе результат 01/02/2026
    ' зависит от региональных настроек Windows и может поменять день с месяцем.
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "^\s*(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2}|\d{4})\s*$"
    Set matches = rx.Execute(valueText)
    If matches.Count > 0 Then
        Set matchObj = matches.Item(0)
        dayNo = VBA.CLng(matchObj.SubMatches(0))
        monthNo = VBA.CLng(matchObj.SubMatches(1))
        yearNo = VBA.CLng(matchObj.SubMatches(2))
        If yearNo < 100 Then yearNo = 2000 + yearNo
        GoTo BuildDate
    End If

    If VBA.IsDate(valueText) Then
        outDate = VBA.CDate(valueText)
        private_TryParseDateValue = True
        Exit Function
    End If

    ' VBA.IsDate обычно не понимает украинские названия месяцев, поэтому
    ' текст приказа преобразуется через явное соответствие месяц -> номер.
    rx.Pattern = "^\s*(\d{1,2})\s+([а-яіїєґ]+)\s+(\d{4})(?:\s+року)?\s*$"
    Set matches = rx.Execute(valueText)
    If matches.Count = 0 Then Exit Function
    Set matchObj = matches.Item(0)
    dayNo = VBA.CLng(matchObj.SubMatches(0))
    monthText = VBA.LCase$(VBA.CStr(matchObj.SubMatches(1)))
    yearNo = VBA.CLng(matchObj.SubMatches(2))

    Select Case monthText
        Case "січня": monthNo = 1
        Case "лютого": monthNo = 2
        Case "березня": monthNo = 3
        Case "квітня": monthNo = 4
        Case "травня": monthNo = 5
        Case "червня": monthNo = 6
        Case "липня": monthNo = 7
        Case "серпня": monthNo = 8
        Case "вересня": monthNo = 9
        Case "жовтня": monthNo = 10
        Case "листопада": monthNo = 11
        Case "грудня": monthNo = 12
        Case Else: Exit Function
    End Select

BuildDate:
    On Error GoTo EH
    outDate = VBA.DateSerial(yearNo, monthNo, dayNo)
    private_TryParseDateValue = _
        (VBA.Day(outDate) = dayNo And VBA.Month(outDate) = monthNo And _
        VBA.Year(outDate) = yearNo)
    Exit Function
EH:
    private_TryParseDateValue = False
End Function

Private Function private_GetTransformReplacement(ByVal transformNode As Object) As String
    Dim replacementToken As String

    replacementToken = VBA.LCase$(VBA.Trim$( _
        ex_XmlCore.fn_NodeAttrText(transformNode, "withToken")))
    Select Case replacementToken
        Case "space"
            private_GetTransformReplacement = " "
        Case "tab"
            private_GetTransformReplacement = VBA.vbTab
        Case "newline"
            private_GetTransformReplacement = VBA.vbCrLf
        Case Else
            private_GetTransformReplacement = _
                ex_XmlCore.fn_NodeAttrText(transformNode, "with")
    End Select
End Function

Private Function private_ResolveValue(ByVal expression As String, ByVal values As Object) As String
    expression = VBA.Trim$(expression)
    If values.Exists(expression) Then
        private_ResolveValue = VBA.CStr(values(expression))
    ElseIf VBA.Left$(expression, 1) <> "$" Then
        private_ResolveValue = expression
    End If
End Function

Private Function private_CreateRegex(ByVal patternText As String, ByVal ignoreCase As Boolean, ByVal multiline As Boolean) As Object
    Dim rx As Object
    Dim resolving As Object
    Dim expandedPattern As String

    Set resolving = VBA.CreateObject("Scripting.Dictionary")
    resolving.CompareMode = 1
    If Not private_ExpandRegexText(patternText, resolving, _
        expandedPattern) Then Exit Function
    On Error GoTo EH
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = ignoreCase
    rx.Multiline = multiline
    rx.Pattern = expandedPattern
    Set private_CreateRegex = rx
    Exit Function
EH:
    private_ShowError "Некорректный regex '" & expandedPattern & "': " & _
        Err.Description
End Function

Private Function private_LoadRegexFragments() As Boolean
    Dim fragmentNodes As Object
    Dim fragmentNode As Object
    Dim fragmentId As String
    Dim resolving As Object
    Dim expandedText As String
    Dim fragmentKey As Variant

    Set m_RegexFragments = VBA.CreateObject("Scripting.Dictionary")
    m_RegexFragments.CompareMode = 1
    Set m_ExpandedRegexFragments = VBA.CreateObject("Scripting.Dictionary")
    m_ExpandedRegexFragments.CompareMode = 1
    Set fragmentNodes = m_Doc.selectNodes( _
        "/p:wordDataExtractor/p:regexFragments/p:fragment")
    For Each fragmentNode In fragmentNodes
        fragmentId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
            fragmentNode, "id"))
        If VBA.Len(fragmentId) = 0 Then
            private_ShowError "Regex fragment должен иметь непустой id."
            Exit Function
        End If
        If Not private_IsValidRegexFragmentId(fragmentId) Then
            private_ShowError "Некорректный id regex fragment: " & fragmentId
            Exit Function
        End If
        If m_RegexFragments.Exists(fragmentId) Then
            private_ShowError "Дублирующийся regex fragment: " & fragmentId
            Exit Function
        End If
        m_RegexFragments.Add fragmentId, VBA.CStr(fragmentNode.Text)
    Next fragmentNode

    For Each fragmentKey In m_RegexFragments.Keys
        Set resolving = VBA.CreateObject("Scripting.Dictionary")
        resolving.CompareMode = 1
        If Not private_ResolveRegexFragment(VBA.CStr(fragmentKey), _
            resolving, expandedText) Then Exit Function
    Next fragmentKey
    private_LoadRegexFragments = True
End Function

Private Function private_IsValidRegexFragmentId( _
    ByVal fragmentId As String _
) As Boolean
    Dim rx As Object

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = "^[A-Za-z0-9_-]+$"
    private_IsValidRegexFragmentId = rx.Test(fragmentId)
End Function

Private Function private_ValidateRegexFragmentUsage() As Boolean
    Dim regexNodes As Object
    Dim regexNode As Object
    Dim patternText As String
    Dim expandedText As String
    Dim resolving As Object

    Set regexNodes = m_Doc.selectNodes( _
        "//p:match | //p:field[@regex] | //p:transform[@pattern]")
    For Each regexNode In regexNodes
        Select Case VBA.LCase$(VBA.CStr(regexNode.nodeName))
            Case "match"
                patternText = VBA.CStr(regexNode.Text)
            Case "field"
                patternText = ex_XmlCore.fn_NodeAttrText(regexNode, "regex")
            Case "transform"
                patternText = ex_XmlCore.fn_NodeAttrText(regexNode, "pattern")
        End Select
        Set resolving = VBA.CreateObject("Scripting.Dictionary")
        resolving.CompareMode = 1
        If Not private_ExpandRegexText(patternText, resolving, _
            expandedText) Then Exit Function
    Next regexNode
    private_ValidateRegexFragmentUsage = True
End Function

Private Function private_ResolveRegexFragment( _
    ByVal fragmentId As String, _
    ByVal resolving As Object, _
    ByRef outText As String _
) As Boolean
    Dim expandedText As String

    If m_ExpandedRegexFragments.Exists(fragmentId) Then
        outText = VBA.CStr(m_ExpandedRegexFragments(fragmentId))
        private_ResolveRegexFragment = True
        Exit Function
    End If
    If Not m_RegexFragments.Exists(fragmentId) Then
        private_ShowError "Не найден regex fragment: " & fragmentId
        Exit Function
    End If
    If resolving.Exists(fragmentId) Then
        private_ShowError "Циклическая ссылка regex fragment: " & fragmentId
        Exit Function
    End If
    resolving.Add fragmentId, True
    If Not private_ExpandRegexText(VBA.CStr(m_RegexFragments(fragmentId)), _
        resolving, expandedText) Then Exit Function
    resolving.Remove fragmentId
    m_ExpandedRegexFragments.Add fragmentId, expandedText
    outText = expandedText
    private_ResolveRegexFragment = True
End Function

Private Function private_ExpandRegexText( _
    ByVal patternText As String, _
    ByVal resolving As Object, _
    ByRef outText As String _
) As Boolean
    Dim tokenRx As Object
    Dim matches As Object
    Dim tokenMatch As Object
    Dim fragmentId As String
    Dim fragmentText As String
    Dim resultText As String

    Set tokenRx = VBA.CreateObject("VBScript.RegExp")
    tokenRx.Global = True
    tokenRx.Pattern = "\{\{([A-Za-z0-9_-]+)\}\}"
    resultText = patternText
    Do
        Set matches = tokenRx.Execute(resultText)
        If matches.Count = 0 Then Exit Do
        Set tokenMatch = matches.Item(0)
        fragmentId = VBA.CStr(tokenMatch.SubMatches(0))
        If Not private_ResolveRegexFragment(fragmentId, resolving, _
            fragmentText) Then Exit Function
        resultText = VBA.Left$(resultText, tokenMatch.FirstIndex) & _
            "(?:" & fragmentText & ")" & _
            VBA.Mid$(resultText, tokenMatch.FirstIndex + _
            tokenMatch.Length + 1)
    Loop
    If VBA.InStr(1, resultText, "{{", VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, resultText, "}}", VBA.vbBinaryCompare) > 0 Then
        private_ShowError "Некорректная ссылка regex fragment в выражении: " & _
            patternText
        Exit Function
    End If
    outText = resultText
    private_ExpandRegexText = True
End Function

Private Function private_ValidateRules() As Boolean
    Dim nodes As Object, node As Object
    Dim ruleNodes As Object, ruleNode As Object
    Dim referenceNodes As Object, referenceNode As Object
    Dim resolvedRuleNode As Object
    Dim linkedStructureNode As Object
    Dim ifNodes As Object
    Dim ifNode As Object
    Dim ruleId As String
    Dim ruleKind As String
    Dim sourceMode As String

    private_ValidateRules = private_ValidatePatternStructure()
    Exit Function

    Set nodes = m_Doc.selectNodes("/p:wordDataExtractor/p:pipelines/p:pipeline")
    If nodes Is Nothing Then
        private_ShowError "Файл rules не содержит pipelines."
        Exit Function
    End If
    If nodes.Length = 0 Then
        private_ShowError "Файл rules не содержит pipelines."
        Exit Function
    End If
    For Each node In nodes
        If VBA.Len(VBA.Trim$(ex_XmlCore.fn_NodeAttrText(node, "id"))) = 0 Then
            private_ShowError "Каждый pipeline должен иметь непустой id."
            Exit Function
        End If
        If node.selectSingleNode("p:structure/p:template") Is Nothing Then
            private_ShowError "Каждый pipeline должен содержать " & _
                "structure/template."
            Exit Function
        End If
        Set referenceNodes = node.selectNodes( _
            "p:structure//*[not(self::p:template or " & _
            "self::p:content or self::p:requiredContent or " & _
            "self::p:items or self::p:requiredItems or " & _
            "self::p:itemTemplate or self::p:boundary or " & _
            "self::p:if or self::p:then or self::p:else)]")
        If referenceNodes.Length > 0 Then
            private_ShowError "Неизвестный узел structure: " & _
                referenceNodes.Item(0).nodeName
            Exit Function
        End If
        Set referenceNodes = node.selectNodes( _
            "p:structure//p:content[@source] | " & _
            "p:structure//p:requiredContent[@source] | " & _
            "p:structure//p:items[@source] | " & _
            "p:structure//p:requiredItems[@source] | " & _
            "p:structure//p:boundary[@source]")
        If referenceNodes.Length = 0 Then
            private_ShowError "Каждый pipeline должен содержать " & _
                "structure с хотя бы одним исполняемым узлом."
            Exit Function
        End If
        For Each referenceNode In referenceNodes
            If Not private_ValidateOccurrenceSpec(referenceNode) Then _
                Exit Function
            ruleId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
                referenceNode, "source"))
            Set resolvedRuleNode = node.selectSingleNode( _
                "p:rules/p:rule[@id=" & _
                ex_XmlCore.fn_XPathLiteral(ruleId) & "]")
            If resolvedRuleNode Is Nothing Then
                private_ShowError "Structure ссылается на неизвестное " & _
                    "правило: " & ruleId
                Exit Function
            End If
            ruleKind = VBA.LCase$(private_AttrOrDefault( _
                resolvedRuleNode, "kind", "extract"))
            Select Case VBA.LCase$(referenceNode.nodeName)
                Case "boundary"
                    If ruleKind <> "boundary" Then
                        private_ShowError "Узел boundary должен ссылаться " & _
                            "на rule kind='boundary': " & ruleId
                        Exit Function
                    End If
                Case "content", "requiredcontent"
                    If ruleKind = "boundary" Then
                        private_ShowError "Узел content не может ссылаться " & _
                            "на boundary-rule: " & ruleId
                        Exit Function
                    End If
                Case "items", "requireditems"
                    If ruleKind = "boundary" Then
                        private_ShowError "Узел items не может ссылаться " & _
                            "на boundary-rule: " & ruleId
                        Exit Function
                    End If
            End Select
            If ruleKind = "section" Then
                Select Case VBA.LCase$(referenceNode.nodeName)
                    Case "content", "requiredcontent"
                        If referenceNode.selectSingleNode( _
                            "p:template") Is Nothing Then
                            private_ShowError "Section-source в content " & _
                                "должен содержать template: " & ruleId
                            Exit Function
                        End If
                    Case "items", "requireditems"
                        If referenceNode.selectSingleNode( _
                            "p:itemTemplate") Is Nothing Then
                            private_ShowError "Section-source в items должен " & _
                                "содержать itemTemplate: " & ruleId
                            Exit Function
                        End If
                End Select
            End If
        Next referenceNode
        Set referenceNodes = node.selectNodes("p:structure//p:if")
        For Each referenceNode In referenceNodes
            ruleId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
                referenceNode, "sourceExists"))
            If VBA.Len(ruleId) = 0 Then
                private_ShowError "if должен содержать непустой sourceExists."
                Exit Function
            End If
            Set resolvedRuleNode = node.selectSingleNode( _
                "p:rules/p:rule[@id=" & _
                ex_XmlCore.fn_XPathLiteral(ruleId) & "]")
            If resolvedRuleNode Is Nothing Then
                private_ShowError "sourceExists ссылается на неизвестное " & _
                    "правило: " & ruleId
                Exit Function
            End If
        Next referenceNode
        Set ifNodes = node.selectNodes("p:structure//p:if")
        For Each ifNode In ifNodes
            If ifNode.selectNodes("p:then").Length <> 1 Then
                private_ShowError "if должен содержать ровно один then."
                Exit Function
            End If
            If ifNode.selectNodes("p:else").Length > 1 Then
                private_ShowError "if может содержать только один else."
                Exit Function
            End If
        Next ifNode
        Set referenceNodes = node.selectNodes( _
            "p:structure//p:then[not(parent::p:if)]")
        If referenceNodes.Length > 0 Then
            private_ShowError "then допустим только как дочерний узел if."
            Exit Function
        End If
        Set referenceNodes = node.selectNodes( _
            "p:structure//p:else[not(parent::p:if)]")
        If referenceNodes.Length > 0 Then
            private_ShowError "else допустим только как дочерний узел if."
            Exit Function
        End If
        Set ruleNodes = node.selectNodes("p:rules/p:rule")
        For Each ruleNode In ruleNodes
            ruleKind = VBA.LCase$(private_AttrOrDefault( _
                ruleNode, "kind", "extract"))
            If ruleKind <> "extract" And ruleKind <> "section" And _
                ruleKind <> "boundary" Then
                private_ShowError "Неизвестный kind правила: " & _
                    private_AttrOrDefault(ruleNode, "id", "?")
                Exit Function
            End If
            If ruleKind = "extract" And _
                ruleNode.selectSingleNode("p:scope") Is Nothing Then
                Set linkedStructureNode = node.selectSingleNode( _
                    "p:structure//*[@source=" & _
                    ex_XmlCore.fn_XPathLiteral(private_AttrOrDefault( _
                    ruleNode, "id", "?")) & "]")
                If Not linkedStructureNode Is Nothing Then
                    If Not linkedStructureNode.selectSingleNode( _
                        "ancestor::p:content | ancestor::p:items | " & _
                        "ancestor::p:requiredContent | " & _
                        "ancestor::p:requiredItems") _
                        Is Nothing Then GoTo NextRuleValidation
                End If
                sourceMode = VBA.LCase$(VBA.Trim$( _
                    ex_XmlCore.fn_NodeAttrText(ruleNode, "source")))
                If sourceMode <> "document" And sourceMode <> "unclaimed" Then
                    private_ShowError "Rule без scope должен задавать " & _
                        "source='document' или source='unclaimed': " & _
                        private_AttrOrDefault(ruleNode, "id", "?")
                    Exit Function
                End If
            End If
NextRuleValidation:
        Next ruleNode
    Next node
    private_ValidateRules = True
End Function

Private Function private_ValidatePatternStructure() As Boolean
    Dim pipelineNodes As Object
    Dim pipelineNode As Object
    Dim invalidNodes As Object
    Dim structureNodes As Object
    Dim structureNode As Object
    Dim pathNodes As Object
    Dim pathNode As Object
    Dim patternNodes As Object
    Dim patternNode As Object
    Dim segments As Collection
    Dim segment As Object
    Dim ruleNode As Object
    Dim locatorNode As Object
    Dim ruleId As String
    Dim ruleKind As String
    Dim i As Long

    Set pipelineNodes = m_Doc.selectNodes( _
        "/p:wordDataExtractor/p:pipelines/p:pipeline")
    If pipelineNodes Is Nothing Or pipelineNodes.Length = 0 Then
        private_ShowError "Файл rules не содержит pipelines."
        Exit Function
    End If

    For Each pipelineNode In pipelineNodes
        If VBA.Len(VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
            pipelineNode, "id"))) = 0 Then
            private_ShowError "Каждый pipeline должен иметь непустой id."
            Exit Function
        End If
        Set structureNode = pipelineNode.selectSingleNode("p:structure")
        If structureNode Is Nothing Then
            private_ShowError "Каждый pipeline должен содержать structure."
            Exit Function
        End If
        Set invalidNodes = structureNode.selectNodes( _
            ".//*[not(self::p:context or self::p:pattern or " & _
            "self::p:requiredPattern)]")
        If invalidNodes.Length > 0 Then
            private_ShowError "Неизвестный узел structure: " & _
                invalidNodes.Item(0).nodeName
            Exit Function
        End If
        Set structureNodes = structureNode.selectNodes( _
            "p:context | p:pattern | p:requiredPattern")
        If structureNodes.Length = 0 Then
            private_ShowError "Structure должна содержать context или pattern."
            Exit Function
        End If

        Set pathNodes = structureNode.selectNodes(".//p:context")
        For Each pathNode In pathNodes
            If Not private_ParseStructurePath( _
                VBA.CStr(ex_XmlCore.fn_NodeAttrText(pathNode, "path")), _
                segments) Then Exit Function
            For i = 1 To segments.Count
                Set segment = segments.Item(i)
                ruleId = VBA.CStr(segment("id"))
                Set ruleNode = pipelineNode.selectSingleNode( _
                    "p:rules/p:rule[@id=" & _
                    ex_XmlCore.fn_XPathLiteral(ruleId) & "]")
                If ruleNode Is Nothing Then
                    private_ShowError "Context ссылается на неизвестное " & _
                        "правило: " & ruleId
                    Exit Function
                End If
                Set locatorNode = private_RuleLocatorNode(ruleNode)
                If locatorNode Is Nothing Then
                    private_ShowError "Структурное правило не содержит " & _
                        "match: " & ruleId
                    Exit Function
                End If
            Next i
        Next pathNode

        Set patternNodes = structureNode.selectNodes( _
            ".//p:pattern | .//p:requiredPattern")
        For Each patternNode In patternNodes
            If VBA.Len(VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
                patternNode, "id"))) = 0 Then
                private_ShowError "Каждый pattern должен иметь непустой id."
                Exit Function
            End If
            If Not private_ParseStructurePath( _
                VBA.CStr(patternNode.Text), segments) Then Exit Function
            If segments.Count = 0 Then
                private_ShowError "Pattern не содержит ссылок на rules: " & _
                    ex_XmlCore.fn_NodeAttrText(patternNode, "id")
                Exit Function
            End If
            For i = 1 To segments.Count
                Set segment = segments.Item(i)
                ruleId = VBA.CStr(segment("id"))
                Set ruleNode = pipelineNode.selectSingleNode( _
                    "p:rules/p:rule[@id=" & _
                    ex_XmlCore.fn_XPathLiteral(ruleId) & "]")
                If ruleNode Is Nothing Then
                    private_ShowError "Pattern ссылается на неизвестное " & _
                        "правило: " & ruleId
                    Exit Function
                End If
                ruleKind = VBA.LCase$(private_AttrOrDefault( _
                    ruleNode, "kind", "extract"))
                If i = segments.Count Then
                    If ruleKind <> "extract" Then
                        private_ShowError "Последняя ссылка pattern должна " & _
                            "указывать на extract-rule: " & ruleId
                        Exit Function
                    End If
                Else
                    Set locatorNode = private_RuleLocatorNode(ruleNode)
                    If locatorNode Is Nothing Then
                        private_ShowError "Промежуточное правило pattern не " & _
                            "содержит match: " & ruleId
                        Exit Function
                    End If
                End If
            Next i
        Next patternNode
    Next pipelineNode
    private_ValidatePatternStructure = True
End Function

Private Sub private_LogScopeNotFound( _
    ByVal datasetNode As Object, _
    ByVal scopeMatchNode As Object, _
    ByVal documentText As String _
)
    Dim previewText As String
    Dim patternText As String

    patternText = VBA.CStr(scopeMatchNode.Text)
    previewText = VBA.Left$(documentText, 1500)
    previewText = VBA.Replace(previewText, VBA.vbCr, "<CR>")
    previewText = VBA.Replace(previewText, VBA.vbLf, "<LF>")
    previewText = VBA.Replace(previewText, "'", "''")

    ex_Core.fn_Diagnostic_LogError _
        "WordDataExtractor: scope-not-found dataset='" & _
        private_AttrOrDefault(datasetNode, "id", "?") & _
        "' documentLength=" & VBA.CStr(VBA.Len(documentText)) & _
        " pattern='" & VBA.Replace(patternText, "'", "''") & _
        "' preview='" & previewText & "'"
End Sub

Private Function private_BoolAttr(ByVal node As Object, ByVal attrName As String, ByVal defaultValue As Boolean) As Boolean
    Dim textValue As String
    textValue = VBA.LCase$(VBA.Trim$(ex_XmlCore.fn_NodeAttrText(node, attrName)))
    If VBA.Len(textValue) = 0 Then private_BoolAttr = defaultValue Else private_BoolAttr = (textValue = "true" Or textValue = "1" Or textValue = "yes")
End Function

Private Function private_AttrOrDefault(ByVal node As Object, ByVal attrName As String, ByVal defaultValue As String) As String
    private_AttrOrDefault = ex_XmlCore.fn_NodeAttrText(node, attrName)
    If VBA.Len(private_AttrOrDefault) = 0 Then private_AttrOrDefault = defaultValue
End Function

Private Sub private_ShowError(ByVal messageText As String)
    ex_Core.fn_Diagnostic_LogError "WordDataExtractor: " & messageText
    VBA.MsgBox messageText, VBA.vbExclamation, "PrototypeNew / WordDataExtractor"
End Sub
