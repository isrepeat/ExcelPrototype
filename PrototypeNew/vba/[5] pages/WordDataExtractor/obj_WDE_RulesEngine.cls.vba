VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_WDE_RulesEngine"
Option Explicit

Private Const RULES_NS As String = "urn:excelprototype:word-data-extractor:v1"
Private m_Doc As Object
Private m_RulesRelPath As String
Private m_IsDisposed As Boolean
Private m_ScopeBoundaryPositions As Collection

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
    m_IsDisposed = False
    Initialize = private_ValidateRules()
End Function

Public Sub Dispose()
    m_IsDisposed = True
    Set m_Doc = Nothing
    Set m_ScopeBoundaryPositions = Nothing
    m_RulesRelPath = VBA.vbNullString
End Sub

Public Function ExtractTables(ByVal pipelineId As String, ByVal documentText As String, ByRef outTables As Collection) As Boolean
    Dim pipelineNode As Object
    Dim datasetNodes As Object
    Dim datasetNode As Object

    Set outTables = Nothing
    If m_IsDisposed Or m_Doc Is Nothing Then Exit Function
    Set pipelineNode = m_Doc.selectSingleNode("/p:wordDataExtractor/p:pipelines/p:pipeline[@id=" & ex_XmlCore.fn_XPathLiteral(pipelineId) & "]")
    If pipelineNode Is Nothing Then
        private_ShowError "Pipeline не найден: " & pipelineId
        Exit Function
    End If
    Set outTables = New Collection
    If Not private_CollectScopeBoundaryPositions( _
        pipelineNode, documentText) Then Exit Function
    Set datasetNodes = pipelineNode.selectNodes("p:dataset")
    For Each datasetNode In datasetNodes
        If Not private_ExtractDataset(datasetNode, documentText, outTables) Then Exit Function
    Next datasetNode
    Set m_ScopeBoundaryPositions = Nothing
    ExtractTables = True
End Function

Private Function private_CollectScopeBoundaryPositions( _
    ByVal pipelineNode As Object, _
    ByVal documentText As String _
) As Boolean
    Dim scopeNodes As Object
    Dim scopeNode As Object
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object

    Set m_ScopeBoundaryPositions = New Collection
    Set scopeNodes = pipelineNode.selectNodes( _
        "p:dataset/p:scope/p:match")
    For Each scopeNode In scopeNodes
        Set rx = private_CreateRegex(VBA.CStr(scopeNode.Text), _
            private_BoolAttr(scopeNode, "ignoreCase", True), _
            private_BoolAttr(scopeNode, "multiline", False))
        If rx Is Nothing Then Exit Function
        Set matches = rx.Execute(documentText)
        For Each matchObj In matches
            private_AddScopeBoundaryPosition _
                VBA.CLng(matchObj.FirstIndex) + 1
        Next matchObj
    Next scopeNode
    private_CollectScopeBoundaryPositions = True
End Function

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

Private Function private_ExtractDataset(ByVal datasetNode As Object, ByVal sourceText As String, ByVal tables As Collection) As Boolean
    Dim rx As Object, matches As Object, matchObj As Object
    Dim tableObj As obj_TableDynamic, rowObj As obj_Row
    Dim fieldValues As Object, columnNodes As Object, columnNode As Object
    Dim fieldNodes As Object, fieldNode As Object
    Dim patternNode As Object, patternText As String, datasetId As String
    Dim matchIndex As Long, datasetMatchIndex As Long, valueText As String
    Dim scopeTexts As Collection, scopeText As Variant
    Dim contextItems As Collection, contextItem As Variant
    Dim contextKey As Variant
    Dim localNodes As Object, commonNodes As Object

    datasetId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(datasetNode, "id"))
    Set patternNode = datasetNode.selectSingleNode("p:match")
    patternText = VBA.CStr(patternNode.Text)
    Set rx = private_CreateRegex(patternText, private_BoolAttr(patternNode, "ignoreCase", True), private_BoolAttr(patternNode, "multiline", True))
    If rx Is Nothing Then Exit Function
    If Not private_CollectScopeTexts(datasetNode, sourceText, scopeTexts) Then Exit Function

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = private_AttrOrDefault(datasetNode, "caption", datasetId)
    tableObj.SourceAlias = datasetId
    tableObj.SourceAliasTemplate = "WordDataExtractor"
    ' Сначала сохраняем специфичные колонки dataset, затем добавляем общие.
    ' XPath union вернул бы узлы в порядке XML-документа и поставил бы
    ' commonColumns перед ПІБ/датой события, поскольку они объявлены выше.
    Set columnNodes = New Collection
    Set localNodes = datasetNode.selectNodes("p:table/p:column")
    For Each columnNode In localNodes
        private_AddColumnNodeOrdered columnNodes, datasetNode, columnNode
    Next columnNode
    Set commonNodes = datasetNode.selectNodes("../p:commonColumns/p:column")
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
    Set commonNodes = datasetNode.selectNodes("../p:commonFields/p:field")
    For Each fieldNode In commonNodes
        fieldNodes.Add fieldNode
    Next fieldNode
    For Each scopeText In scopeTexts
        If Not private_CollectContextItems( _
            datasetNode, VBA.CStr(scopeText), contextItems) Then Exit Function
        For Each contextItem In contextItems
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
                fieldValues("$matchIndex") = datasetMatchIndex
                For matchIndex = 0 To matchObj.SubMatches.Count - 1
                    fieldValues("$" & VBA.CStr(matchIndex + 1)) = VBA.CStr(matchObj.SubMatches(matchIndex))
                Next matchIndex
                For Each fieldNode In fieldNodes
                    If Not private_ExtractField(fieldNode, VBA.CStr(matchObj.Value), fieldValues) Then Exit Function
                Next fieldNode

                Set rowObj = New obj_Row
                For Each columnNode In columnNodes
                    valueText = private_ResolveValue(ex_XmlCore.fn_NodeAttrText(columnNode, "value"), fieldValues)
                    valueText = private_ApplyTransforms( _
                        valueText, columnNode, fieldValues)
                    valueText = private_ApplyColumnRules( _
                        valueText, datasetNode, columnNode, fieldValues)
                    rowObj.PushCellRaw valueText
                Next columnNode
                If Not tableObj.PushRow(rowObj) Then Exit Function
            Next matchObj
        Next contextItem
    Next scopeText
    ' Пустые таблицы также возвращаются вызывающему коду: UI самостоятельно
    ' решает, показывать все datasets или только содержащие строки.
    tables.Add tableObj
    private_ExtractDataset = True
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
    Set ruleNodes = datasetNode.selectNodes("../p:columnRules/p:rule")

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
    ' Підстава всегда замыкает строку, включая datasets с дополнительными
    ' специализированными колонками, неизвестными общему движку.
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

    contextId = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(contextNode, "id"))
    If VBA.Len(contextId) = 0 Then
        private_ShowError "У каждого context должен быть непустой id."
        Exit Function
    End If
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
    ByVal documentText As String, _
    ByRef outScopeTexts As Collection _
) As Boolean
    Dim scopeMatchNode As Object
    Dim rx As Object, matches As Object, matchObj As Object
    Dim groupIndex As Long, scopeText As String
    Dim relativeScopeStart As Long
    Dim absoluteScopeStart As Long

    Set outScopeTexts = New Collection
    Set scopeMatchNode = datasetNode.selectSingleNode("p:scope/p:match")
    If scopeMatchNode Is Nothing Then
        outScopeTexts.Add documentText
        private_CollectScopeTexts = True
        Exit Function
    End If

    Set rx = private_CreateRegex(VBA.CStr(scopeMatchNode.Text), _
        private_BoolAttr(scopeMatchNode, "ignoreCase", True), _
        private_BoolAttr(scopeMatchNode, "multiline", False))
    If rx Is Nothing Then Exit Function
    Set matches = rx.Execute(documentText)
    groupIndex = VBA.CLng(VBA.Val(private_AttrOrDefault(scopeMatchNode, "group", "0")))
    For Each matchObj In matches
        scopeText = VBA.vbNullString
        If groupIndex = 0 Then
            scopeText = VBA.CStr(matchObj.Value)
        ElseIf groupIndex <= matchObj.SubMatches.Count Then
            scopeText = VBA.CStr(matchObj.SubMatches(groupIndex - 1))
        End If
        If VBA.Len(scopeText) > 0 Then
            ' Scope regex отвечает за распознавание собственного заголовка.
            ' Фактический конец централизованно ограничивается ближайшим
            ' началом любой другой присутствующей секции из того же pipeline.
            relativeScopeStart = VBA.InStr(1, VBA.CStr(matchObj.Value), _
                scopeText, VBA.vbBinaryCompare)
            If relativeScopeStart = 0 Then relativeScopeStart = 1
            absoluteScopeStart = VBA.CLng(matchObj.FirstIndex) + _
                relativeScopeStart
            scopeText = private_ClipScopeAtNextBoundary( _
                scopeText, absoluteScopeStart)
            outScopeTexts.Add scopeText
        End If
    Next matchObj

    If outScopeTexts.Count = 0 Then
        ' Optional применяется только к отсутствующей секции целиком. Если
        ' заголовок найден, но вложенный context повреждён, это по-прежнему
        ' явная ошибка, а не молчаливо пропущенные данные.
        If private_BoolAttr(datasetNode, "optional", False) Then
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

Private Function private_ClipScopeAtNextBoundary( _
    ByVal scopeText As String, _
    ByVal absoluteScopeStart As Long _
) As String
    Dim boundaryPosition As Variant
    Dim scopeEnd As Long

    private_ClipScopeAtNextBoundary = scopeText
    If m_ScopeBoundaryPositions Is Nothing Then Exit Function
    scopeEnd = absoluteScopeStart + VBA.Len(scopeText)
    For Each boundaryPosition In m_ScopeBoundaryPositions
        If VBA.CLng(boundaryPosition) > absoluteScopeStart Then
            If VBA.CLng(boundaryPosition) < scopeEnd Then
                private_ClipScopeAtNextBoundary = VBA.Left$(scopeText, _
                    VBA.CLng(boundaryPosition) - absoluteScopeStart)
            End If
            Exit Function
        End If
    Next boundaryPosition
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
    If values.Exists(expression) Then private_ResolveValue = VBA.CStr(values(expression)) Else private_ResolveValue = expression
End Function

Private Function private_CreateRegex(ByVal patternText As String, ByVal ignoreCase As Boolean, ByVal multiline As Boolean) As Object
    Dim rx As Object
    On Error GoTo EH
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = ignoreCase
    rx.Multiline = multiline
    rx.Pattern = patternText
    Set private_CreateRegex = rx
    Exit Function
EH:
    private_ShowError "Некорректный regex '" & patternText & "': " & Err.Description
End Function

Private Function private_ValidateRules() As Boolean
    Dim nodes As Object, node As Object
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
    Next node
    private_ValidateRules = True
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
