Attribute VB_Name = "ex_ConfigVirtualSources"
Option Explicit

Private Const RUNTIME_EXPANDED_FLAG_KEY As String = "__Runtime.VirtualSources.Expanded"
Private Const ERR_INVALID_CONFIG As Long = 1760
Private Const ERR_RESOLVER_FAILED As Long = 1761
Private Const ERR_EMPTY_TABLE_ALIAS As Long = 1335
Private Const ERR_DUPLICATE_TABLE_ALIAS As Long = 1340
Private Const ERR_TABLE_ALIAS_NOT_FOUND As Long = 1341
Private Const ERR_NO_SOURCES As Long = 1350

Public Sub m_ExpandVirtualSourcesAndOutput( _
    ByVal cfg As Object, _
    Optional ByVal errSource As String = "ex_ConfigVirtualSources" _
)
    Dim templateAliases As Object
    Dim generatedByTemplate As Object
    Dim sourceAlias As Variant
    Dim templateAlias As String
    Dim fileKey As String
    Dim resolverKey As String
    Dim resolverArgsKey As String
    Dim sheetAliasesKey As String
    Dim hasFilePath As Boolean
    Dim hasResolver As Boolean
    Dim filePattern As String
    Dim resolverName As String
    Dim resolverArgs As String
    Dim sheetAliasesRaw As String
    Dim aliasPlaceholders As Object
    Dim pathPlaceholders As Object
    Dim candidatePaths As Variant
    Dim concreteAliases As Collection
    Dim concreteSeen As Object
    Dim i As Long
    Dim candidatePath As String
    Dim candidateAbsPath As String
    Dim placeholderValues As Object
    Dim concreteAlias As String

    If cfg Is Nothing Then Exit Sub

    If cfg.Exists(RUNTIME_EXPANDED_FLAG_KEY) Then
        If StrComp(CStr(cfg(RUNTIME_EXPANDED_FLAG_KEY)), "true", vbTextCompare) = 0 Then Exit Sub
    End If

    Set templateAliases = mp_CollectTemplateSourceAliases(cfg)
    Set generatedByTemplate = CreateObject("Scripting.Dictionary")
    generatedByTemplate.CompareMode = 1

    If templateAliases.Count = 0 Then
        cfg(RUNTIME_EXPANDED_FLAG_KEY) = "true"
        Exit Sub
    End If

    For Each sourceAlias In templateAliases.Keys
        templateAlias = CStr(sourceAlias)
        fileKey = "Source." & templateAlias & ".FilePath"
        resolverKey = "Source." & templateAlias & ".FileResolver"
        resolverArgsKey = "Source." & templateAlias & ".FileResolverArgs"
        sheetAliasesKey = "Source." & templateAlias & ".SheetAliases"

        hasFilePath = cfg.Exists(fileKey)
        hasResolver = cfg.Exists(resolverKey)
        If hasFilePath Xor hasResolver Then
            Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                "Template source '" & templateAlias & "' must define both '" & fileKey & "' and '" & resolverKey & "'."
        End If
        If Not hasFilePath Then GoTo ContinueTemplate

        filePattern = Trim$(CStr(cfg(fileKey)))
        resolverName = Trim$(CStr(cfg(resolverKey)))
        resolverArgs = mp_GetCfgOptional(cfg, resolverArgsKey, vbNullString)
        sheetAliasesRaw = mp_GetCfgRequired(cfg, sheetAliasesKey, errSource)

        If Len(filePattern) = 0 Then
            Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, "Config key '" & fileKey & "' is empty."
        End If
        If Len(resolverName) = 0 Then
            Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, "Config key '" & resolverKey & "' is empty."
        End If

        Set aliasPlaceholders = mp_CollectNamedPlaceholders(templateAlias)
        Set pathPlaceholders = mp_CollectNamedPlaceholders(filePattern)
        If Not mp_AreSamePlaceholderSets(aliasPlaceholders, pathPlaceholders) Then
            Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                "Placeholder set mismatch between source alias template '" & templateAlias & "' and file pattern '" & filePattern & "'."
        End If

        candidatePaths = mp_RunSourceResolverToPathArray(filePattern, resolverName, resolverArgs, errSource)
        If mp_IsEmptyVariantArray(candidatePaths) Then
            Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                "Source resolver '" & resolverName & "' returned no paths for template source '" & templateAlias & "'."
        End If

        Set concreteAliases = New Collection
        Set concreteSeen = CreateObject("Scripting.Dictionary")
        concreteSeen.CompareMode = 1

        For i = LBound(candidatePaths) To UBound(candidatePaths)
            candidatePath = Trim$(CStr(candidatePaths(i)))
            If Len(candidatePath) = 0 Then GoTo ContinueCandidate

            candidateAbsPath = mp_ResolvePathLocal(candidatePath)
            If Len(candidateAbsPath) = 0 Then GoTo ContinueCandidate
            If Dir(candidateAbsPath) = vbNullString Then
                Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                    "Resolver returned path that does not exist for source '" & templateAlias & "': " & candidateAbsPath
            End If

            If Not mp_TryExtractPlaceholderValuesByPattern(filePattern, candidateAbsPath, placeholderValues) Then
                Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                    "Resolver returned path that does not satisfy template pattern for source '" & templateAlias & "': " & candidateAbsPath
            End If

            concreteAlias = mp_ReplaceNamedPlaceholders(templateAlias, placeholderValues)
            If mp_IsSourceAliasTemplate(concreteAlias) Then
                Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                    "Failed to resolve source alias template '" & templateAlias & "' for path '" & candidateAbsPath & "'."
            End If

            If concreteSeen.Exists(concreteAlias) Then GoTo ContinueCandidate
            concreteSeen(concreteAlias) = candidateAbsPath

            cfg("Source." & concreteAlias & ".FilePath") = candidateAbsPath
            cfg("Source." & concreteAlias & ".SheetAliases") = sheetAliasesRaw
            mp_CloneTemplateSheetConfig cfg, templateAlias, concreteAlias
            concreteAliases.Add concreteAlias
ContinueCandidate:
        Next i

        If concreteAliases.Count = 0 Then
            Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                "No concrete sources were generated for template source '" & templateAlias & "'."
        End If

        generatedByTemplate.Add templateAlias, concreteAliases
ContinueTemplate:
    Next sourceAlias

    mp_ExpandOutputSheetsTemplates cfg, generatedByTemplate, errSource
    cfg(RUNTIME_EXPANDED_FLAG_KEY) = "true"
End Sub

Public Function m_BuildOutputEntries( _
    ByVal cfg As Object, _
    Optional ByVal errSource As String = "ex_ConfigVirtualSources" _
) As Collection
    Dim result As Collection
    Dim outputAliases As Variant
    Dim i As Long
    Dim tokenText As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim entry As Object

    m_ExpandVirtualSourcesAndOutput cfg, errSource

    Set result = New Collection
    outputAliases = mp_GetListRequired(cfg, "Output.Sheets", errSource)

    For i = LBound(outputAliases) To UBound(outputAliases)
        tokenText = Trim$(CStr(outputAliases(i)))
        If Len(tokenText) = 0 Then GoTo ContinueToken

        If Not mp_TryResolveOutputSheetEntry(cfg, tokenText, sourceAlias, tableAlias, errSource) Then
            Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                "Invalid Output.Sheets entry '" & tokenText & "'. Expected table alias or '<source>.Sheet[TableAlias]'."
        End If

        Set entry = CreateObject("Scripting.Dictionary")
        entry.CompareMode = 1
        entry("RawToken") = tokenText
        entry("SourceAlias") = sourceAlias
        entry("TableAlias") = tableAlias
        result.Add entry
ContinueToken:
    Next i

    Set m_BuildOutputEntries = result
End Function

Public Function m_TryParseQualifiedOutputSheetRef( _
    ByVal tokenText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String _
) As Boolean
    m_TryParseQualifiedOutputSheetRef = mp_TryParseQualifiedOutputSheetRef(tokenText, outSourceAlias, outTableAlias)
End Function

Public Function m_TryParseSourceKey( _
    ByVal keyText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTail As String _
) As Boolean
    m_TryParseSourceKey = mp_TryParseSourceKey(keyText, outSourceAlias, outTail)
End Function

Public Function m_IsSourceAliasTemplate(ByVal sourceAlias As String) As Boolean
    m_IsSourceAliasTemplate = mp_IsSourceAliasTemplate(sourceAlias)
End Function

Public Function m_FindSourceAliasForTable( _
    ByVal cfg As Object, _
    ByVal tableAlias As String, _
    Optional ByVal errSource As String = "ex_ConfigVirtualSources" _
) As String
    m_FindSourceAliasForTable = mp_FindSourceAliasForTable(cfg, tableAlias, errSource)
End Function

Public Function m_GetSourceAliases( _
    ByVal cfg As Object, _
    Optional ByVal errSource As String = "ex_ConfigVirtualSources" _
) As Variant
    m_GetSourceAliases = mp_GetSourceAliases(cfg, errSource)
End Function

Private Function mp_TryResolveOutputSheetEntry( _
    ByVal cfg As Object, _
    ByVal tokenText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByVal errSource As String _
) As Boolean
    outSourceAlias = vbNullString
    outTableAlias = vbNullString

    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then Exit Function

    If mp_TryParseQualifiedOutputSheetRef(tokenText, outSourceAlias, outTableAlias) Then
        mp_TryResolveOutputSheetEntry = True
        Exit Function
    End If

    outTableAlias = tokenText
    outSourceAlias = mp_FindSourceAliasForTable(cfg, outTableAlias, errSource)
    mp_TryResolveOutputSheetEntry = True
End Function

Private Function mp_TryParseQualifiedOutputSheetRef( _
    ByVal tokenText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String _
) As Boolean
    Dim workText As String
    Dim sheetPos As Long
    Dim closePos As Long
    Dim sourceAlias As String
    Dim tableAlias As String

    outSourceAlias = vbNullString
    outTableAlias = vbNullString

    workText = Trim$(tokenText)
    If Len(workText) = 0 Then Exit Function

    If Left$(workText, 1) = "{" Then
        If Not mp_TryFindMatchingBrace(workText, 1, closePos) Then Exit Function
        sourceAlias = Mid$(workText, 1, closePos)
        If LCase$(Mid$(workText, closePos + 1, Len(".Sheet["))) <> LCase$(".Sheet[") Then Exit Function
        workText = Mid$(workText, closePos + 1)
        sheetPos = 1
    Else
        sheetPos = InStr(1, workText, ".Sheet[", vbTextCompare)
        If sheetPos <= 1 Then Exit Function
        sourceAlias = Trim$(Left$(workText, sheetPos - 1))
        workText = Mid$(workText, sheetPos)
        sheetPos = 1
    End If

    If LCase$(Left$(workText, Len(".Sheet["))) <> LCase$(".Sheet[") Then Exit Function
    closePos = InStr(Len(".Sheet[") + 1, workText, "]", vbBinaryCompare)
    If closePos <= Len(".Sheet[") + 1 Then Exit Function
    If closePos <> Len(workText) Then Exit Function

    tableAlias = Trim$(Mid$(workText, Len(".Sheet[") + 1, closePos - Len(".Sheet[") - 1))
    If Len(sourceAlias) = 0 Or Len(tableAlias) = 0 Then Exit Function

    outSourceAlias = sourceAlias
    outTableAlias = tableAlias
    mp_TryParseQualifiedOutputSheetRef = True
End Function

Private Function mp_TryParseSourceKey( _
    ByVal keyText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTail As String _
) As Boolean
    Dim workText As String
    Dim closePos As Long
    Dim dotPos As Long

    outSourceAlias = vbNullString
    outTail = vbNullString

    workText = Trim$(keyText)
    If Len(workText) = 0 Then Exit Function
    If LCase$(Left$(workText, Len("Source."))) <> LCase$("Source.") Then Exit Function

    workText = Mid$(workText, Len("Source.") + 1)
    If Len(workText) = 0 Then Exit Function

    If Left$(workText, 1) = "{" Then
        If Not mp_TryFindMatchingBrace(workText, 1, closePos) Then Exit Function
        outSourceAlias = Mid$(workText, 1, closePos)
        If closePos = Len(workText) Then
            outTail = vbNullString
            mp_TryParseSourceKey = True
            Exit Function
        End If
        If Mid$(workText, closePos + 1, 1) <> "." Then Exit Function
        outTail = Mid$(workText, closePos + 2)
    Else
        dotPos = InStr(1, workText, ".", vbBinaryCompare)
        If dotPos <= 1 Then Exit Function
        outSourceAlias = Left$(workText, dotPos - 1)
        outTail = Mid$(workText, dotPos + 1)
    End If

    outSourceAlias = Trim$(outSourceAlias)
    outTail = Trim$(outTail)
    If Len(outSourceAlias) = 0 Then Exit Function

    mp_TryParseSourceKey = True
End Function

Private Function mp_TryFindMatchingBrace( _
    ByVal textValue As String, _
    ByVal openPos As Long, _
    ByRef outClosePos As Long _
) As Boolean
    Dim i As Long
    Dim depth As Long
    Dim ch As String

    outClosePos = 0
    If openPos < 1 Or openPos > Len(textValue) Then Exit Function
    If Mid$(textValue, openPos, 1) <> "{" Then Exit Function

    depth = 0
    For i = openPos To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch = "{" Then
            depth = depth + 1
        ElseIf ch = "}" Then
            depth = depth - 1
            If depth = 0 Then
                outClosePos = i
                mp_TryFindMatchingBrace = True
                Exit Function
            End If
            If depth < 0 Then Exit Function
        End If
    Next i
End Function

Private Function mp_CollectTemplateSourceAliases(ByVal cfg As Object) As Object
    Dim result As Object
    Dim key As Variant
    Dim sourceAlias As String
    Dim tail As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If cfg Is Nothing Then
        Set mp_CollectTemplateSourceAliases = result
        Exit Function
    End If

    For Each key In cfg.Keys
        If mp_TryParseSourceKey(CStr(key), sourceAlias, tail) Then
            If mp_IsSourceAliasTemplate(sourceAlias) Then
                result(sourceAlias) = True
            End If
        End If
    Next key

    Set mp_CollectTemplateSourceAliases = result
End Function

Private Function mp_IsSourceAliasTemplate(ByVal sourceAlias As String) As Boolean
    Dim placeholders As Object

    Set placeholders = mp_CollectNamedPlaceholders(sourceAlias)
    mp_IsSourceAliasTemplate = (Not placeholders Is Nothing And placeholders.Count > 0)
End Function

Private Function mp_CollectNamedPlaceholders(ByVal textValue As String) As Object
    Dim result As Object
    Dim rgx As Object
    Dim matches As Object
    Dim m As Object
    Dim placeholderName As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    Set rgx = CreateObject("VBScript.RegExp")
    rgx.Global = True
    rgx.IgnoreCase = True
    rgx.Pattern = "\{([A-Za-z_][A-Za-z0-9_]*)\}"
    Set matches = rgx.Execute(CStr(textValue))

    For Each m In matches
        placeholderName = LCase$(Trim$(CStr(m.SubMatches(0))))
        If Len(placeholderName) > 0 Then
            result(placeholderName) = True
        End If
    Next m

    Set mp_CollectNamedPlaceholders = result
End Function

Private Function mp_AreSamePlaceholderSets(ByVal leftSet As Object, ByVal rightSet As Object) As Boolean
    Dim key As Variant

    If leftSet Is Nothing Or rightSet Is Nothing Then Exit Function
    If leftSet.Count <> rightSet.Count Then Exit Function

    For Each key In leftSet.Keys
        If Not rightSet.Exists(CStr(key)) Then Exit Function
    Next key

    mp_AreSamePlaceholderSets = True
End Function

Private Function mp_RunSourceResolverToPathArray( _
    ByVal filePathPattern As String, _
    ByVal resolverName As String, _
    ByVal resolverArgs As String, _
    ByVal errSource As String _
) As Variant
    Dim resolverCallName As String
    Dim resolvedValue As Variant
    Dim resolvedObject As Object
    Dim hasObjectResult As Boolean
    Dim values As Collection
    Dim item As Variant
    Dim parsed As Variant
    Dim i As Long

    If InStr(1, resolverName, "!", vbBinaryCompare) > 0 Then
        resolverCallName = resolverName
    Else
        resolverCallName = "'" & ThisWorkbook.Name & "'!" & resolverName
    End If

    hasObjectResult = False
    On Error Resume Next
    Set resolvedObject = Application.Run(resolverCallName, filePathPattern, resolverArgs)
    If Err.Number = 0 Then
        hasObjectResult = True
    ElseIf Err.Number = 13 Then
        Err.Clear
    Else
        GoTo ResolverEH
    End If
    On Error GoTo 0

    If Not hasObjectResult Then
        On Error GoTo ResolverEH
        resolvedValue = Application.Run(resolverCallName, filePathPattern, resolverArgs)
        On Error GoTo 0
    End If

    Set values = New Collection
    If hasObjectResult Then
        If TypeName(resolvedObject) = "Collection" Then
            For Each item In resolvedObject
                If Len(Trim$(CStr(item))) > 0 Then values.Add Trim$(CStr(item))
            Next item
        Else
            Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                "Unsupported resolver return type '" & TypeName(resolvedObject) & "'. Expected String or Collection."
        End If
    Else
        parsed = mp_ParseResolverPathList(CStr(resolvedValue))
        If Not mp_IsEmptyVariantArray(parsed) Then
            For i = LBound(parsed) To UBound(parsed)
                If Len(Trim$(CStr(parsed(i)))) > 0 Then values.Add Trim$(CStr(parsed(i)))
            Next i
        End If
    End If

    mp_RunSourceResolverToPathArray = mp_StringCollectionToArray(values)
    Exit Function

ResolverEH:
    On Error GoTo 0
    Err.Raise vbObjectError + ERR_RESOLVER_FAILED, errSource, _
        "Source file resolver failed (resolver='" & resolverName & "'): " & Err.Description
End Function

Private Function mp_ParseResolverPathList(ByVal rawText As String) As Variant
    Dim normalized As String
    Dim values As Collection
    Dim items As Variant
    Dim i As Long
    Dim token As String

    normalized = CStr(rawText)
    normalized = Replace$(normalized, vbCrLf, vbLf)
    normalized = Replace$(normalized, vbCr, vbLf)

    Set values = New Collection
    If InStr(1, normalized, vbLf, vbBinaryCompare) > 0 Then
        items = Split(normalized, vbLf)
        For i = LBound(items) To UBound(items)
            token = Trim$(CStr(items(i)))
            If Len(token) > 0 Then values.Add token
        Next i
    ElseIf InStr(1, normalized, ";", vbBinaryCompare) > 0 Then
        items = Split(normalized, ";")
        For i = LBound(items) To UBound(items)
            token = Trim$(CStr(items(i)))
            If Len(token) > 0 Then values.Add token
        Next i
    Else
        token = Trim$(normalized)
        If Len(token) > 0 Then values.Add token
    End If

    mp_ParseResolverPathList = mp_StringCollectionToArray(values)
End Function

Private Function mp_StringCollectionToArray(ByVal values As Collection) As Variant
    Dim arr() As String
    Dim i As Long

    If values Is Nothing Then
        mp_StringCollectionToArray = Array()
        Exit Function
    End If
    If values.Count = 0 Then
        mp_StringCollectionToArray = Array()
        Exit Function
    End If

    ReDim arr(0 To values.Count - 1)
    For i = 1 To values.Count
        arr(i - 1) = CStr(values(i))
    Next i
    mp_StringCollectionToArray = arr
End Function

Private Function mp_TryExtractPlaceholderValuesByPattern( _
    ByVal filePattern As String, _
    ByVal candidatePath As String, _
    ByRef outValues As Object _
) As Boolean
    Dim normalizedPattern As String
    Dim normalizedCandidate As String
    Dim regexPattern As String
    Dim tokenOrder As Collection
    Dim rgx As Object
    Dim matches As Object
    Dim m As Object
    Dim i As Long
    Dim tokenName As String
    Dim tokenValue As String

    normalizedPattern = mp_ResolvePathLocal(filePattern)
    normalizedCandidate = mp_ResolvePathLocal(candidatePath)
    normalizedPattern = Replace$(normalizedPattern, "/", "\")
    normalizedCandidate = Replace$(normalizedCandidate, "/", "\")

    If Not mp_BuildPlaceholderPatternRegex(normalizedPattern, regexPattern, tokenOrder) Then Exit Function

    Set rgx = CreateObject("VBScript.RegExp")
    rgx.Global = False
    rgx.IgnoreCase = True
    rgx.Pattern = regexPattern
    Set matches = rgx.Execute(normalizedCandidate)
    If matches.Count = 0 Then Exit Function

    Set outValues = CreateObject("Scripting.Dictionary")
    outValues.CompareMode = 1
    Set m = matches(0)
    For i = 1 To tokenOrder.Count
        tokenName = LCase$(CStr(tokenOrder(i)))
        tokenValue = CStr(m.SubMatches(i - 1))
        If outValues.Exists(tokenName) Then
            If StrComp(CStr(outValues(tokenName)), tokenValue, vbBinaryCompare) <> 0 Then Exit Function
        Else
            outValues(tokenName) = tokenValue
        End If
    Next i

    mp_TryExtractPlaceholderValuesByPattern = True
End Function

Private Function mp_BuildPlaceholderPatternRegex( _
    ByVal patternText As String, _
    ByRef outRegex As String, _
    ByRef outTokenOrder As Collection _
) As Boolean
    Dim i As Long
    Dim closePos As Long
    Dim ch As String
    Dim tokenName As String
    Dim resultRegex As String

    Set outTokenOrder = New Collection
    resultRegex = "^"

    i = 1
    Do While i <= Len(patternText)
        ch = Mid$(patternText, i, 1)
        If ch = "{" Then
            closePos = InStr(i + 1, patternText, "}", vbBinaryCompare)
            If closePos <= i + 1 Then Exit Function
            tokenName = Trim$(Mid$(patternText, i + 1, closePos - i - 1))
            If Not ex_ScriptParserCore.m_IsIdentifier(tokenName) Then Exit Function
            outTokenOrder.Add tokenName
            Select Case LCase$(tokenName)
                Case "dd", "mm"
                    resultRegex = resultRegex & "(\d{2})"
                Case "yyyy"
                    resultRegex = resultRegex & "(\d{4})"
                Case Else
                    resultRegex = resultRegex & "(.+?)"
            End Select
            i = closePos + 1
        Else
            resultRegex = resultRegex & mp_EscapeRegexLiteral(ch)
            i = i + 1
        End If
    Loop

    resultRegex = resultRegex & "$"
    outRegex = resultRegex
    mp_BuildPlaceholderPatternRegex = True
End Function

Private Function mp_EscapeRegexLiteral(ByVal rawText As String) As String
    Dim i As Long
    Dim ch As String
    Dim resultText As String

    For i = 1 To Len(rawText)
        ch = Mid$(rawText, i, 1)
        Select Case ch
            Case "\", ".", "+", "*", "?", "^", "$", "(", ")", "[", "]", "{", "}", "|"
                resultText = resultText & "\" & ch
            Case Else
                resultText = resultText & ch
        End Select
    Next i

    mp_EscapeRegexLiteral = resultText
End Function

Private Function mp_ReplaceNamedPlaceholders(ByVal templateText As String, ByVal valuesByName As Object) As String
    Dim resultText As String
    Dim key As Variant

    resultText = CStr(templateText)
    If valuesByName Is Nothing Then
        mp_ReplaceNamedPlaceholders = resultText
        Exit Function
    End If

    For Each key In valuesByName.Keys
        resultText = Replace$(resultText, "{" & CStr(key) & "}", CStr(valuesByName(key)))
    Next key

    mp_ReplaceNamedPlaceholders = resultText
End Function

Private Sub mp_CloneTemplateSheetConfig( _
    ByVal cfg As Object, _
    ByVal templateAlias As String, _
    ByVal concreteAlias As String _
)
    Dim keys As Variant
    Dim i As Long
    Dim keyText As String
    Dim prefix As String
    Dim newKey As String

    If cfg Is Nothing Then Exit Sub

    prefix = LCase$(templateAlias & ".Sheet[")
    keys = cfg.Keys

    For i = LBound(keys) To UBound(keys)
        keyText = CStr(keys(i))
        If LCase$(Left$(keyText, Len(prefix))) = prefix Then
            newKey = concreteAlias & Mid$(keyText, Len(templateAlias) + 1)
            cfg(newKey) = cfg(keyText)
        End If
    Next i
End Sub

Private Sub mp_ExpandOutputSheetsTemplates( _
    ByVal cfg As Object, _
    ByVal generatedByTemplate As Object, _
    ByVal errSource As String _
)
    Dim outputAliases As Variant
    Dim expanded As Collection
    Dim i As Long
    Dim tokenText As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim concreteAliases As Collection
    Dim concreteAlias As Variant

    outputAliases = mp_GetListRequired(cfg, "Output.Sheets", errSource)
    Set expanded = New Collection

    For i = LBound(outputAliases) To UBound(outputAliases)
        tokenText = Trim$(CStr(outputAliases(i)))
        If Len(tokenText) = 0 Then GoTo ContinueToken

        If mp_TryParseQualifiedOutputSheetRef(tokenText, sourceAlias, tableAlias) Then
            If mp_IsSourceAliasTemplate(sourceAlias) Then
                If generatedByTemplate Is Nothing Or Not generatedByTemplate.Exists(sourceAlias) Then
                    Err.Raise vbObjectError + ERR_INVALID_CONFIG, errSource, _
                        "Output.Sheets entry '" & tokenText & "' references unresolved template source '" & sourceAlias & "'."
                End If
                Set concreteAliases = generatedByTemplate(sourceAlias)
                For Each concreteAlias In concreteAliases
                    expanded.Add CStr(concreteAlias) & ".Sheet[" & tableAlias & "]"
                Next concreteAlias
            Else
                expanded.Add sourceAlias & ".Sheet[" & tableAlias & "]"
            End If
        Else
            expanded.Add tokenText
        End If
ContinueToken:
    Next i

    cfg("Output.Sheets") = mp_JoinCollectionItems(expanded, "; ")
End Sub

Private Function mp_JoinCollectionItems(ByVal values As Collection, ByVal delimiter As String) As String
    Dim i As Long

    If values Is Nothing Then Exit Function

    For i = 1 To values.Count
        If i > 1 Then mp_JoinCollectionItems = mp_JoinCollectionItems & delimiter
        mp_JoinCollectionItems = mp_JoinCollectionItems & CStr(values(i))
    Next i
End Function

Private Function mp_FindSourceAliasForTable( _
    ByVal cfg As Object, _
    ByVal tableAlias As String, _
    ByVal errSource As String _
) As String
    Dim sourceAliases As Variant
    Dim aliases As Variant
    Dim i As Long
    Dim src As String
    Dim found As String

    tableAlias = Trim$(tableAlias)
    If Len(tableAlias) = 0 Then
        Err.Raise vbObjectError + ERR_EMPTY_TABLE_ALIAS, errSource, "Output.Sheets contains an empty table alias."
    End If

    sourceAliases = mp_GetSourceAliases(cfg, errSource)
    For i = LBound(sourceAliases) To UBound(sourceAliases)
        src = CStr(sourceAliases(i))
        If mp_IsSourceAliasTemplate(src) Then GoTo ContinueSource
        aliases = mp_GetListRequired(cfg, "Source." & src & ".SheetAliases", errSource)
        If mp_ArrayContainsText(aliases, tableAlias) Then
            If Len(found) > 0 Then
                Err.Raise vbObjectError + ERR_DUPLICATE_TABLE_ALIAS, errSource, _
                    "Sheet alias '" & tableAlias & "' is declared in multiple sources: '" & found & "' and '" & src & "'."
            End If
            found = src
        End If
ContinueSource:
    Next i

    If Len(found) = 0 Then
        Err.Raise vbObjectError + ERR_TABLE_ALIAS_NOT_FOUND, errSource, _
            "Sheet alias '" & tableAlias & "' is not declared in any Source.*.SheetAliases."
    End If

    mp_FindSourceAliasForTable = found
End Function

Private Function mp_GetSourceAliases( _
    ByVal cfg As Object, _
    ByVal errSource As String _
) As Variant
    Dim d As Object
    Dim key As Variant
    Dim srcAlias As String
    Dim tail As String
    Dim arr() As String
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    For Each key In cfg.Keys
        If mp_TryParseSourceKey(CStr(key), srcAlias, tail) Then
            If Len(srcAlias) > 0 Then
                d(srcAlias) = srcAlias
            End If
        End If
    Next key

    If d.Count = 0 Then
        Err.Raise vbObjectError + ERR_NO_SOURCES, errSource, "No Source.* keys found in config."
    End If

    ReDim arr(0 To d.Count - 1)
    i = 0
    For Each key In d.Keys
        arr(i) = CStr(key)
        i = i + 1
    Next key

    mp_GetSourceAliases = arr
End Function

Private Function mp_GetListRequired( _
    ByVal cfg As Object, _
    ByVal keyName As String, _
    ByVal errSource As String _
) As Variant
    Dim values As Variant

    values = mp_SplitList(mp_GetCfgRequired(cfg, keyName, errSource))
    If mp_IsEmptyVariantArray(values) Then
        Err.Raise vbObjectError + 1380, errSource, "List is empty for config key: " & keyName
    End If

    mp_GetListRequired = values
End Function

Private Function mp_SplitList(ByVal rawText As String) As Variant
    Dim s As String

    s = Replace$(rawText, vbCr, ";")
    s = Replace$(s, vbLf, ";")
    s = Replace$(s, "|", ";")
    s = Replace$(s, ",", ";")

    mp_SplitList = Split(s, ";")
End Function

Private Function mp_ArrayContainsText(ByVal arr As Variant, ByVal wanted As String) As Boolean
    Dim i As Long

    For i = LBound(arr) To UBound(arr)
        If StrComp(Trim$(CStr(arr(i))), Trim$(wanted), vbTextCompare) = 0 Then
            mp_ArrayContainsText = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_IsEmptyVariantArray(ByVal arr As Variant) As Boolean
    If Not IsArray(arr) Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    On Error GoTo EH
    If UBound(arr) < LBound(arr) Then
        mp_IsEmptyVariantArray = True
    End If
    Exit Function
EH:
    mp_IsEmptyVariantArray = True
End Function

Private Function mp_GetCfgRequired( _
    ByVal cfg As Object, _
    ByVal keyName As String, _
    ByVal errSource As String _
) As String
    Dim valueText As String

    If cfg Is Nothing Or Not cfg.Exists(keyName) Then
        Err.Raise vbObjectError + 1370, errSource, "Missing config key: " & keyName
    End If

    valueText = Trim$(CStr(cfg(keyName)))
    If Len(valueText) = 0 Then
        Err.Raise vbObjectError + 1371, errSource, "Empty config value: " & keyName
    End If

    mp_GetCfgRequired = valueText
End Function

Private Function mp_GetCfgOptional( _
    ByVal cfg As Object, _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    If cfg Is Nothing Then
        mp_GetCfgOptional = defaultValue
        Exit Function
    End If

    If Not cfg.Exists(keyName) Then
        mp_GetCfgOptional = defaultValue
        Exit Function
    End If

    mp_GetCfgOptional = Trim$(CStr(cfg(keyName)))
End Function

Private Function mp_ResolvePathLocal(ByVal inputPath As String) As String
    Dim basePath As String
    Dim resolvedPath As String

    inputPath = Trim$(inputPath)
    If Len(inputPath) = 0 Then Exit Function

    If Left$(inputPath, 2) = "\\" Or InStr(1, inputPath, ":\", vbTextCompare) > 0 Then
        mp_ResolvePathLocal = mp_CanonicalizePath(inputPath)
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        mp_ResolvePathLocal = mp_CanonicalizePath(inputPath)
        Exit Function
    End If

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    resolvedPath = basePath & inputPath
    mp_ResolvePathLocal = mp_CanonicalizePath(resolvedPath)
End Function

Private Function mp_CanonicalizePath(ByVal rawPath As String) As String
    Dim fso As Object

    rawPath = Trim$(rawPath)
    If Len(rawPath) = 0 Then Exit Function

    On Error GoTo Fallback
    Set fso = CreateObject("Scripting.FileSystemObject")
    mp_CanonicalizePath = CStr(fso.GetAbsolutePathName(rawPath))
    If Len(Trim$(mp_CanonicalizePath)) = 0 Then mp_CanonicalizePath = rawPath
    Exit Function

Fallback:
    On Error GoTo 0
    mp_CanonicalizePath = rawPath
End Function
