Attribute VB_Name = "ex_ResultTemplatesParser"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"

' Reserved keywords/tokens supported by this parser.
' Reserved numeric offset token:
' - <NUMERIC>{+N}
' - <NUMERIC>{-N}
' Reserved line-join token:
' - {#^}  (ignores a line break around token)
' - #^    (same behavior, shorthand)
' Legacy line-join token (backward compatibility):
' - {#_}
' Reserved trim-indentation token:
' - #_    (removes token and horizontal whitespace after it)
' Reserved forced line-break token:
' - #newline
' - {#newline}
' Reserved computed-value token:
' - #let varName = $ModuleName.MethodName(arg1, arg2);
' Reserved if-condition unary operator:
' - #not
' Reserved if-condition numeric comparison operators:
' - ==, !=, >, <, >=, <=
' Reserved loop block:
' - {#for item in Collection}
' - {#endfor}
' Reserved shared-template include token:
' - {#include SharedTemplateId}
Private Const NUMERIC_OFFSET_TOKEN_PATTERN As String = "\{([+-]\d+)\}"
Private Const LEGACY_DAY_TOKEN_PATTERN As String = "\{#dd(?:[+-]\d+)?\}"
Private Const RESERVED_JOINLINE_TOKEN As String = "{#^}"
Private Const RESERVED_JOINLINE_TOKEN_SHORT As String = "#^"
Private Const RESERVED_JOINLINE_TOKEN_LEGACY As String = "{#_}"
Private Const RESERVED_TRIMINDENT_TOKEN_SHORT As String = "#_"
Private Const RESERVED_NEWLINE_TOKEN As String = "{#newline}"
Private Const RESERVED_NEWLINE_TOKEN_SHORT As String = "#newline"
Private Const IF_BLOCK_OPEN As String = "{#if"
Private Const IF_BLOCK_CLOSE As String = "{#endif}"
Private Const FOR_BLOCK_OPEN As String = "{#for"
Private Const FOR_BLOCK_CLOSE As String = "{#endfor}"
Private Const INCLUDE_BLOCK_OPEN As String = "{#include"
Private Const BOOLEAN_TRUE As String = "true"
Private Const BOOLEAN_FALSE As String = "false"

Private Const FORMATTER_UPPER As String = "upper"
Private Const FORMATTER_LOWER As String = "lower"
Private Const FORMATTER_CAPITALIZE As String = "capitalize"
Private Const FORMATTER_FIRSTCHAR As String = "firstchar"
Private Const FORMATTER_UPPERFIRSTWORD As String = "upperfirstword"
Private Const FORMATTER_UPPERFIRSTLETTER As String = "upperfirstletter"
Private Const FORMATTER_LOWERFIRSTWORD As String = "lowerfirstword"
Private Const FORMATTER_LOWERFIRSTLETTER As String = "lowerfirstletter"
Private Const FORMATTER_GENITIVE As String = "genitive"
Private Const FORMATTER_ACCUSATIVE As String = "accusative"
Private Const FORMATTER_DATIVE As String = "dative"
Private Const FORMATTER_TRUNCATE As String = "truncate"
Private Const FORMATTER_REPLACE As String = "replace"
Private Const FORMATTER_REGEX_REPLACE As String = "regexreplace"
Private Const FORMATTER_DATEFORMAT As String = "dateformat"
Private Const FORMATTER_TO_DATE_DAY As String = "todate_day"
Private Const FORMATTER_TO_DATE_DAY_WITH_MONTH As String = "todate_daywithmonth"
Private Const FORMATTER_CALENDAR_DAYS_UA As String = "calendardaysua"
Private Const FORMATTER_SURNAME_INITIALS As String = "surnameinitials"
Private Const FORMATTER_FIO_SURNAME As String = "fiosurname"
Private Const FORMATTER_FIO_INITIALS As String = "fioinitials"

Private Const CASE_GENITIVE As String = "genitive"
Private Const CASE_ACCUSATIVE As String = "accusative"
Private Const CASE_DATIVE As String = "dative"

Private Const NBSP_CODE_POINT As Long = 160
Private Const NARROW_NBSP_CODE_POINT As Long = 8239
Private Const TEMPLATE_ERROR_PREFIX As String = "[TEMPLATE ERROR]"
Private Const HIGHLIGHT_MARKER_START As String = "/Start"
Private Const HIGHLIGHT_MARKER_END As String = "/End"
Private Const DEFAULT_HIGHLIGHT_COLOR_HEX As String = "#66CCFF"

Private g_TemplateCollections As Object

Public Function m_GetTemplateText( _
    ByVal templateId As String, _
    ByVal resultTemplatesRelPath As String _
) As String
    Dim doc As Object
    Dim node As Object
    Dim xpath As String
    Dim templateText As String
    Dim templatesPath As String

    templateId = mp_TrimWhitespace(templateId)
    If Len(templateId) = 0 Then
        Err.Raise vbObjectError + 1760, "ex_ResultTemplatesParser", "Template id is empty."
    End If
    templatesPath = mp_TrimWhitespace(resultTemplatesRelPath)
    If Len(templatesPath) = 0 Then
        Err.Raise vbObjectError + 1819, "ex_ResultTemplatesParser", "Result templates path is empty."
    End If

    On Error GoTo EH

    Set doc = ex_XmlCore.m_LoadDomByRelativePath( _
        ThisWorkbook, _
        templatesPath, _
        PROFILES_NS, _
        "Missing result templates file: ", _
        "Failed to parse result templates file: " _
    )
    If doc Is Nothing Then
        Err.Raise vbObjectError + 1761, "ex_ResultTemplatesParser", "Unable to load result templates xml."
    End If

    xpath = "/p:resultTemplates/p:template[@id=" & ex_XmlCore.m_XPathLiteral(templateId) & "]/p:text"
    Set node = doc.selectSingleNode(xpath)
    If node Is Nothing Then
        Err.Raise vbObjectError + 1762, "ex_ResultTemplatesParser", "Template not found: '" & templateId & "'."
    End If

    templateText = CStr(node.Text)
    templateText = mp_ExpandSharedTemplateIncludes(templateText, doc, templateId)
    m_GetTemplateText = mp_NormalizeTemplateText(templateText)
    Exit Function

EH:
    m_GetTemplateText = mp_PrependTemplateError(vbNullString, "m_GetTemplateText('" & templateId & "', '" & templatesPath & "')")
End Function

Private Function mp_ExpandSharedTemplateIncludes( _
    ByVal sourceText As String, _
    ByVal doc As Object, _
    ByVal templateId As String _
) As String
    Dim includeChain As Collection

    Set includeChain = New Collection
    mp_ExpandSharedTemplateIncludes = mp_ExpandSharedTemplateIncludesRecursive( _
        CStr(sourceText), _
        doc, _
        CStr(templateId), _
        includeChain _
    )
End Function

Private Function mp_ExpandSharedTemplateIncludesRecursive( _
    ByVal sourceText As String, _
    ByVal doc As Object, _
    ByVal ownerName As String, _
    ByVal includeChain As Collection _
) As String
    Dim resultText As String
    Dim rx As Object
    Dim matches As Object
    Dim includeId As String
    Dim includeText As String
    Dim expandedIncludeText As String
    Dim matchStart As Long
    Dim matchLen As Long

    resultText = CStr(sourceText)
    If Len(resultText) = 0 Then
        mp_ExpandSharedTemplateIncludesRecursive = resultText
        Exit Function
    End If

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.MultiLine = True
    rx.Pattern = "\{#include\s+([A-Za-z_][A-Za-z0-9_.-]*)\s*\}"

    Do
        Set matches = rx.Execute(resultText)
        If matches Is Nothing Then Exit Do
        If matches.Count = 0 Then Exit Do

        includeId = mp_TrimWhitespace(CStr(matches(0).SubMatches(0)))
        If Len(includeId) = 0 Then
            Err.Raise vbObjectError + 1846, "ex_ResultTemplatesParser", _
                "Invalid #include syntax in '" & CStr(ownerName) & "'. Use '{#include SharedTemplateId}'."
        End If

        If mp_IncludeChainContains(includeChain, includeId) Then
            Err.Raise vbObjectError + 1847, "ex_ResultTemplatesParser", _
                "Circular shared template include detected: " & mp_BuildIncludeChainText(includeChain, includeId) & "."
        End If

        includeText = mp_GetSharedTemplateTextById(doc, includeId)
        includeChain.Add includeId
        expandedIncludeText = mp_ExpandSharedTemplateIncludesRecursive(includeText, doc, "sharedTemplate '" & includeId & "'", includeChain)
        includeChain.Remove includeChain.Count

        matchStart = CLng(matches(0).FirstIndex)
        matchLen = CLng(matches(0).Length)
        resultText = Left$(resultText, matchStart) & expandedIncludeText & Mid$(resultText, matchStart + matchLen + 1)
    Loop

    If InStr(1, resultText, INCLUDE_BLOCK_OPEN, vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 1848, "ex_ResultTemplatesParser", _
            "Invalid #include directive in '" & CStr(ownerName) & "'. Use '{#include SharedTemplateId}'."
    End If

    mp_ExpandSharedTemplateIncludesRecursive = resultText
End Function

Private Function mp_GetSharedTemplateTextById( _
    ByVal doc As Object, _
    ByVal sharedTemplateId As String _
) As String
    Dim node As Object
    Dim xpath As String

    sharedTemplateId = mp_TrimWhitespace(CStr(sharedTemplateId))
    If Len(sharedTemplateId) = 0 Then
        Err.Raise vbObjectError + 1849, "ex_ResultTemplatesParser", "Shared template id is empty."
    End If

    xpath = "/p:resultTemplates/p:sharedTemplates/p:sharedTemplate[@id=" & ex_XmlCore.m_XPathLiteral(sharedTemplateId) & "]/p:text"
    Set node = doc.selectSingleNode(xpath)
    If node Is Nothing Then
        Err.Raise vbObjectError + 1850, "ex_ResultTemplatesParser", "Shared template not found: '" & sharedTemplateId & "'."
    End If

    mp_GetSharedTemplateTextById = CStr(node.Text)
End Function

Private Function mp_IncludeChainContains( _
    ByVal includeChain As Collection, _
    ByVal includeId As String _
) As Boolean
    Dim i As Long

    If includeChain Is Nothing Then Exit Function

    For i = 1 To includeChain.Count
        If StrComp(CStr(includeChain(i)), CStr(includeId), vbBinaryCompare) = 0 Then
            mp_IncludeChainContains = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_BuildIncludeChainText( _
    ByVal includeChain As Collection, _
    ByVal closingIncludeId As String _
) As String
    Dim i As Long

    If Not includeChain Is Nothing Then
        For i = 1 To includeChain.Count
            If Len(mp_BuildIncludeChainText) > 0 Then mp_BuildIncludeChainText = mp_BuildIncludeChainText & " -> "
            mp_BuildIncludeChainText = mp_BuildIncludeChainText & "'" & CStr(includeChain(i)) & "'"
        Next i
    End If

    If Len(mp_BuildIncludeChainText) > 0 Then
        mp_BuildIncludeChainText = mp_BuildIncludeChainText & " -> "
    End If
    mp_BuildIncludeChainText = mp_BuildIncludeChainText & "'" & CStr(closingIncludeId) & "'"
End Function

Public Function m_ReplaceToken( _
    ByVal sourceText As String, _
    ByVal tokenText As String, _
    ByVal replacementText As String _
) As String
    On Error GoTo EH
    m_ReplaceToken = Replace(CStr(sourceText), CStr(tokenText), CStr(replacementText))
    Exit Function

EH:
    m_ReplaceToken = mp_PrependTemplateError(CStr(sourceText), "m_ReplaceToken('" & CStr(tokenText) & "')")
End Function

Public Function m_ReplacePlaceholder( _
    ByVal sourceText As String, _
    ByVal placeholderName As String, _
    ByVal replacementText As String, _
    Optional ByVal highlightColorHex As String = vbNullString _
) As String
    Dim normalizedName As String
    Dim resultText As String

    On Error GoTo EH

    normalizedName = mp_TrimWhitespace(placeholderName)
    If Len(normalizedName) = 0 Then
        m_ReplacePlaceholder = CStr(sourceText)
        Exit Function
    End If

    If mp_ShouldCollapsePlaceholderValue(replacementText) Then
        resultText = mp_CollapseNamedPlaceholderTokens(CStr(sourceText), normalizedName)
        resultText = mp_ReplaceIfConditionForPlaceholder(resultText, normalizedName, vbNullString)
        m_ReplacePlaceholder = resultText
        Exit Function
    End If

    resultText = CStr(sourceText)
    resultText = mp_ReplaceNamedPlaceholderTokens(resultText, normalizedName, CStr(replacementText), highlightColorHex)

    resultText = mp_ReplaceIfConditionForPlaceholder(resultText, normalizedName, CStr(replacementText))
    m_ReplacePlaceholder = resultText
    Exit Function

EH:
    If Len(resultText) = 0 Then resultText = CStr(sourceText)
    m_ReplacePlaceholder = mp_PrependTemplateError(resultText, "m_ReplacePlaceholder('" & CStr(placeholderName) & "')")
End Function

Public Function m_ReplaceMultiPlaceholders( _
    ByVal sourceText As String, _
    ByVal placeholderMapText As String, _
    Optional ByVal highlightColorHex As String = vbNullString _
) As String
    Dim resultText As String
    Dim normalized As String
    Dim lines As Variant
    Dim parts As Variant
    Dim itemSeparator As String
    Dim i As Long

    On Error GoTo EH

    resultText = CStr(sourceText)
    normalized = CStr(placeholderMapText)
    normalized = Replace(normalized, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    If Len(Trim$(normalized)) = 0 Then
        m_ReplaceMultiPlaceholders = resultText
        Exit Function
    End If

    If InStr(1, normalized, vbLf, vbBinaryCompare) > 0 Then
        lines = Split(normalized, vbLf)
        For i = LBound(lines) To UBound(lines)
            mp_ApplyPlaceholderMapEntry resultText, CStr(lines(i)), True, highlightColorHex
        Next i
        m_ReplaceMultiPlaceholders = resultText
        Exit Function
    End If

    itemSeparator = ";"
    If InStr(1, normalized, ";", vbBinaryCompare) = 0 Then
        If InStr(1, normalized, ",", vbBinaryCompare) > 0 Then itemSeparator = ","
    End If

    parts = Split(normalized, itemSeparator)
    For i = LBound(parts) To UBound(parts)
        mp_ApplyPlaceholderMapEntry resultText, CStr(parts(i)), False, highlightColorHex
    Next i

    m_ReplaceMultiPlaceholders = resultText
    Exit Function

EH:
    If Len(resultText) = 0 Then resultText = CStr(sourceText)
    m_ReplaceMultiPlaceholders = mp_PrependTemplateError(resultText, "m_ReplaceMultiPlaceholders")
End Function

' Backward-compatible alias. Prefer m_ReplaceMultiPlaceholders in new scripts.
Public Function m_ReplacePlaceholder2( _
    ByVal sourceText As String, _
    ByVal placeholderMapText As String, _
    Optional ByVal highlightColorHex As String = vbNullString _
) As String
    m_ReplacePlaceholder2 = m_ReplaceMultiPlaceholders(sourceText, placeholderMapText, highlightColorHex)
End Function

Public Sub m_ClearCollections()
    If Not g_TemplateCollections Is Nothing Then
        g_TemplateCollections.RemoveAll
    End If
End Sub

Public Sub m_AddCollectionItem( _
    ByVal collectionName As String, _
    ByVal itemMapText As String _
)
    Dim normalizedName As String
    Dim collectionItems As Collection
    Dim itemValues As Object

    normalizedName = mp_TrimWhitespace(CStr(collectionName))
    If Len(normalizedName) = 0 Then
        Err.Raise vbObjectError + 1838, "ex_ResultTemplatesParser", "Collection name is empty for m_AddCollectionItem."
    End If

    Set itemValues = mp_ParsePlaceholderMapTextToDictionary(itemMapText)
    Set collectionItems = mp_GetOrCreateTemplateCollection(normalizedName)
    collectionItems.Add itemValues
End Sub

Public Function m_GetFirstCollectionItemFieldByRegex( _
    ByVal collectionName As String, _
    ByVal matchFieldName As String, _
    ByVal regexPattern As String, _
    Optional ByVal returnFieldName As String = vbNullString _
) As String
    Dim collectionItems As Collection
    Dim itemValues As Object
    Dim i As Long
    Dim normalizedMatchFieldName As String
    Dim normalizedReturnFieldName As String
    Dim candidateValue As String
    Dim regexFirstMatch As String

    normalizedMatchFieldName = mp_TrimWhitespace(CStr(matchFieldName))
    normalizedReturnFieldName = mp_TrimWhitespace(CStr(returnFieldName))
    If Len(normalizedMatchFieldName) = 0 Then Exit Function
    If Len(normalizedReturnFieldName) = 0 Then normalizedReturnFieldName = normalizedMatchFieldName

    Set collectionItems = mp_GetTemplateCollectionItems(collectionName)
    If collectionItems Is Nothing Then Exit Function
    If collectionItems.Count = 0 Then Exit Function

    For i = 1 To collectionItems.Count
        Set itemValues = collectionItems(i)
        candidateValue = mp_GetCollectionItemFieldValue(itemValues, normalizedMatchFieldName)
        regexFirstMatch = ex_Helpers.m_RegexFirstMatch(candidateValue, CStr(regexPattern))
        If Len(regexFirstMatch) > 0 Then
            m_GetFirstCollectionItemFieldByRegex = mp_GetCollectionItemFieldValue(itemValues, normalizedReturnFieldName)
            Exit Function
        End If
    Next i
End Function

Public Function m_HasCollectionItemFieldByRegex( _
    ByVal collectionName As String, _
    ByVal matchFieldName As String, _
    ByVal regexPattern As String _
) As String
    Dim collectionItems As Collection
    Dim itemValues As Object
    Dim i As Long
    Dim normalizedMatchFieldName As String
    Dim candidateValue As String
    Dim regexFirstMatch As String

    normalizedMatchFieldName = mp_TrimWhitespace(CStr(matchFieldName))
    If Len(normalizedMatchFieldName) = 0 Then
        m_HasCollectionItemFieldByRegex = BOOLEAN_FALSE
        Exit Function
    End If

    Set collectionItems = mp_GetTemplateCollectionItems(collectionName)
    If collectionItems Is Nothing Then
        m_HasCollectionItemFieldByRegex = BOOLEAN_FALSE
        Exit Function
    End If
    If collectionItems.Count = 0 Then
        m_HasCollectionItemFieldByRegex = BOOLEAN_FALSE
        Exit Function
    End If

    For i = 1 To collectionItems.Count
        Set itemValues = collectionItems(i)
        candidateValue = mp_GetCollectionItemFieldValue(itemValues, normalizedMatchFieldName)
        regexFirstMatch = ex_Helpers.m_RegexFirstMatch(candidateValue, CStr(regexPattern))
        If Len(regexFirstMatch) > 0 Then
            m_HasCollectionItemFieldByRegex = BOOLEAN_TRUE
            Exit Function
        End If
    Next i

    m_HasCollectionItemFieldByRegex = BOOLEAN_FALSE
End Function

Public Function m_ExtractFirstCollectionItemFieldByRegex( _
    ByVal collectionName As String, _
    ByVal matchFieldName As String, _
    ByVal regexPattern As String, _
    Optional ByVal returnFieldName As String = vbNullString _
) As String
    Dim collectionItems As Collection
    Dim itemValues As Object
    Dim i As Long
    Dim normalizedMatchFieldName As String
    Dim normalizedReturnFieldName As String
    Dim candidateValue As String
    Dim regexFirstMatch As String

    normalizedMatchFieldName = mp_TrimWhitespace(CStr(matchFieldName))
    normalizedReturnFieldName = mp_TrimWhitespace(CStr(returnFieldName))
    If Len(normalizedMatchFieldName) = 0 Then Exit Function
    If Len(normalizedReturnFieldName) = 0 Then normalizedReturnFieldName = normalizedMatchFieldName

    Set collectionItems = mp_GetTemplateCollectionItems(collectionName)
    If collectionItems Is Nothing Then Exit Function
    If collectionItems.Count = 0 Then Exit Function

    For i = 1 To collectionItems.Count
        Set itemValues = collectionItems(i)
        candidateValue = mp_GetCollectionItemFieldValue(itemValues, normalizedMatchFieldName)
        regexFirstMatch = ex_Helpers.m_RegexFirstMatch(candidateValue, CStr(regexPattern))
        If Len(regexFirstMatch) > 0 Then
            m_ExtractFirstCollectionItemFieldByRegex = mp_GetCollectionItemFieldValue(itemValues, normalizedReturnFieldName)
            collectionItems.Remove i
            Exit Function
        End If
    Next i
End Function

Private Sub mp_ApplyPlaceholderMapEntry( _
    ByRef sourceText As String, _
    ByVal rawEntry As String, _
    ByVal allowTrailingSemicolon As Boolean, _
    ByVal highlightColorHex As String _
)
    Dim entryText As String
    Dim sepPos As Long
    Dim nameText As String
    Dim valueText As String

    entryText = Trim$(CStr(rawEntry))
    If Len(entryText) = 0 Then Exit Sub

    If allowTrailingSemicolon Then
        If Right$(entryText, 1) = ";" Then
            entryText = Trim$(Left$(entryText, Len(entryText) - 1))
            If Len(entryText) = 0 Then Exit Sub
        End If
    End If

    sepPos = InStr(1, entryText, ":", vbBinaryCompare)
    If sepPos <= 1 Then
        Err.Raise vbObjectError + 1835, "ex_ResultTemplatesParser", _
            "Invalid placeholder map entry '" & entryText & "'. Expected 'Name:Value'."
    End If

    nameText = Trim$(Left$(entryText, sepPos - 1))
    If Len(nameText) = 0 Then
        Err.Raise vbObjectError + 1836, "ex_ResultTemplatesParser", "Placeholder name cannot be empty in map entry."
    End If

    valueText = Trim$(Mid$(entryText, sepPos + 1))
    sourceText = m_ReplacePlaceholder(sourceText, nameText, valueText, highlightColorHex)
End Sub

Private Function mp_ParsePlaceholderMapTextToDictionary(ByVal placeholderMapText As String) As Object
    Dim normalized As String
    Dim lines As Variant
    Dim parts As Variant
    Dim itemSeparator As String
    Dim i As Long
    Dim valuesByKey As Object

    normalized = CStr(placeholderMapText)
    normalized = Replace(normalized, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    normalized = mp_TrimWhitespace(normalized)

    Set valuesByKey = CreateObject("Scripting.Dictionary")
    valuesByKey.CompareMode = 1 ' vbTextCompare

    If Len(normalized) = 0 Then
        Set mp_ParsePlaceholderMapTextToDictionary = valuesByKey
        Exit Function
    End If

    If InStr(1, normalized, vbLf, vbBinaryCompare) > 0 Then
        lines = Split(normalized, vbLf)
        For i = LBound(lines) To UBound(lines)
            mp_ApplyPlaceholderMapEntryToDictionary valuesByKey, CStr(lines(i)), True
        Next i
        Set mp_ParsePlaceholderMapTextToDictionary = valuesByKey
        Exit Function
    End If

    itemSeparator = ";"
    If InStr(1, normalized, ";", vbBinaryCompare) = 0 Then
        If InStr(1, normalized, ",", vbBinaryCompare) > 0 Then itemSeparator = ","
    End If

    parts = Split(normalized, itemSeparator)
    For i = LBound(parts) To UBound(parts)
        mp_ApplyPlaceholderMapEntryToDictionary valuesByKey, CStr(parts(i)), False
    Next i

    Set mp_ParsePlaceholderMapTextToDictionary = valuesByKey
End Function

Private Sub mp_ApplyPlaceholderMapEntryToDictionary( _
    ByVal valuesByKey As Object, _
    ByVal rawEntry As String, _
    ByVal allowTrailingSemicolon As Boolean _
)
    Dim entryText As String
    Dim sepPos As Long
    Dim nameText As String
    Dim valueText As String

    If valuesByKey Is Nothing Then
        Err.Raise vbObjectError + 1839, "ex_ResultTemplatesParser", "Internal error: placeholder map dictionary is not initialized."
    End If

    entryText = Trim$(CStr(rawEntry))
    If Len(entryText) = 0 Then Exit Sub

    If allowTrailingSemicolon Then
        If Right$(entryText, 1) = ";" Then
            entryText = Trim$(Left$(entryText, Len(entryText) - 1))
            If Len(entryText) = 0 Then Exit Sub
        End If
    End If

    sepPos = InStr(1, entryText, ":", vbBinaryCompare)
    If sepPos <= 1 Then
        Err.Raise vbObjectError + 1840, "ex_ResultTemplatesParser", _
            "Invalid collection item map entry '" & entryText & "'. Expected 'Name:Value'."
    End If

    nameText = Trim$(Left$(entryText, sepPos - 1))
    If Len(nameText) = 0 Then
        Err.Raise vbObjectError + 1841, "ex_ResultTemplatesParser", "Collection item field name cannot be empty in map entry."
    End If

    valueText = Trim$(Mid$(entryText, sepPos + 1))
    valuesByKey(nameText) = valueText
End Sub

Private Function mp_GetOrCreateTemplateCollection(ByVal collectionName As String) As Collection
    Dim normalizedName As String
    Dim collectionItems As Collection

    normalizedName = mp_TrimWhitespace(CStr(collectionName))
    If Len(normalizedName) = 0 Then
        Err.Raise vbObjectError + 1842, "ex_ResultTemplatesParser", "Collection name is empty."
    End If

    mp_EnsureTemplateCollectionsStore
    If g_TemplateCollections.Exists(normalizedName) Then
        Set mp_GetOrCreateTemplateCollection = g_TemplateCollections(normalizedName)
        Exit Function
    End If

    Set collectionItems = New Collection
    g_TemplateCollections.Add normalizedName, collectionItems
    Set mp_GetOrCreateTemplateCollection = collectionItems
End Function

Private Function mp_GetTemplateCollectionItems(ByVal collectionName As String) As Collection
    Dim normalizedName As String

    normalizedName = mp_TrimWhitespace(CStr(collectionName))
    If Len(normalizedName) = 0 Then
        Set mp_GetTemplateCollectionItems = New Collection
        Exit Function
    End If

    If g_TemplateCollections Is Nothing Then
        Set mp_GetTemplateCollectionItems = New Collection
        Exit Function
    End If

    If Not g_TemplateCollections.Exists(normalizedName) Then
        Set mp_GetTemplateCollectionItems = New Collection
        Exit Function
    End If

    Set mp_GetTemplateCollectionItems = g_TemplateCollections(normalizedName)
End Function

Private Function mp_GetCollectionItemFieldValue( _
    ByVal itemValues As Object, _
    ByVal fieldName As String _
) As String
    Dim normalizedFieldName As String
    normalizedFieldName = mp_TrimWhitespace(CStr(fieldName))
    If Len(normalizedFieldName) = 0 Then Exit Function
    If itemValues Is Nothing Then Exit Function
    If Not itemValues.Exists(normalizedFieldName) Then Exit Function
    mp_GetCollectionItemFieldValue = CStr(itemValues(normalizedFieldName))
End Function

Private Sub mp_EnsureTemplateCollectionsStore()
    If g_TemplateCollections Is Nothing Then
        Set g_TemplateCollections = CreateObject("Scripting.Dictionary")
        g_TemplateCollections.CompareMode = 1 ' vbTextCompare
    End If
End Sub

Private Function mp_ResolveNumericOffsetForPlaceholderValue( _
    ByVal sourceValue As String, _
    ByVal offsetText As String, _
    ByVal placeholderName As String _
) As String
    Dim baseNumber As Double
    Dim baseIsInteger As Boolean
    Dim offsetValue As Double
    Dim resultNumber As Double

    sourceValue = CStr(sourceValue)
    offsetText = mp_TrimWhitespace(CStr(offsetText))
    If Len(offsetText) = 0 Then
        mp_ResolveNumericOffsetForPlaceholderValue = sourceValue
        Exit Function
    End If

    If Not mp_TryParseTemplateNumeric(sourceValue, baseNumber, baseIsInteger) Then
        Err.Raise vbObjectError + 1824, "ex_ResultTemplatesParser", _
            "Placeholder '" & CStr(placeholderName) & "' value '" & CStr(sourceValue) & "' is not numeric for offset '{" & CStr(offsetText) & "}'."
    End If
    If Not mp_TryParseSignedInteger(offsetText, offsetValue) Then
        Err.Raise vbObjectError + 1825, "ex_ResultTemplatesParser", _
            "Invalid numeric offset '{" & CStr(offsetText) & "}' after placeholder '" & CStr(placeholderName) & "'. Use '{+N}' or '{-N}'."
    End If

    resultNumber = baseNumber + offsetValue
    If baseIsInteger And Fix(resultNumber) = resultNumber Then
        mp_ResolveNumericOffsetForPlaceholderValue = mp_FormatIntegerWithBasePadding(CLng(resultNumber), sourceValue)
    Else
        mp_ResolveNumericOffsetForPlaceholderValue = CStr(resultNumber)
    End If
End Function

Public Function m_ResolveTemplate( _
    ByVal sourceText As String, _
    Optional ByVal baseDateText As String = vbNullString, _
    Optional ByVal highlightColorHex As String = vbNullString _
) As String
    On Error GoTo EH
    ' Final pass for template text.
    m_ResolveTemplate = mp_ResolveTemplateLetBindings(CStr(sourceText), highlightColorHex)
    m_ResolveTemplate = mp_ResolveForBlocks(m_ResolveTemplate, highlightColorHex)
    m_ResolveTemplate = mp_ResolveConditionalBlocks(m_ResolveTemplate)
    m_ResolveTemplate = mp_ResolveJoinLineTokens(m_ResolveTemplate)
    m_ResolveTemplate = mp_ResolveTrimIndentTokens(m_ResolveTemplate)
    m_ResolveTemplate = mp_ResolveNewLineTokens(m_ResolveTemplate)
    m_ResolveTemplate = mp_ResolveDateExpressions(m_ResolveTemplate, baseDateText)
    m_ResolveTemplate = mp_CollapseUnresolvedPlaceholders(m_ResolveTemplate)
    m_ResolveTemplate = mp_CollapseHorizontalWhitespaceRuns(m_ResolveTemplate)
    Exit Function

EH:
    m_ResolveTemplate = mp_PrependTemplateError(CStr(sourceText), "m_ResolveTemplate")
End Function

Public Function m_ExtractHighlightSegments( _
    ByVal sourceText As String, _
    ByRef outSegments As Collection, _
    Optional ByVal fallbackColorHex As String = vbNullString, _
    Optional ByVal includeNamedColorTags As Boolean = True _
) As String
    Dim resultText As String
    Dim scanPos As Long
    Dim nextStartPos As Long
    Dim nextEndPos As Long
    Dim markerPos As Long
    Dim markerTailPos As Long
    Dim markerEndPos As Long
    Dim isStartMarker As Boolean
    Dim textChunk As String
    Dim colorHex As String
    Dim colorStartPos As Long
    Dim colorEndPos As Long
    Dim segmentStart As Long
    Dim segmentLength As Long
    Dim segment As Object
    Dim markerStartLen As Long
    Dim markerEndLen As Long
    Dim stackSize As Long
    Dim stackStarts() As Long
    Dim stackColors() As String

    On Error GoTo EH

    Set outSegments = New Collection
    resultText = vbNullString
    sourceText = CStr(sourceText)
    fallbackColorHex = mp_TrimWhitespace(CStr(fallbackColorHex))
    If Len(fallbackColorHex) = 0 Then fallbackColorHex = DEFAULT_HIGHLIGHT_COLOR_HEX
    If includeNamedColorTags Then
        sourceText = mp_ConvertNamedColorTagsToHighlightMarkers(sourceText)
    End If

    markerStartLen = Len(HIGHLIGHT_MARKER_START)
    markerEndLen = Len(HIGHLIGHT_MARKER_END)
    scanPos = 1
    Do While scanPos <= Len(sourceText)
        nextStartPos = InStr(scanPos, sourceText, HIGHLIGHT_MARKER_START, vbBinaryCompare)
        nextEndPos = InStr(scanPos, sourceText, HIGHLIGHT_MARKER_END, vbBinaryCompare)

        If nextStartPos = 0 And nextEndPos = 0 Then
            resultText = resultText & Mid$(sourceText, scanPos)
            Exit Do
        End If

        isStartMarker = False
        markerPos = 0
        If nextStartPos > 0 Then
            If nextEndPos = 0 Or nextStartPos <= nextEndPos Then
                isStartMarker = True
                markerPos = nextStartPos
            Else
                markerPos = nextEndPos
            End If
        Else
            markerPos = nextEndPos
        End If

        If markerPos > scanPos Then
            textChunk = Mid$(sourceText, scanPos, markerPos - scanPos)
            resultText = resultText & textChunk
        End If

        If isStartMarker Then
            markerTailPos = markerPos + markerStartLen
            colorHex = fallbackColorHex

            If markerTailPos <= Len(sourceText) Then
                If Mid$(sourceText, markerTailPos, 1) = "(" Then
                    colorStartPos = markerTailPos + 1
                    colorEndPos = InStr(colorStartPos, sourceText, ")", vbBinaryCompare)
                    If colorEndPos > 0 Then
                        colorHex = mp_TrimWhitespace(Mid$(sourceText, colorStartPos, colorEndPos - colorStartPos))
                        If Len(colorHex) = 0 Then colorHex = fallbackColorHex
                        markerTailPos = colorEndPos + 1
                    End If
                End If
            End If

            stackSize = stackSize + 1
            ReDim Preserve stackStarts(1 To stackSize)
            ReDim Preserve stackColors(1 To stackSize)
            stackStarts(stackSize) = CLng(Len(resultText) + 1)
            stackColors(stackSize) = CStr(colorHex)

            scanPos = markerTailPos
        Else
            markerEndPos = markerPos + markerEndLen
            If stackSize > 0 Then
                segmentStart = CLng(stackStarts(stackSize))
                colorHex = CStr(stackColors(stackSize))
                segmentLength = Len(resultText) - segmentStart + 1
                If segmentStart > 0 And segmentLength > 0 Then
                    Set segment = CreateObject("Scripting.Dictionary")
                    segment.CompareMode = 1 ' vbTextCompare
                    segment("Start") = segmentStart
                    segment("Length") = CLng(segmentLength)
                    segment("ColorHex") = CStr(colorHex)
                    outSegments.Add segment
                End If

                stackSize = stackSize - 1
                If stackSize > 0 Then
                    ReDim Preserve stackStarts(1 To stackSize)
                    ReDim Preserve stackColors(1 To stackSize)
                Else
                    Erase stackStarts
                    Erase stackColors
                End If
            Else
                resultText = resultText & HIGHLIGHT_MARKER_END
            End If

            scanPos = markerEndPos
        End If
    Loop

    Do While stackSize > 0
        segmentStart = CLng(stackStarts(stackSize))
        colorHex = CStr(stackColors(stackSize))
        segmentLength = Len(resultText) - segmentStart + 1
        If segmentStart > 0 And segmentLength > 0 Then
            Set segment = CreateObject("Scripting.Dictionary")
            segment.CompareMode = 1 ' vbTextCompare
            segment("Start") = segmentStart
            segment("Length") = CLng(segmentLength)
            segment("ColorHex") = CStr(colorHex)
            outSegments.Add segment
        End If

        stackSize = stackSize - 1
        If stackSize > 0 Then
            ReDim Preserve stackStarts(1 To stackSize)
            ReDim Preserve stackColors(1 To stackSize)
        Else
            Erase stackStarts
            Erase stackColors
        End If
    Loop

    m_ExtractHighlightSegments = resultText
    Exit Function

EH:
    Set outSegments = New Collection
    m_ExtractHighlightSegments = CStr(sourceText)
End Function

Public Function m_RemoveHighlightMarkers(ByVal sourceText As String) As String
    Dim segments As Collection

    m_RemoveHighlightMarkers = m_ExtractHighlightSegments(CStr(sourceText), segments, DEFAULT_HIGHLIGHT_COLOR_HEX, False)
End Function

Private Function mp_ConvertNamedColorTagsToHighlightMarkers(ByVal sourceText As String) As String
    Dim rxBegin As Object
    Dim rxEnd As Object
    Dim matches As Object
    Dim colorToken As String
    Dim colorHex As String
    Dim matchStart As Long
    Dim matchLen As Long
    Dim resultText As String

    resultText = CStr(sourceText)
    If Len(resultText) = 0 Then
        mp_ConvertNamedColorTagsToHighlightMarkers = resultText
        Exit Function
    End If

    Set rxBegin = CreateObject("VBScript.RegExp")
    rxBegin.Global = False
    rxBegin.IgnoreCase = True
    rxBegin.MultiLine = True
    rxBegin.Pattern = "\{#color_([A-Za-z0-9_-]+)_begin\}"

    Do
        Set matches = rxBegin.Execute(resultText)
        If matches Is Nothing Then Exit Do
        If matches.Count = 0 Then Exit Do

        colorToken = CStr(matches(0).SubMatches(0))
        If Not mp_TryResolveNamedColorHex(colorToken, colorHex) Then
            Err.Raise vbObjectError + 1851, "ex_ResultTemplatesParser", _
                "Unknown named color token '#color_" & colorToken & "_begin'. Use known color name (red/green/blue/...) or hex form 'hex_RRGGBB'."
        End If

        matchStart = CLng(matches(0).FirstIndex) + 1
        matchLen = CLng(matches(0).Length)
        resultText = Left$(resultText, matchStart - 1) & HIGHLIGHT_MARKER_START & "(" & colorHex & ")" & Mid$(resultText, matchStart + matchLen)
    Loop

    Set rxEnd = CreateObject("VBScript.RegExp")
    rxEnd.Global = True
    rxEnd.IgnoreCase = True
    rxEnd.MultiLine = True
    rxEnd.Pattern = "\{#color_[A-Za-z0-9_-]+_end\}"
    resultText = rxEnd.Replace(resultText, HIGHLIGHT_MARKER_END)

    rxEnd.Pattern = "\{#color_end\}"
    resultText = rxEnd.Replace(resultText, HIGHLIGHT_MARKER_END)

    mp_ConvertNamedColorTagsToHighlightMarkers = resultText
End Function

Private Function mp_TryResolveNamedColorHex( _
    ByVal colorToken As String, _
    ByRef outColorHex As String _
) As Boolean
    Dim normalized As String

    normalized = LCase$(mp_TrimWhitespace(CStr(colorToken)))
    If Len(normalized) = 0 Then Exit Function

    Select Case normalized
        Case "red"
            outColorHex = "#FF0000"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "green"
            outColorHex = "#008000"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "blue"
            outColorHex = "#0000FF"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "cyan"
            outColorHex = "#00B7FF"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "orange"
            outColorHex = "#FFA500"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "yellow"
            outColorHex = "#FFD700"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "purple"
            outColorHex = "#800080"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "teal"
            outColorHex = "#008080"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "gray", "grey"
            outColorHex = "#808080"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "black"
            outColorHex = "#000000"
            mp_TryResolveNamedColorHex = True
            Exit Function
        Case "white"
            outColorHex = "#FFFFFF"
            mp_TryResolveNamedColorHex = True
            Exit Function
    End Select

    If Left$(normalized, 4) = "hex_" Then
        normalized = Mid$(normalized, 5)
    End If
    If Left$(normalized, 1) = "#" Then
        normalized = Mid$(normalized, 2)
    End If

    If mp_IsHexColorLiteral(normalized) Then
        outColorHex = "#" & UCase$(normalized)
        mp_TryResolveNamedColorHex = True
    End If
End Function

Private Function mp_IsHexColorLiteral(ByVal colorText As String) As Boolean
    Dim rx As Object
    Dim matches As Object

    colorText = mp_TrimWhitespace(CStr(colorText))
    If Len(colorText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.MultiLine = False
    rx.Pattern = "^[0-9A-F]{6}([0-9A-F]{2})?$"

    Set matches = rx.Execute(UCase$(colorText))
    If matches Is Nothing Then Exit Function
    If matches.Count > 0 Then mp_IsHexColorLiteral = True
End Function

Private Function mp_ShouldCollapsePlaceholderValue(ByVal replacementText As String) As Boolean
    Dim normalized As String

    normalized = mp_TrimWhitespace(CStr(replacementText))
    If Len(normalized) = 0 Then
        mp_ShouldCollapsePlaceholderValue = True
        Exit Function
    End If

    If normalized = "-" Or normalized = "?" Then
        mp_ShouldCollapsePlaceholderValue = True
    End If
End Function

Private Function mp_WrapHighlightMarkers( _
    ByVal valueText As String, _
    ByVal highlightColorHex As String _
) As String
    highlightColorHex = mp_TrimWhitespace(CStr(highlightColorHex))
    valueText = CStr(valueText)

    If Len(valueText) = 0 Or Len(highlightColorHex) = 0 Then
        mp_WrapHighlightMarkers = valueText
        Exit Function
    End If

    mp_WrapHighlightMarkers = HIGHLIGHT_MARKER_START & "(" & highlightColorHex & ")" & valueText & HIGHLIGHT_MARKER_END
End Function

Private Function mp_CollapseNamedPlaceholderTokens( _
    ByVal sourceText As String, _
    ByVal placeholderName As String _
) As String
    Dim src As String
    Dim resultText As String
    Dim i As Long
    Dim phEndPos As Long
    Dim tokenEndPos As Long
    Dim formatterPipeline As String
    Dim offsetText As String

    src = CStr(sourceText)
    i = 1
    Do While i <= Len(src)
        If Mid$(src, i, 1) = "{" Then
            If mp_TryParseNamedPlaceholderTokenAt(src, i, CStr(placeholderName), phEndPos, formatterPipeline, offsetText, tokenEndPos) Then
                i = tokenEndPos + 1
                GoTo ContinueLoop
            End If
        End If

        resultText = resultText & Mid$(src, i, 1)
        i = i + 1
ContinueLoop:
    Loop

    mp_CollapseNamedPlaceholderTokens = resultText
End Function

Private Function mp_ReplaceNamedPlaceholderTokens( _
    ByVal sourceText As String, _
    ByVal placeholderName As String, _
    ByVal replacementText As String, _
    ByVal highlightColorHex As String _
) As String
    Dim src As String
    Dim resultText As String
    Dim i As Long
    Dim phEndPos As Long
    Dim tokenEndPos As Long
    Dim formatterPipeline As String
    Dim offsetText As String
    Dim resolvedValue As String
    Dim formattedValue As String

    src = CStr(sourceText)
    i = 1
    Do While i <= Len(src)
        If Mid$(src, i, 1) = "{" Then
            If mp_TryParseNamedPlaceholderTokenAt(src, i, CStr(placeholderName), phEndPos, formatterPipeline, offsetText, tokenEndPos) Then
                formattedValue = CStr(replacementText)
                If Len(formatterPipeline) > 0 Then
                    formattedValue = mp_ApplyFormatterPipeline(formattedValue, formatterPipeline)
                End If

                resolvedValue = mp_ResolveNumericOffsetForPlaceholderValue(formattedValue, offsetText, CStr(placeholderName))
                resolvedValue = mp_WrapHighlightMarkers(resolvedValue, highlightColorHex)
                resultText = resultText & resolvedValue

                i = tokenEndPos + 1
                GoTo ContinueLoop
            End If
        End If

        resultText = resultText & Mid$(src, i, 1)
        i = i + 1
ContinueLoop:
    Loop

    mp_ReplaceNamedPlaceholderTokens = resultText
End Function

Private Function mp_TryParseNamedPlaceholderTokenAt( _
    ByVal sourceText As String, _
    ByVal startPos As Long, _
    ByVal placeholderName As String, _
    ByRef outPlaceholderEndPos As Long, _
    ByRef outFormatterPipeline As String, _
    ByRef outOffsetText As String, _
    ByRef outTokenEndPos As Long _
) As Boolean
    Dim nameLen As Long
    Dim nameStartPos As Long
    Dim afterNamePos As Long
    Dim tailPos As Long
    Dim firstTailChar As String
    Dim closePos As Long
    Dim offsetEndPos As Long

    sourceText = CStr(sourceText)
    placeholderName = CStr(placeholderName)
    nameLen = Len(placeholderName)
    If nameLen = 0 Then Exit Function
    If startPos < 1 Or startPos > Len(sourceText) Then Exit Function

    If Mid$(sourceText, startPos, 1) <> "{" Then Exit Function

    nameStartPos = startPos + 1
    If nameStartPos + nameLen - 1 > Len(sourceText) Then Exit Function
    If StrComp(Mid$(sourceText, nameStartPos, nameLen), placeholderName, vbBinaryCompare) <> 0 Then Exit Function

    afterNamePos = nameStartPos + nameLen
    If afterNamePos > Len(sourceText) Then Exit Function

    tailPos = mp_FindNextNonWhitespacePosition(sourceText, afterNamePos)
    If tailPos = 0 Then Exit Function

    firstTailChar = Mid$(sourceText, tailPos, 1)
    If firstTailChar = "}" Then
        closePos = tailPos
        outFormatterPipeline = vbNullString
    ElseIf firstTailChar = "|" Then
        If Not mp_TryFindPlaceholderTokenClose(sourceText, startPos, closePos) Then Exit Function
        outFormatterPipeline = Mid$(sourceText, tailPos + 1, closePos - tailPos - 1)
    Else
        Exit Function
    End If

    outPlaceholderEndPos = closePos
    outOffsetText = vbNullString
    outTokenEndPos = closePos
    If mp_TryParseNumericOffsetTokenAt(sourceText, closePos + 1, outOffsetText, offsetEndPos) Then
        outTokenEndPos = offsetEndPos
    End If

    mp_TryParseNamedPlaceholderTokenAt = True
End Function

Private Function mp_TryFindPlaceholderTokenClose( _
    ByVal sourceText As String, _
    ByVal openBracePos As Long, _
    ByRef outClosePos As Long _
) As Boolean
    Dim i As Long
    Dim ch As String
    Dim quoteChar As String
    Dim depth As Long
    Dim sourceLen As Long

    sourceText = CStr(sourceText)
    sourceLen = Len(sourceText)
    If openBracePos < 1 Or openBracePos > sourceLen Then Exit Function
    If Mid$(sourceText, openBracePos, 1) <> "{" Then Exit Function

    depth = 1
    i = openBracePos + 1
    Do While i <= sourceLen
        ch = Mid$(sourceText, i, 1)
        If Len(quoteChar) = 0 Then
            If ch = """" Or ch = "'" Then
                quoteChar = ch
                i = i + 1
            ElseIf ch = "{" Then
                depth = depth + 1
                i = i + 1
            ElseIf ch = "}" Then
                depth = depth - 1
                If depth = 0 Then
                    outClosePos = i
                    mp_TryFindPlaceholderTokenClose = True
                    Exit Function
                End If
                i = i + 1
            Else
                i = i + 1
            End If
        Else
            If ch = "\" Then
                If i < sourceLen Then
                    i = i + 2
                Else
                    i = i + 1
                End If
            ElseIf ch = quoteChar Then
                quoteChar = vbNullString
                i = i + 1
            Else
                i = i + 1
            End If
        End If
    Loop
End Function

Private Function mp_TryParseNumericOffsetTokenAt( _
    ByVal sourceText As String, _
    ByVal startPos As Long, _
    ByRef outOffsetText As String, _
    ByRef outEndPos As Long _
) As Boolean
    Dim i As Long
    Dim signChar As String
    Dim sourceLen As Long
    Dim digitStart As Long
    Dim ch As String

    sourceText = CStr(sourceText)
    sourceLen = Len(sourceText)
    If startPos < 1 Or startPos > sourceLen Then Exit Function
    If Mid$(sourceText, startPos, 1) <> "{" Then Exit Function

    If startPos + 2 > sourceLen Then Exit Function
    signChar = Mid$(sourceText, startPos + 1, 1)
    If signChar <> "+" And signChar <> "-" Then Exit Function

    digitStart = startPos + 2
    i = digitStart
    Do While i <= sourceLen
        ch = Mid$(sourceText, i, 1)
        If ch < "0" Or ch > "9" Then Exit Do
        i = i + 1
    Loop

    If i = digitStart Then Exit Function
    If i > sourceLen Then Exit Function
    If Mid$(sourceText, i, 1) <> "}" Then Exit Function

    outOffsetText = Mid$(sourceText, startPos + 1, i - startPos - 1)
    outEndPos = i
    mp_TryParseNumericOffsetTokenAt = True
End Function

Private Function mp_CollapseUnresolvedPlaceholders(ByVal sourceText As String) As String
    Dim src As String
    Dim resultText As String
    Dim i As Long
    Dim tokenEndPos As Long

    src = CStr(sourceText)
    i = 1
    Do While i <= Len(src)
        If Mid$(src, i, 1) = "{" Then
            If mp_TryParseAnyPlaceholderTokenAt(src, i, tokenEndPos) Then
                i = tokenEndPos + 1
                GoTo ContinueLoop
            End If
        End If

        resultText = resultText & Mid$(src, i, 1)
        i = i + 1
ContinueLoop:
    Loop

    mp_CollapseUnresolvedPlaceholders = resultText
End Function

Private Function mp_TryParseAnyPlaceholderTokenAt( _
    ByVal sourceText As String, _
    ByVal startPos As Long, _
    ByRef outTokenEndPos As Long _
) As Boolean
    Dim sourceLen As Long
    Dim i As Long
    Dim tailPos As Long
    Dim ch As String
    Dim nameEndPos As Long
    Dim tailChar As String
    Dim closePos As Long
    Dim offsetText As String
    Dim offsetEndPos As Long

    sourceText = CStr(sourceText)
    sourceLen = Len(sourceText)
    If startPos < 1 Or startPos > sourceLen Then Exit Function
    If Mid$(sourceText, startPos, 1) <> "{" Then Exit Function
    If startPos + 1 > sourceLen Then Exit Function

    ch = Mid$(sourceText, startPos + 1, 1)
    If Not mp_IsPlaceholderNameStartChar(ch) Then Exit Function

    i = startPos + 2
    Do While i <= sourceLen
        ch = Mid$(sourceText, i, 1)
        If Not mp_IsPlaceholderNamePartChar(ch) Then Exit Do
        i = i + 1
    Loop
    nameEndPos = i - 1
    If nameEndPos < startPos + 1 Then Exit Function
    If i > sourceLen Then Exit Function

    tailPos = mp_FindNextNonWhitespacePosition(sourceText, i)
    If tailPos = 0 Then Exit Function

    tailChar = Mid$(sourceText, tailPos, 1)
    If tailChar = "}" Then
        closePos = tailPos
    ElseIf tailChar = "|" Then
        If Not mp_TryFindPlaceholderTokenClose(sourceText, startPos, closePos) Then Exit Function
    Else
        Exit Function
    End If

    outTokenEndPos = closePos
    If mp_TryParseNumericOffsetTokenAt(sourceText, closePos + 1, offsetText, offsetEndPos) Then
        outTokenEndPos = offsetEndPos
    End If

    mp_TryParseAnyPlaceholderTokenAt = True
End Function

Private Function mp_IsPlaceholderNameStartChar(ByVal ch As String) As Boolean
    ch = CStr(ch)
    If Len(ch) <> 1 Then Exit Function
    mp_IsPlaceholderNameStartChar = _
        ((ch >= "A" And ch <= "Z") Or _
         (ch >= "a" And ch <= "z") Or _
         ch = "_")
End Function

Private Function mp_IsPlaceholderNamePartChar(ByVal ch As String) As Boolean
    ch = CStr(ch)
    If Len(ch) <> 1 Then Exit Function
    mp_IsPlaceholderNamePartChar = _
        ((ch >= "A" And ch <= "Z") Or _
         (ch >= "a" And ch <= "z") Or _
         (ch >= "0" And ch <= "9") Or _
         ch = "_" Or _
         ch = ".")
End Function

Private Function mp_FindNextNonWhitespacePosition( _
    ByVal sourceText As String, _
    ByVal startPos As Long _
) As Long
    Dim i As Long
    Dim sourceLen As Long
    Dim ch As String

    sourceText = CStr(sourceText)
    sourceLen = Len(sourceText)
    If startPos < 1 Then startPos = 1
    If startPos > sourceLen Then Exit Function

    For i = startPos To sourceLen
        ch = Mid$(sourceText, i, 1)
        If Not mp_IsWhitespaceChar(ch) Then
            mp_FindNextNonWhitespacePosition = i
            Exit Function
        End If
    Next i
End Function

Private Function mp_CollapseHorizontalWhitespaceRuns(ByVal sourceText As String) As String
    Dim rx As Object

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "[ \t]{2,}"

    mp_CollapseHorizontalWhitespaceRuns = rx.Replace(CStr(sourceText), " ")
End Function

Public Function m_FormatValue( _
    ByVal sourceValue As String, _
    ByVal formatterName As String _
) As String
    Dim normalizedFormatter As String

    On Error GoTo EH

    normalizedFormatter = mp_TrimWhitespace(formatterName)
    If Len(normalizedFormatter) = 0 Then
        m_FormatValue = CStr(sourceValue)
        Exit Function
    End If

    m_FormatValue = mp_ApplyFormatterPipeline(CStr(sourceValue), normalizedFormatter)
    Exit Function

EH:
    m_FormatValue = mp_PrependTemplateError(CStr(sourceValue), "m_FormatValue('" & CStr(formatterName) & "')")
End Function

Private Function mp_FormatUaDateDay(ByVal sourceDateText As String) As String
    mp_FormatUaDateDay = ex_DateHelpers.m_FormatDateDay(CStr(sourceDateText))
End Function

Private Function mp_FormatUaDateDayWithMonth(ByVal sourceDateText As String) As String
    mp_FormatUaDateDayWithMonth = ex_DateHelpers.m_FormatDateDayWithMonth(CStr(sourceDateText))
End Function

Private Function mp_FormatSurnameInitials(ByVal sourceValue As String) As String
    mp_FormatSurnameInitials = ex_MorphUaLite.m_ToShortFioNormalized(CStr(sourceValue))
End Function

Private Function mp_FormatFioSurname(ByVal sourceValue As String) As String
    mp_FormatFioSurname = ex_MorphUaLite.m_ToFioSurnameNormalized(CStr(sourceValue))
End Function

Private Function mp_FormatFioInitials(ByVal sourceValue As String) As String
    mp_FormatFioInitials = ex_MorphUaLite.m_ToFioInitials(CStr(sourceValue))
End Function

Private Function mp_CapitalizeText(ByVal textValue As String) As String
    textValue = CStr(textValue)
    If Len(textValue) = 0 Then Exit Function
    If Len(textValue) = 1 Then
        mp_CapitalizeText = UCase$(textValue)
        Exit Function
    End If

    mp_CapitalizeText = UCase$(Left$(textValue, 1)) & LCase$(Mid$(textValue, 2))
End Function

Private Function mp_ApplyFormatter(ByVal sourceValue As String, ByVal formatterName As String) As String
    Dim normalizedFormatter As String

    normalizedFormatter = LCase$(mp_TrimWhitespace(formatterName))

    Select Case normalizedFormatter
        Case FORMATTER_UPPER
            mp_ApplyFormatter = UCase$(CStr(sourceValue))
        Case FORMATTER_LOWER
            mp_ApplyFormatter = LCase$(CStr(sourceValue))
        Case FORMATTER_CAPITALIZE
            mp_ApplyFormatter = mp_CapitalizeText(CStr(sourceValue))
        Case FORMATTER_FIRSTCHAR
            mp_ApplyFormatter = mp_FirstNonSpaceChar(CStr(sourceValue))
        Case FORMATTER_UPPERFIRSTLETTER
            mp_ApplyFormatter = mp_UppercaseFirstLetter(CStr(sourceValue))
        Case FORMATTER_UPPERFIRSTWORD
            mp_ApplyFormatter = mp_UppercaseFirstWord(CStr(sourceValue))
        Case FORMATTER_LOWERFIRSTLETTER
            mp_ApplyFormatter = ex_MorphUaLite.m_LowercaseFirstLetter(CStr(sourceValue))
        Case FORMATTER_LOWERFIRSTWORD
            mp_ApplyFormatter = mp_LowercaseFirstWord(CStr(sourceValue))
        Case FORMATTER_GENITIVE
            mp_ApplyFormatter = mp_InflectPhraseToCase(CStr(sourceValue), CASE_GENITIVE)
        Case FORMATTER_ACCUSATIVE
            mp_ApplyFormatter = mp_InflectPhraseToCase(CStr(sourceValue), CASE_ACCUSATIVE)
        Case FORMATTER_DATIVE
            mp_ApplyFormatter = mp_InflectPhraseToCase(CStr(sourceValue), CASE_DATIVE)
        Case FORMATTER_TO_DATE_DAY
            mp_ApplyFormatter = mp_FormatUaDateDay(CStr(sourceValue))
        Case FORMATTER_TO_DATE_DAY_WITH_MONTH
            mp_ApplyFormatter = mp_FormatUaDateDayWithMonth(CStr(sourceValue))
        Case FORMATTER_CALENDAR_DAYS_UA
            mp_ApplyFormatter = ex_DateHelpers.m_FormatCalendarDaysUa(CStr(sourceValue))
        Case FORMATTER_SURNAME_INITIALS
            mp_ApplyFormatter = mp_FormatSurnameInitials(CStr(sourceValue))
        Case FORMATTER_FIO_SURNAME
            mp_ApplyFormatter = mp_FormatFioSurname(CStr(sourceValue))
        Case FORMATTER_FIO_INITIALS
            mp_ApplyFormatter = mp_FormatFioInitials(CStr(sourceValue))
        Case Else
            Err.Raise vbObjectError + 1766, "ex_ResultTemplatesParser", _
                "Unsupported formatter '" & formatterName & "'."
    End Select
End Function

Private Function mp_UppercaseFirstLetter(ByVal sourceValue As String) As String
    Dim textValue As String

    textValue = CStr(sourceValue)
    If Len(textValue) = 0 Then
        mp_UppercaseFirstLetter = textValue
        Exit Function
    End If

    mp_UppercaseFirstLetter = UCase$(Left$(textValue, 1)) & Mid$(textValue, 2)
End Function

Private Function mp_UppercaseFirstWord(ByVal sourceValue As String) As String
    Dim textValue As String
    Dim wordStart As Long
    Dim wordEnd As Long
    Dim ch As String

    textValue = CStr(sourceValue)
    If Len(textValue) = 0 Then
        mp_UppercaseFirstWord = textValue
        Exit Function
    End If

    wordStart = 1
    Do While wordStart <= Len(textValue)
        ch = Mid$(textValue, wordStart, 1)
        If Not mp_IsWhitespaceChar(ch) Then Exit Do
        wordStart = wordStart + 1
    Loop

    If wordStart > Len(textValue) Then
        mp_UppercaseFirstWord = textValue
        Exit Function
    End If

    wordEnd = wordStart
    Do While wordEnd <= Len(textValue)
        ch = Mid$(textValue, wordEnd, 1)
        If mp_IsWhitespaceChar(ch) Then Exit Do
        wordEnd = wordEnd + 1
    Loop

    mp_UppercaseFirstWord = _
        Left$(textValue, wordStart - 1) & _
        UCase$(Mid$(textValue, wordStart, wordEnd - wordStart)) & _
        Mid$(textValue, wordEnd)
End Function

Private Function mp_LowercaseFirstWord(ByVal sourceValue As String) As String
    Dim textValue As String
    Dim wordStart As Long
    Dim wordEnd As Long
    Dim ch As String

    textValue = CStr(sourceValue)
    If Len(textValue) = 0 Then
        mp_LowercaseFirstWord = textValue
        Exit Function
    End If

    wordStart = 1
    Do While wordStart <= Len(textValue)
        ch = Mid$(textValue, wordStart, 1)
        If Not mp_IsWhitespaceChar(ch) Then Exit Do
        wordStart = wordStart + 1
    Loop

    If wordStart > Len(textValue) Then
        mp_LowercaseFirstWord = textValue
        Exit Function
    End If

    wordEnd = wordStart
    Do While wordEnd <= Len(textValue)
        ch = Mid$(textValue, wordEnd, 1)
        If mp_IsWhitespaceChar(ch) Then Exit Do
        wordEnd = wordEnd + 1
    Loop

    mp_LowercaseFirstWord = _
        Left$(textValue, wordStart - 1) & _
        LCase$(Mid$(textValue, wordStart, wordEnd - wordStart)) & _
        Mid$(textValue, wordEnd)
End Function

Private Function mp_ApplyFormatterPipeline(ByVal sourceValue As String, ByVal formatterPipeline As String) As String
    Dim i As Long
    Dim actionSpec As String
    Dim formattedValue As String
    Dim actions As Collection

    formatterPipeline = CStr(formatterPipeline)
    actionSpec = mp_TrimWhitespace(formatterPipeline)
    If Len(actionSpec) = 0 Then
        mp_ApplyFormatterPipeline = CStr(sourceValue)
        Exit Function
    End If

    formattedValue = CStr(sourceValue)
    Set actions = mp_SplitByDelimiterOutsideQuotedLiterals(actionSpec, "|")

    For i = 1 To actions.Count
        actionSpec = mp_TrimWhitespace(CStr(actions(i)))
        If Len(actionSpec) = 0 Then
            Err.Raise vbObjectError + 1770, "ex_ResultTemplatesParser", "Empty formatter action in '" & formatterPipeline & "'."
        End If
        formattedValue = mp_ApplyFormatterAction(formattedValue, actionSpec)
    Next i

    mp_ApplyFormatterPipeline = formattedValue
End Function

Private Function mp_ApplyFormatterAction(ByVal sourceValue As String, ByVal actionSpec As String) As String
    Dim colonPos As Long
    Dim actionName As String
    Dim actionArgs As String
    Dim commaPos As Long
    Dim replaceFrom As String
    Dim replaceTo As String
    Dim regexPattern As String
    Dim regexReplaceTo As String
    Dim maxLen As Long

    actionSpec = mp_TrimWhitespace(CStr(actionSpec))
    colonPos = InStr(1, actionSpec, ":", vbBinaryCompare)
    If colonPos > 0 Then
        actionName = LCase$(mp_TrimWhitespace(Left$(actionSpec, colonPos - 1)))
        actionArgs = Mid$(actionSpec, colonPos + 1)
    Else
        actionName = LCase$(actionSpec)
        actionArgs = vbNullString
    End If

    Select Case actionName
        Case FORMATTER_TRUNCATE
            actionArgs = mp_TrimWhitespace(actionArgs)
            If Not mp_TryParseNonNegativeLong(actionArgs, maxLen) Then
                Err.Raise vbObjectError + 1771, "ex_ResultTemplatesParser", "truncate requires non-negative integer argument: '" & actionSpec & "'."
            End If
            If maxLen <= 0 Then
                mp_ApplyFormatterAction = vbNullString
            ElseIf Len(sourceValue) <= maxLen Then
                mp_ApplyFormatterAction = CStr(sourceValue)
            Else
                mp_ApplyFormatterAction = Left$(CStr(sourceValue), maxLen)
            End If
            Exit Function

        Case FORMATTER_REPLACE
            commaPos = mp_IndexOfDelimiterOutsideQuotedLiterals(actionArgs, ",")
            If commaPos <= 0 Then
                Err.Raise vbObjectError + 1772, "ex_ResultTemplatesParser", "replace requires two args 'from,to': '" & actionSpec & "'."
            End If
            replaceFrom = Left$(actionArgs, commaPos - 1)
            replaceTo = Mid$(actionArgs, commaPos + 1)
            If Len(replaceFrom) = 0 Then
                Err.Raise vbObjectError + 1773, "ex_ResultTemplatesParser", "replace 'from' argument cannot be empty: '" & actionSpec & "'."
            End If
            mp_ApplyFormatterAction = Replace(CStr(sourceValue), replaceFrom, replaceTo)
            Exit Function

        Case FORMATTER_REGEX_REPLACE
            commaPos = mp_IndexOfDelimiterOutsideQuotedLiterals(actionArgs, ",")
            If commaPos <= 0 Then
                Err.Raise vbObjectError + 1828, "ex_ResultTemplatesParser", "regexreplace requires two args 'pattern,replacement': '" & actionSpec & "'."
            End If

            regexPattern = mp_UnquoteFormatterArgument(Left$(actionArgs, commaPos - 1))
            regexReplaceTo = mp_UnquoteFormatterArgument(Mid$(actionArgs, commaPos + 1))
            If Len(regexPattern) = 0 Then
                Err.Raise vbObjectError + 1829, "ex_ResultTemplatesParser", "regexreplace 'pattern' cannot be empty: '" & actionSpec & "'."
            End If

            mp_ApplyFormatterAction = mp_RegexReplaceText(CStr(sourceValue), regexPattern, regexReplaceTo)
            Exit Function

        Case FORMATTER_DATEFORMAT
            actionArgs = mp_UnquoteFormatterArgument(actionArgs)
            If Len(actionArgs) = 0 Then
                Err.Raise vbObjectError + 1801, "ex_ResultTemplatesParser", "dateformat requires non-empty format argument: '" & actionSpec & "'."
            End If
            mp_ApplyFormatterAction = ex_DateHelpers.m_FormatDateByPattern(CStr(sourceValue), actionArgs)
            Exit Function

        Case Else
            If Len(actionArgs) > 0 Then
                Err.Raise vbObjectError + 1774, "ex_ResultTemplatesParser", "Formatter '" & actionName & "' does not support arguments."
            End If
            mp_ApplyFormatterAction = mp_ApplyFormatter(CStr(sourceValue), actionName)
            Exit Function
    End Select
End Function

Private Function mp_SplitByDelimiterOutsideQuotedLiterals( _
    ByVal sourceText As String, _
    ByVal delimiterChar As String _
) As Collection
    Dim parts As Collection
    Dim cursor As Long
    Dim delimiterPos As Long
    Dim tokenText As String

    sourceText = CStr(sourceText)
    delimiterChar = CStr(delimiterChar)
    If Len(delimiterChar) <> 1 Then
        Err.Raise vbObjectError + 1831, "ex_ResultTemplatesParser", "Delimiter must be exactly one character."
    End If

    Set parts = New Collection
    cursor = 1

    Do
        delimiterPos = mp_IndexOfDelimiterOutsideQuotedLiterals(sourceText, delimiterChar, cursor)
        If delimiterPos <= 0 Then
            parts.Add Mid$(sourceText, cursor)
            Exit Do
        End If

        tokenText = Mid$(sourceText, cursor, delimiterPos - cursor)
        parts.Add tokenText
        cursor = delimiterPos + 1

        If cursor > Len(sourceText) + 1 Then
            parts.Add vbNullString
            Exit Do
        End If
    Loop

    Set mp_SplitByDelimiterOutsideQuotedLiterals = parts
End Function

Private Function mp_IndexOfDelimiterOutsideQuotedLiterals( _
    ByVal sourceText As String, _
    ByVal delimiterChar As String, _
    Optional ByVal startPos As Long = 1 _
) As Long
    Dim i As Long
    Dim textLen As Long
    Dim ch As String
    Dim quoteChar As String

    sourceText = CStr(sourceText)
    delimiterChar = CStr(delimiterChar)
    If Len(delimiterChar) <> 1 Then
        Err.Raise vbObjectError + 1832, "ex_ResultTemplatesParser", "Delimiter must be exactly one character."
    End If

    textLen = Len(sourceText)
    If startPos < 1 Then startPos = 1
    If startPos > textLen Then Exit Function

    i = startPos
    Do While i <= textLen
        ch = Mid$(sourceText, i, 1)

        If Len(quoteChar) = 0 Then
            If ch = """" Or ch = "'" Then
                quoteChar = ch
                i = i + 1
                GoTo ContinueLoop
            End If
            If ch = delimiterChar Then
                mp_IndexOfDelimiterOutsideQuotedLiterals = i
                Exit Function
            End If
            i = i + 1
            GoTo ContinueLoop
        End If

        If ch = "\" Then
            If i < textLen Then
                i = i + 2
            Else
                i = i + 1
            End If
            GoTo ContinueLoop
        End If

        If ch = quoteChar Then
            If i < textLen Then
                If Mid$(sourceText, i + 1, 1) = quoteChar Then
                    If mp_IsQuoteTerminatorForDelimiter(sourceText, i + 1, delimiterChar) Then
                        quoteChar = vbNullString
                    End If
                    i = i + 2
                    GoTo ContinueLoop
                End If
            End If

            If mp_IsQuoteTerminatorForDelimiter(sourceText, i, delimiterChar) Then
                quoteChar = vbNullString
            End If
            i = i + 1
            GoTo ContinueLoop
        End If

        i = i + 1
ContinueLoop:
    Loop

    If Len(quoteChar) > 0 Then
        Err.Raise vbObjectError + 1833, "ex_ResultTemplatesParser", "Unclosed string literal in formatter expression: '" & sourceText & "'."
    End If
End Function

Private Function mp_IsQuoteTerminatorForDelimiter( _
    ByVal sourceText As String, _
    ByVal quotePos As Long, _
    ByVal delimiterChar As String _
) As Boolean
    Dim nextPos As Long
    Dim nextChar As String
    Dim textLen As Long

    sourceText = CStr(sourceText)
    delimiterChar = CStr(delimiterChar)
    If Len(delimiterChar) <> 1 Then Exit Function
    If quotePos < 1 Then Exit Function

    textLen = Len(sourceText)
    nextPos = quotePos + 1
    If nextPos > textLen Then
        mp_IsQuoteTerminatorForDelimiter = True
        Exit Function
    End If

    nextChar = Mid$(sourceText, nextPos, 1)
    If nextChar = delimiterChar Then
        mp_IsQuoteTerminatorForDelimiter = True
        Exit Function
    End If
    If Not mp_IsWhitespaceChar(nextChar) Then Exit Function

    nextPos = mp_FindNextNonWhitespacePosition(sourceText, nextPos + 1)
    If nextPos = 0 Then
        mp_IsQuoteTerminatorForDelimiter = True
        Exit Function
    End If

    nextChar = Mid$(sourceText, nextPos, 1)
    If nextChar = delimiterChar Then
        mp_IsQuoteTerminatorForDelimiter = True
    End If
End Function

Private Function mp_RegexReplaceText( _
    ByVal sourceValue As String, _
    ByVal regexPattern As String, _
    ByVal replacementText As String _
) As String
    Dim rx As Object

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.MultiLine = True

    On Error GoTo PatternErr
    rx.Pattern = CStr(regexPattern)
    On Error GoTo 0

    mp_RegexReplaceText = rx.Replace(CStr(sourceValue), CStr(replacementText))
    Exit Function

PatternErr:
    Err.Raise vbObjectError + 1830, "ex_ResultTemplatesParser", "Invalid regex pattern '" & CStr(regexPattern) & "' in formatter regexreplace."
End Function

Private Function mp_UnquoteFormatterArgument(ByVal argText As String) As String
    Dim normalized As String

    normalized = mp_TrimWhitespace(CStr(argText))
    If Len(normalized) >= 2 Then
        If (Left$(normalized, 1) = """" And Right$(normalized, 1) = """") Or _
           (Left$(normalized, 1) = "'" And Right$(normalized, 1) = "'") Then
            normalized = Mid$(normalized, 2, Len(normalized) - 2)
        End If
    End If

    mp_UnquoteFormatterArgument = normalized
End Function

Private Function mp_TryParseNonNegativeLong(ByVal textValue As String, ByRef outValue As Long) As Boolean
    Dim i As Long
    Dim ch As String
    Dim parsed As Double

    textValue = mp_TrimWhitespace(CStr(textValue))
    If Len(textValue) = 0 Then Exit Function

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    parsed = CDbl(textValue)
    If parsed < 0# Or parsed > 2147483647# Then Exit Function

    outValue = CLng(parsed)
    mp_TryParseNonNegativeLong = True
End Function

Private Function mp_FirstNonSpaceChar(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String

    textValue = CStr(textValue)
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If Not mp_IsWhitespaceChar(ch) Then
            mp_FirstNonSpaceChar = ch
            Exit Function
        End If
    Next i
End Function

Private Function mp_IsWhitespaceChar(ByVal ch As String) As Boolean
    Dim codePoint As Long

    If Len(ch) = 0 Then Exit Function

    codePoint = AscW(Left$(ch, 1))
    mp_IsWhitespaceChar = _
        (codePoint = 32) Or _
        (codePoint = 9) Or _
        (codePoint = 10) Or _
        (codePoint = 13) Or _
        (codePoint = NBSP_CODE_POINT) Or _
        (codePoint = NARROW_NBSP_CODE_POINT)
End Function

Private Function mp_TrimWhitespace(ByVal textValue As String) As String
    Dim startPos As Long
    Dim endPos As Long

    textValue = CStr(textValue)
    startPos = 1
    endPos = Len(textValue)

    Do While startPos <= endPos
        If Not mp_IsWhitespaceChar(Mid$(textValue, startPos, 1)) Then Exit Do
        startPos = startPos + 1
    Loop

    Do While endPos >= startPos
        If Not mp_IsWhitespaceChar(Mid$(textValue, endPos, 1)) Then Exit Do
        endPos = endPos - 1
    Loop

    If startPos > endPos Then
        mp_TrimWhitespace = vbNullString
    Else
        mp_TrimWhitespace = Mid$(textValue, startPos, endPos - startPos + 1)
    End If
End Function

Private Function mp_InflectPhraseToCase(ByVal sourceValue As String, ByVal caseName As String) As String
    Dim convertedText As String

    sourceValue = CStr(sourceValue)
    convertedText = ex_MorphUaLite.m_InflectPhraseToCase(sourceValue, caseName)
    If Len(convertedText) = 0 Then
        mp_InflectPhraseToCase = sourceValue
    Else
        mp_InflectPhraseToCase = convertedText
    End If
End Function

Private Function mp_EscapeRegex(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String
    Dim escaped As String

    escaped = vbNullString
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        Select Case ch
            Case "\", ".", "^", "$", "|", "(", ")", "[", "]", "{", "}", "*", "+", "?"
                escaped = escaped & "\" & ch
            Case Else
                escaped = escaped & ch
        End Select
    Next i

    mp_EscapeRegex = escaped
End Function

Private Function mp_ReplaceIfConditionForPlaceholder( _
    ByVal sourceText As String, _
    ByVal placeholderName As String, _
    ByVal replacementText As String _
) As String
    Dim rx As Object
    Dim rxExpr As Object
    Dim replacementCondition As String
    Dim resultText As String
    Dim updatedText As String

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "\{#if\s+" & mp_EscapeRegex(placeholderName) & "\s*\}"

    replacementCondition = IF_BLOCK_OPEN & " " & mp_BooleanTextFromValue(replacementText) & "}"
    resultText = rx.Replace(CStr(sourceText), replacementCondition)

    Set rxExpr = CreateObject("VBScript.RegExp")
    rxExpr.Global = True
    rxExpr.IgnoreCase = False
    rxExpr.Pattern = "(\{#if\s+[^}]*)\b" & mp_EscapeRegex(placeholderName) & "\b"

    Do
        updatedText = rxExpr.Replace(resultText, "$1" & CStr(replacementText))
        If StrComp(updatedText, resultText, vbBinaryCompare) = 0 Then Exit Do
        resultText = updatedText
    Loop

    mp_ReplaceIfConditionForPlaceholder = resultText
End Function

Private Function mp_BooleanTextFromValue(ByVal textValue As String) As String
    If mp_IsTruthyConditionValue(textValue) Then
        mp_BooleanTextFromValue = BOOLEAN_TRUE
    Else
        mp_BooleanTextFromValue = BOOLEAN_FALSE
    End If
End Function

Private Function mp_IsTruthyConditionValue(ByVal textValue As String) As Boolean
    Dim normalized As String
    Dim tokens As Variant
    Dim tokenCount As Long
    Dim tokenLBound As Long
    Dim tokenUBound As Long
    Dim tokenIndex As Long

    normalized = mp_NormalizeConditionWhitespace(CStr(textValue))
    If Len(normalized) = 0 Then Exit Function

    tokens = Split(normalized, " ")
    tokenCount = mp_GetArrayItemCount(tokens)
    If tokenCount <= 0 Then Exit Function

    tokenLBound = LBound(tokens)
    tokenUBound = UBound(tokens)
    tokenIndex = tokenLBound

    mp_IsTruthyConditionValue = mp_EvaluateConditionOr(tokens, tokenIndex, tokenLBound, tokenUBound, CStr(textValue))
    If tokenIndex <= tokenUBound Then
        Err.Raise vbObjectError + 1812, "ex_ResultTemplatesParser", _
            "Unsupported if-condition '" & CStr(textValue) & "' near token '" & CStr(tokens(tokenIndex)) & "'."
    End If
End Function

Private Function mp_EvaluateConditionOr( _
    ByVal tokens As Variant, _
    ByRef tokenIndex As Long, _
    ByVal tokenLBound As Long, _
    ByVal tokenUBound As Long, _
    ByVal sourceCondition As String _
) As Boolean
    Dim resultValue As Boolean
    Dim nextValue As Boolean
    Dim tokenText As String

    resultValue = mp_EvaluateConditionAnd(tokens, tokenIndex, tokenLBound, tokenUBound, sourceCondition)
    Do While tokenIndex <= tokenUBound
        tokenText = LCase$(CStr(tokens(tokenIndex)))
        If tokenText <> "#or" Then Exit Do

        tokenIndex = tokenIndex + 1
        nextValue = mp_EvaluateConditionAnd(tokens, tokenIndex, tokenLBound, tokenUBound, sourceCondition)
        resultValue = (resultValue Or nextValue)
    Loop

    mp_EvaluateConditionOr = resultValue
End Function

Private Function mp_EvaluateConditionAnd( _
    ByVal tokens As Variant, _
    ByRef tokenIndex As Long, _
    ByVal tokenLBound As Long, _
    ByVal tokenUBound As Long, _
    ByVal sourceCondition As String _
) As Boolean
    Dim resultValue As Boolean
    Dim nextValue As Boolean
    Dim tokenText As String

    resultValue = mp_EvaluateConditionUnary(tokens, tokenIndex, tokenLBound, tokenUBound, sourceCondition)
    Do While tokenIndex <= tokenUBound
        tokenText = LCase$(CStr(tokens(tokenIndex)))
        If tokenText <> "#and" Then Exit Do

        tokenIndex = tokenIndex + 1
        nextValue = mp_EvaluateConditionUnary(tokens, tokenIndex, tokenLBound, tokenUBound, sourceCondition)
        resultValue = (resultValue And nextValue)
    Loop

    mp_EvaluateConditionAnd = resultValue
End Function

Private Function mp_EvaluateConditionUnary( _
    ByVal tokens As Variant, _
    ByRef tokenIndex As Long, _
    ByVal tokenLBound As Long, _
    ByVal tokenUBound As Long, _
    ByVal sourceCondition As String _
) As Boolean
    Dim isNegated As Boolean
    Dim tokenText As String
    Dim atomStart As Long
    Dim atomText As String
    Dim atomValue As Boolean

    Do While tokenIndex <= tokenUBound
        tokenText = LCase$(CStr(tokens(tokenIndex)))
        If tokenText <> "#not" Then Exit Do
        isNegated = Not isNegated
        tokenIndex = tokenIndex + 1
    Loop

    If tokenIndex > tokenUBound Then
        atomValue = False
    Else
        tokenText = LCase$(CStr(tokens(tokenIndex)))
        If tokenText = "#and" Or tokenText = "#or" Then
            ' Missing operand is treated as False to handle empty placeholders in expressions.
            atomValue = False
        Else
            atomStart = tokenIndex
            Do While tokenIndex <= tokenUBound
                tokenText = LCase$(CStr(tokens(tokenIndex)))
                If tokenText = "#and" Or tokenText = "#or" Then Exit Do
                tokenIndex = tokenIndex + 1
            Loop

            atomText = mp_JoinConditionTokens(tokens, atomStart, tokenIndex - 1)
            atomValue = mp_EvaluateAtomicCondition(atomText, sourceCondition)
        End If
    End If

    If isNegated Then atomValue = Not atomValue
    mp_EvaluateConditionUnary = atomValue
End Function

Private Function mp_EvaluateAtomicCondition(ByVal atomicText As String, ByVal sourceCondition As String) As Boolean
    Dim normalized As String
    Dim hasNumericComparison As Boolean

    normalized = LCase$(mp_TrimWhitespace(CStr(atomicText)))
    If Len(normalized) = 0 Then Exit Function

    hasNumericComparison = mp_HasNumericComparisonOperator(normalized)
    If hasNumericComparison Then
        If Not mp_TryEvaluateNumericComparisonCondition(normalized, mp_EvaluateAtomicCondition) Then
            Err.Raise vbObjectError + 1810, "ex_ResultTemplatesParser", _
                "Unsupported numeric if-condition '" & CStr(sourceCondition) & "'. Use '<NUMBER> <OP> <NUMBER>' where OP is ==, !=, >, <, >=, <=."
        End If
        Exit Function
    End If

    If normalized = BOOLEAN_FALSE Then
        mp_EvaluateAtomicCondition = False
    Else
        mp_EvaluateAtomicCondition = True
    End If
End Function

Private Function mp_JoinConditionTokens(ByVal tokens As Variant, ByVal startIndex As Long, ByVal endIndex As Long) As String
    Dim i As Long
    Dim resultText As String

    If endIndex < startIndex Then Exit Function

    For i = startIndex To endIndex
        If Len(resultText) > 0 Then resultText = resultText & " "
        resultText = resultText & CStr(tokens(i))
    Next i

    mp_JoinConditionTokens = resultText
End Function

Private Function mp_NormalizeConditionWhitespace(ByVal conditionText As String) As String
    Dim i As Long
    Dim ch As String
    Dim resultText As String
    Dim needSeparator As Boolean

    conditionText = CStr(conditionText)
    For i = 1 To Len(conditionText)
        ch = Mid$(conditionText, i, 1)
        If mp_IsWhitespaceChar(ch) Then
            If Len(resultText) > 0 Then needSeparator = True
        Else
            If needSeparator Then
                resultText = resultText & " "
                needSeparator = False
            End If
            resultText = resultText & ch
        End If
    Next i

    mp_NormalizeConditionWhitespace = mp_TrimWhitespace(resultText)
End Function

Private Function mp_TryStripNotPrefix(ByRef conditionText As String) As Boolean
    Dim nextCh As String

    conditionText = LCase$(mp_TrimWhitespace(CStr(conditionText)))
    If Len(conditionText) < 4 Then Exit Function
    If Left$(conditionText, 4) <> "#not" Then Exit Function

    If Len(conditionText) = 4 Then
        conditionText = vbNullString
        mp_TryStripNotPrefix = True
        Exit Function
    End If

    nextCh = Mid$(conditionText, 5, 1)
    If Not mp_IsWhitespaceChar(nextCh) Then Exit Function

    conditionText = LCase$(mp_TrimWhitespace(Mid$(conditionText, 5)))
    mp_TryStripNotPrefix = True
End Function

Private Function mp_HasNumericComparisonOperator(ByVal conditionText As String) As Boolean
    conditionText = CStr(conditionText)
    If InStr(1, conditionText, "==", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, "!=", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, ">=", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, "<=", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, ">", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, "<", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
End Function

Private Function mp_TryEvaluateNumericComparisonCondition( _
    ByVal conditionText As String, _
    ByRef outResult As Boolean _
) As Boolean
    Dim rx As Object
    Dim matches As Object
    Dim leftText As String
    Dim rightText As String
    Dim operatorText As String
    Dim leftValue As Double
    Dim rightValue As Double
    Dim leftIsInteger As Boolean
    Dim rightIsInteger As Boolean

    conditionText = mp_TrimWhitespace(CStr(conditionText))
    If Len(conditionText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = "^\s*(.+?)\s*(==|!=|>=|<=|>|<)\s*(.+?)\s*$"

    Set matches = rx.Execute(conditionText)
    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    leftText = mp_TrimWhitespace(CStr(matches(0).SubMatches(0)))
    operatorText = CStr(matches(0).SubMatches(1))
    rightText = mp_TrimWhitespace(CStr(matches(0).SubMatches(2)))

    If Not mp_TryParseTemplateNumeric(leftText, leftValue, leftIsInteger) Then Exit Function
    If Not mp_TryParseTemplateNumeric(rightText, rightValue, rightIsInteger) Then Exit Function

    Select Case operatorText
        Case "=="
            outResult = (leftValue = rightValue)
        Case "!="
            outResult = (leftValue <> rightValue)
        Case ">"
            outResult = (leftValue > rightValue)
        Case "<"
            outResult = (leftValue < rightValue)
        Case ">="
            outResult = (leftValue >= rightValue)
        Case "<="
            outResult = (leftValue <= rightValue)
        Case Else
            Exit Function
    End Select

    mp_TryEvaluateNumericComparisonCondition = True
End Function

Private Function mp_ResolveTemplateLetBindings( _
    ByVal sourceText As String, _
    Optional ByVal highlightColorHex As String = vbNullString _
) As String
    Dim resultText As String
    Dim rx As Object
    Dim matches As Object
    Dim i As Long
    Dim letVarName As String
    Dim letExpression As String
    Dim letValue As String
    Dim matchStart As Long
    Dim matchLen As Long
    Dim valuesByVar As Object
    Dim key As Variant

    resultText = CStr(sourceText)

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "#let\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^;]+?)\s*;"

    Set matches = rx.Execute(resultText)
    If matches Is Nothing Then
        mp_ResolveTemplateLetBindings = resultText
        Exit Function
    End If
    If matches.Count = 0 Then
        mp_ResolveTemplateLetBindings = resultText
        Exit Function
    End If

    Set valuesByVar = CreateObject("Scripting.Dictionary")
    valuesByVar.CompareMode = 1 ' vbTextCompare

    For i = matches.Count - 1 To 0 Step -1
        letVarName = mp_TrimWhitespace(CStr(matches(i).SubMatches(0)))
        letExpression = mp_TrimWhitespace(CStr(matches(i).SubMatches(1)))
        letValue = mp_EvaluateTemplateLetExpression(letExpression)

        If valuesByVar.Exists(letVarName) Then
            Err.Raise vbObjectError + 1778, "ex_ResultTemplatesParser", "Template let variable '" & letVarName & "' is already declared."
        End If
        valuesByVar.Add letVarName, letValue

        matchStart = CLng(matches(i).FirstIndex)
        matchLen = CLng(matches(i).Length)
        resultText = Left$(resultText, matchStart) & Mid$(resultText, matchStart + matchLen + 1)
    Next i

    If InStr(1, resultText, "#let", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 1809, "ex_ResultTemplatesParser", "Invalid #let syntax. Use '#let <VAR> = <EXPR>;' format."
    End If

    resultText = mp_RemoveEmptyLetContainers(resultText)

    For Each key In valuesByVar.Keys
        resultText = m_ReplacePlaceholder(resultText, CStr(key), CStr(valuesByVar(CStr(key))), highlightColorHex)
    Next key

    mp_ResolveTemplateLetBindings = resultText
End Function

Private Function mp_RemoveEmptyLetContainers(ByVal sourceText As String) As String
    Dim resultText As String
    Dim updatedText As String
    Dim rx As Object

    resultText = CStr(sourceText)

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "\{[ \t\r\n]*\}"

    Do
        updatedText = rx.Replace(resultText, vbNullString)
        If StrComp(updatedText, resultText, vbBinaryCompare) = 0 Then Exit Do
        resultText = updatedText
    Loop

    mp_RemoveEmptyLetContainers = resultText
End Function

Private Function mp_EvaluateTemplateLetExpression(ByVal expressionText As String) As String
    Dim normalizedExpression As String
    Dim openPos As Long
    Dim closePos As Long
    Dim helperRef As String
    Dim argsText As String
    Dim args As Variant

    normalizedExpression = mp_TrimWhitespace(CStr(expressionText))
    If Right$(normalizedExpression, 1) = ";" Then
        normalizedExpression = mp_TrimWhitespace(Left$(normalizedExpression, Len(normalizedExpression) - 1))
    End If
    If Len(normalizedExpression) = 0 Then
        Err.Raise vbObjectError + 1779, "ex_ResultTemplatesParser", "Template let expression is empty."
    End If

    openPos = InStr(1, normalizedExpression, "(", vbBinaryCompare)
    closePos = InStrRev(normalizedExpression, ")", -1, vbBinaryCompare)
    If openPos <= 1 Or closePos <= openPos Then
        Err.Raise vbObjectError + 1780, "ex_ResultTemplatesParser", "Unsupported template let expression: '" & normalizedExpression & "'."
    End If

    helperRef = mp_TrimWhitespace(Left$(normalizedExpression, openPos - 1))
    argsText = Mid$(normalizedExpression, openPos + 1, closePos - openPos - 1)

    args = mp_SplitLetExpressionArgs(argsText)

    If Left$(helperRef, 1) <> "$" Then
        Err.Raise vbObjectError + 1783, "ex_ResultTemplatesParser", _
            "Template let helper must use '$<MODULE>.<METHOD>(...)' syntax: '" & normalizedExpression & "'."
    End If

    mp_EvaluateTemplateLetExpression = mp_RunExternalTemplateHelper(helperRef, args)
End Function

Private Function mp_SplitLetExpressionArgs(ByVal argsText As String) As Variant
    Dim rawParts As Collection
    Dim parts() As String
    Dim i As Long

    argsText = mp_TrimWhitespace(CStr(argsText))
    If Len(argsText) = 0 Then
        mp_SplitLetExpressionArgs = Array()
        Exit Function
    End If

    Set rawParts = mp_SplitByDelimiterOutsideQuotedLiterals(argsText, ",")
    If rawParts Is Nothing Then
        mp_SplitLetExpressionArgs = Array()
        Exit Function
    End If
    If rawParts.Count = 0 Then
        mp_SplitLetExpressionArgs = Array()
        Exit Function
    End If

    ReDim parts(0 To rawParts.Count - 1)
    For i = 1 To rawParts.Count
        parts(i - 1) = mp_TrimWhitespace(CStr(rawParts(i)))
    Next i

    mp_SplitLetExpressionArgs = parts
End Function

Private Function mp_GetArrayItemCount(ByVal values As Variant) As Long
    On Error GoTo EmptyArray
    If IsArray(values) Then
        mp_GetArrayItemCount = UBound(values) - LBound(values) + 1
    End If
    Exit Function
EmptyArray:
    mp_GetArrayItemCount = 0
End Function

Private Function mp_RunExternalTemplateHelper(ByVal helperRef As String, ByVal args As Variant) As String
    Dim methodRef As String
    Dim argCount As Long
    Dim parsedArgs() As Variant
    Dim i As Long
    Dim invokeResult As Variant
    Dim argsPreview As String

    methodRef = mp_TrimWhitespace(CStr(helperRef))
    If Left$(methodRef, 1) = "$" Then methodRef = Mid$(methodRef, 2)
    methodRef = mp_TrimWhitespace(methodRef)

    If Len(methodRef) = 0 Then
        Err.Raise vbObjectError + 1784, "ex_ResultTemplatesParser", "Template helper reference is empty."
    End If
    If InStr(1, methodRef, ".", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 1785, "ex_ResultTemplatesParser", "Template helper must use '<MODULE>.<METHOD>' syntax: '" & helperRef & "'."
    End If

    argCount = mp_GetArrayItemCount(args)
    If argCount > 5 Then
        Err.Raise vbObjectError + 1786, "ex_ResultTemplatesParser", "Template helper supports at most 5 arguments: '" & helperRef & "'."
    End If

    If argCount > 0 Then
        ReDim parsedArgs(0 To argCount - 1)
        For i = 0 To argCount - 1
            mp_ValidateTemplateHelperArgumentRaw CStr(args(i)), helperRef, i + 1
            parsedArgs(i) = mp_ParseTemplateHelperArgument(CStr(args(i)))
        Next i
    End If

    On Error GoTo InvokeErr
    Select Case argCount
        Case 0
            invokeResult = Application.Run(methodRef)
        Case 1
            invokeResult = Application.Run(methodRef, parsedArgs(0))
        Case 2
            invokeResult = Application.Run(methodRef, parsedArgs(0), parsedArgs(1))
        Case 3
            invokeResult = Application.Run(methodRef, parsedArgs(0), parsedArgs(1), parsedArgs(2))
        Case 4
            invokeResult = Application.Run(methodRef, parsedArgs(0), parsedArgs(1), parsedArgs(2), parsedArgs(3))
        Case 5
            invokeResult = Application.Run(methodRef, parsedArgs(0), parsedArgs(1), parsedArgs(2), parsedArgs(3), parsedArgs(4))
    End Select
    On Error GoTo 0

    mp_RunExternalTemplateHelper = mp_NormalizeTemplateHelperResult(invokeResult, helperRef)
    Exit Function

InvokeErr:
    argsPreview = mp_FormatTemplateHelperArgsForError(args)
    Err.Raise vbObjectError + 1834, "ex_ResultTemplatesParser", _
        "Template helper '" & helperRef & "' failed with " & CStr(argCount) & " args (" & argsPreview & "): " & _
        "[" & CStr(Err.Source) & " #" & CStr(Err.Number) & "] " & CStr(Err.Description)
End Function

Private Function mp_ParseTemplateHelperArgument(ByVal argText As String) As Variant
    Dim normalized As String
    Dim numberValue As Double
    Dim unquoted As String

    normalized = mp_TrimWhitespace(CStr(argText))
    normalized = mp_RemoveInlineHighlightMarkers(normalized)
    normalized = mp_TrimWhitespace(normalized)
    If Len(normalized) = 0 Then
        mp_ParseTemplateHelperArgument = vbNullString
        Exit Function
    End If

    If (Left$(normalized, 1) = """" And Right$(normalized, 1) = """") Or _
       (Left$(normalized, 1) = "'" And Right$(normalized, 1) = "'") Then
        unquoted = Mid$(normalized, 2, Len(normalized) - 2)
        unquoted = mp_RemoveInlineHighlightMarkers(unquoted)
        mp_ParseTemplateHelperArgument = unquoted
        Exit Function
    End If

    Select Case LCase$(normalized)
        Case "true"
            mp_ParseTemplateHelperArgument = True
            Exit Function
        Case "false"
            mp_ParseTemplateHelperArgument = False
            Exit Function
    End Select

    If IsNumeric(normalized) Then
        numberValue = CDbl(normalized)
        If Fix(numberValue) = numberValue Then
            mp_ParseTemplateHelperArgument = CLng(numberValue)
        Else
            mp_ParseTemplateHelperArgument = numberValue
        End If
        Exit Function
    End If

    mp_ParseTemplateHelperArgument = normalized
End Function

Private Function mp_FormatTemplateHelperArgsForError(ByVal args As Variant) As String
    Dim argCount As Long
    Dim i As Long
    Dim parts() As String

    argCount = mp_GetArrayItemCount(args)
    If argCount <= 0 Then Exit Function

    ReDim parts(0 To argCount - 1)
    For i = 0 To argCount - 1
        parts(i) = """" & Replace(CStr(args(i)), """", """""") & """"
    Next i

    mp_FormatTemplateHelperArgsForError = Join(parts, ", ")
End Function

Private Sub mp_ValidateTemplateHelperArgumentRaw( _
    ByVal rawArgText As String, _
    ByVal helperRef As String, _
    ByVal argIndex As Long _
)
    Dim normalized As String
    Dim checkText As String
    Dim placeholderToken As String

    normalized = mp_TrimWhitespace(CStr(rawArgText))
    If Len(normalized) = 0 Then Exit Sub

    ' Escaped braces are allowed: \{ and \}
    checkText = Replace(normalized, "\{", vbNullString)
    checkText = Replace(checkText, "\}", vbNullString)

    placeholderToken = mp_FindFirstPlaceholderLikeToken(checkText)
    If Len(placeholderToken) > 0 Then
        Err.Raise vbObjectError + 1810, "ex_ResultTemplatesParser", _
            "Template helper '" & helperRef & "' argument #" & CStr(argIndex) & _
            " contains unresolved placeholder '" & placeholderToken & "'. " & _
            "Ensure placeholders are resolved before '#let $<MODULE>.<METHOD>(...)'."
    End If

    If InStr(1, checkText, "{", vbBinaryCompare) > 0 Or InStr(1, checkText, "}", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 1811, "ex_ResultTemplatesParser", _
            "Template helper '" & helperRef & "' argument #" & CStr(argIndex) & _
            " contains unescaped '{' or '}'. Use '\{' and '\}' for literal braces."
    End If
End Sub

Private Function mp_FindFirstPlaceholderLikeToken(ByVal sourceText As String) As String
    Dim rx As Object
    Dim matches As Object

    sourceText = CStr(sourceText)
    If Len(sourceText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = "\{[A-Za-z_#][^{}]*\}"

    Set matches = rx.Execute(sourceText)
    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    mp_FindFirstPlaceholderLikeToken = CStr(matches(0).Value)
End Function

Private Function mp_NormalizeTemplateHelperResult(ByVal resultValue As Variant, ByVal helperRef As String) As String
    If IsObject(resultValue) Then
        Err.Raise vbObjectError + 1787, "ex_ResultTemplatesParser", "Template helper '" & helperRef & "' returned object result, string/number/boolean expected."
    End If
    If IsError(resultValue) Then
        Err.Raise vbObjectError + 1788, "ex_ResultTemplatesParser", "Template helper '" & helperRef & "' returned error value."
    End If

    If VarType(resultValue) = vbBoolean Then
        mp_NormalizeTemplateHelperResult = mp_BooleanTextFromValue(CStr(resultValue))
        Exit Function
    End If

    If IsNull(resultValue) Then
        mp_NormalizeTemplateHelperResult = vbNullString
    Else
        mp_NormalizeTemplateHelperResult = CStr(resultValue)
    End If
End Function

Private Function mp_ResolveForBlocks( _
    ByVal sourceText As String, _
    Optional ByVal highlightColorHex As String = vbNullString _
) As String
    Dim resultText As String
    Dim forOpenPos As Long
    Dim headerEndPos As Long
    Dim closeStartPos As Long
    Dim closeEndPos As Long
    Dim itemVarName As String
    Dim collectionName As String
    Dim innerTemplateText As String
    Dim replacementText As String
    Dim collectionItems As Collection

    resultText = CStr(sourceText)

    Do
        forOpenPos = InStr(1, resultText, FOR_BLOCK_OPEN, vbTextCompare)
        If forOpenPos = 0 Then Exit Do

        If Not mp_TryParseForHeader(resultText, forOpenPos, headerEndPos, itemVarName, collectionName) Then
            Err.Raise vbObjectError + 1843, "ex_ResultTemplatesParser", "Invalid for-block syntax near position " & CStr(forOpenPos) & ". Use '{#for item in Collection}'."
        End If

        If Not mp_TryFindMatchingForClose(resultText, headerEndPos + 1, closeStartPos, closeEndPos) Then
            Err.Raise vbObjectError + 1844, "ex_ResultTemplatesParser", "Missing closing {#endfor} for for-block near position " & CStr(forOpenPos) & "."
        End If

        innerTemplateText = Mid$(resultText, headerEndPos + 1, closeStartPos - headerEndPos - 1)
        Set collectionItems = mp_GetTemplateCollectionItems(collectionName)
        replacementText = mp_RenderForBlock(innerTemplateText, itemVarName, collectionItems, highlightColorHex)

        resultText = Left$(resultText, forOpenPos - 1) & replacementText & Mid$(resultText, closeEndPos + 1)
    Loop

    If InStr(1, resultText, FOR_BLOCK_CLOSE, vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 1845, "ex_ResultTemplatesParser", "Unexpected {#endfor} without matching {#for ...}."
    End If

    mp_ResolveForBlocks = resultText
End Function

Private Function mp_RenderForBlock( _
    ByVal innerTemplateText As String, _
    ByVal itemVarName As String, _
    ByVal collectionItems As Collection, _
    ByVal highlightColorHex As String _
) As String
    Dim i As Long
    Dim itemValues As Object
    Dim key As Variant
    Dim iterationText As String
    Dim itemPrefix As String
    Dim loopFlag As String

    itemPrefix = CStr(itemVarName) & "."

    If collectionItems Is Nothing Then Exit Function
    If collectionItems.Count = 0 Then Exit Function

    For i = 1 To collectionItems.Count
        iterationText = CStr(innerTemplateText)

        Set itemValues = collectionItems(i)
        If Not itemValues Is Nothing Then
            For Each key In itemValues.Keys
                iterationText = m_ReplacePlaceholder( _
                    iterationText, _
                    itemPrefix & CStr(key), _
                    CStr(itemValues(CStr(key))), _
                    highlightColorHex _
                )
            Next key
        End If

        iterationText = m_ReplacePlaceholder(iterationText, itemPrefix & "__index", CStr(i - 1), highlightColorHex)
        loopFlag = BOOLEAN_FALSE
        If i = 1 Then loopFlag = BOOLEAN_TRUE
        iterationText = m_ReplacePlaceholder(iterationText, itemPrefix & "__first", loopFlag, highlightColorHex)
        loopFlag = BOOLEAN_FALSE
        If i = collectionItems.Count Then loopFlag = BOOLEAN_TRUE
        iterationText = m_ReplacePlaceholder(iterationText, itemPrefix & "__last", loopFlag, highlightColorHex)

        iterationText = mp_ResolveForBlocks(iterationText, highlightColorHex)
        mp_RenderForBlock = mp_RenderForBlock & iterationText
    Next i
End Function

Private Function mp_TryParseForHeader( _
    ByVal sourceText As String, _
    ByVal forOpenPos As Long, _
    ByRef outHeaderEndPos As Long, _
    ByRef outItemVarName As String, _
    ByRef outCollectionName As String _
) As Boolean
    Dim closePos As Long
    Dim rawHeader As String
    Dim rx As Object
    Dim matches As Object

    If StrComp(Mid$(sourceText, forOpenPos, Len(FOR_BLOCK_OPEN)), FOR_BLOCK_OPEN, vbTextCompare) <> 0 Then Exit Function

    closePos = InStr(forOpenPos + Len(FOR_BLOCK_OPEN), sourceText, "}", vbBinaryCompare)
    If closePos = 0 Then Exit Function

    rawHeader = Mid$(sourceText, forOpenPos + Len(FOR_BLOCK_OPEN), closePos - forOpenPos - Len(FOR_BLOCK_OPEN))
    rawHeader = mp_TrimWhitespace(rawHeader)
    If Len(rawHeader) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.MultiLine = False
    rx.Pattern = "^([A-Za-z_][A-Za-z0-9_]*)\s+in\s+([A-Za-z_][A-Za-z0-9_]*)$"

    Set matches = rx.Execute(rawHeader)
    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    outItemVarName = CStr(matches(0).SubMatches(0))
    outCollectionName = CStr(matches(0).SubMatches(1))
    outHeaderEndPos = closePos
    mp_TryParseForHeader = True
End Function

Private Function mp_TryFindMatchingForClose( _
    ByVal sourceText As String, _
    ByVal searchFromPos As Long, _
    ByRef outCloseStartPos As Long, _
    ByRef outCloseEndPos As Long _
) As Boolean
    Dim depth As Long
    Dim nextOpenPos As Long
    Dim nextClosePos As Long

    depth = 1
    Do While searchFromPos <= Len(sourceText)
        nextOpenPos = InStr(searchFromPos, sourceText, FOR_BLOCK_OPEN, vbTextCompare)
        nextClosePos = InStr(searchFromPos, sourceText, FOR_BLOCK_CLOSE, vbTextCompare)

        If nextClosePos = 0 Then Exit Function

        If nextOpenPos > 0 And nextOpenPos < nextClosePos Then
            depth = depth + 1
            searchFromPos = nextOpenPos + Len(FOR_BLOCK_OPEN)
        Else
            depth = depth - 1
            If depth = 0 Then
                outCloseStartPos = nextClosePos
                outCloseEndPos = nextClosePos + Len(FOR_BLOCK_CLOSE) - 1
                mp_TryFindMatchingForClose = True
                Exit Function
            End If
            searchFromPos = nextClosePos + Len(FOR_BLOCK_CLOSE)
        End If
    Loop
End Function

Private Function mp_ResolveConditionalBlocks(ByVal sourceText As String) As String
    Dim resultText As String
    Dim openPos As Long
    Dim headerEndPos As Long
    Dim closeStartPos As Long
    Dim closeEndPos As Long
    Dim conditionText As String
    Dim innerText As String
    Dim replacementText As String

    resultText = CStr(sourceText)

    Do
        openPos = InStr(1, resultText, IF_BLOCK_OPEN, vbTextCompare)
        If openPos = 0 Then Exit Do

        If Not mp_TryParseIfHeader(resultText, openPos, headerEndPos, conditionText) Then
            Err.Raise vbObjectError + 1767, "ex_ResultTemplatesParser", "Invalid if-block syntax near position " & CStr(openPos) & "."
        End If

        If Not mp_TryFindMatchingIfClose(resultText, headerEndPos + 1, closeStartPos, closeEndPos) Then
            Err.Raise vbObjectError + 1768, "ex_ResultTemplatesParser", "Missing closing {#endif} for if-block near position " & CStr(openPos) & "."
        End If

        innerText = Mid$(resultText, headerEndPos + 1, closeStartPos - headerEndPos - 1)
        innerText = mp_ResolveConditionalBlocks(innerText)

        If mp_IsTruthyConditionValue(conditionText) Then
            replacementText = innerText
        Else
            replacementText = vbNullString
        End If

        resultText = Left$(resultText, openPos - 1) & replacementText & Mid$(resultText, closeEndPos + 1)
    Loop

    If InStr(1, resultText, IF_BLOCK_CLOSE, vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 1769, "ex_ResultTemplatesParser", "Unexpected {#endif} without matching {#if ...}."
    End If

    mp_ResolveConditionalBlocks = resultText
End Function

Private Function mp_ResolveJoinLineTokens(ByVal sourceText As String) As String
    Dim resultText As String

    resultText = CStr(sourceText)
    resultText = Replace(resultText, vbCrLf, vbLf)
    resultText = Replace(resultText, vbCr, vbLf)

    resultText = mp_ResolveJoinLineToken(resultText, RESERVED_JOINLINE_TOKEN)
    resultText = mp_ResolveJoinLineToken(resultText, RESERVED_JOINLINE_TOKEN_SHORT)
    resultText = mp_ResolveJoinLineToken(resultText, RESERVED_JOINLINE_TOKEN_LEGACY)
    mp_ResolveJoinLineTokens = resultText
End Function

Private Function mp_ResolveTrimIndentTokens(ByVal sourceText As String) As String
    Dim resultText As String

    resultText = CStr(sourceText)
    resultText = mp_ResolveTrimIndentToken(resultText, RESERVED_TRIMINDENT_TOKEN_SHORT)
    mp_ResolveTrimIndentTokens = resultText
End Function

Private Function mp_ResolveNewLineTokens(ByVal sourceText As String) As String
    Dim resultText As String

    resultText = CStr(sourceText)
    resultText = Replace(resultText, RESERVED_NEWLINE_TOKEN, vbLf)
    resultText = Replace(resultText, RESERVED_NEWLINE_TOKEN_SHORT, vbLf)
    mp_ResolveNewLineTokens = resultText
End Function

Private Function mp_ResolveTrimIndentToken(ByVal sourceText As String, ByVal tokenText As String) As String
    Dim resultText As String
    Dim rx As Object
    Dim tokenPattern As String

    resultText = CStr(sourceText)
    tokenPattern = mp_EscapeRegex(CStr(tokenText))

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False

    ' Remove token and all horizontal whitespace after it.
    rx.Pattern = tokenPattern & "[ \t]*"
    resultText = rx.Replace(resultText, vbNullString)

    mp_ResolveTrimIndentToken = resultText
End Function

Private Function mp_ResolveJoinLineToken(ByVal sourceText As String, ByVal tokenText As String) As String
    Dim resultText As String
    Dim rx As Object
    Dim tokenPattern As String

    resultText = CStr(sourceText)
    tokenPattern = mp_EscapeRegex(CStr(tokenText))

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False

    ' Join with next line: "...TOKEN\n    ..."
    rx.Pattern = tokenPattern & "[ \t]*" & vbLf & "[ \t]*"
    resultText = rx.Replace(resultText, vbNullString)

    ' Join with previous line: "...\n    TOKEN..."
    rx.Pattern = "[ \t]*" & vbLf & "[ \t]*" & tokenPattern
    resultText = rx.Replace(resultText, vbNullString)

    ' Fallback: strip standalone token if no line-break pattern was matched.
    resultText = Replace(resultText, CStr(tokenText), vbNullString)
    mp_ResolveJoinLineToken = resultText
End Function

Private Function mp_TryParseIfHeader( _
    ByVal sourceText As String, _
    ByVal ifOpenPos As Long, _
    ByRef outHeaderEndPos As Long, _
    ByRef outConditionText As String _
) As Boolean
    Dim closePos As Long
    Dim rawCondition As String

    If StrComp(Mid$(sourceText, ifOpenPos, Len(IF_BLOCK_OPEN)), IF_BLOCK_OPEN, vbTextCompare) <> 0 Then Exit Function

    closePos = InStr(ifOpenPos + Len(IF_BLOCK_OPEN), sourceText, "}", vbBinaryCompare)
    If closePos = 0 Then Exit Function

    rawCondition = Mid$(sourceText, ifOpenPos + Len(IF_BLOCK_OPEN), closePos - ifOpenPos - Len(IF_BLOCK_OPEN))
    outConditionText = mp_TrimWhitespace(rawCondition)
    If Len(outConditionText) = 0 Then Exit Function

    outHeaderEndPos = closePos
    mp_TryParseIfHeader = True
End Function

Private Function mp_TryFindMatchingIfClose( _
    ByVal sourceText As String, _
    ByVal searchFromPos As Long, _
    ByRef outCloseStartPos As Long, _
    ByRef outCloseEndPos As Long _
) As Boolean
    Dim depth As Long
    Dim nextOpenPos As Long
    Dim nextClosePos As Long

    depth = 1
    Do While searchFromPos <= Len(sourceText)
        nextOpenPos = InStr(searchFromPos, sourceText, IF_BLOCK_OPEN, vbTextCompare)
        nextClosePos = InStr(searchFromPos, sourceText, IF_BLOCK_CLOSE, vbTextCompare)

        If nextClosePos = 0 Then Exit Function

        If nextOpenPos > 0 And nextOpenPos < nextClosePos Then
            depth = depth + 1
            searchFromPos = nextOpenPos + Len(IF_BLOCK_OPEN)
        Else
            depth = depth - 1
            If depth = 0 Then
                outCloseStartPos = nextClosePos
                outCloseEndPos = nextClosePos + Len(IF_BLOCK_CLOSE) - 1
                mp_TryFindMatchingIfClose = True
                Exit Function
            End If
            searchFromPos = nextClosePos + Len(IF_BLOCK_CLOSE)
        End If
    Loop
End Function

Private Function mp_ResolveDateExpressions( _
    ByVal sourceText As String, _
    Optional ByVal baseDateText As String = vbNullString _
) As String
    Dim rx As Object
    Dim matches As Object
    Dim tokenStart As Long
    Dim tokenLen As Long
    Dim tokenOffsetText As String
    Dim prefixText As String
    Dim valueStart As Long
    Dim valueLen As Long
    Dim valueText As String
    Dim baseNumber As Double
    Dim baseIsInteger As Boolean
    Dim offsetValue As Double
    Dim resultNumber As Double
    Dim resolvedValue As String
    Dim leftPart As String
    Dim rightPart As String
    Dim resultText As String
    Dim hasHighlight As Boolean
    Dim highlightColorHex As String

    resultText = CStr(sourceText)
    baseDateText = CStr(baseDateText) ' reserved for signature compatibility.

    mp_EnsureLegacyDayTokensNotUsed resultText

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = NUMERIC_OFFSET_TOKEN_PATTERN

    Do
        Set matches = rx.Execute(resultText)
        If matches Is Nothing Then Exit Do
        If matches.Count = 0 Then Exit Do

        tokenStart = CLng(matches(0).FirstIndex)
        tokenLen = CLng(matches(0).Length)
        tokenOffsetText = CStr(matches(0).SubMatches(0))
        prefixText = Left$(resultText, tokenStart)

        If Not mp_TryFindImmediateLeftValueSpan(prefixText, valueStart, valueLen, valueText) Then
            Err.Raise vbObjectError + 1764, "ex_ResultTemplatesParser", _
                "Numeric offset token '{" & tokenOffsetText & "}' has no left value."
        End If
        hasHighlight = mp_TryExtractInlineHighlightColor(valueText, highlightColorHex)

        If Not mp_TryParseTemplateNumeric(valueText, baseNumber, baseIsInteger) Then
            Err.Raise vbObjectError + 1765, "ex_ResultTemplatesParser", _
                "Left value '" & valueText & "' before token '{" & tokenOffsetText & "}' is not numeric."
        End If
        If Not mp_TryParseSignedInteger(tokenOffsetText, offsetValue) Then
            Err.Raise vbObjectError + 1766, "ex_ResultTemplatesParser", _
                "Invalid numeric offset '{" & tokenOffsetText & "}'. Use '{+N}' or '{-N}'."
        End If

        resultNumber = baseNumber + offsetValue
        If baseIsInteger And Fix(resultNumber) = resultNumber Then
            resolvedValue = mp_FormatIntegerWithBasePadding(CLng(resultNumber), valueText)
        Else
            resolvedValue = CStr(resultNumber)
        End If
        If hasHighlight Then
            resolvedValue = mp_WrapHighlightMarkers(resolvedValue, highlightColorHex)
        End If

        leftPart = Left$(resultText, valueStart - 1)
        rightPart = Mid$(resultText, tokenStart + tokenLen + 1)
        resultText = leftPart & resolvedValue & rightPart
    Loop

    mp_ResolveDateExpressions = resultText
End Function

Private Function mp_NormalizeTemplateText(ByVal templateText As String) As String
    templateText = Replace(templateText, vbCrLf, vbLf)
    templateText = Replace(templateText, vbCr, vbLf)
    If Left$(templateText, 1) = vbLf Then templateText = Mid$(templateText, 2)
    If Right$(templateText, 1) = vbLf Then templateText = Left$(templateText, Len(templateText) - 1)
    mp_NormalizeTemplateText = templateText
End Function

Private Function mp_PrependTemplateError(ByVal sourceText As String, ByVal operationName As String) As String
    Dim errorLine As String
    Dim fullText As String

    errorLine = TEMPLATE_ERROR_PREFIX & " " & operationName & ": [" & CStr(Err.Source) & " #" & CStr(Err.Number) & "] " & CStr(Err.Description)
    fullText = CStr(sourceText)

    If Len(fullText) = 0 Then
        mp_PrependTemplateError = errorLine
        Exit Function
    End If

    If StrComp(Left$(fullText, Len(errorLine)), errorLine, vbBinaryCompare) = 0 Then
        mp_PrependTemplateError = fullText
        Exit Function
    End If

    mp_PrependTemplateError = errorLine & vbLf & fullText
End Function

Private Sub mp_EnsureLegacyDayTokensNotUsed(ByVal sourceText As String)
    Dim rx As Object
    Dim matches As Object

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = LEGACY_DAY_TOKEN_PATTERN

    Set matches = rx.Execute(CStr(sourceText))
    If matches Is Nothing Then Exit Sub
    If matches.Count = 0 Then Exit Sub

    Err.Raise vbObjectError + 1763, "ex_ResultTemplatesParser", _
        "Legacy token '" & CStr(matches(0).Value) & "' is not supported. Use '<NUMERIC>{+N}' or '<NUMERIC>{-N}'."
End Sub

Private Function mp_TryFindImmediateLeftValueSpan( _
    ByVal textValue As String, _
    ByRef outStartPos As Long, _
    ByRef outLength As Long, _
    ByRef outValueText As String _
) As Boolean
    Dim endPos As Long
    Dim startPos As Long

    textValue = CStr(textValue)
    endPos = Len(textValue)

    Do While endPos > 0
        If Not mp_IsWhitespaceChar(Mid$(textValue, endPos, 1)) Then Exit Do
        endPos = endPos - 1
    Loop
    If endPos <= 0 Then Exit Function

    startPos = endPos
    Do While startPos > 1
        If mp_IsWhitespaceChar(Mid$(textValue, startPos - 1, 1)) Then Exit Do
        startPos = startPos - 1
    Loop

    outStartPos = startPos
    outLength = endPos - startPos + 1
    outValueText = Mid$(textValue, outStartPos, outLength)
    mp_TryFindImmediateLeftValueSpan = (outLength > 0)
End Function

Private Function mp_TryParseTemplateNumeric(ByVal numberText As String, ByRef outValue As Double, ByRef outIsInteger As Boolean) As Boolean
    Dim rx As Object
    Dim normalized As String

    numberText = mp_RemoveInlineHighlightMarkers(CStr(numberText))
    numberText = mp_TrimWhitespace(CStr(numberText))
    If Len(numberText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = "^[+-]?\d+(?:[.,]\d+)?$"

    If Not rx.Test(numberText) Then Exit Function

    normalized = Replace(numberText, ",", ".")
    outValue = Val(normalized)
    outIsInteger = (InStr(1, normalized, ".", vbBinaryCompare) = 0)
    mp_TryParseTemplateNumeric = True
End Function

Private Function mp_RemoveInlineHighlightMarkers(ByVal sourceText As String) As String
    Dim rxStart As Object
    Dim resultText As String

    resultText = CStr(sourceText)
    If Len(resultText) = 0 Then
        mp_RemoveInlineHighlightMarkers = vbNullString
        Exit Function
    End If

    Set rxStart = CreateObject("VBScript.RegExp")
    rxStart.Global = True
    rxStart.IgnoreCase = False
    rxStart.Pattern = mp_EscapeRegex(HIGHLIGHT_MARKER_START) & "(\([^)]*\))?"
    resultText = rxStart.Replace(resultText, vbNullString)
    resultText = Replace(resultText, HIGHLIGHT_MARKER_END, vbNullString)

    mp_RemoveInlineHighlightMarkers = resultText
End Function

Private Function mp_TryExtractInlineHighlightColor(ByVal sourceText As String, ByRef outColorHex As String) As Boolean
    Dim rx As Object
    Dim matches As Object

    outColorHex = vbNullString
    sourceText = CStr(sourceText)
    If Len(sourceText) = 0 Then Exit Function
    If InStr(1, sourceText, HIGHLIGHT_MARKER_START, vbBinaryCompare) = 0 Then Exit Function
    If InStr(1, sourceText, HIGHLIGHT_MARKER_END, vbBinaryCompare) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = mp_EscapeRegex(HIGHLIGHT_MARKER_START) & "(?:\(([^)]*)\))?"

    Set matches = rx.Execute(sourceText)
    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    outColorHex = mp_TrimWhitespace(CStr(matches(0).SubMatches(0)))
    If Len(outColorHex) = 0 Then outColorHex = DEFAULT_HIGHLIGHT_COLOR_HEX
    mp_TryExtractInlineHighlightColor = True
End Function

Private Function mp_TryParseSignedInteger(ByVal numberText As String, ByRef outValue As Double) As Boolean
    Dim rx As Object

    numberText = mp_TrimWhitespace(CStr(numberText))
    If Len(numberText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = "^[+-]\d+$"

    If Not rx.Test(numberText) Then Exit Function

    outValue = CDbl(numberText)
    mp_TryParseSignedInteger = True
End Function

Private Function mp_FormatIntegerWithBasePadding(ByVal resultValue As Long, ByVal baseNumberText As String) As String
    Dim signText As String
    Dim absResult As Long
    Dim baseDigits As String
    Dim width As Long
    Dim paddedText As String
    Dim shouldPad As Boolean

    signText = vbNullString
    If resultValue < 0 Then signText = "-"

    absResult = resultValue
    If absResult < 0 Then absResult = -absResult

    baseDigits = mp_TrimWhitespace(CStr(baseNumberText))
    If Left$(baseDigits, 1) = "+" Or Left$(baseDigits, 1) = "-" Then
        baseDigits = Mid$(baseDigits, 2)
    End If

    width = Len(baseDigits)
    shouldPad = (width > 1 And Left$(baseDigits, 1) = "0")
    If shouldPad Then
        paddedText = Format$(absResult, String$(width, "0"))
    Else
        paddedText = CStr(absResult)
    End If

    mp_FormatIntegerWithBasePadding = signText & paddedText
End Function
