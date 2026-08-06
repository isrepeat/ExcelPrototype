Attribute VB_Name = "ex_Helpers"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:ex_Helpers.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function fn_NormalizeText(ByVal valueText As String) As String
    valueText = VBA.Replace(VBA.CStr(valueText), VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    fn_NormalizeText = VBA.LCase$(VBA.Trim$(valueText))
End Function


Public Function m_RegexIsMatch( _
    ByVal textValue As String, _
    ByVal regexPattern As String _
) As Boolean
    Dim rx As Object

    Set rx = private_CreateRegex(regexPattern)
    m_RegexIsMatch = rx.Test(VBA.CStr(textValue))
End Function


Public Function m_RegexFirstMatch( _
    ByVal textValue As String, _
    ByVal regexPattern As String _
) As String
    Dim rx As Object
    Dim matches As Object

    Set rx = private_CreateRegex(regexPattern)
    Set matches = rx.Execute(VBA.CStr(textValue))
    If matches.Count > 0 Then m_RegexFirstMatch = VBA.CStr(matches(0).Value)
End Function


Public Function m_RegexGetGroup( _
    ByVal textValue As String, _
    ByVal regexPattern As String, _
    Optional ByVal groupIndex As Long = 1 _
) As String
    Dim rx As Object
    Dim matches As Object
    Dim firstMatch As Object

    If groupIndex < 0 Then
        Err.Raise VBA.vbObjectError + 1811, "ex_Helpers", "Regex group index cannot be negative."
    End If

    Set rx = private_CreateRegex(regexPattern)
    Set matches = rx.Execute(VBA.CStr(textValue))
    If matches.Count = 0 Then Exit Function

    Set firstMatch = matches(0)
    If groupIndex = 0 Then
        m_RegexGetGroup = VBA.CStr(firstMatch.Value)
        Exit Function
    End If

    If groupIndex > firstMatch.SubMatches.Count Then Exit Function
    m_RegexGetGroup = VBA.CStr(firstMatch.SubMatches(groupIndex - 1))
End Function


Public Function m_RegexReplace( _
    ByVal textValue As String, _
    ByVal regexPattern As String, _
    ByVal replacementText As String _
) As String
    Dim rx As Object

    Set rx = private_CreateRegex(regexPattern, True)
    m_RegexReplace = rx.Replace(VBA.CStr(textValue), VBA.CStr(replacementText))
End Function


Public Function fn_TryResolveDateWithContext( _
    ByVal rawDateValue As Variant, _
    ByVal contextDateValue As Variant, _
    ByRef outDateValue As Date _
) As Boolean
    Dim valueText As String
    Dim normalizedText As String
    Dim parts() As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim contextDate As Date
    Dim serialValue As Double

    valueText = VBA.Trim$(VBA.CStr(rawDateValue))
    If VBA.Len(valueText) = 0 Then Exit Function

    If VBA.VarType(rawDateValue) = VBA.vbDate Then
        outDateValue = VBA.CDate(rawDateValue)
        fn_TryResolveDateWithContext = True
        Exit Function
    End If

    normalizedText = VBA.Replace(valueText, "-", ".")
    normalizedText = VBA.Replace(normalizedText, "/", ".")
    normalizedText = VBA.Replace(normalizedText, " ", VBA.vbNullString)
    If VBA.Len(normalizedText) = 0 Then Exit Function

    parts = VBA.Split(normalizedText, ".")
    If UBound(parts) = 0 And VBA.IsNumeric(normalizedText) Then
        serialValue = VBA.CDbl(normalizedText)
        If serialValue > 31 Then
            outDateValue = VBA.DateSerial(1899, 12, 30) + serialValue
            fn_TryResolveDateWithContext = True
            Exit Function
        End If
    End If

    If UBound(parts) = 2 Then
        dayValue = VBA.CLng(VBA.Val(parts(0)))
        monthValue = VBA.CLng(VBA.Val(parts(1)))
        yearValue = private_NormalizeYear(VBA.CLng(VBA.Val(parts(2))))
        fn_TryResolveDateWithContext = private_TryBuildDate(dayValue, monthValue, yearValue, outDateValue)
        Exit Function
    End If

    If Not fn_TryResolveDateWithContext(contextDateValue, VBA.Date, contextDate) Then contextDate = VBA.Date

    If UBound(parts) = 1 Then
        dayValue = VBA.CLng(VBA.Val(parts(0)))
        monthValue = VBA.CLng(VBA.Val(parts(1)))
        yearValue = VBA.Year(contextDate)
        fn_TryResolveDateWithContext = private_TryBuildDate(dayValue, monthValue, yearValue, outDateValue)
        Exit Function
    End If

    If UBound(parts) = 0 Then
        dayValue = VBA.CLng(VBA.Val(parts(0)))
        monthValue = VBA.Month(contextDate)
        yearValue = VBA.Year(contextDate)
        fn_TryResolveDateWithContext = private_TryBuildDate(dayValue, monthValue, yearValue, outDateValue)
    End If
End Function


Public Function fn_IsShortDateValue(ByVal rawDateValue As Variant) As Boolean
    Dim valueText As String
    Dim normalizedText As String
    Dim parts() As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim serialValue As Double

    If VBA.VarType(rawDateValue) = VBA.vbDate Then Exit Function

    valueText = VBA.Trim$(VBA.CStr(rawDateValue))
    If VBA.Len(valueText) = 0 Then Exit Function

    normalizedText = VBA.Replace(valueText, "-", ".")
    normalizedText = VBA.Replace(normalizedText, "/", ".")
    normalizedText = VBA.Replace(normalizedText, " ", VBA.vbNullString)
    If VBA.Len(normalizedText) = 0 Then Exit Function

    parts = VBA.Split(normalizedText, ".")
    If UBound(parts) = 0 And VBA.IsNumeric(normalizedText) Then
        serialValue = VBA.CDbl(normalizedText)
        fn_IsShortDateValue = (serialValue >= 1 And serialValue <= 31)
        Exit Function
    End If

    If UBound(parts) = 1 Then
        If Not VBA.IsNumeric(parts(0)) Or Not VBA.IsNumeric(parts(1)) Then Exit Function
        dayValue = VBA.CLng(VBA.Val(parts(0)))
        monthValue = VBA.CLng(VBA.Val(parts(1)))
        fn_IsShortDateValue = (dayValue >= 1 And dayValue <= 31 And monthValue >= 1 And monthValue <= 12)
    End If
End Function


Public Function fn_FormatUaDateLong(ByVal dateValue As Date) As String
    fn_FormatUaDateLong = VBA.Format$(dateValue, "dd") & VBA.ChrW$(160) & _
        private_GetUaMonthGenitiveName(VBA.Month(dateValue)) & " " & _
        VBA.CStr(VBA.Year(dateValue)) & VBA.ChrW$(160) & "року"
End Function


Public Function fn_FormatUaDatePattern(ByVal dateValue As Date, ByVal formatPattern As String) As String
    Dim resultText As String

    resultText = VBA.CStr(formatPattern)
    If VBA.Len(resultText) = 0 Then
        fn_FormatUaDatePattern = fn_FormatUaDateLong(dateValue)
        Exit Function
    End If

    ' DSL dateformat uses escaped tokens, for example:
    '   \dd \month \yyyy року -> 01 березня 2026 року
    '   \dd \month            -> 01 березня
    resultText = VBA.Replace(resultText, "\month", private_GetUaMonthGenitiveName(VBA.Month(dateValue)))
    resultText = VBA.Replace(resultText, "\yyyy", VBA.Format$(dateValue, "yyyy"))
    resultText = VBA.Replace(resultText, "\yy", VBA.Format$(dateValue, "yy"))
    resultText = VBA.Replace(resultText, "\dd", VBA.Format$(dateValue, "dd"))
    resultText = VBA.Replace(resultText, "\d", VBA.CStr(VBA.Day(dateValue)))
    resultText = VBA.Replace(resultText, "\mm", VBA.Format$(dateValue, "mm"))
    resultText = VBA.Replace(resultText, "\m", VBA.CStr(VBA.Month(dateValue)))

    fn_FormatUaDatePattern = resultText
End Function


Public Function fn_EscapeXmlAttr(ByVal valueText As String) As String
    valueText = VBA.Replace$(valueText, "&", "&amp;")
    valueText = VBA.Replace$(valueText, "<", "&lt;")
    valueText = VBA.Replace$(valueText, ">", "&gt;")
    valueText = VBA.Replace$(valueText, """", "&quot;")
    valueText = VBA.Replace$(valueText, "'", "&apos;")
    fn_EscapeXmlAttr = valueText
End Function


Public Function fn_ReadSnapshotLongAttr(ByVal sourceNode As Object, ByVal attrName As String, ByVal defaultValue As Long) As Long
    Dim rawText As String

    If sourceNode Is Nothing Then
        fn_ReadSnapshotLongAttr = defaultValue
        Exit Function
    End If

    rawText = VBA.Trim$(VBA.CStr(sourceNode.getAttribute(attrName)))
    If VBA.Len(rawText) = 0 Then
        fn_ReadSnapshotLongAttr = defaultValue
        Exit Function
    End If
    If Not VBA.IsNumeric(rawText) Then
        fn_ReadSnapshotLongAttr = defaultValue
        Exit Function
    End If

    fn_ReadSnapshotLongAttr = VBA.CLng(rawText)
End Function


Public Function fn_ReadSnapshotDoubleAttr(ByVal sourceNode As Object, ByVal attrName As String, ByVal defaultValue As Double) As Double
    Dim rawText As String

    If sourceNode Is Nothing Then
        fn_ReadSnapshotDoubleAttr = defaultValue
        Exit Function
    End If

    rawText = VBA.Trim$(VBA.CStr(sourceNode.getAttribute(attrName)))
    If VBA.Len(rawText) = 0 Then
        fn_ReadSnapshotDoubleAttr = defaultValue
        Exit Function
    End If
    If Not private_TryParseFlexibleDouble(rawText, fn_ReadSnapshotDoubleAttr) Then
        fn_ReadSnapshotDoubleAttr = defaultValue
    End If
End Function


Public Function fn_ReadSnapshotBooleanAttr(ByVal sourceNode As Object, ByVal attrName As String, ByVal defaultValue As Boolean) As Boolean
    Dim rawText As String

    If sourceNode Is Nothing Then
        fn_ReadSnapshotBooleanAttr = defaultValue
        Exit Function
    End If

    rawText = VBA.Trim$(VBA.CStr(sourceNode.getAttribute(attrName)))
    If VBA.Len(rawText) = 0 Then
        fn_ReadSnapshotBooleanAttr = defaultValue
        Exit Function
    End If

    If Not fn_TryGetBooleanFromVariant(rawText, fn_ReadSnapshotBooleanAttr) Then
        fn_ReadSnapshotBooleanAttr = defaultValue
    End If
End Function


Public Function fn_TryGetBooleanFromVariant(ByVal valueCandidate As Variant, ByRef outValue As Boolean) As Boolean
    Dim textValue As String

    outValue = False

    If VBA.IsObject(valueCandidate) Then
        outValue = Not valueCandidate Is Nothing
        fn_TryGetBooleanFromVariant = True
        Exit Function
    End If

    If VBA.IsError(valueCandidate) Then Exit Function
    If VBA.IsNull(valueCandidate) Then Exit Function
    If VBA.IsEmpty(valueCandidate) Then Exit Function

    Select Case VBA.VarType(valueCandidate)
        Case VBA.vbBoolean
            outValue = VBA.CBool(valueCandidate)
            fn_TryGetBooleanFromVariant = True
            Exit Function
        Case VBA.vbByte, VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
            outValue = (VBA.CDbl(valueCandidate) <> 0)
            fn_TryGetBooleanFromVariant = True
            Exit Function
    End Select

    textValue = VBA.LCase$(VBA.Trim$(VBA.CStr(valueCandidate)))
    Select Case textValue
        Case "1", "true", "yes", "y", "on"
            outValue = True
            fn_TryGetBooleanFromVariant = True
        Case "0", "false", "no", "n", "off"
            outValue = False
            fn_TryGetBooleanFromVariant = True
    End Select
End Function


Public Function fn_GetSnapshotRawValueText(ByVal rawItems As Collection, ByVal idx As Long, ByVal fallbackText As String) As String
    Dim rawObject As Object
    Dim valueCandidate As Variant

    fn_GetSnapshotRawValueText = VBA.CStr(fallbackText)
    If rawItems Is Nothing Then Exit Function
    If idx <= 0 Or idx > rawItems.Count Then Exit Function

    Set rawObject = Nothing
    On Error Resume Next
    Set rawObject = rawItems(idx)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    If rawObject Is Nothing Then
        On Error Resume Next
        valueCandidate = rawItems(idx)
        If Err.Number = 0 Then
            If Not VBA.IsObject(valueCandidate) Then fn_GetSnapshotRawValueText = VBA.CStr(valueCandidate)
        Else
            Err.Clear
        End If
        On Error GoTo 0
        Exit Function
    End If

    If VBA.LCase$(VBA.TypeName(rawObject)) = "dictionary" Then
        If rawObject.Exists("RawValue") Then
            fn_GetSnapshotRawValueText = VBA.CStr(rawObject("RawValue"))
            Exit Function
        End If
        If rawObject.Exists("Id") Then
            fn_GetSnapshotRawValueText = VBA.CStr(rawObject("Id"))
            Exit Function
        End If
    End If

    On Error Resume Next
    valueCandidate = VBA.CallByName(rawObject, "RawValue", VbGet)
    If Err.Number = 0 Then
        If Not VBA.IsObject(valueCandidate) Then
            fn_GetSnapshotRawValueText = VBA.CStr(valueCandidate)
            On Error GoTo 0
            Exit Function
        End If
    Else
        Err.Clear
    End If

    valueCandidate = VBA.CallByName(rawObject, "Id", VbGet)
    If Err.Number = 0 Then
        If Not VBA.IsObject(valueCandidate) Then fn_GetSnapshotRawValueText = VBA.CStr(valueCandidate)
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function fn_TextStartsWith(ByVal sourceText As String, ByVal prefixText As String) As Boolean
    If VBA.Len(prefixText) = 0 Then Exit Function
    If VBA.Len(sourceText) < VBA.Len(prefixText) Then Exit Function
    fn_TextStartsWith = (VBA.StrComp(VBA.Left$(sourceText, VBA.Len(prefixText)), prefixText, VBA.vbBinaryCompare) = 0)
End Function

Public Function fn_IsStringEquals( _
    ByVal leftText As String, _
    ByVal rightText As String, _
    Optional ByVal compareMethod As VbCompareMethod = VBA.vbTextCompare _
) As Boolean
    fn_IsStringEquals = (VBA.StrComp(VBA.CStr(leftText), VBA.CStr(rightText), compareMethod) = 0)
End Function

Public Function fn_CreateDictionaryTextCompare() As Object
    Dim dict As Object

    ' Универсальный регистронезависимый словарь.
    Set dict = VBA.CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Set fn_CreateDictionaryTextCompare = dict
End Function


Public Sub fn_QuickSortLongArray(ByRef values() As Long, ByVal leftIndex As Long, ByVal rightIndex As Long)
    Dim i As Long
    Dim j As Long
    Dim pivotValue As Long
    Dim tempValue As Long

    i = leftIndex
    j = rightIndex
    pivotValue = values((leftIndex + rightIndex) \ 2)

    Do While i <= j
        Do While values(i) < pivotValue
            i = i + 1
        Loop

        Do While values(j) > pivotValue
            j = j - 1
        Loop

        If i <= j Then
            tempValue = values(i)
            values(i) = values(j)
            values(j) = tempValue
            i = i + 1
            j = j - 1
        End If
    Loop

    If leftIndex < j Then fn_QuickSortLongArray values, leftIndex, j
    If i < rightIndex Then fn_QuickSortLongArray values, i, rightIndex
End Sub

' //
' // Internal
' //
Private Function private_TryParseFlexibleDouble(ByVal rawText As String, ByRef outValue As Double) As Boolean
    Dim normalized As String
    Dim decimalSep As String

    rawText = VBA.Trim$(rawText)
    If VBA.Len(rawText) = 0 Then Exit Function

    decimalSep = VBA.CStr(Application.International(xlDecimalSeparator))
    normalized = rawText

    If decimalSep = "," Then
        normalized = VBA.Replace$(normalized, ".", ",")
    Else
        normalized = VBA.Replace$(normalized, ",", ".")
    End If

    If Not VBA.IsNumeric(normalized) Then Exit Function
    outValue = VBA.CDbl(normalized)
    private_TryParseFlexibleDouble = True
End Function


Private Function private_CreateRegex(ByVal regexPattern As String, Optional ByVal globalMatch As Boolean = False) As Object
    Dim rx As Object

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = globalMatch
    rx.IgnoreCase = False
    rx.MultiLine = True
    rx.Pattern = VBA.CStr(regexPattern)
    Set private_CreateRegex = rx
End Function


Private Function private_NormalizeYear(ByVal yearValue As Long) As Long
    If yearValue >= 100 Then
        private_NormalizeYear = yearValue
    ElseIf yearValue >= 0 Then
        private_NormalizeYear = 2000 + yearValue
    End If
End Function


Private Function private_TryBuildDate( _
    ByVal dayValue As Long, _
    ByVal monthValue As Long, _
    ByVal yearValue As Long, _
    ByRef outDateValue As Date _
) As Boolean
    Dim candidateDate As Date

    If yearValue < 1900 Then Exit Function
    If monthValue < 1 Or monthValue > 12 Then Exit Function
    If dayValue < 1 Or dayValue > 31 Then Exit Function

    On Error Resume Next
    candidateDate = VBA.DateSerial(yearValue, monthValue, dayValue)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If VBA.Day(candidateDate) <> dayValue Then Exit Function
    If VBA.Month(candidateDate) <> monthValue Then Exit Function
    If VBA.Year(candidateDate) <> yearValue Then Exit Function

    outDateValue = candidateDate
    private_TryBuildDate = True
End Function


Private Function private_GetUaMonthGenitiveName(ByVal monthValue As Long) As String
    Select Case monthValue
        Case 1: private_GetUaMonthGenitiveName = "січня"
        Case 2: private_GetUaMonthGenitiveName = "лютого"
        Case 3: private_GetUaMonthGenitiveName = "березня"
        Case 4: private_GetUaMonthGenitiveName = "квітня"
        Case 5: private_GetUaMonthGenitiveName = "травня"
        Case 6: private_GetUaMonthGenitiveName = "червня"
        Case 7: private_GetUaMonthGenitiveName = "липня"
        Case 8: private_GetUaMonthGenitiveName = "серпня"
        Case 9: private_GetUaMonthGenitiveName = "вересня"
        Case 10: private_GetUaMonthGenitiveName = "жовтня"
        Case 11: private_GetUaMonthGenitiveName = "листопада"
        Case 12: private_GetUaMonthGenitiveName = "грудня"
    End Select
End Function
