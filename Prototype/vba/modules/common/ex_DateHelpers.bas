Attribute VB_Name = "ex_DateHelpers"
Option Explicit

Public Function m_ToFullDate( _
    ByVal sourceDateText As String, _
    Optional ByVal baseDateText As String = vbNullString _
) As String
    Dim normalized As String
    Dim baseDate As Date
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim parts() As String
    Dim dt As Date

    normalized = Trim$(CStr(sourceDateText))
    If Len(normalized) = 0 Then
        Err.Raise vbObjectError + 1788, "ex_DateHelpers", "ToFullDate: source date is empty."
    End If

    baseDate = mp_ParseBaseDate(baseDateText)

    normalized = Replace(normalized, "-", ".")
    normalized = Replace(normalized, "/", ".")

    If InStr(1, normalized, ".", vbBinaryCompare) = 0 Then
        If Not mp_TryParsePositiveInteger(normalized, dayValue) Then
            Err.Raise vbObjectError + 1789, "ex_DateHelpers", "ToFullDate: invalid day value '" & sourceDateText & "'."
        End If
        monthValue = Month(baseDate)
        yearValue = Year(baseDate)
    Else
        parts = Split(normalized, ".")
        Select Case UBound(parts) - LBound(parts) + 1
            Case 2
                If Not mp_TryParsePositiveInteger(parts(0), dayValue) Then
                    Err.Raise vbObjectError + 1790, "ex_DateHelpers", "ToFullDate: invalid day value '" & sourceDateText & "'."
                End If
                If Not mp_TryParsePositiveInteger(parts(1), monthValue) Then
                    Err.Raise vbObjectError + 1791, "ex_DateHelpers", "ToFullDate: invalid month value '" & sourceDateText & "'."
                End If
                yearValue = Year(baseDate)

            Case 3
                If Not mp_TryParsePositiveInteger(parts(0), dayValue) Then
                    Err.Raise vbObjectError + 1792, "ex_DateHelpers", "ToFullDate: invalid day value '" & sourceDateText & "'."
                End If
                If Not mp_TryParsePositiveInteger(parts(1), monthValue) Then
                    Err.Raise vbObjectError + 1793, "ex_DateHelpers", "ToFullDate: invalid month value '" & sourceDateText & "'."
                End If
                If Not mp_TryParsePositiveInteger(parts(2), yearValue) Then
                    Err.Raise vbObjectError + 1794, "ex_DateHelpers", "ToFullDate: invalid year value '" & sourceDateText & "'."
                End If
                If Len(Trim$(parts(2))) = 2 Then yearValue = 2000 + yearValue

            Case Else
                Err.Raise vbObjectError + 1795, "ex_DateHelpers", "ToFullDate: unsupported date format '" & sourceDateText & "'."
        End Select
    End If

    If dayValue < 1 Or dayValue > 31 Then
        Err.Raise vbObjectError + 1796, "ex_DateHelpers", "ToFullDate: day is out of range '" & sourceDateText & "'."
    End If
    If monthValue < 1 Or monthValue > 12 Then
        Err.Raise vbObjectError + 1797, "ex_DateHelpers", "ToFullDate: month is out of range '" & sourceDateText & "'."
    End If
    If yearValue < 1900 Or yearValue > 9999 Then
        Err.Raise vbObjectError + 1798, "ex_DateHelpers", "ToFullDate: year is out of range '" & sourceDateText & "'."
    End If

    On Error GoTo InvalidDate
    dt = DateSerial(yearValue, monthValue, dayValue)
    If Day(dt) <> dayValue Or Month(dt) <> monthValue Or Year(dt) <> yearValue Then GoTo InvalidDate
    m_ToFullDate = Format$(dt, "dd.mm.yyyy")
    Exit Function

InvalidDate:
    Err.Raise vbObjectError + 1799, "ex_DateHelpers", "ToFullDate: invalid calendar date '" & sourceDateText & "'."
End Function

Public Function m_FormatDateDay(ByVal sourceDateText As String) As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim hasYear As Boolean

    If Not mp_TryParseDayMonthYear(sourceDateText, dayValue, monthValue, yearValue, hasYear) Then
        m_FormatDateDay = Trim$(CStr(sourceDateText))
        Exit Function
    End If

    m_FormatDateDay = Format$(dayValue, "00")
End Function

Public Function m_FormatDateDayWithMonth(ByVal sourceDateText As String) As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim hasYear As Boolean

    If Not mp_TryParseDayMonthYear(sourceDateText, dayValue, monthValue, yearValue, hasYear) Then
        m_FormatDateDayWithMonth = Trim$(CStr(sourceDateText))
        Exit Function
    End If

    m_FormatDateDayWithMonth = Format$(dayValue, "00") & " " & mp_GetUaMonthGenitiveName(monthValue)
End Function

Public Function m_FormatDateByPattern(ByVal sourceDateText As String, ByVal patternText As String) As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim hasYear As Boolean
    Dim normalizedDate As String
    Dim resultText As String
    Dim rx As Object
    Dim matches As Object
    Dim i As Long
    Dim tokenValue As String
    Dim tokenReplacement As String
    Dim tokenStart As Long
    Dim tokenLen As Long

    normalizedDate = Trim$(CStr(sourceDateText))
    If Len(normalizedDate) = 0 Then
        Err.Raise vbObjectError + 1801, "ex_DateHelpers", "FormatDateByPattern: source date is empty."
    End If

    If Not mp_TryParseDayMonthYear(normalizedDate, dayValue, monthValue, yearValue, hasYear) Or Not hasYear Then
        On Error GoTo DateNormalizationError
        normalizedDate = m_ToFullDate(normalizedDate)
        On Error GoTo DateParseError
        If Not mp_TryParseDayMonthYear(normalizedDate, dayValue, monthValue, yearValue, hasYear) Or Not hasYear Then GoTo DateParseError
    End If
    On Error GoTo 0

    resultText = CStr(patternText)
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "\\(dd|mm|yyyy|month)"
    Set matches = rx.Execute(resultText)

    If matches Is Nothing Then
        Err.Raise vbObjectError + 1802, "ex_DateHelpers", "FormatDateByPattern: failed to parse format '" & patternText & "'."
    End If
    If matches.Count = 0 Then
        Err.Raise vbObjectError + 1803, "ex_DateHelpers", "FormatDateByPattern: format must contain at least one token (\dd, \mm, \month, \yyyy): '" & patternText & "'."
    End If

    For i = matches.Count - 1 To 0 Step -1
        tokenValue = LCase$(CStr(matches(i).Value))
        Select Case tokenValue
            Case "\dd"
                tokenReplacement = Format$(dayValue, "00")
            Case "\mm"
                tokenReplacement = Format$(monthValue, "00")
            Case "\yyyy"
                tokenReplacement = CStr(yearValue)
            Case "\month"
                tokenReplacement = mp_GetUaMonthGenitiveName(monthValue)
            Case Else
                Err.Raise vbObjectError + 1804, "ex_DateHelpers", "FormatDateByPattern: unsupported token '" & tokenValue & "'."
        End Select

        tokenStart = CLng(matches(i).FirstIndex)
        tokenLen = CLng(matches(i).Length)
        resultText = Left$(resultText, tokenStart) & tokenReplacement & Mid$(resultText, tokenStart + tokenLen + 1)
    Next i

    m_FormatDateByPattern = resultText
    Exit Function

DateNormalizationError:
    Err.Raise vbObjectError + 1805, "ex_DateHelpers", "FormatDateByPattern: failed to normalize date '" & CStr(sourceDateText) & "': " & CStr(Err.Description)

DateParseError:
    Err.Raise vbObjectError + 1806, "ex_DateHelpers", "FormatDateByPattern: invalid date value '" & CStr(sourceDateText) & "'."
End Function

Public Function m_IsSameMonth(ByVal leftDateText As String, ByVal rightDateText As String) As Boolean
    Dim leftMonth As Long
    Dim rightMonth As Long
    Dim normalizedLeft As String
    Dim normalizedRight As String

    normalizedLeft = Trim$(CStr(leftDateText))
    normalizedRight = Trim$(CStr(rightDateText))

    If mp_IsUnresolvedTemplateToken(normalizedLeft) Then
        Err.Raise vbObjectError + 1810, "ex_DateHelpers", "IsSameMonth: first date contains unresolved placeholder '" & normalizedLeft & "'. Expected concrete date (dd.mm or dd.mm.yyyy). Check DateFrom/OutDate placeholder replacement in post-process script."
    End If
    If mp_IsUnresolvedTemplateToken(normalizedRight) Then
        Err.Raise vbObjectError + 1811, "ex_DateHelpers", "IsSameMonth: second date contains unresolved placeholder '" & normalizedRight & "'. Expected concrete date (dd.mm or dd.mm.yyyy). Check DateTo/ReturnDate placeholder replacement in post-process script."
    End If

    If Len(normalizedLeft) = 0 Or normalizedLeft = "?" Then
        Err.Raise vbObjectError + 1812, "ex_DateHelpers", "IsSameMonth: first date is empty or '?'. Expected concrete date (dd.mm or dd.mm.yyyy)."
    End If
    If Len(normalizedRight) = 0 Or normalizedRight = "?" Then
        Err.Raise vbObjectError + 1813, "ex_DateHelpers", "IsSameMonth: second date is empty or '?'. Expected concrete date (dd.mm or dd.mm.yyyy)."
    End If

    If Not mp_TryParseMonth(normalizedLeft, leftMonth) Then
        Err.Raise vbObjectError + 1786, "ex_DateHelpers", "IsSameMonth: invalid first date '" & normalizedLeft & "'. Expected format dd.mm or dd.mm.yyyy."
    End If
    If Not mp_TryParseMonth(normalizedRight, rightMonth) Then
        Err.Raise vbObjectError + 1787, "ex_DateHelpers", "IsSameMonth: invalid second date '" & normalizedRight & "'. Expected format dd.mm or dd.mm.yyyy."
    End If

    m_IsSameMonth = (leftMonth = rightMonth)
End Function

Public Function m_FormatCalendarDaysUa(ByVal dayCountText As String) As String
    Dim dayCount As Long
    Dim normalized As String

    normalized = Trim$(CStr(dayCountText))
    If Len(normalized) = 0 Then
        Err.Raise vbObjectError + 1808, "ex_DateHelpers", "FormatCalendarDaysUa: day count is empty."
    End If
    If Not ex_XmlCore.m_TryParseLong(normalized, dayCount) Then
        Err.Raise vbObjectError + 1809, "ex_DateHelpers", "FormatCalendarDaysUa: invalid integer day count '" & normalized & "'."
    End If

    m_FormatCalendarDaysUa = CStr(dayCount) & " " & mp_GetUaCalendarDaysPhrase(dayCount)
End Function

Private Function mp_TryParseMonth(ByVal sourceDateText As String, ByRef outMonth As Long) As Boolean
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim hasYear As Boolean

    If Not mp_TryParseDayMonthYear(sourceDateText, dayValue, monthValue, yearValue, hasYear) Then Exit Function

    outMonth = monthValue
    mp_TryParseMonth = True
End Function

Private Function mp_ParseBaseDate(ByVal baseDateText As String) As Date
    Dim normalized As String

    normalized = Trim$(CStr(baseDateText))
    If Len(normalized) = 0 Then
        mp_ParseBaseDate = Date
        Exit Function
    End If

    On Error GoTo ParseErr
    mp_ParseBaseDate = CDate(normalized)
    Exit Function

ParseErr:
    Err.Raise vbObjectError + 1800, "ex_DateHelpers", "ToFullDate: invalid base date '" & baseDateText & "'."
End Function

Private Function mp_TryParsePositiveInteger(ByVal textValue As String, ByRef outValue As Long) As Boolean
    Dim i As Long
    Dim ch As String
    Dim normalized As String
    Dim parsed As Double

    normalized = Trim$(CStr(textValue))
    If Len(normalized) = 0 Then Exit Function

    For i = 1 To Len(normalized)
        ch = Mid$(normalized, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    parsed = CDbl(normalized)
    If parsed < 0# Or parsed > 2147483647# Then Exit Function

    outValue = CLng(parsed)
    mp_TryParsePositiveInteger = True
End Function

Private Function mp_TryParseDayMonthYear( _
    ByVal sourceDateText As String, _
    ByRef outDay As Long, _
    ByRef outMonth As Long, _
    ByRef outYear As Long, _
    ByRef outHasYear As Boolean _
) As Boolean
    Dim rx As Object
    Dim matches As Object
    Dim yearText As String
    Dim parsedYear As Long

    sourceDateText = Trim$(CStr(sourceDateText))
    If Len(sourceDateText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = "^(\d{1,2})[.\-/](\d{1,2})(?:[.\-/](\d{2,4}))?$"

    Set matches = rx.Execute(sourceDateText)
    If matches Is Nothing Then Exit Function
    If matches.Count <> 1 Then Exit Function

    outDay = CLng(matches(0).SubMatches(0))
    outMonth = CLng(matches(0).SubMatches(1))

    If outDay < 1 Or outDay > 31 Then Exit Function
    If outMonth < 1 Or outMonth > 12 Then Exit Function

    yearText = Trim$(CStr(matches(0).SubMatches(2)))
    If Len(yearText) > 0 Then
        parsedYear = CLng(yearText)
        If Len(yearText) = 2 Then parsedYear = 2000 + parsedYear
        If parsedYear < 1900 Or parsedYear > 9999 Then Exit Function
        outYear = parsedYear
        outHasYear = True
    Else
        outYear = 0
        outHasYear = False
    End If

    mp_TryParseDayMonthYear = True
End Function

Private Function mp_GetUaMonthGenitiveName(ByVal monthNumber As Long) As String
    Select Case monthNumber
        Case 1: mp_GetUaMonthGenitiveName = "січня"
        Case 2: mp_GetUaMonthGenitiveName = "лютого"
        Case 3: mp_GetUaMonthGenitiveName = "березня"
        Case 4: mp_GetUaMonthGenitiveName = "квітня"
        Case 5: mp_GetUaMonthGenitiveName = "травня"
        Case 6: mp_GetUaMonthGenitiveName = "червня"
        Case 7: mp_GetUaMonthGenitiveName = "липня"
        Case 8: mp_GetUaMonthGenitiveName = "серпня"
        Case 9: mp_GetUaMonthGenitiveName = "вересня"
        Case 10: mp_GetUaMonthGenitiveName = "жовтня"
        Case 11: mp_GetUaMonthGenitiveName = "листопада"
        Case 12: mp_GetUaMonthGenitiveName = "грудня"
        Case Else
            Err.Raise vbObjectError + 1807, "ex_DateHelpers", "Invalid month number '" & CStr(monthNumber) & "' for date formatting."
    End Select
End Function

Private Function mp_GetUaCalendarDaysPhrase(ByVal dayCount As Long) As String
    Dim absDayCount As Long
    Dim mod10 As Long
    Dim mod100 As Long

    absDayCount = Abs(dayCount)
    mod10 = absDayCount Mod 10
    mod100 = absDayCount Mod 100

    If mod100 >= 11 And mod100 <= 14 Then
        mp_GetUaCalendarDaysPhrase = "календарних днів"
        Exit Function
    End If

    Select Case mod10
        Case 1
            mp_GetUaCalendarDaysPhrase = "календарний день"
        Case 2
            mp_GetUaCalendarDaysPhrase = "календарних дня"
        Case 3
            mp_GetUaCalendarDaysPhrase = "календарних дні"
        Case 4
            mp_GetUaCalendarDaysPhrase = "календарних дня"
        Case Else
            mp_GetUaCalendarDaysPhrase = "календарних днів"
    End Select
End Function

Private Function mp_IsUnresolvedTemplateToken(ByVal sourceText As String) As Boolean
    Dim normalized As String

    normalized = Trim$(CStr(sourceText))
    If Len(normalized) < 3 Then Exit Function

    mp_IsUnresolvedTemplateToken = _
        (Left$(normalized, 1) = "{") And (Right$(normalized, 1) = "}")
End Function
