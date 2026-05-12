Attribute VB_Name = "ex_HelpersCSS"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_HelpersCSS.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function fn_TryReadStyleValue( _
    ByVal styleText As String, _
    ByVal keyName As String, _
    ByRef outValue As String _
) As Boolean
    Dim normalized As String
    Dim parts As Variant
    Dim part As Variant
    Dim pairText As String
    Dim sepPos As Long
    Dim keyText As String

    normalized = VBA.Replace$(styleText, VBA.vbCr, VBA.vbNullString)
    normalized = VBA.Replace$(normalized, VBA.vbLf, VBA.vbNullString)
    normalized = VBA.Replace$(normalized, "{", VBA.vbNullString)
    normalized = VBA.Replace$(normalized, "}", VBA.vbNullString)

    parts = VBA.Split(normalized, ";")
    For Each part In parts
        pairText = VBA.Trim$(VBA.CStr(part))
        If VBA.Len(pairText) = 0 Then GoTo ContinuePart

        sepPos = VBA.InStr(1, pairText, ":", VBA.vbBinaryCompare)
        If sepPos <= 1 Then
            sepPos = VBA.InStr(1, pairText, "=", VBA.vbBinaryCompare)
        End If
        If sepPos <= 1 Or sepPos >= VBA.Len(pairText) Then GoTo ContinuePart

        keyText = VBA.LCase$(VBA.Trim$(VBA.Left$(pairText, sepPos - 1)))
        If VBA.StrComp(keyText, VBA.LCase$(VBA.Trim$(keyName)), VBA.vbBinaryCompare) <> 0 Then GoTo ContinuePart

        outValue = VBA.Trim$(VBA.Mid$(pairText, sepPos + 1))
        If VBA.Len(outValue) = 0 Then Exit Function

        fn_TryReadStyleValue = True
        Exit Function

ContinuePart:
    Next part
End Function


Public Function fn_TryParsePositiveDouble(ByVal valueText As String, ByRef outValue As Double) As Boolean
    Dim normalized As String
    Dim decimalSep As String

    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    If VBA.IsNumeric(valueText) Then
        outValue = VBA.CDbl(valueText)
        If outValue <= 0# Then Exit Function
        fn_TryParsePositiveDouble = True
        Exit Function
    End If

    ' Locale-safe parse for values like "0.75" on systems with decimal separator ",".
    decimalSep = Application.DecimalSeparator
    If VBA.Len(decimalSep) = 0 Then decimalSep = "."

    normalized = VBA.Replace$(valueText, ".", decimalSep)
    normalized = VBA.Replace$(normalized, ",", decimalSep)

    If Not VBA.IsNumeric(normalized) Then Exit Function
    outValue = VBA.CDbl(normalized)
    If outValue <= 0# Then Exit Function
    fn_TryParsePositiveDouble = True
End Function


Public Function fn_TryParseColor(ByVal valueText As String, ByRef outColor As Long) As Boolean
    Dim r As Long
    Dim g As Long
    Dim b As Long

    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    If VBA.Left$(valueText, 1) = "#" Then
        If VBA.Len(valueText) <> 7 Then Exit Function
        If Not private_IsHexPair(VBA.Mid$(valueText, 2, 2)) Then Exit Function
        If Not private_IsHexPair(VBA.Mid$(valueText, 4, 2)) Then Exit Function
        If Not private_IsHexPair(VBA.Mid$(valueText, 6, 2)) Then Exit Function

        r = VBA.CLng("&H" & VBA.Mid$(valueText, 2, 2))
        g = VBA.CLng("&H" & VBA.Mid$(valueText, 4, 2))
        b = VBA.CLng("&H" & VBA.Mid$(valueText, 6, 2))
        outColor = VBA.RGB(r, g, b)
        fn_TryParseColor = True
        Exit Function
    End If

    If VBA.IsNumeric(valueText) Then
        outColor = VBA.CLng(valueText)
        fn_TryParseColor = True
    End If
End Function


Public Function fn_TryParseShapeBorderWeight(ByVal valueText As String, ByRef outValue As Double) As Boolean
    valueText = VBA.LCase$(VBA.Trim$(valueText))

    Select Case valueText
        Case "hairline"
            outValue = 0.25
            fn_TryParseShapeBorderWeight = True
            Exit Function
        Case "thin"
            outValue = 0.75
            fn_TryParseShapeBorderWeight = True
            Exit Function
        Case "medium"
            outValue = 1.5
            fn_TryParseShapeBorderWeight = True
            Exit Function
        Case "thick"
            outValue = 2.25
            fn_TryParseShapeBorderWeight = True
            Exit Function
    End Select

    fn_TryParseShapeBorderWeight = fn_TryParsePositiveDouble(valueText, outValue)
End Function


Public Function fn_TryParseCellBorderWeight(ByVal valueText As String, ByRef outValue As Variant) As Boolean
    Dim numericValue As Double

    valueText = VBA.LCase$(VBA.Trim$(valueText))

    Select Case valueText
        Case "hairline"
            outValue = xlHairline
            fn_TryParseCellBorderWeight = True
            Exit Function
        Case "thin"
            outValue = xlThin
            fn_TryParseCellBorderWeight = True
            Exit Function
        Case "medium"
            outValue = xlMedium
            fn_TryParseCellBorderWeight = True
            Exit Function
        Case "thick"
            outValue = xlThick
            fn_TryParseCellBorderWeight = True
            Exit Function
    End Select

    If Not fn_TryParsePositiveDouble(valueText, numericValue) Then Exit Function

    outValue = numericValue
    fn_TryParseCellBorderWeight = True
End Function

' //
' // Internal
' //
Private Function private_IsHexPair(ByVal pairText As String) As Boolean
    Dim value As Long

    On Error GoTo EH
    If VBA.Len(pairText) <> 2 Then Exit Function

    value = VBA.CLng("&H" & pairText)
    If value < 0 Or value > 255 Then Exit Function

    private_IsHexPair = True
    Exit Function

EH:
    private_IsHexPair = False
End Function
