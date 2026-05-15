Attribute VB_Name = "ex_Helpers"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_Helpers.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
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

    rawText = VBA.LCase$(VBA.Trim$(VBA.CStr(sourceNode.getAttribute(attrName))))
    If VBA.Len(rawText) = 0 Then
        fn_ReadSnapshotBooleanAttr = defaultValue
        Exit Function
    End If

    Select Case rawText
        Case "true", "1", "yes"
            fn_ReadSnapshotBooleanAttr = True
        Case "false", "0", "no"
            fn_ReadSnapshotBooleanAttr = False
        Case Else
            fn_ReadSnapshotBooleanAttr = defaultValue
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
