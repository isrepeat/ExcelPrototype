Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
' Формирует имя документа по контексту и создаёт Word-файл из шаблона.
Public Function ex_TryGenerateWordDocument( _
    ByVal templatePath As String, _
    ByVal documentNamePattern As String, _
    ByVal documentNameValues As Object, _
    ByVal placeholderNames As Variant, _
    ByVal placeholderValues As Variant, _
    ByRef outDocumentPath As String _
) As Boolean
    Dim documentName As String

    outDocumentPath = VBA.vbNullString
    If documentNameValues Is Nothing Then
        ex_Helpers.LogError "Document name context is not initialized"
        VBA.MsgBox "Document name context is not initialized.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not ex_Helpers.private_Text_TryFormat( _
        documentNamePattern, documentNameValues, documentName) Then Exit Function
    ex_Helpers.LogDebug "Generated document name: " & documentName
    ex_TryGenerateWordDocument = ex_Helpers.private_Word_TryGenerateDocument( _
        templatePath, documentName, placeholderNames, placeholderValues, _
        outDocumentPath)
End Function

' Читает обязательное поле по alias из карты адресов формы.
Public Function ex_TryReadRequired( _
    ByVal inputSheetName As String, _
    ByVal inputCellMap As Object, _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String, _
    ByRef outValue As String _
) As Boolean
    Dim sourceSheet As Worksheet
    Dim cellAddress As String

    If Not private_TryGetInputCellAddress(inputCellMap, fieldAlias, cellAddress) Then Exit Function
    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Worksheets(inputSheetName)
    On Error GoTo 0
    If sourceSheet Is Nothing Then
        ex_Helpers.LogError "Input sheet was not found: " & inputSheetName
        VBA.MsgBox "Input sheet was not found: " & inputSheetName, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outValue = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(outValue) = 0 Then
        VBA.MsgBox "Enter " & fieldCaption & " in cell " & cellAddress & ".", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    ex_TryReadRequired = True
End Function

' Читает целое неотрицательное число; пустое optional-поле означает ноль.
Public Function ex_TryReadNonNegativeDays( _
    ByVal inputSheetName As String, _
    ByVal inputCellMap As Object, _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String, _
    ByVal isOptional As Boolean, _
    ByRef outDays As Long _
) As Boolean
    Dim daysText As String

    If isOptional Then
        If Not private_TryReadOptional(inputSheetName, inputCellMap, fieldAlias, daysText) Then Exit Function
        If VBA.Len(daysText) = 0 Then
            outDays = 0
            ex_TryReadNonNegativeDays = True
            Exit Function
        End If
    Else
        If Not ex_TryReadRequired( _
            inputSheetName, inputCellMap, fieldAlias, fieldCaption, daysText) Then Exit Function
    End If
    If Not ex_Helpers.private_Text_IsDigits(daysText) Then
        VBA.MsgBox "Enter a non-negative whole number for " & fieldCaption & ".", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    On Error GoTo EH
    outDays = VBA.CLng(daysText)
    ex_TryReadNonNegativeDays = True
    Exit Function
EH:
    VBA.MsgBox "The value for " & fieldCaption & " is out of range.", _
        VBA.vbExclamation, "Document Generation"
End Function

' Читает optional-поле по alias; пустое значение возвращается как пустая строка.
Public Function ex_TryReadOptional( _
    ByVal inputSheetName As String, _
    ByVal inputCellMap As Object, _
    ByVal fieldAlias As String, _
    ByRef outValue As String _
) As Boolean
    ex_TryReadOptional = private_TryReadOptional( _
        inputSheetName, inputCellMap, fieldAlias, outValue)
End Function

' Логирует книгу, листы и фактические привязки конфигурационной формы.
Public Sub ex_LogWorkbookContext( _
    ByVal inputSheetName As String, _
    ByVal inputCellMap As Object _
)
    Dim worksheetIndex As Long
    Dim worksheetObj As Worksheet
    Dim activeSheetText As String
    Dim fieldAlias As Variant
    Dim cellAddress As String
    Dim bindingText As String
    Dim valuesText As String

    On Error Resume Next
    activeSheetText = Application.ActiveSheet.Name
    On Error GoTo 0
    ex_Helpers.LogDebug "Workbook path: " & ThisWorkbook.FullName
    ex_Helpers.LogDebug "Worksheet count: " & VBA.CStr(ThisWorkbook.Worksheets.Count)
    ex_Helpers.LogDebug "Active sheet Unicode: " & _
        ex_Helpers.private_Text_ToUnicodeDebug(activeSheetText)
    For worksheetIndex = 1 To ThisWorkbook.Worksheets.Count
        Set worksheetObj = ThisWorkbook.Worksheets(worksheetIndex)
        ex_Helpers.LogDebug "Worksheet | Index=" & VBA.CStr(worksheetIndex) & _
            " | CodeName=" & worksheetObj.CodeName & " | NameUnicode=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(worksheetObj.Name)
    Next worksheetIndex

    Set worksheetObj = Nothing
    On Error Resume Next
    Set worksheetObj = ThisWorkbook.Worksheets(inputSheetName)
    On Error GoTo 0
    If worksheetObj Is Nothing Then
        ex_Helpers.LogError "Input binding failed | ExpectedNameUnicode=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(inputSheetName)
        Exit Sub
    End If
    If inputCellMap Is Nothing Then
        ex_Helpers.LogError "Input cell mapper is not initialized"
        Exit Sub
    End If
    For Each fieldAlias In inputCellMap.Keys
        cellAddress = VBA.CStr(inputCellMap(fieldAlias))
        bindingText = bindingText & " | " & fieldAlias & "=" & cellAddress
        valuesText = valuesText & " | " & fieldAlias & "=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(VBA.CStr( _
                worksheetObj.Range(cellAddress).Text))
    Next fieldAlias
    ex_Helpers.LogDebug "Input binding | SheetNameUnicode=" & _
        ex_Helpers.private_Text_ToUnicodeDebug(inputSheetName) & bindingText
    ex_Helpers.LogDebug "Input raw values Unicode" & valuesText
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_TryReadOptional( _
    ByVal inputSheetName As String, _
    ByVal inputCellMap As Object, _
    ByVal fieldAlias As String, _
    ByRef outValue As String _
) As Boolean
    Dim sourceSheet As Worksheet
    Dim cellAddress As String

    If Not private_TryGetInputCellAddress(inputCellMap, fieldAlias, cellAddress) Then Exit Function
    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Worksheets(inputSheetName)
    On Error GoTo 0
    If sourceSheet Is Nothing Then
        ex_Helpers.LogError "Input sheet was not found: " & inputSheetName
        VBA.MsgBox "Input sheet was not found: " & inputSheetName, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outValue = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    private_TryReadOptional = True
End Function

Private Function private_TryGetInputCellAddress( _
    ByVal inputCellMap As Object, _
    ByVal fieldAlias As String, _
    ByRef outCellAddress As String _
) As Boolean
    If inputCellMap Is Nothing Then
        ex_Helpers.LogError "Input cell mapper is not initialized"
        VBA.MsgBox "Input cell mapper is not initialized.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not inputCellMap.Exists(fieldAlias) Then
        ex_Helpers.LogError "Input field alias is not mapped: " & fieldAlias
        VBA.MsgBox "Input field alias is not mapped: " & fieldAlias, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outCellAddress = VBA.CStr(inputCellMap(fieldAlias))
    private_TryGetInputCellAddress = True
End Function