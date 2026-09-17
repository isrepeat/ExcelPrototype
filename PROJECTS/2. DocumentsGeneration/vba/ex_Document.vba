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
    ByRef outDocumentPath As String, _
    Optional ByVal outputFolderPath As String = "", _
    Optional ByVal overwriteDocumentPath As String = "", _
    Optional ByVal failIfDocumentExists As Boolean = False, _
    Optional ByVal renameUpdatedDocument As Boolean = False _
) As Boolean
    Dim documentName As String

    outDocumentPath = VBA.vbNullString
    If documentNameValues Is Nothing Then
        ex_Helpers.LogError "Document name context is not initialized"
        ex_Helpers.ex_ShowErrorMessage "Document name context is not initialized.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not ex_Helpers.private_Text_TryFormat( _
        documentNamePattern, documentNameValues, documentName) Then Exit Function
    ex_Helpers.LogDebug "Generated document name: " & documentName
    ex_TryGenerateWordDocument = ex_Helpers.private_Word_TryGenerateDocument( _
        templatePath, documentName, placeholderNames, placeholderValues, _
        outDocumentPath, outputFolderPath, overwriteDocumentPath, _
        failIfDocumentExists, renameUpdatedDocument)
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
        ex_Helpers.ex_ShowErrorMessage "Input sheet was not found: " & inputSheetName, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outValue = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(outValue) = 0 Then
        ex_Helpers.LogError "Required input is empty | Field=" & fieldAlias & _
            " | Cell=" & cellAddress
        ex_Helpers.ex_ShowErrorMessage "Enter " & fieldCaption & " in cell " & cellAddress & ".", _
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
        ex_Helpers.LogError "Input must be a non-negative whole number | Field=" & _
            fieldAlias & " | Value=" & daysText
        ex_Helpers.ex_ShowErrorMessage "Enter a non-negative whole number for " & fieldCaption & ".", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    On Error GoTo EH
    outDays = VBA.CLng(daysText)
    ex_TryReadNonNegativeDays = True
    Exit Function
EH:
    ex_Helpers.LogError "Input value is out of range | Field=" & fieldAlias & _
        " | Value=" & daysText
    ex_Helpers.ex_ShowErrorMessage "The value for " & fieldCaption & " is out of range.", _
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

' Находит единственную открытую умную таблицу по её имени во всех книгах Excel.
Public Function ex_TryFindOpenTable( _
    ByVal tableName As String, _
    ByRef outTable As ListObject _
) As Boolean
    Dim workbookObj As Workbook
    Dim worksheetObj As Worksheet
    Dim tableObj As ListObject
    Dim matchCount As Long
    Dim matchLocations As String

    Set outTable = Nothing
    tableName = ex_Helpers.private_Text_Normalize(tableName)
    If VBA.Len(tableName) = 0 Then
        ex_Helpers.LogError "Target table name is empty"
        ex_Helpers.ex_ShowErrorMessage "Target table name is empty.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    For Each workbookObj In Application.Workbooks
        For Each worksheetObj In workbookObj.Worksheets
            For Each tableObj In worksheetObj.ListObjects
                If VBA.StrComp(tableObj.Name, tableName, VBA.vbTextCompare) = 0 Then
                    matchCount = matchCount + 1
                    matchLocations = matchLocations & VBA.vbCrLf & "- " & _
                        workbookObj.FullName & " | " & worksheetObj.Name
                    If matchCount = 1 Then Set outTable = tableObj
                End If
            Next tableObj
        Next worksheetObj
    Next workbookObj

    If matchCount = 0 Then
        ex_Helpers.LogError "Open target table was not found: " & tableName
        ex_Helpers.ex_ShowErrorMessage "Open a workbook containing the table '" & tableName & "'.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If matchCount > 1 Then
        ex_Helpers.LogError "Target table is ambiguous: " & tableName & _
            " | Matches=" & VBA.CStr(matchCount)
        ex_Helpers.ex_ShowErrorMessage "More than one open table named '" & tableName & _
            "' was found:" & matchLocations, _
            VBA.vbExclamation, "Document Generation"
        Set outTable = Nothing
        Exit Function
    End If

    ex_Helpers.LogDebug "Open target table found: " & tableName & _
        " | Workbook=" & outTable.Parent.Parent.FullName & _
        " | Worksheet=" & outTable.Parent.Name
    ex_TryFindOpenTable = True
End Function

' Добавляет строку в умную таблицу, записывая значения по именам её колонок.
Public Function ex_TryAppendTableRow( _
    ByVal tableName As String, _
    ByVal columnValues As Object _
) As Boolean
    Dim targetTable As ListObject
    Dim targetRow As ListRow
    Dim columnName As Variant

    If columnValues Is Nothing Then
        ex_Helpers.LogError "Table row values are not initialized"
        ex_Helpers.ex_ShowErrorMessage "Table row values are not initialized.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not ex_TryFindOpenTable(tableName, targetTable) Then Exit Function
    For Each columnName In columnValues.Keys
        If Not private_TableHasColumn(targetTable, VBA.CStr(columnName)) Then Exit Function
    Next columnName

    On Error GoTo EH
    Set targetRow = targetTable.ListRows.Add
    For Each columnName In columnValues.Keys
        targetRow.Range.Cells(1, targetTable.ListColumns( _
            VBA.CStr(columnName)).Index).Value = columnValues(columnName)
    Next columnName
    ex_Helpers.LogDebug "Table row appended | Table=" & tableName & _
        " | Row=" & VBA.CStr(targetRow.Index)
    ex_TryAppendTableRow = True
    Exit Function
EH:
    If Not targetRow Is Nothing Then targetRow.Delete
    ex_Helpers.LogError "Failed to append table row | Table=" & tableName & _
        " | Number=" & VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to append a row to table '" & tableName & "': " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
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

Private Function private_TableHasColumn( _
    ByVal targetTable As ListObject, _
    ByVal columnName As String _
) As Boolean
    Dim tableColumn As ListColumn

    For Each tableColumn In targetTable.ListColumns
        If VBA.StrComp(tableColumn.Name, columnName, VBA.vbTextCompare) = 0 Then
            private_TableHasColumn = True
            Exit Function
        End If
    Next tableColumn
    ex_Helpers.LogError "Required table column was not found | Table=" & _
        targetTable.Name & " | Column=" & columnName
    ex_Helpers.ex_ShowErrorMessage "Required column '" & columnName & "' was not found in table '" & _
        targetTable.Name & "'.", VBA.vbExclamation, "Document Generation"
End Function

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
        ex_Helpers.ex_ShowErrorMessage "Input sheet was not found: " & inputSheetName, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outValue = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    private_TryReadOptional = True
End Function

Private Function private_TryGetInputCellAddress(ByVal inputCellMap As Object, ByVal fieldAlias As String, ByRef outCellAddress As String) As Boolean
    If inputCellMap Is Nothing Then
        ex_Helpers.LogError "Input cell mapper is not initialized"
        ex_Helpers.ex_ShowErrorMessage "Input cell mapper is not initialized.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not inputCellMap.Exists(fieldAlias) Then
        ex_Helpers.LogError "Input field alias is not mapped: " & fieldAlias
        ex_Helpers.ex_ShowErrorMessage "Input field alias is not mapped: " & fieldAlias, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outCellAddress = VBA.CStr(inputCellMap(fieldAlias))
    private_TryGetInputCellAddress = True
End Function