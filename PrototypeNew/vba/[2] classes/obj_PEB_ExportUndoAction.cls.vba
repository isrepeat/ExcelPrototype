VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExportUndoAction"
Option Explicit

Implements obj_IUndoAction

Private Const KIND_MOVEMENT As String = "movement"
Private Const KIND_WORD As String = "word"

Private m_ActionId As String
Private m_Caption As String
Private m_Kind As String

Private m_WorkbookPath As String
Private m_WorkbookName As String
Private m_WorksheetName As String
Private m_TableName As String
Private m_MovementRowCountAfter As Long
Private m_ChangedRowIndex As Long
Private m_ChangedBeforeFormula As Variant
Private m_ChangedAfterFormula As Variant
Private m_InsertedRowIndex As Long
Private m_InsertedAfterFormula As Variant
Private m_InsertedRowWasAdded As Boolean
Private m_InsertedBeforeFormula As Variant

Private m_WordDocumentPath As String
Private m_WordBookmarkName As String
Private m_WordMetadataBookmarkName As String
Private m_WordGroupBookmarkName As String
Private m_WordGroupStart As Long
Private m_WordNestedGroupBookmarkName As String
Private m_WordNestedGroupStart As Long
Private m_WordInsertedStart As Long
Private m_WordInsertedText As String

Private Sub Class_Initialize()
    m_ActionId = "peb-export-" & VBA.Format$(VBA.Now, "yyyymmdd-hhnnss") & "-" & VBA.CStr(VBA.CLng(VBA.Timer * 1000#))
End Sub

Private Function obj_IUndoAction_GetActionId() As String
    obj_IUndoAction_GetActionId = m_ActionId
End Function

Private Function obj_IUndoAction_GetCaption() As String
    obj_IUndoAction_GetCaption = m_Caption
End Function

Private Function obj_IUndoAction_GetScopeKey() As String
    obj_IUndoAction_GetScopeKey = "PEB.Export." & m_Kind
End Function

Private Function obj_IUndoAction_IsValid() As Boolean
    obj_IUndoAction_IsValid = Me.IsValid()
End Function

Private Function obj_IUndoAction_Execute(ByRef outErrorText As String) As Boolean
    obj_IUndoAction_Execute = Me.Redo(outErrorText)
End Function

Private Function obj_IUndoAction_Undo(ByRef outErrorText As String) As Boolean
    obj_IUndoAction_Undo = Me.Undo(outErrorText)
End Function

Private Function obj_IUndoAction_Redo(ByRef outErrorText As String) As Boolean
    obj_IUndoAction_Redo = Me.Redo(outErrorText)
End Function

Public Function InitializeMovement( _
    ByVal workbookPath As String, _
    ByVal workbookName As String, _
    ByVal worksheetName As String, _
    ByVal tableName As String, _
    ByVal rowCountAfter As Long, _
    ByVal changedRowIndex As Long, _
    ByVal changedBeforeFormula As Variant, _
    ByVal changedAfterFormula As Variant, _
    ByVal insertedRowIndex As Long, _
    ByVal insertedAfterFormula As Variant, _
    Optional ByVal insertedRowWasAdded As Boolean = True, _
    Optional ByVal insertedBeforeFormula As Variant _
) As Boolean
    m_Kind = KIND_MOVEMENT
    m_Caption = "Undo last Movement export"
    m_WorkbookPath = VBA.Trim$(workbookPath)
    m_WorkbookName = VBA.Trim$(workbookName)
    m_WorksheetName = VBA.Trim$(worksheetName)
    m_TableName = VBA.Trim$(tableName)
    m_MovementRowCountAfter = rowCountAfter
    m_ChangedRowIndex = changedRowIndex
    m_ChangedBeforeFormula = changedBeforeFormula
    m_ChangedAfterFormula = changedAfterFormula
    m_InsertedRowIndex = insertedRowIndex
    m_InsertedAfterFormula = insertedAfterFormula
    m_InsertedRowWasAdded = insertedRowWasAdded
    m_InsertedBeforeFormula = insertedBeforeFormula

    If VBA.Len(m_WorkbookPath) = 0 Then Exit Function
    If VBA.Len(m_WorksheetName) = 0 Or VBA.Len(m_TableName) = 0 Then Exit Function
    If m_MovementRowCountAfter < 0 Then Exit Function
    If m_ChangedRowIndex <= 0 And m_InsertedRowIndex <= 0 Then Exit Function
    InitializeMovement = True
End Function

Public Function InitializeWord( _
    ByVal documentPath As String, _
    ByVal bookmarkName As String, _
    ByVal insertedStart As Long, _
    ByVal insertedText As String, _
    Optional ByVal metadataBookmarkName As String = VBA.vbNullString, _
    Optional ByVal groupBookmarkName As String = VBA.vbNullString, _
    Optional ByVal groupStart As Long = 0, _
    Optional ByVal nestedGroupBookmarkName As String = VBA.vbNullString, _
    Optional ByVal nestedGroupStart As Long = 0 _
) As Boolean
    m_Kind = KIND_WORD
    m_Caption = "Undo last WORD export"
    m_WordDocumentPath = VBA.Trim$(documentPath)
    m_WordBookmarkName = VBA.Trim$(bookmarkName)
    m_WordMetadataBookmarkName = VBA.Trim$(metadataBookmarkName)
    m_WordGroupBookmarkName = VBA.Trim$(groupBookmarkName)
    m_WordGroupStart = groupStart
    m_WordNestedGroupBookmarkName = VBA.Trim$(nestedGroupBookmarkName)
    m_WordNestedGroupStart = nestedGroupStart
    m_WordInsertedStart = insertedStart
    m_WordInsertedText = insertedText

    If VBA.Len(m_WordDocumentPath) = 0 Then Exit Function
    If VBA.Len(m_WordBookmarkName) = 0 Then Exit Function
    If m_WordInsertedStart < 0 Then Exit Function
    InitializeWord = True
End Function

Public Function IsValid() As Boolean
    Select Case m_Kind
        Case KIND_MOVEMENT
            IsValid = (VBA.Len(m_WorkbookPath) > 0 And VBA.Len(m_WorksheetName) > 0 And VBA.Len(m_TableName) > 0)
        Case KIND_WORD
            IsValid = (VBA.Len(m_WordDocumentPath) > 0 And VBA.Len(m_WordBookmarkName) > 0)
    End Select
End Function

Public Function Undo(ByRef outErrorText As String) As Boolean
    Select Case m_Kind
        Case KIND_MOVEMENT
            Undo = private_ApplyMovement(True, outErrorText)
        Case KIND_WORD
            Undo = private_ApplyWord(True, outErrorText)
    End Select
End Function

Public Function Redo(ByRef outErrorText As String) As Boolean
    Select Case m_Kind
        Case KIND_MOVEMENT
            Redo = private_ApplyMovement(False, outErrorText)
        Case KIND_WORD
            Redo = private_ApplyWord(False, outErrorText)
    End Select
End Function

Private Function private_ApplyMovement(ByVal isUndo As Boolean, ByRef outErrorText As String) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetTable As ListObject
    Dim openedHere As Boolean
    Dim insertedRow As ListRow

    On Error GoTo EH
    outErrorText = VBA.vbNullString
    If Not private_TryOpenWorkbook(wb, openedHere) Then
        outErrorText = "Movement workbook is unavailable: " & m_WorkbookPath
        Exit Function
    End If
    Set ws = wb.Worksheets(m_WorksheetName)
    Set targetTable = ws.ListObjects(m_TableName)

    If isUndo Then
        ' Валидация выполняется целиком до первой мутации. Помимо индексов
        ' сравниваем число строк и полный Formula-снимок: ручное редактирование,
        ' сортировка, добавление или удаление другой сущностью блокируют rollback.
        If targetTable.ListRows.Count <> m_MovementRowCountAfter Then GoTo StateChanged
        If m_InsertedRowIndex > 0 Then
            If m_InsertedRowIndex > targetTable.ListRows.Count Then GoTo StateChanged
            If Not private_AreFormulaSnapshotsEqual( _
                targetTable.ListRows(m_InsertedRowIndex).Range.Formula, _
                m_InsertedAfterFormula) Then GoTo StateChanged
        End If
        If m_ChangedRowIndex > 0 Then
            If m_ChangedRowIndex > targetTable.ListRows.Count Then GoTo StateChanged
            If Not private_AreFormulaSnapshotsEqual( _
                targetTable.ListRows(m_ChangedRowIndex).Range.Formula, _
                m_ChangedAfterFormula) Then GoTo StateChanged
        End If

        ' Mirror undo выполняется в обратном порядке: сначала убираем новую
        ' запись, затем возвращаем исходное состояние закрытой строки.
        If m_InsertedRowIndex > 0 Then
            If m_InsertedRowIndex > targetTable.ListRows.Count Then GoTo InvalidRow
            If m_InsertedRowWasAdded Then
                targetTable.ListRows(m_InsertedRowIndex).Delete
            Else
                targetTable.ListRows(m_InsertedRowIndex).Range.Formula = m_InsertedBeforeFormula
            End If
        End If
        If m_ChangedRowIndex > 0 Then
            If m_ChangedRowIndex > targetTable.ListRows.Count Then GoTo InvalidRow
            targetTable.ListRows(m_ChangedRowIndex).Range.Formula = m_ChangedBeforeFormula
        End If
    Else
        If targetTable.ListRows.Count <> _
            m_MovementRowCountAfter - VBA.IIf(m_InsertedRowIndex > 0 And m_InsertedRowWasAdded, 1, 0) Then GoTo StateChanged
        If m_ChangedRowIndex > 0 Then
            If m_ChangedRowIndex > targetTable.ListRows.Count Then GoTo StateChanged
            If Not private_AreFormulaSnapshotsEqual( _
                targetTable.ListRows(m_ChangedRowIndex).Range.Formula, _
                m_ChangedBeforeFormula) Then GoTo StateChanged
        End If
        If m_ChangedRowIndex > 0 Then
            If m_ChangedRowIndex > targetTable.ListRows.Count Then GoTo InvalidRow
            targetTable.ListRows(m_ChangedRowIndex).Range.Formula = m_ChangedAfterFormula
        End If
        If m_InsertedRowIndex > 0 Then
            If m_InsertedRowWasAdded Then
                Set insertedRow = targetTable.ListRows.Add(Position:=m_InsertedRowIndex)
                insertedRow.Range.Formula = m_InsertedAfterFormula
            Else
                If m_InsertedRowIndex > targetTable.ListRows.Count Then GoTo InvalidRow
                If Not private_AreFormulaSnapshotsEqual( _
                    targetTable.ListRows(m_InsertedRowIndex).Range.Formula, _
                    m_InsertedBeforeFormula) Then GoTo StateChanged
                targetTable.ListRows(m_InsertedRowIndex).Range.Formula = m_InsertedAfterFormula
            End If
        End If
    End If

    ' Для уже открытой пользователем книги сохраняем прежнюю модель exporter-а:
    ' меняем данные в памяти, но не делаем неявный Save. Открытую нами книгу
    ' закрываем с сохранением, иначе rollback не попадёт на диск.
    If openedHere Then wb.Close SaveChanges:=True
    private_ApplyMovement = True
    Exit Function

InvalidRow:
    outErrorText = "Movement table structure changed; recorded row index is unavailable."
    If openedHere Then wb.Close SaveChanges:=False
    Exit Function
StateChanged:
    outErrorText = "Movement undo was blocked because the target table changed after export."
    If openedHere Then wb.Close SaveChanges:=False
    Exit Function
EH:
    outErrorText = "Movement export rollback failed. " & Err.Description
    On Error Resume Next
    If openedHere Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Function private_AreFormulaSnapshotsEqual( _
    ByVal currentFormula As Variant, _
    ByVal expectedFormula As Variant _
) As Boolean
    Dim currentIsArray As Boolean
    Dim expectedIsArray As Boolean
    Dim rowIndex As Long
    Dim columnIndex As Long

    currentIsArray = VBA.IsArray(currentFormula)
    expectedIsArray = VBA.IsArray(expectedFormula)
    If currentIsArray <> expectedIsArray Then Exit Function
    If Not currentIsArray Then
        private_AreFormulaSnapshotsEqual = private_AreSnapshotValuesEqual( _
            currentFormula, expectedFormula)
        Exit Function
    End If

    On Error GoTo NotEqual
    If LBound(currentFormula, 1) <> LBound(expectedFormula, 1) Then Exit Function
    If UBound(currentFormula, 1) <> UBound(expectedFormula, 1) Then Exit Function
    If LBound(currentFormula, 2) <> LBound(expectedFormula, 2) Then Exit Function
    If UBound(currentFormula, 2) <> UBound(expectedFormula, 2) Then Exit Function
    For rowIndex = LBound(currentFormula, 1) To UBound(currentFormula, 1)
        For columnIndex = LBound(currentFormula, 2) To UBound(currentFormula, 2)
            If Not private_AreSnapshotValuesEqual( _
                currentFormula(rowIndex, columnIndex), _
                expectedFormula(rowIndex, columnIndex)) Then Exit Function
        Next columnIndex
    Next rowIndex
    private_AreFormulaSnapshotsEqual = True
NotEqual:
End Function

Private Function private_AreSnapshotValuesEqual( _
    ByVal leftValue As Variant, _
    ByVal rightValue As Variant _
) As Boolean
    On Error GoTo NotEqual

    If VBA.IsError(leftValue) Or VBA.IsError(rightValue) Then
        If Not (VBA.IsError(leftValue) And VBA.IsError(rightValue)) Then Exit Function
        private_AreSnapshotValuesEqual = (VBA.CStr(leftValue) = VBA.CStr(rightValue))
        Exit Function
    End If
    If VBA.IsNull(leftValue) Or VBA.IsNull(rightValue) Then
        private_AreSnapshotValuesEqual = (VBA.IsNull(leftValue) And VBA.IsNull(rightValue))
        Exit Function
    End If
    If VBA.IsEmpty(leftValue) Or VBA.IsEmpty(rightValue) Then
        private_AreSnapshotValuesEqual = (VBA.IsEmpty(leftValue) And VBA.IsEmpty(rightValue))
        Exit Function
    End If
    private_AreSnapshotValuesEqual = _
        (VBA.VarType(leftValue) = VBA.VarType(rightValue)) And _
        (VBA.CStr(leftValue) = VBA.CStr(rightValue))
NotEqual:
End Function

Private Function private_ApplyWord(ByVal isUndo As Boolean, ByRef outErrorText As String) As Boolean
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim targetRange As Object
    Dim documentOpened As Boolean
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim nestedGroupStart As Long
    Dim insertedStart As Long
    Dim insertedEnd As Long

    On Error GoTo EH
    outErrorText = VBA.vbNullString
    If VBA.Len(VBA.Dir$(m_WordDocumentPath, VBA.vbNormal Or VBA.vbReadOnly Or VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        outErrorText = "WORD result document is unavailable: " & m_WordDocumentPath
        Exit Function
    End If
    If Not rt_WordExportRuntime.fn_TryAcquireWordDocument( _
        m_WordDocumentPath, wordApp, wordDoc, documentOpened) Then
        outErrorText = "WORD application is unavailable."
        Exit Function
    End If

    If isUndo Then
        If Not wordDoc.Bookmarks.Exists(m_WordBookmarkName) Then
            outErrorText = "WORD export bookmark is missing: " & m_WordBookmarkName
            GoTo CleanFail
        End If
        ' PEB-закладка служит локатором начала вставки. Word может расширить
        ' её End при последующих вставках в ту же секцию, поэтому удаление
        ' всего Bookmark.Range способно захватить соседние пункты. Длина
        ' исходного текста хранится в самой undo-action и определяет точную
        ' правую границу только этой операции.
        Set targetRange = wordDoc.Bookmarks(m_WordBookmarkName).Range
        insertedStart = targetRange.Start
        insertedEnd = insertedStart + VBA.Len(m_WordInsertedText)
        If insertedEnd > wordDoc.Content.End Then
            outErrorText = "WORD document structure changed; recorded export text is unavailable."
            GoTo CleanFail
        End If
        wordDoc.Bookmarks(m_WordBookmarkName).Delete
        Set targetRange = wordDoc.Range(insertedStart, insertedEnd)
        targetRange.Delete
        ' Если удалена последняя запись группы, Word оставляет PEG/PEN как
        ' схлопнутые закладки. Следующий экспорт ошибочно решил бы, что
        ' заголовки даты и больницы уже присутствуют в документе.
        private_DeleteCollapsedWordBookmark _
            wordDoc, m_WordNestedGroupBookmarkName
        private_DeleteCollapsedWordBookmark _
            wordDoc, m_WordGroupBookmarkName
    Else
        If m_WordInsertedStart > wordDoc.Content.End Then
            outErrorText = "WORD document structure changed; recorded insertion point is unavailable."
            GoTo CleanFail
        End If
        groupStart = m_WordGroupStart
        ' Перед redo временно снимаем обе групповые закладки. После вставки
        ' восстанавливаем PEN для больницы и охватывающую её PEG для даты;
        ' иначе Word может оставить старые границы перед добавленным текстом.
        nestedGroupStart = m_WordNestedGroupStart
        If VBA.Len(m_WordNestedGroupBookmarkName) > 0 And _
            wordDoc.Bookmarks.Exists(m_WordNestedGroupBookmarkName) Then
            nestedGroupStart = wordDoc.Bookmarks( _
                m_WordNestedGroupBookmarkName).Range.Start
            wordDoc.Bookmarks(m_WordNestedGroupBookmarkName).Delete
        End If
        If VBA.Len(m_WordGroupBookmarkName) > 0 And _
            wordDoc.Bookmarks.Exists(m_WordGroupBookmarkName) Then
            groupStart = wordDoc.Bookmarks(m_WordGroupBookmarkName).Range.Start
            groupEnd = wordDoc.Bookmarks(m_WordGroupBookmarkName).Range.End
            wordDoc.Bookmarks(m_WordGroupBookmarkName).Delete
        End If
        Set targetRange = wordDoc.Range(m_WordInsertedStart, m_WordInsertedStart)
        targetRange.Text = m_WordInsertedText
        Set targetRange = wordDoc.Range(m_WordInsertedStart, m_WordInsertedStart + VBA.Len(m_WordInsertedText))
        targetRange.HighlightColorIndex = 0
        wordDoc.Bookmarks.Add m_WordBookmarkName, targetRange
        If VBA.Len(m_WordGroupBookmarkName) > 0 Then
            If groupEnd > 0 Then
                groupEnd = groupEnd + VBA.Len(m_WordInsertedText)
            Else
                groupEnd = m_WordInsertedStart + VBA.Len(m_WordInsertedText)
            End If
            Set targetRange = wordDoc.Range( _
                groupStart, groupEnd)
            wordDoc.Bookmarks.Add m_WordGroupBookmarkName, targetRange
        End If
        If VBA.Len(m_WordNestedGroupBookmarkName) > 0 Then
            Set targetRange = wordDoc.Range( _
                nestedGroupStart, _
                m_WordInsertedStart + VBA.Len(m_WordInsertedText))
            wordDoc.Bookmarks.Add m_WordNestedGroupBookmarkName, targetRange
        End If
        If VBA.Len(m_WordMetadataBookmarkName) > 0 And VBA.Len(m_WordInsertedText) > 0 Then
            ' Grouping metadata bookmark охватывает первый символ той же вставки.
            ' Основной delete удаляет его автоматически; redo восстанавливает.
            Set targetRange = wordDoc.Range(m_WordInsertedStart, m_WordInsertedStart + 1)
            wordDoc.Bookmarks.Add m_WordMetadataBookmarkName, targetRange
        End If
    End If

    wordDoc.Save
    If documentOpened Then
        wordDoc.Close False
        documentOpened = False
    End If
    private_ApplyWord = True
    Exit Function
CleanFail:
    If documentOpened Then wordDoc.Close False
    Exit Function
EH:
    outErrorText = "WORD export rollback failed. " & Err.Description
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
End Function

Private Sub private_DeleteCollapsedWordBookmark( _
    ByVal wordDoc As Object, _
    ByVal bookmarkName As String _
)
    Dim bookmarkRange As Object

    If wordDoc Is Nothing Then Exit Sub
    bookmarkName = VBA.Trim$(bookmarkName)
    If VBA.Len(bookmarkName) = 0 Then Exit Sub
    If Not wordDoc.Bookmarks.Exists(bookmarkName) Then Exit Sub

    Set bookmarkRange = wordDoc.Bookmarks(bookmarkName).Range
    If bookmarkRange.Start >= bookmarkRange.End Then
        wordDoc.Bookmarks(bookmarkName).Delete
    End If
End Sub

Private Function private_TryOpenWorkbook(ByRef outWorkbook As Workbook, ByRef outOpenedHere As Boolean) As Boolean
    Dim wb As Workbook

    Set outWorkbook = Nothing
    outOpenedHere = False
    For Each wb In Application.Workbooks
        If VBA.StrComp(VBA.Trim$(wb.FullName), m_WorkbookPath, VBA.vbTextCompare) = 0 Then
            Set outWorkbook = wb
            private_TryOpenWorkbook = True
            Exit Function
        End If
    Next wb
    If VBA.Len(VBA.Dir$(m_WorkbookPath)) = 0 Then Exit Function
    Set outWorkbook = Application.Workbooks.Open(m_WorkbookPath)
    outOpenedHere = Not outWorkbook Is Nothing
    private_TryOpenWorkbook = Not outWorkbook Is Nothing
End Function
