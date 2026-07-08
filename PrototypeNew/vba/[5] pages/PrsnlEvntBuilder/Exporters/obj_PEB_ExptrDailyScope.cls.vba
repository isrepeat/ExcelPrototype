VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExptrDailyScope"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IDataExporter

Private m_IsDisposed As Boolean
Private m_Base As obj_DataExporterBase
Private m_Data As obj_PrsnlEvntBuilderData

Private Const SAVE_ALREADY_OPEN_WORKBOOK As Boolean = False

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
    
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IDataExporter_Export( _
    ByVal sourceTables As Collection, _
    Optional ByVal context As Object = Nothing _
) As Boolean
    obj_IDataExporter_Export = Me.Export(sourceTables, context)
End Function

' //
' // API
' //
Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    private_LogMethodEntry "Initialize"

    m_IsDisposed = False
    Set m_Base = New obj_DataExporterBase
    Set m_Data = New obj_PrsnlEvntBuilderData

    If Not m_Base.Initialize(configTable, "DailyScope", "PrototypeNew / DailyScope export") Then Exit Function

    Initialize = True
End Function

Public Function TryGetSectionTypeOptions(ByRef outSectionTypeOptions As Collection) As Boolean
    private_LogMethodEntry "TryGetSectionTypeOptions"
    Set outSectionTypeOptions = m_Data.SectionTypeNames
    If outSectionTypeOptions Is Nothing Then Exit Function
    TryGetSectionTypeOptions = (outSectionTypeOptions.Count > 0)
End Function

Public Sub Dispose()
    private_LogMethodEntry "Dispose"
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_Base Is Nothing Then m_Base.Dispose
    Set m_Base = Nothing
    Set m_Data = Nothing

    On Error GoTo 0
End Sub

Public Function Export( _
    ByVal sourceTables As Collection, _
    Optional ByVal context As Object = Nothing _
) As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim targetTable As ListObject
    Dim insertedRow As ListRow
    Dim targetRowRange As Range
    Dim targetSectionCaption As String
    Dim targetSheetName As String
    Dim openedByExporter As Boolean
    Dim fastModeStarted As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalculation As XlCalculation

    On Error GoTo EH
    private_LogMethodEntry "Export"
    If m_IsDisposed Then
        VBA.MsgBox "PrototypeNew: DailyScope exporter is disposed.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If
    If Not m_Base.TryGetMainSourceTable(sourceTables, sourceTable) Then Exit Function
    If Not private_TryResolveTargetSectionCaption(sourceTable, context, targetSectionCaption) Then Exit Function

    m_Base.BeginFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    fastModeStarted = True

    If Not m_Base.TryOpenTargetWorkbook(targetWb, openedByExporter) Then GoTo CleanFail
    targetSheetName = m_Base.ResolveTargetWorksheetName()
    If VBA.Len(targetSheetName) = 0 Then GoTo CleanFail
    If Not m_Base.TryGetWorksheet(targetWb, targetSheetName, targetWs) Then GoTo CleanFail
    If Not m_Base.TryFindConfiguredTargetTable(targetWs, targetTable) Then GoTo CleanFail

    If Not private_TryGetSectionWriteRowRange(targetTable, targetSectionCaption, targetRowRange, insertedRow) Then GoTo CleanFail

    If Not private_TryWriteSourceRow(sourceTable, targetTable, targetRowRange) Then GoTo CleanFail

    If Not openedByExporter And SAVE_ALREADY_OPEN_WORKBOOK Then targetWb.Save
    Export = True
    GoTo CleanExit

CleanFail:
    Export = False
    If Not insertedRow Is Nothing Then
        On Error Resume Next
        insertedRow.Delete
        On Error GoTo 0
    End If

CleanExit:
    If openedByExporter Then
        On Error Resume Next
        targetWb.Close SaveChanges:=Export
        On Error GoTo 0
    End If
    If fastModeStarted Then m_Base.RestoreFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    Exit Function

EH:
    VBA.MsgBox "PrototypeNew: DailyScope test export failed. " & Err.Description, VBA.vbExclamation, "PrototypeNew / DailyScope export"
    On Error Resume Next
    If openedByExporter Then targetWb.Close SaveChanges:=False
    If fastModeStarted Then m_Base.RestoreFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    On Error GoTo 0
End Function

Private Function private_TryResolveTargetSectionCaption( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal context As Object, _
    ByRef outSectionCaption As String _
) As Boolean
    Dim sectionKey As String

    outSectionCaption = VBA.vbNullString
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    sectionKey = private_GetContextText(context, "SectionType")
    If VBA.Len(sectionKey) = 0 Then sectionKey = VBA.Trim$(sourceTable.SectionTitle)
    If VBA.Len(sectionKey) = 0 Then
        VBA.MsgBox "PrototypeNew: DailyScope export requires SectionType in export context or source table SectionTitle.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    If private_TryMapSectionKeyToCaption(sectionKey, outSectionCaption) Then
        private_TryResolveTargetSectionCaption = True
        Exit Function
    End If

    VBA.MsgBox "PrototypeNew: unknown DailyScope section key: " & sectionKey, VBA.vbExclamation, "PrototypeNew / DailyScope export"
End Function

Private Function private_GetContextText(ByVal context As Object, ByVal keyText As String) As String
    If context Is Nothing Then Exit Function
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Function

    On Error Resume Next
    If context.Exists(keyText) Then private_GetContextText = VBA.Trim$(VBA.CStr(context(keyText)))
    If Err.Number <> 0 Then
        Err.Clear
        private_GetContextText = VBA.Trim$(VBA.CStr(VBA.CallByName(context, keyText, VbGet)))
    End If
    On Error GoTo 0
End Function

Private Function private_TryWriteSourceRow( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range _
) As Boolean
    Dim sourceRow As obj_Row
    Dim sourceColumn As obj_Column
    Dim sourceColumnIndex As Long
    Dim targetColumnIndex As Long
    Dim sourceValue As Variant

    If sourceTable Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    For sourceColumnIndex = 1 To sourceTable.ColumnCount
        Set sourceColumn = sourceTable.Columns.Item(sourceColumnIndex)
        If sourceColumn Is Nothing Then GoTo ContinueColumn

        targetColumnIndex = private_FindTargetColumnIndex(targetTable, sourceColumn.Name)
        If targetColumnIndex <= 0 Then GoTo ContinueColumn

        sourceValue = sourceRow.GetCellValue(sourceColumnIndex)
        If Not private_TryWriteCellValueWithFormulaPolicy(rowRange.Cells(1, targetColumnIndex), sourceValue) Then Exit Function

ContinueColumn:
    Next sourceColumnIndex

    private_TryWriteSourceRow = True
End Function

Private Function private_FindTargetColumnIndex(ByVal targetTable As ListObject, ByVal targetColumnName As String) As Long
    Dim columnObj As ListColumn
    Dim expectedName As String
    Dim candidateName As String

    If targetTable Is Nothing Then Exit Function

    expectedName = private_NormalizeText(targetColumnName)
    If VBA.Len(expectedName) = 0 Then Exit Function

    For Each columnObj In targetTable.ListColumns
        candidateName = private_NormalizeText(VBA.CStr(columnObj.Name))
        If VBA.StrComp(candidateName, expectedName, VBA.vbTextCompare) = 0 Then
            private_FindTargetColumnIndex = columnObj.Index
            Exit Function
        End If
    Next columnObj
End Function

Private Function private_TryWriteCellValueWithFormulaPolicy( _
    ByVal targetCell As Range, _
    ByVal incomingValue As Variant _
) As Boolean
    Dim incomingText As String

    If targetCell Is Nothing Then Exit Function

    incomingText = VBA.Trim$(VBA.CStr(incomingValue))
    If VBA.Len(incomingText) = 0 And targetCell.HasFormula Then
        private_TryWriteCellValueWithFormulaPolicy = True
        Exit Function
    End If

    targetCell.Value2 = incomingValue
    private_TryWriteCellValueWithFormulaPolicy = True
End Function

Private Function private_TryGetSectionWriteRowRange( _
    ByVal targetTable As ListObject, _
    ByVal sectionCaption As String, _
    ByRef outRowRange As Range, _
    ByRef outInsertedRow As ListRow _
) As Boolean
    Dim sectionRowIndex As Long
    Dim nextSectionRowIndex As Long
    Dim sectionFirstDataIndex As Long
    Dim sectionLastDataIndex As Long
    Dim writeRowIndex As Long
    Dim insertPosition As Long

    Set outRowRange = Nothing
    Set outInsertedRow = Nothing
    If targetTable Is Nothing Then Exit Function
    If targetTable.DataBodyRange Is Nothing Then Exit Function

    If Not private_TryFindSectionRowIndex(targetTable, sectionCaption, sectionRowIndex) Then
        VBA.MsgBox "PrototypeNew: DailyScope section was not found: " & sectionCaption, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    nextSectionRowIndex = private_FindNextSectionRowIndex(targetTable, sectionRowIndex + 1)
    sectionFirstDataIndex = sectionRowIndex + 1
    If nextSectionRowIndex > 0 Then
        sectionLastDataIndex = nextSectionRowIndex - 1
    Else
        sectionLastDataIndex = targetTable.ListRows.Count
    End If

    writeRowIndex = private_FindNextEmptySectionRowIndex(targetTable, sectionFirstDataIndex, sectionLastDataIndex)
    If writeRowIndex > 0 Then
        Set outRowRange = targetTable.ListRows.Item(writeRowIndex).Range
        private_TryGetSectionWriteRowRange = Not outRowRange Is Nothing
        Exit Function
    End If

    If nextSectionRowIndex > 0 Then
        insertPosition = nextSectionRowIndex
    Else
        insertPosition = targetTable.ListRows.Count + 1
    End If

    Set outInsertedRow = targetTable.ListRows.Add(Position:=insertPosition)
    If outInsertedRow Is Nothing Then
        VBA.MsgBox "PrototypeNew: failed to insert row into DailyScope section: " & sectionCaption, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    If Not private_TryApplyInsertedRowSectionFormat(targetTable, outInsertedRow, sectionFirstDataIndex, sectionLastDataIndex) Then Exit Function

    Set outRowRange = outInsertedRow.Range
    private_TryGetSectionWriteRowRange = Not outRowRange Is Nothing
End Function

Private Function private_TryApplyInsertedRowSectionFormat( _
    ByVal targetTable As ListObject, _
    ByVal insertedRow As ListRow, _
    ByVal sectionFirstDataIndex As Long, _
    ByVal sectionLastDataIndex As Long _
) As Boolean
    Dim templateRowIndex As Long
    Dim templateRange As Range

    private_TryApplyInsertedRowSectionFormat = True
    If targetTable Is Nothing Then Exit Function
    If insertedRow Is Nothing Then Exit Function
    If insertedRow.Range Is Nothing Then Exit Function

    templateRowIndex = private_FindSectionTemplateRowIndex(targetTable, sectionFirstDataIndex, sectionLastDataIndex)
    If templateRowIndex <= 0 Then Exit Function

    Set templateRange = targetTable.ListRows.Item(templateRowIndex).Range
    If templateRange Is Nothing Then Exit Function

    On Error GoTo EH_APPLY_FORMAT
    ' Keep export as values-only while still preserving visual section template.
    templateRange.Copy
    insertedRow.Range.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    private_TryApplyInsertedRowSectionFormat = True
    Exit Function

EH_APPLY_FORMAT:
    Application.CutCopyMode = False
    VBA.MsgBox "PrototypeNew: failed to apply section row format for inserted DailyScope row.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
    private_TryApplyInsertedRowSectionFormat = False
End Function

Private Function private_FindSectionTemplateRowIndex( _
    ByVal targetTable As ListObject, _
    ByVal sectionFirstDataIndex As Long, _
    ByVal sectionLastDataIndex As Long _
) As Long
    Dim rowIndex As Long

    If targetTable Is Nothing Then Exit Function
    If sectionFirstDataIndex <= 0 Then Exit Function
    If sectionLastDataIndex < sectionFirstDataIndex Then Exit Function
    If targetTable.ListRows.Count <= 0 Then Exit Function

    For rowIndex = sectionFirstDataIndex To sectionLastDataIndex
        If rowIndex >= 1 And rowIndex <= targetTable.ListRows.Count Then
            private_FindSectionTemplateRowIndex = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function private_TryFindSectionRowIndex( _
    ByVal targetTable As ListObject, _
    ByVal sectionCaption As String, _
    ByRef outRowIndex As Long _
) As Boolean
    Dim rowIndex As Long
    Dim rowText As String
    Dim expectedText As String

    outRowIndex = 0
    expectedText = private_NormalizeText(sectionCaption)
    If VBA.Len(expectedText) = 0 Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If targetTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To targetTable.ListRows.Count
        rowText = private_GetSectionTextFromRow(targetTable.ListRows.Item(rowIndex).Range)
        If VBA.StrComp(private_NormalizeText(rowText), expectedText, VBA.vbTextCompare) = 0 Then
            outRowIndex = rowIndex
            private_TryFindSectionRowIndex = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function private_FindNextSectionRowIndex(ByVal targetTable As ListObject, ByVal firstRowIndex As Long) As Long
    Dim rowIndex As Long
    Dim rowText As String

    If targetTable Is Nothing Then Exit Function
    If targetTable.DataBodyRange Is Nothing Then Exit Function
    If firstRowIndex < 1 Then firstRowIndex = 1

    For rowIndex = firstRowIndex To targetTable.ListRows.Count
        rowText = private_GetSectionTextFromRow(targetTable.ListRows.Item(rowIndex).Range)
        If private_IsKnownSectionCaption(rowText) Then
            private_FindNextSectionRowIndex = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function private_FindNextEmptySectionRowIndex( _
    ByVal targetTable As ListObject, _
    ByVal firstRowIndex As Long, _
    ByVal lastRowIndex As Long _
) As Long
    Dim rowIndex As Long
    Dim lastUsedRowIndex As Long
    Dim candidateRowIndex As Long

    If targetTable Is Nothing Then Exit Function
    If firstRowIndex <= 0 Then Exit Function
    If lastRowIndex < firstRowIndex Then Exit Function

    lastUsedRowIndex = firstRowIndex - 1
    For rowIndex = firstRowIndex To lastRowIndex
        If Not private_IsRowEmpty(targetTable.ListRows.Item(rowIndex).Range) Then lastUsedRowIndex = rowIndex
    Next rowIndex

    candidateRowIndex = lastUsedRowIndex + 1
    If candidateRowIndex <= lastRowIndex Then
        If private_IsRowEmpty(targetTable.ListRows.Item(candidateRowIndex).Range) Then private_FindNextEmptySectionRowIndex = candidateRowIndex
    End If
End Function

Private Function private_IsRowEmpty(ByVal rowRange As Range) As Boolean
    Dim cellObj As Range
    Dim valueText As String

    If rowRange Is Nothing Then Exit Function

    For Each cellObj In rowRange.Cells
        valueText = VBA.Trim$(VBA.CStr(cellObj.Value2))
        If VBA.Len(valueText) > 0 Then Exit Function
    Next cellObj

    private_IsRowEmpty = True
End Function

Private Function private_GetSectionTextFromRow(ByVal rowRange As Range) As String
    If rowRange Is Nothing Then Exit Function

    On Error Resume Next
    private_GetSectionTextFromRow = VBA.CStr(rowRange.Cells(1, 1).Value2)
    On Error GoTo 0
End Function

Private Function private_TryGetSourceTextByAnyColumn( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByRef outText As String, _
    ParamArray columnNames() As Variant _
) As Boolean
    Dim columnName As Variant
    Dim columnIndex As Long

    outText = VBA.vbNullString
    If sourceTable Is Nothing Then Exit Function
    If sourceRow Is Nothing Then Exit Function

    For Each columnName In columnNames
        columnIndex = private_GetSourceColumnIndex(sourceTable, VBA.CStr(columnName))
        If columnIndex > 0 Then
            outText = VBA.Trim$(VBA.CStr(sourceRow.GetCellValue(columnIndex)))
            private_TryGetSourceTextByAnyColumn = VBA.Len(outText) > 0
            Exit Function
        End If
    Next columnName
End Function

Private Function private_GetSourceColumnIndex(ByVal sourceTable As obj_TableDynamic, ByVal columnName As String) As Long
    Dim sourceColIndex As Long
    Dim sourceColumn As obj_Column
    Dim expectedName As String

    If sourceTable Is Nothing Then Exit Function
    expectedName = private_NormalizeText(columnName)
    If VBA.Len(expectedName) = 0 Then Exit Function

    For sourceColIndex = 1 To sourceTable.ColumnCount
        Set sourceColumn = sourceTable.Columns.Item(sourceColIndex)
        If sourceColumn Is Nothing Then GoTo ContinueColumn
        If VBA.StrComp(private_NormalizeText(sourceColumn.Name), expectedName, VBA.vbTextCompare) = 0 Then
            private_GetSourceColumnIndex = sourceColIndex
            Exit Function
        End If

ContinueColumn:
    Next sourceColIndex
End Function

Private Function private_TryMapSectionKeyToCaption(ByVal sectionKey As String, ByRef outCaption As String) As Boolean
    Dim normalizedKey As String

    outCaption = VBA.vbNullString
    normalizedKey = private_NormalizeText(sectionKey)
    If VBA.Len(normalizedKey) = 0 Then Exit Function

    If Not private_IsSupportedSectionType(normalizedKey) Then
        Exit Function
    End If

    outCaption = private_GetKnownSectionCaptionByText(normalizedKey)

    private_TryMapSectionKeyToCaption = VBA.Len(outCaption) > 0
End Function

Private Function private_IsSupportedSectionType(ByVal valueText As String) As Boolean
    Dim normalizedText As String
    Dim sectionTypes As Collection
    Dim sectionTypeValue As Variant

    normalizedText = private_NormalizeText(valueText)
    If VBA.Len(normalizedText) = 0 Then Exit Function

    Set sectionTypes = m_Data.SectionTypeNames
    If sectionTypes Is Nothing Then Exit Function

    For Each sectionTypeValue In sectionTypes
        If VBA.StrComp(private_NormalizeText(VBA.CStr(sectionTypeValue)), normalizedText, VBA.vbTextCompare) = 0 Then
            private_IsSupportedSectionType = True
            Exit Function
        End If
    Next sectionTypeValue
End Function

Private Function private_IsKnownSectionCaption(ByVal valueText As String) As Boolean
    Dim normalizedCandidate As String
    Dim knownCaptions As Collection
    Dim captionObj As Variant
    Dim normalizedKnownCaption As String

    normalizedCandidate = private_NormalizeText(valueText)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "exporter:section-caption-check start raw='" & VBA.Replace(VBA.CStr(valueText), "'", "''") & "' normalized='" & VBA.Replace(normalizedCandidate, "'", "''") & "'"
#End If
    If VBA.Len(normalizedCandidate) = 0 Then Exit Function

    Set knownCaptions = private_BuildKnownSectionCaptions()
    If knownCaptions Is Nothing Then Exit Function

    For Each captionObj In knownCaptions
        normalizedKnownCaption = private_NormalizeText(VBA.CStr(captionObj))
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "exporter:section-caption-check compare candidate='" & VBA.Replace(normalizedCandidate, "'", "''") & "' known='" & VBA.Replace(normalizedKnownCaption, "'", "''") & "'"
#End If
        If VBA.StrComp(normalizedCandidate, normalizedKnownCaption, VBA.vbTextCompare) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo "exporter:section-caption-check match known='" & VBA.Replace(normalizedKnownCaption, "'", "''") & "'"
#End If
            private_IsKnownSectionCaption = True
            Exit Function
        End If
    Next captionObj

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "exporter:section-caption-check no-match candidate='" & VBA.Replace(normalizedCandidate, "'", "''") & "'"
#End If
End Function

Private Function private_BuildKnownSectionCaptions() As Collection
    Dim sectionTypes As Collection
    Dim sectionTypeValue As Variant
    Dim captionText As String
    Dim captionKey As String
    Dim knownCaptionMap As Object
    Dim captions As Collection

    Set captions = New Collection
    Set knownCaptionMap = ex_Helpers.fn_CreateDictionaryTextCompare()
    If knownCaptionMap Is Nothing Then Exit Function

    Set sectionTypes = m_Data.SectionTypeNames
    If sectionTypes Is Nothing Then Exit Function

    For Each sectionTypeValue In sectionTypes
        captionText = private_GetKnownSectionCaptionByText(VBA.CStr(sectionTypeValue))
        captionKey = private_NormalizeText(captionText)
        If VBA.Len(captionKey) = 0 Then GoTo ContinueSectionType
        If knownCaptionMap.Exists(captionKey) Then GoTo ContinueSectionType

        knownCaptionMap(captionKey) = True
        captions.Add captionText

ContinueSectionType:
    Next sectionTypeValue

    Set private_BuildKnownSectionCaptions = captions
End Function

Private Function private_GetKnownSectionCaptionByText(ByVal valueText As String) As String
    Dim normalizedText As String

    normalizedText = private_NormalizeText(valueText)
    Select Case normalizedText
        Case "з лікування"
            private_GetKnownSectionCaptionByText = "З лікування:"
        Case "з відпустки для лікування"
            private_GetKnownSectionCaptionByText = "З відпустки для лікування:"
        Case "з щорічної основної відпустки"
            private_GetKnownSectionCaptionByText = "З щорічної основної відпустки:"
        Case "з відпустки за сімейними обставинами"
            private_GetKnownSectionCaptionByText = "З відпустки за сімейними обставинами:"
        Case "з лікування медична рота"
            private_GetKnownSectionCaptionByText = "З лікування (медична рота):"
        Case "з амбулаторного обстеження влк"
            private_GetKnownSectionCaptionByText = "З амбулаторного обстеження / ВЛК:"
        Case "на лікування"
            private_GetKnownSectionCaptionByText = "На лікування:" & VBA.vbLf & "(давальний відмінок)"
        Case "у частину щорічної основної відпустки"
            private_GetKnownSectionCaptionByText = "У частину щорічної основної відпустки:"
        Case "у відпустку за сімейними обставинами"
            private_GetKnownSectionCaptionByText = "У відпустку за сімейними обставинами:"
        Case "у відпустку для лікування"
            private_GetKnownSectionCaptionByText = "У відпустку для лікування:"
        Case "на лікування медична рота"
            private_GetKnownSectionCaptionByText = "На лікування (медична рота):"
        Case "на амбулаторне обстеження влк"
            private_GetKnownSectionCaptionByText = "На амбулаторне обстеження / ВЛК:"
        Case "зміна місця перебування лікування => відпустка для лік"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(Лікування / Відпустка лік.):"
        Case "зміна місця перебування відпустка для лік => відпустка для лік"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(Відпустка лік / Відпустка лік.):"
        Case "зміна місця перебування відпустка для лік => лікування"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(Відпустка лік. / Лікування):"
        Case "зміна місця перебування відпустка для лік => влк"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(Відпустка лік. / ВЛК):"
        Case "зміна місця перебування влк => відпустка для лік"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(ВЛК / Відпустка лік.):"
        Case "зміна місця перебування влк => лікування"
            private_GetKnownSectionCaptionByText = "Зміна місця перебування" & VBA.vbLf & "(ВЛК / Лікування):"
        Case "у відрядження"
            private_GetKnownSectionCaptionByText = "У відрядження"
        Case "у відрядження сзч"
            private_GetKnownSectionCaptionByText = "У відрядження (СЗЧ):"
    End Select
End Function

Private Function private_NormalizeText(ByVal valueText As String) As String
    valueText = VBA.LCase$(VBA.Trim$(VBA.CStr(valueText)))
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    valueText = VBA.Replace(valueText, ":", VBA.vbNullString)
    valueText = VBA.Replace(valueText, ".", VBA.vbNullString)
    valueText = VBA.Replace(valueText, "/", " ")
    valueText = VBA.Replace(valueText, "(", " ")
    valueText = VBA.Replace(valueText, ")", " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    private_NormalizeText = VBA.Trim$(valueText)
End Function

Private Sub private_LogMethodEntry(ByVal methodName As String)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "exporter:enter " & VBA.Trim$(methodName)
#End If
End Sub
