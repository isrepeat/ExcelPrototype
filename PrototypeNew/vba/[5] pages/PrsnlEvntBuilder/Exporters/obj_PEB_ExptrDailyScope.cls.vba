VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExptrDailyScope"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IDataExporter

Private m_IsDisposed As Boolean
Private m_Base As obj_DataExporterBase
Private m_Data As obj_PrsnlEvntBuilderData
Private m_ExporterDataProvider As obj_PEB_ExptrDataPrvdr

Private Const SAVE_ALREADY_OPEN_WORKBOOK As Boolean = False
Private Const SOURCE_ALIAS_HOSPITAL As String = "Hospital"
Private Const SOURCE_ALIAS_HOSPITAL_SHORT As String = "HospitalShort"
Private Const SOURCE_ALIAS_VACATION As String = "Vacation"
Private Const SOURCE_ALIAS_RANK As String = "Rank"
Private Const SOURCE_ALIAS_FIO As String = "FIO"
Private Const SOURCE_ALIAS_IPN As String = "IPN"
Private Const EXPORT_CONTEXT_VALIDATION_ENABLED As String = "ValidateDailyScope"
Private Const SOURCE_ALIAS_POSITION_CODE As String = "PositionCode"
Private Const SOURCE_ALIAS_POSITION_NAME As String = "PositionName"
Private Const SOURCE_ALIAS_REPORT_RANK As String = "ReportRank"
Private Const SOURCE_ALIAS_REPORT_PERSON As String = "ReportPerson"
Private Const SOURCE_ALIAS_REPORT_POSITION_CODE As String = "ReportPositionCode"
Private Const SOURCE_ALIAS_INCOMING_NO As String = "IncomingNo"
Private Const SOURCE_ALIAS_INCOMING_DATE As String = "IncomingDate"
Private Const SOURCE_ALIAS_DOCUMENT_NOTE As String = "DocumentNote"
Private Const SOURCE_ALIAS_DOC_NO As String = "DocNo"
Private Const SOURCE_ALIAS_DOC_DATE As String = "DocDate"
Private Const SOURCE_ALIAS_DURATION_DAYS As String = "DurationDays"
Private Const SOURCE_ALIAS_DATE_FROM As String = "DateFrom"
Private Const SOURCE_ALIAS_DATE_TO As String = "DateTo"
Private Const SOURCE_ALIAS_VACATION_TICKET_NO As String = "VacationTicketNo"
Private Const SOURCE_ALIAS_VACATION_TICKET_DATE As String = "VacationTicketDate"
Private Const SOURCE_ALIAS_VLK_NO As String = "VlkNo"
Private Const SOURCE_ALIAS_VLK_DATE As String = "VlkDate"
Private Const TARGET_COLUMN_HOSPITAL As String = "Лікарня"
Private Const TARGET_COLUMN_HOSPITAL_SHORT As String = "Лікарня скорочена назва"
Private Const TARGET_COLUMN_VACATION As String = "Відпустка"
Private Const TARGET_COLUMN_RANK As String = "Звання"
Private Const TARGET_COLUMN_FIO As String = "ПІБ"
Private Const TARGET_COLUMN_IPN As String = "ІПН"
Private Const TARGET_COLUMN_POSITION_CODE As String = "Код посади"
Private Const TARGET_COLUMN_POSITION_NAME As String = "Посада"
Private Const TARGET_COLUMN_REPORT_TVO As String = "Рапорт ТВО"
Private Const TARGET_COLUMN_REPORT_PERSON As String = "Рапорт кого"
Private Const TARGET_COLUMN_INCOMING_NO As String = "Вх.№"
Private Const TARGET_COLUMN_INCOMING_NO_ALT As String = "Вх. №"
Private Const TARGET_COLUMN_INCOMING_DATE As String = "Вх.Дата"
Private Const TARGET_COLUMN_INCOMING_DATE_ALT As String = "Вх. дата"
Private Const TARGET_COLUMN_DOCUMENT_NOTE As String = "Документ / Замітки"
Private Const TARGET_COLUMN_DOC_NO As String = "Док.№"
Private Const TARGET_COLUMN_DOC_NO_ALT As String = "Док. №"
Private Const TARGET_COLUMN_DOC_DATE As String = "Док.дата"
Private Const TARGET_COLUMN_DOC_DATE_ALT As String = "Док. дата"
Private Const TARGET_COLUMN_DURATION_DAYS As String = "На скільки"
Private Const TARGET_COLUMN_DATE_FROM As String = "З"
Private Const TARGET_COLUMN_DATE_TO As String = "По"
Private Const TARGET_COLUMN_VACATION_TICKET_NO As String = "Квит. №"
Private Const TARGET_COLUMN_VACATION_TICKET_DATE As String = "Квит. дата"
Private Const TARGET_COLUMN_VLK_NO As String = "ВЛК №"
Private Const TARGET_COLUMN_VLK_DATE As String = "ВЛК дата"
Private Const REPORT_TVO_TEXT As String = "тимчасово виконуючого обов'язки"
Private Const META_SECTION_TYPE_TVO As String = "Мета: ТВО"
Private Const TVO_ROW_MARKER As String = "ТВО"
Private Const SPECIAL_POSITION_PREFIX_ROZP As String = "A1A"
Private Const SPECIAL_POSITION_PREFIX_SPIS As String = "A1B"
Private Const SPECIAL_POSITION_CODE_ROZP As String = "РОЗП"
Private Const SPECIAL_POSITION_CODE_SPIS As String = "СПИС"
Private Const SPECIAL_POSITION_NAME_ROZP As String = "який перебуває у розпорядженні командира військової частини А7383"
Private Const SPECIAL_POSITION_NAME_SPIS As String = "який зарахований до списків військової частини А7383"

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
Public Function Initialize( _
    ByVal configTable As obj_ConfigTable, _
    Optional ByVal profileConfigTable As obj_ConfigTable = Nothing _
) As Boolean
    Dim dataProviderConfigTable As obj_ConfigTable

    private_LogMethodEntry "Initialize"

    m_IsDisposed = False
    Set m_Base = New obj_DataExporterBase
    Set m_Data = New obj_PrsnlEvntBuilderData
    Set m_ExporterDataProvider = New obj_PEB_ExptrDataPrvdr
    Set dataProviderConfigTable = configTable
    If Not profileConfigTable Is Nothing Then Set dataProviderConfigTable = profileConfigTable

    If Not m_Base.Initialize(configTable, "DailyScope", "PrototypeNew / DailyScope export") Then Exit Function
    If Not m_ExporterDataProvider.Initialize(dataProviderConfigTable) Then Exit Function

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
    If Not m_ExporterDataProvider Is Nothing Then m_ExporterDataProvider.Dispose
    Set m_Base = Nothing
    Set m_Data = Nothing
    Set m_ExporterDataProvider = Nothing

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
    Dim sectionKey As String
    Dim targetSheetName As String
    Dim openedByExporter As Boolean
    Dim fastModeStarted As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalculation As XlCalculation
    Dim exportValidationError As String
    Dim latestMovementTvoChain As Collection
    Dim latestMovementRecord As Object
    Dim insertedTvoRows As Collection
    Dim validationEnabled As Boolean

    On Error GoTo EH
    private_LogMethodEntry "Export"
    Set insertedTvoRows = New Collection
    If m_IsDisposed Then
        VBA.MsgBox "PrototypeNew: DailyScope exporter is disposed.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If
    If Not m_Base.TryGetMainSourceTable(sourceTables, sourceTable) Then Exit Function
    If Not private_TryResolveTargetSectionCaption(sourceTable, context, targetSectionCaption, sectionKey) Then Exit Function
    validationEnabled = (VBA.StrComp(private_GetContextText(context, EXPORT_CONTEXT_VALIDATION_ENABLED), "False", VBA.vbTextCompare) <> 0)
    If Not m_ExporterDataProvider.IsExportAllowed(sourceTable, sectionKey, exportValidationError, latestMovementTvoChain, latestMovementRecord, validationEnabled) Then
        VBA.MsgBox exportValidationError, VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    m_Base.BeginFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    fastModeStarted = True

    If Not m_Base.TryOpenTargetWorkbook(targetWb, openedByExporter) Then GoTo CleanFail
    targetSheetName = m_Base.ResolveTargetWorksheetName()
    If VBA.Len(targetSheetName) = 0 Then GoTo CleanFail
    If Not m_Base.TryGetWorksheet(targetWb, targetSheetName, targetWs) Then GoTo CleanFail
    If Not m_Base.TryFindConfiguredTargetTable(targetWs, targetTable) Then GoTo CleanFail

    If Not private_TryGetSectionWriteRowRange(targetTable, targetSectionCaption, targetRowRange, insertedRow) Then GoTo CleanFail

    If Not private_TryWriteSourceRow(sourceTable, targetTable, targetRowRange, sectionKey) Then GoTo CleanFail
    If Not private_TryAppendTvoRows(sourceTables, sourceTable, sectionKey, latestMovementTvoChain, targetTable, targetRowRange, insertedTvoRows) Then GoTo CleanFail
    If Not openedByExporter And SAVE_ALREADY_OPEN_WORKBOOK Then targetWb.Save
    Export = True
    GoTo CleanExit

CleanFail:
    Export = False
    private_DeleteInsertedRows insertedTvoRows
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
    private_DeleteInsertedRows insertedTvoRows
    If Not insertedRow Is Nothing Then insertedRow.Delete
    If openedByExporter Then targetWb.Close SaveChanges:=False
    If fastModeStarted Then m_Base.RestoreFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    On Error GoTo 0
End Function

' Добавляет под основной строкой отдельную строку для каждого участника цепочки ТВО.
' Явные meta-ТВО строки формы имеют приоритет. Для события возвращения, где форма
' обычно их не содержит, используется snapshot цепочки из последней Movement-строки.
Private Function private_TryAppendTvoRows( _
    ByVal sourceTables As Collection, _
    ByVal mainSourceTable As obj_TableDynamic, _
    ByVal sectionKey As String, _
    ByVal latestMovementTvoChain As Collection, _
    ByVal targetTable As ListObject, _
    ByVal mainRowRange As Range, _
    ByVal insertedRows As Collection _
) As Boolean
    Dim tvoItems As Collection
    Dim itemValue As Variant
    Dim itemObj As Object
    Dim insertedRow As ListRow
    Dim insertPosition As Long


    If sourceTables Is Nothing Then Exit Function
    If mainSourceTable Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If mainRowRange Is Nothing Then Exit Function
    If insertedRows Is Nothing Then Exit Function

    If Not private_TryBuildTvoItems(sourceTables, mainSourceTable, sectionKey, latestMovementTvoChain, tvoItems) Then Exit Function
    If tvoItems Is Nothing Then Exit Function

    insertPosition = mainRowRange.Row - targetTable.DataBodyRange.Row + 2
    For Each itemValue In tvoItems
        Set itemObj = itemValue
        Set insertedRow = targetTable.ListRows.Add(Position:=insertPosition)
        If insertedRow Is Nothing Then
            VBA.MsgBox "PrototypeNew: failed to insert a TVO row into DailyScope.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
            Exit Function
        End If
        insertedRows.Add insertedRow

        If Not private_TryApplyTvoRowFormat(mainRowRange, insertedRow.Range) Then Exit Function
        If Not private_TryWriteTvoRow(targetTable, insertedRow.Range, itemObj) Then Exit Function
        insertPosition = insertPosition + 1
    Next itemValue

    private_TryAppendTvoRows = True
End Function

Private Function private_TryBuildTvoItems( _
    ByVal sourceTables As Collection, _
    ByVal mainSourceTable As obj_TableDynamic, _
    ByVal sectionKey As String, _
    ByVal latestMovementTvoChain As Collection, _
    ByRef outItems As Collection _
) As Boolean
    Dim tableIndex As Long
    Dim tvoTable As obj_TableDynamic
    Dim tvoRow As obj_Row
    Dim itemObj As Object
    Dim chainValue As Variant
    Dim chainItem As Object


    Set outItems = New Collection

    ' Сначала переносим явно заданные пользователем meta-ТВО строки.
    For tableIndex = 2 To sourceTables.Count
        Set tvoTable = sourceTables.Item(tableIndex)
        If Not tvoTable Is Nothing Then
            If VBA.StrComp(private_NormalizeText(tvoTable.SectionTitle), private_NormalizeText(META_SECTION_TYPE_TVO), VBA.vbTextCompare) = 0 Then
                If tvoTable.RowCount > 0 Then
                    Set tvoRow = tvoTable.Rows.Item(1)
                    If tvoRow Is Nothing Then Exit Function
                    If Not private_TryBuildTvoItemFromSource(tvoTable, tvoRow, itemObj) Then Exit Function
                    outItems.Add itemObj
                End If
            End If
        End If
    Next tableIndex

    If outItems.Count > 0 Then
        private_TryBuildTvoItems = True
        Exit Function
    End If

    ' Историческая цепочка нужна только при закрытии Movement-события.
    If Not m_Data.IsMovementClosingSectionType(sectionKey) Then
        private_TryBuildTvoItems = True
        Exit Function
    End If

    ' IsExportAllowed уже прочитал последнюю Movement-строку вместе с TVO-полями.
    ' Используем переданный snapshot и не выполняем второй scan открытого листа.
    If latestMovementTvoChain Is Nothing Then
        private_TryBuildTvoItems = True
        Exit Function
    End If

    For Each chainValue In latestMovementTvoChain
        Set chainItem = chainValue
        If Not private_TryBuildTvoItem( _
            VBA.CStr(chainItem("FIO")), _
            VBA.CStr(chainItem("IPN")), _
            VBA.CStr(chainItem("PositionCode")), _
            VBA.vbNullString, _
            VBA.vbNullString, _
            itemObj) Then Exit Function
        outItems.Add itemObj
    Next chainValue

    private_TryBuildTvoItems = True
End Function

Private Function private_TryBuildTvoItemFromSource( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByRef outItem As Object _
) As Boolean
    Dim fioText As String
    Dim ipnText As String
    Dim positionCodeText As String
    Dim positionText As String
    Dim rankText As String

    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, fioText, SOURCE_ALIAS_FIO, TARGET_COLUMN_FIO) Then Exit Function
    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, ipnText, SOURCE_ALIAS_IPN, TARGET_COLUMN_IPN) Then Exit Function
    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, positionCodeText, SOURCE_ALIAS_POSITION_CODE, TARGET_COLUMN_POSITION_CODE) Then Exit Function
    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, positionText, "Position", SOURCE_ALIAS_POSITION_NAME, TARGET_COLUMN_POSITION_NAME) Then positionText = VBA.vbNullString
    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, rankText, SOURCE_ALIAS_RANK, TARGET_COLUMN_RANK) Then rankText = VBA.vbNullString

    private_TryBuildTvoItemFromSource = private_TryBuildTvoItem( _
        fioText, ipnText, positionCodeText, positionText, rankText, outItem)
End Function

Private Function private_TryBuildTvoItem( _
    ByVal fioText As String, _
    ByVal ipnText As String, _
    ByVal positionCodeText As String, _
    ByVal positionText As String, _
    ByVal rankText As String, _
    ByRef outItem As Object _
) As Boolean

    Set outItem = Nothing
    If m_ExporterDataProvider Is Nothing Then Exit Function
    If m_ExporterDataProvider.CommonData Is Nothing Then Exit Function

    If VBA.Len(VBA.Trim$(rankText)) = 0 Then
        If Not m_ExporterDataProvider.CommonData.TryResolveRankByIpn(ipnText, rankText) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(positionText)) = 0 Then
        If Not m_ExporterDataProvider.CommonData.TryResolvePositionDefault(positionCodeText, positionText) Then Exit Function
    End If

    Set outItem = VBA.CreateObject("Scripting.Dictionary")
    outItem.CompareMode = 1
    outItem("Rank") = rankText
    outItem("FIO") = fioText
    outItem("IPN") = ipnText
    outItem("PositionCode") = positionCodeText
    outItem("Position") = positionText
    private_TryBuildTvoItem = True
End Function

Private Function private_TryWriteTvoRow( _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByVal itemObj As Object _
) As Boolean
    If itemObj Is Nothing Then Exit Function
    If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_RANK, VBA.CStr(itemObj("Rank"))) Then Exit Function
    If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_FIO, VBA.CStr(itemObj("FIO"))) Then Exit Function
    If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_IPN, VBA.CStr(itemObj("IPN"))) Then Exit Function
    If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_POSITION_CODE, VBA.CStr(itemObj("PositionCode"))) Then Exit Function
    If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_POSITION_NAME, VBA.CStr(itemObj("Position"))) Then Exit Function
    If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_DOCUMENT_NOTE, TVO_ROW_MARKER) Then Exit Function
    private_TryWriteTvoRow = True
End Function

Private Function private_TryApplyTvoRowFormat(ByVal templateRange As Range, ByVal targetRange As Range) As Boolean
    On Error GoTo EH
    If templateRange Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function
    templateRange.Copy
    targetRange.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    private_TryApplyTvoRowFormat = True
    Exit Function
EH:
    Application.CutCopyMode = False
    VBA.MsgBox "PrototypeNew: failed to apply DailyScope TVO row format.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
End Function

Private Sub private_DeleteInsertedRows(ByVal insertedRows As Collection)
    Dim rowIndex As Long

    If insertedRows Is Nothing Then Exit Sub
    On Error Resume Next
    For rowIndex = insertedRows.Count To 1 Step -1
        insertedRows.Item(rowIndex).Delete
    Next rowIndex
    On Error GoTo 0
End Sub

Private Function private_TryResolveTargetSectionCaption( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal context As Object, _
    ByRef outSectionCaption As String, _
    ByRef outSectionKey As String _
) As Boolean
    Dim sectionKey As String

    outSectionCaption = VBA.vbNullString
    outSectionKey = VBA.vbNullString
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    sectionKey = private_GetContextText(context, "SectionType")
    If VBA.Len(sectionKey) = 0 Then sectionKey = VBA.Trim$(sourceTable.SectionTitle)
    If VBA.Len(sectionKey) = 0 Then
        VBA.MsgBox "PrototypeNew: DailyScope export requires SectionType in export context or source table SectionTitle.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
        Exit Function
    End If

    If private_TryMapSectionKeyToCaption(sectionKey, outSectionCaption) Then
        outSectionKey = sectionKey
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
    ByVal rowRange As Range, _
    ByVal sectionKey As String _
) As Boolean
    Dim sourceRow As obj_Row

    If sourceTable Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    ' Базовый export-контракт DailyScope: source читаем по стабильным alias-ам
    ' DynamicTable, а target ищем по фактическим заголовкам целевой таблицы.
    ' Это намеренно не copy-by-caption: UI формы может менять подписи колонок,
    ' не ломая экспорт в DailyScope.
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_HOSPITAL, TARGET_COLUMN_HOSPITAL) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_HOSPITAL_SHORT, TARGET_COLUMN_HOSPITAL_SHORT) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_VACATION, TARGET_COLUMN_VACATION) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_RANK, TARGET_COLUMN_RANK) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_FIO, TARGET_COLUMN_FIO) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_IPN, TARGET_COLUMN_IPN) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_POSITION_CODE, TARGET_COLUMN_POSITION_CODE) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_POSITION_NAME, TARGET_COLUMN_POSITION_NAME) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_INCOMING_NO, TARGET_COLUMN_INCOMING_NO, TARGET_COLUMN_INCOMING_NO_ALT) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_INCOMING_DATE, TARGET_COLUMN_INCOMING_DATE, TARGET_COLUMN_INCOMING_DATE_ALT) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_DOCUMENT_NOTE, TARGET_COLUMN_DOCUMENT_NOTE) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_DOC_NO, TARGET_COLUMN_DOC_NO, TARGET_COLUMN_DOC_NO_ALT) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_DOC_DATE, TARGET_COLUMN_DOC_DATE, TARGET_COLUMN_DOC_DATE_ALT) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_DURATION_DAYS, TARGET_COLUMN_DURATION_DAYS) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_DATE_FROM, TARGET_COLUMN_DATE_FROM) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_DATE_TO, TARGET_COLUMN_DATE_TO) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_VACATION_TICKET_NO, TARGET_COLUMN_VACATION_TICKET_NO) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_VACATION_TICKET_DATE, TARGET_COLUMN_VACATION_TICKET_DATE) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_VLK_NO, TARGET_COLUMN_VLK_NO) Then Exit Function
    If Not private_TryWriteDirectSourceValue(sourceTable, sourceRow, targetTable, rowRange, SOURCE_ALIAS_VLK_DATE, TARGET_COLUMN_VLK_DATE) Then Exit Function

    ' Обычный copy-by-caption не подходит для target-колонки "Рапорт кого":
    ' в форме теперь хранятся раздельные поля рапортующего
    ' (звание/ФІО/код посади), а DailyScope ждет уже собранную строку в
    ' родительном падеже. Поэтому после общего копирования точечно
    ' перезаписываем эту колонку вычисленным значением.
    If Not private_TryWriteReporterGenitiveValue(sourceTable, sourceRow, targetTable, rowRange) Then Exit Function
    If Not private_TryWriteHospitalDeclensionValue(sourceTable, sourceRow, targetTable, rowRange, sectionKey) Then Exit Function
    If Not private_TryWriteSpecialPositionValue(sourceTable, sourceRow, targetTable, rowRange) Then Exit Function

    private_TryWriteSourceRow = True
End Function

Private Function private_TryWriteDirectSourceValue( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByVal sourceAlias As String, _
    ParamArray targetColumnNames() As Variant _
) As Boolean
    Dim sourceColumnIndex As Long
    Dim targetColumnName As Variant
    Dim targetColumnIndex As Long
    Dim sourceValue As Variant

    private_TryWriteDirectSourceValue = True
    If sourceTable Is Nothing Then Exit Function
    If sourceRow Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    sourceColumnIndex = private_GetSourceColumnIndex(sourceTable, sourceAlias)
    If sourceColumnIndex <= 0 Then Exit Function

    For Each targetColumnName In targetColumnNames
        targetColumnIndex = private_FindTargetColumnIndex(targetTable, VBA.CStr(targetColumnName))
        If targetColumnIndex > 0 Then Exit For
    Next targetColumnName
    If targetColumnIndex <= 0 Then Exit Function

    sourceValue = sourceRow.GetCellValue(sourceColumnIndex)
    private_TryWriteDirectSourceValue = private_TryWriteCellValueWithFormulaPolicy( _
        rowRange.Cells(1, targetColumnIndex), _
        sourceValue)
End Function

Private Function private_TryWriteHospitalDeclensionValue( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByVal sectionKey As String _
) As Boolean
    Dim hospitalShortText As String
    Dim hospitalFullText As String
    Dim hospitalValueText As String

    private_TryWriteHospitalDeclensionValue = True
    If sourceTable Is Nothing Then Exit Function
    If sourceRow Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function
    If m_ExporterDataProvider Is Nothing Then Exit Function
    If m_ExporterDataProvider.CommonData Is Nothing Then Exit Function

    If Not private_TryGetSourceTextByAnyColumn( _
        sourceTable, sourceRow, hospitalShortText, _
        SOURCE_ALIAS_HOSPITAL_SHORT, "Лікарня скорочена назва") Then hospitalShortText = VBA.vbNullString
    If Not private_TryGetSourceTextByAnyColumn( _
        sourceTable, sourceRow, hospitalFullText, _
        SOURCE_ALIAS_HOSPITAL, TARGET_COLUMN_HOSPITAL) Then hospitalFullText = VBA.vbNullString

    If VBA.Len(hospitalShortText) = 0 And VBA.Len(hospitalFullText) = 0 Then Exit Function

    ' Для секций "На лікування..." DailyScope ожидает лечебное заведение
    ' в знахідному падеже; для остальных секций остается родительный.
    If private_ShouldUseHospitalAccusative(sectionKey) Then
        If Not m_ExporterDataProvider.CommonData.TryResolveHospitalAccusative(hospitalShortText, hospitalValueText) Then
            private_TryWriteHospitalDeclensionValue = False
            Exit Function
        End If
    Else
        If Not m_ExporterDataProvider.CommonData.TryResolveHospitalGenitive(hospitalShortText, hospitalValueText) Then
            private_TryWriteHospitalDeclensionValue = False
            Exit Function
        End If
    End If

    If VBA.Len(hospitalValueText) = 0 Then hospitalValueText = hospitalFullText
    If VBA.Len(hospitalValueText) = 0 Then hospitalValueText = hospitalShortText
    If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_HOSPITAL, hospitalValueText) Then
        private_TryWriteHospitalDeclensionValue = False
    End If
End Function

Private Function private_TryWriteSpecialPositionValue( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range _
) As Boolean
    Dim positionCodeText As String
    Dim targetPositionCodeText As String
    Dim targetPositionNameText As String

    private_TryWriteSpecialPositionValue = True
    If sourceTable Is Nothing Then Exit Function
    If sourceRow Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    If Not private_TryGetSourceTextByAnyColumn( _
        sourceTable, sourceRow, positionCodeText, _
        SOURCE_ALIAS_POSITION_CODE, TARGET_COLUMN_POSITION_CODE) Then Exit Function

    ' Некоторые служебные коды ШПС в целевых таблицах должны выглядеть как
    ' состояние военнослужащего, а не как исходный код должности.
    If Not private_TryResolveSpecialPositionMapping( _
        positionCodeText, targetPositionCodeText, targetPositionNameText) Then Exit Function

    If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_POSITION_CODE, targetPositionCodeText) Then
        private_TryWriteSpecialPositionValue = False
        Exit Function
    End If
    If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_POSITION_NAME, targetPositionNameText) Then
        private_TryWriteSpecialPositionValue = False
    End If
End Function

Private Function private_TryResolveSpecialPositionMapping( _
    ByVal sourcePositionCodeText As String, _
    ByRef outTargetPositionCodeText As String, _
    ByRef outTargetPositionNameText As String _
) As Boolean
    Dim normalizedCodeText As String

    outTargetPositionCodeText = VBA.vbNullString
    outTargetPositionNameText = VBA.vbNullString

    normalizedCodeText = private_NormalizeSpecialPositionPrefix(sourcePositionCodeText)
    If VBA.Left$(normalizedCodeText, VBA.Len(SPECIAL_POSITION_PREFIX_ROZP)) = SPECIAL_POSITION_PREFIX_ROZP Then
        outTargetPositionCodeText = SPECIAL_POSITION_CODE_ROZP
        outTargetPositionNameText = SPECIAL_POSITION_NAME_ROZP
        private_TryResolveSpecialPositionMapping = True
        Exit Function
    End If
    If VBA.Left$(normalizedCodeText, VBA.Len(SPECIAL_POSITION_PREFIX_SPIS)) = SPECIAL_POSITION_PREFIX_SPIS Then
        outTargetPositionCodeText = SPECIAL_POSITION_CODE_SPIS
        outTargetPositionNameText = SPECIAL_POSITION_NAME_SPIS
        private_TryResolveSpecialPositionMapping = True
    End If
End Function

Private Function private_NormalizeSpecialPositionPrefix(ByVal sourcePositionCodeText As String) As String
    Dim normalizedCodeText As String

    normalizedCodeText = VBA.UCase$(VBA.Trim$(sourcePositionCodeText))
    normalizedCodeText = VBA.Replace(normalizedCodeText, "А", "A")
    normalizedCodeText = VBA.Replace(normalizedCodeText, "В", "B")
    normalizedCodeText = VBA.Replace(normalizedCodeText, " ", VBA.vbNullString)
    private_NormalizeSpecialPositionPrefix = normalizedCodeText
End Function

Private Function private_TryWriteReporterGenitiveValue( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range _
) As Boolean
    Dim targetColumnIndex As Long
    Dim reportRankText As String
    Dim reportPersonText As String
    Dim reportPositionCodeText As String
    Dim reportRankGenitive As String
    Dim reportPersonInitialsGenitive As String
    Dim reportPositionGenitive As String
    Dim reportTvoPositionGenitive As String
    Dim reporterText As String
    Dim isReporterTvo As Boolean

    private_TryWriteReporterGenitiveValue = True
    If sourceTable Is Nothing Then Exit Function
    If sourceRow Is Nothing Then Exit Function
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    targetColumnIndex = private_FindTargetColumnIndex(targetTable, TARGET_COLUMN_REPORT_PERSON)
    If targetColumnIndex <= 0 Then Exit Function

    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, reportRankText, SOURCE_ALIAS_REPORT_RANK, "Звання (рапорт)") Then reportRankText = VBA.vbNullString
    If Not private_TryGetSourceTextByAnyColumn( _
        sourceTable, sourceRow, reportPersonText, SOURCE_ALIAS_REPORT_PERSON, _
        "ФІО (рапорт)", "Рапорт кого") Then reportPersonText = VBA.vbNullString
    If Not private_TryGetSourceTextByAnyColumn( _
        sourceTable, sourceRow, reportPositionCodeText, SOURCE_ALIAS_REPORT_POSITION_CODE, _
        "Код посади (рапорт)") Then reportPositionCodeText = VBA.vbNullString

    If VBA.Len(reportRankText) = 0 And VBA.Len(reportPersonText) = 0 And VBA.Len(reportPositionCodeText) = 0 Then Exit Function
    ' "сам" не ищем в АЛФ и не склоняем как командира: в DailyScope должна
    ' остаться явная отметка, что рапорт от самого военнослужащего.
    If private_IsSelfReportText(reportPersonText) Then
        If Not private_TryWriteCellValueWithFormulaPolicy(rowRange.Cells(1, targetColumnIndex), private_LowerFirstLetter(reportPersonText)) Then
            private_TryWriteReporterGenitiveValue = False
        End If
        Exit Function
    End If
    If m_ExporterDataProvider Is Nothing Then Exit Function
    If m_ExporterDataProvider.CommonData Is Nothing Then Exit Function

    If Not m_ExporterDataProvider.TryResolveReporterTvoPositionGenitive(reportPersonText, reportTvoPositionGenitive, isReporterTvo) Then
        private_TryWriteReporterGenitiveValue = False
        Exit Function
    End If
    If isReporterTvo Then
        reportPositionGenitive = reportTvoPositionGenitive
    Else
        If Not m_ExporterDataProvider.CommonData.TryResolvePositionGenitive(reportPositionCodeText, reportPositionGenitive) Then
            private_TryWriteReporterGenitiveValue = False
            Exit Function
        End If
    End If
    If Not m_ExporterDataProvider.CommonData.TryResolveRankGenitive(reportRankText, reportRankGenitive) Then
        private_TryWriteReporterGenitiveValue = False
        Exit Function
    End If
    If Not m_ExporterDataProvider.CommonData.TryResolveFioInitialsGenitiveByName(reportPersonText, reportPersonInitialsGenitive) Then
        private_TryWriteReporterGenitiveValue = False
        Exit Function
    End If

    ' Итоговый формат соответствует приказной формулировке:
    ' <посада в родовом> <звание в родовом> <ФИО-инициалы в родовом>.
    ' Например: "командира 1 механізованого батальйону майора КАСТЄРОВА С.О."
    reporterText = private_JoinNonEmptyParts( _
        private_JoinNonEmptyParts(reportPositionGenitive, reportRankGenitive), _
        reportPersonInitialsGenitive)

    If VBA.Len(reporterText) = 0 Then reporterText = reportPersonText
    If VBA.Len(reporterText) = 0 Then Exit Function
    reporterText = private_LowerFirstLetter(reporterText)

    If Not private_TryWriteCellValueWithFormulaPolicy(rowRange.Cells(1, targetColumnIndex), reporterText) Then
        private_TryWriteReporterGenitiveValue = False
        Exit Function
    End If

    If isReporterTvo Then
        If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_REPORT_TVO, REPORT_TVO_TEXT) Then
            private_TryWriteReporterGenitiveValue = False
            Exit Function
        End If
    Else
        If Not private_TryWriteTargetColumnText(targetTable, rowRange, TARGET_COLUMN_REPORT_TVO, VBA.vbNullString) Then
            private_TryWriteReporterGenitiveValue = False
            Exit Function
        End If
    End If
End Function

Private Function private_TryWriteTargetColumnText( _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByVal targetColumnName As String, _
    ByVal valueText As String _
) As Boolean
    Dim targetColumnIndex As Long

    private_TryWriteTargetColumnText = True
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    targetColumnIndex = private_FindTargetColumnIndex(targetTable, targetColumnName)
    If targetColumnIndex <= 0 Then
        Exit Function
    End If

    private_TryWriteTargetColumnText = private_TryWriteCellValueWithFormulaPolicy( _
        rowRange.Cells(1, targetColumnIndex), _
        valueText)
End Function

Private Function private_LowerFirstLetter(ByVal valueText As String) As String
    If VBA.Len(valueText) = 0 Then Exit Function
    private_LowerFirstLetter = VBA.LCase$(VBA.Left$(valueText, 1)) & VBA.Mid$(valueText, 2)
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

    If sourceTable.TryGetColumnIndexByAlias(columnName, private_GetSourceColumnIndex) Then Exit Function
    If sourceTable.TryGetColumnIndexByName(columnName, private_GetSourceColumnIndex) Then Exit Function

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

Private Function private_JoinNonEmptyParts(ByVal leftText As String, ByVal rightText As String) As String
    leftText = VBA.Trim$(leftText)
    rightText = VBA.Trim$(rightText)
    If VBA.Len(leftText) = 0 Then
        private_JoinNonEmptyParts = rightText
    ElseIf VBA.Len(rightText) = 0 Then
        private_JoinNonEmptyParts = leftText
    Else
        private_JoinNonEmptyParts = leftText & " " & rightText
    End If
End Function

Private Function private_IsSelfReportText(ByVal valueText As String) As Boolean
    valueText = private_NormalizeText(valueText)
    private_IsSelfReportText = (VBA.StrComp(valueText, "сам", VBA.vbTextCompare) = 0)
End Function

Private Function private_ShouldUseHospitalAccusative(ByVal sectionKey As String) As Boolean
    sectionKey = private_NormalizeText(sectionKey)
    private_ShouldUseHospitalAccusative = (VBA.Left$(sectionKey, VBA.Len("на лікування")) = "на лікування")
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
        Case "з щорічної відпустки"
            private_GetKnownSectionCaptionByText = "З щорічної відпустки:"
        Case "з відпустки за сімейними обставинами"
            private_GetKnownSectionCaptionByText = "З відпустки за сімейними обставинами:"
        Case "з медичної роти"
            private_GetKnownSectionCaptionByText = "З медичної роти:"
        Case "з амбулаторного влк"
            private_GetKnownSectionCaptionByText = "З амбулаторного ВЛК:"
        Case "зі стаціонарного влк"
            private_GetKnownSectionCaptionByText = "Зі стаціонарного ВЛК:"
        Case "на лікування"
            private_GetKnownSectionCaptionByText = "На лікування:" & VBA.vbLf & "(давальний відмінок)"
        Case "у щорічну відпустку"
            private_GetKnownSectionCaptionByText = "У щорічну відпустку:"
        Case "у відпустку за сімейними обставинами"
            private_GetKnownSectionCaptionByText = "У відпустку за сімейними обставинами:"
        Case "у відпустку для лікування"
            private_GetKnownSectionCaptionByText = "У відпустку для лікування:"
        Case "у медичну роту"
            private_GetKnownSectionCaptionByText = "У медичну роту:"
        Case "на амбулаторне влк"
            private_GetKnownSectionCaptionByText = "На амбулаторне ВЛК:"
        Case "лікування => відпустка для лікування"
            private_GetKnownSectionCaptionByText = "Лікування / Відпустка для лікування:"
        Case "лікування => стаціонарне влк"
            private_GetKnownSectionCaptionByText = "Лікування / Стаціонарне ВЛК:"
        Case "щорічна відпустка => лікування"
            private_GetKnownSectionCaptionByText = "Щорічна відпустка / Лікування:"
        Case "відпустка за сімейними => лікування"
            private_GetKnownSectionCaptionByText = "Відпустка за сімейними / Лікування:"
        Case "відпустка для лікування => лікування"
            private_GetKnownSectionCaptionByText = "Відпустка для лікування / Лікування:"
        Case "відпустка для лікування => відпустка для лікування"
            private_GetKnownSectionCaptionByText = "Відпустка для лікування / Відпустка для лікування:"
        Case "відпустка для лікування => стаціонарне влк"
            private_GetKnownSectionCaptionByText = "Відпустка для лікування / Стаціонарне ВЛК:"
        Case "амбулаторне влк => лікування"
            private_GetKnownSectionCaptionByText = "Амбулаторне ВЛК / Лікування:"
        Case "амбулаторне влк => відпустка для лікування"
            private_GetKnownSectionCaptionByText = "Амбулаторне ВЛК / Відпустка для лікування:"
        Case "стаціонарне влк => відпустка для лікування"
            private_GetKnownSectionCaptionByText = "Стаціонарне ВЛК / Відпустка для лікування:"
        Case "медична рота => лікування"
            private_GetKnownSectionCaptionByText = "Медична рота / Лікування:"
        Case "медична рота => відпустка для лікування"
            private_GetKnownSectionCaptionByText = "Медична рота / Відпустка для лікування:"
        Case "лікування у медичній роті => відпустка для лікування"
            private_GetKnownSectionCaptionByText = "Медична рота (лікування) / Відпустка для лікування:"
        Case "відпустка для лікування у медичній роті => лікування"
            private_GetKnownSectionCaptionByText = "Медична рота (відпустка для лікування) / Лікування:"
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
