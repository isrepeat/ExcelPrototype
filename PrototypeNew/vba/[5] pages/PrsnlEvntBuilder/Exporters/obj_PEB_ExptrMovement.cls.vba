VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExptrMovement"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IDataExporter

Private m_IsDisposed As Boolean
Private m_Base As obj_DataExporterBase
Private m_Data As obj_PrsnlEvntBuilderData
Private m_DataProvider As obj_PEB_ExptrDataPrvdr

Private Const SAVE_ALREADY_OPEN_WORKBOOK As Boolean = False
Private Const MOVEMENT_TARGET_COLUMN_COUNT As Long = 6
Private Const MOVEMENT_TRAILING_EMPTY_LOOKBACK_ROWS As Long = 20
Private Const MOVEMENT_SOURCE_INCOMING_NO As String = "Вх. №"
Private Const MOVEMENT_CONTEXT_MANUAL_ORDER_NO As String = "ManualOrderNo"
Private Const MOVEMENT_CONTEXT_SECTION_TYPE As String = "SectionType"
Private Const MOVEMENT_CONTEXT_VALIDATION_ENABLED As String = "ValidateMovement"
Private Const MOVEMENT_SOURCE_EVENT As String = "Подія"
Private Const MOVEMENT_SOURCE_INCOMING_DATE As String = "Вх. дата"
Private Const MOVEMENT_SOURCE_DEPARTURE_DATE As String = "З"
Private Const MOVEMENT_SOURCE_DURATION_DAYS As String = "На скільки"
Private Const MOVEMENT_SOURCE_VK_NO As String = "В/к №"
Private Const MOVEMENT_SOURCE_REPORT_RANK As String = "ReportRank"
Private Const MOVEMENT_SOURCE_REPORT_PERSON As String = "ReportPerson"
Private Const MOVEMENT_SOURCE_REPORT_POSITION_CODE As String = "ReportPositionCode"
Private Const MOVEMENT_SOURCE_PERSON_RANK As String = "Rank"
Private Const MOVEMENT_SOURCE_PERSON_IPN As String = "IPN"
Private Const MOVEMENT_TARGET_ORDER_NO As String = "Наказ вибуття"
Private Const MOVEMENT_TARGET_FOOD_FROM As String = "З продовольчого"
Private Const MOVEMENT_TARGET_DEPARTURE As String = "Вибуття"
Private Const MOVEMENT_TARGET_ARRIVAL_ORDER_NO As String = "Наказ прибуття"
Private Const MOVEMENT_TARGET_ON_FOOD As String = "На продовольче"
Private Const MOVEMENT_TARGET_ARRIVAL As String = "Прибуття"
Private Const MOVEMENT_TARGET_OUT_REASON As String = "Підстава вибуття"
Private Const MOVEMENT_TARGET_RETURN_REASON As String = "Підстава прибуття"
Private Const MOVEMENT_TARGET_IPN As String = "ІПН"
Private Const MOVEMENT_TARGET_DURATION_TERM As String = "Термін вибуття"
Private Const MOVEMENT_TARGET_ESCORT_DOCUMENT As String = "Супровідний документ"
Private Const MOVEMENT_TARGET_EVENT As String = "Подія"
Private Const MOVEMENT_TARGET_TVO_FIO As String = "ТВО ПІБ"
Private Const MOVEMENT_TARGET_TVO_IPN As String = "ТВО ІПН"
Private Const MOVEMENT_TARGET_TVO_POSITION As String = "ТВО Посада"
Private Const META_SECTION_TYPE_TVO As String = "Мета: ТВО"
Private Const SOURCE_ALIAS_FIO As String = "FIO"
Private Const SOURCE_ALIAS_IPN As String = "IPN"
Private Const SOURCE_ALIAS_POSITION_CODE As String = "PositionCode"
Private Const SENTINEL_SHORT_DATE As Date = #1/1/1900#
Private Const SPECIAL_POSITION_PREFIX_ROZP As String = "A1A"
Private Const SPECIAL_POSITION_PREFIX_SPIS As String = "A1B"
Private Const SPECIAL_POSITION_CODE_ROZP As String = "РОЗП"
Private Const SPECIAL_POSITION_CODE_SPIS As String = "СПИС"
Private Const MEDICAL_COMPANY_DESTINATION As String = "Медична рота А7383"

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
    Set m_DataProvider = New obj_PEB_ExptrDataPrvdr
    Set dataProviderConfigTable = configTable
    If Not profileConfigTable Is Nothing Then Set dataProviderConfigTable = profileConfigTable

    If Not m_Base.Initialize(configTable, "Movement", "PrototypeNew / Movement export") Then Exit Function
    If Not m_DataProvider.Initialize(dataProviderConfigTable) Then Exit Function
    private_LogInfo "movement:init workbook='" & private_EscapeForLog(m_Base.TargetWorkbookPath) & _
        "' sheet='" & private_EscapeForLog(m_Base.TargetSheetName) & _
        "' start='" & private_EscapeForLog(m_Base.TargetRangeStartMarker) & _
        "' end='" & private_EscapeForLog(m_Base.TargetRangeEndMarker) & "'"

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
    If Not m_DataProvider Is Nothing Then m_DataProvider.Dispose
    Set m_Base = Nothing
    Set m_Data = Nothing
    Set m_DataProvider = Nothing
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
    Dim targetValues As Variant
    Dim openedByExporter As Boolean
    Dim fastModeStarted As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalculation As XlCalculation
    Dim outgoingOrderNo As Variant
    Dim outgoingFoodFromDate As Variant
    Dim outgoingDepartureDate As Variant
    Dim closingOrderNo As Variant
    Dim closingOnFoodDate As Variant
    Dim closingArrivalDate As Variant
    Dim mirrorOpeningOrderNo As Variant
    Dim mirrorOpeningFoodFromDate As Variant
    Dim mirrorOpeningDepartureDate As Variant
    Dim writeSpecialOpeningFields As Boolean
    Dim specialDurationValue As Variant
    Dim specialVkNoValue As Variant
    Dim sectionTypeRaw As String
    Dim sectionTypeNormalized As String
    Dim mappedEventText As String
    Dim shouldWriteMappedEvent As Boolean
    Dim isClosingEvent As Boolean
    Dim isMirrorTransferEvent As Boolean
    Dim closingTargetIpn As String
    Dim basisSummaryText As String
    Dim exportValidationError As String
    Dim latestMovementTvoChain As Collection
    Dim latestMovementRecord As Object
    Dim validationEnabled As Boolean
    Dim tvoFioText As String
    Dim tvoIpnText As String
    Dim tvoPositionText As String

    On Error GoTo EH

    ' //
    ' // Flow экспорта Movement
    ' //
    ' // 1. Контроллер передает коллекцию sourceTables и общий context.
    ' //    Для Movement сейчас используется первая таблица: основная строка "Формы экспорта".
    ' //    Из context берутся общие значения, например тип секции и номер приказа.
    ' // 2. По типу секции определяется режим записи:
    ' //    - обычное выбытие/opening: проверить закрытие прошлой строки и дописать новую;
    ' //    - closing: найти последнюю открытую строку по ІПН и закрыть ее;
    ' //    - mirror transfer: сначала закрыть старую строку, затем сразу открыть новую.
    ' // 3. Для opening/mirror заранее собирается массив первых 6 целевых колонок
    ' //    таблицы Movement: звание, ПІБ, ІПН, код посади, подія, напрям/місце.
    ' // 4. После открытия target workbook/sheet/table выполняется одна из трех веток ниже.
    ' // 5. При ошибке после вставки новой строки CleanFail удаляет вставленную строку,
    ' //    чтобы частичный экспорт не оставлял мусор в целевой таблице.
    '
    If m_IsDisposed Then
        private_LogError "Movement exporter is disposed."
        VBA.MsgBox "PrototypeNew: Movement exporter is disposed.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    ' Основная source-таблица содержит данные строки события, а meta-ТВО
    ' таблицы ниже собираются в синхронные многострочные Movement-поля.
    If Not m_Base.TryGetMainSourceTable(sourceTables, sourceTable) Then Exit Function
    If Not private_TryValidateManualOrderNoSpecified(context) Then Exit Function

    ' SectionType определяет семантику операции. Для mirror transfer флаг closing
    ' принудительно сбрасывается, потому что mirror-ветка сама делает оба действия:
    ' закрытие старой записи и открытие новой.
    sectionTypeRaw = private_GetSectionTypeTextFromContext(context, sourceTable)
    sectionTypeNormalized = private_NormalizeText(sectionTypeRaw)
    isClosingEvent = private_IsClosingSectionType(sectionTypeNormalized)
    isMirrorTransferEvent = private_IsMirrorTransferSectionType(sectionTypeRaw)
    If isMirrorTransferEvent Then isClosingEvent = False
    validationEnabled = (VBA.StrComp(private_GetContextText(context, MOVEMENT_CONTEXT_VALIDATION_ENABLED), "False", VBA.vbTextCompare) <> 0)
    If Not m_DataProvider.IsExportAllowed(sourceTable, sectionTypeRaw, exportValidationError, latestMovementTvoChain, latestMovementRecord, validationEnabled) Then
        private_LogError "Movement export blocked by last-event validation: " & private_EscapeForLog(exportValidationError)
        VBA.MsgBox exportValidationError, VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    ' Некоторые типы выбытия пишут дополнительные поля открывающей записи:
    ' срок выбытия и сопровождающий документ. Helper сам решает, нужны ли эти поля для sectionType.
    If Not private_TryBuildSpecialOpeningValues(sourceTable, context, sectionTypeRaw, writeSpecialOpeningFields, specialDurationValue, specialVkNoValue) Then Exit Function
    If Not private_TryBuildMovementTvoValues(sourceTables, tvoFioText, tvoIpnText, tvoPositionText) Then Exit Function

    ' Для части профилей значение "Подія" берется не из формы напрямую,
    ' а мапится из типа секции в справочнике PrsnlEvntBuilderData.
    shouldWriteMappedEvent = private_TryMapSectionTypeToEventText(sectionTypeRaw, mappedEventText)
    If Not private_TryBuildMovementBasisSummary(sourceTable, context, basisSummaryText) Then Exit Function

    ' Closing-режим только закрывает существующую строку, поэтому values для новой
    ' строки ему не нужны. Opening и mirror transfer будут писать новую строку.
    If Not isClosingEvent Then
        If Not private_TryBuildMovementRowValues(sourceTable, sectionTypeRaw, targetValues) Then Exit Function
    End If

    m_Base.BeginFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    fastModeStarted = True

    If Not m_Base.TryOpenTargetWorkbook(targetWb, openedByExporter) Then GoTo CleanFail

    If Not m_Base.TryGetWorksheet(targetWb, m_Base.ResolveTargetWorksheetName(), targetWs) Then GoTo CleanFail
    If Not m_Base.TryFindConfiguredTargetTable(targetWs, targetTable) Then GoTo CleanFail

    If isClosingEvent Then
        ' Closing: закрываем уже существующее движение.
        ' Ищем последнюю строку целевой таблицы по ІПН и заполняем поля прибытия:
        ' наказ прибуття, на продовольче, прибуття.
        If Not private_TryBuildMovementClosingValues(sourceTable, context, closingOrderNo, closingOnFoodDate, closingArrivalDate) Then GoTo CleanFail

        If Not private_TryGetRequiredSourceText(sourceTable, sourceTable.Rows.Item(1), MOVEMENT_TARGET_IPN, closingTargetIpn) Then GoTo CleanFail
        If Not private_TryFindLastRowByIpn(targetTable, closingTargetIpn, targetRowRange) Then GoTo CleanFail
        If Not private_TryWriteMovementClosingRow(targetTable, targetRowRange, closingOrderNo, closingOnFoodDate, closingArrivalDate, basisSummaryText) Then GoTo CleanFail
    ElseIf isMirrorTransferEvent Then
        ' Mirror transfer: "зеркальный перевод" внутри таблицы Movement.
        ' Смысл: одним событием закрываем предыдущую открытую строку военнослужащего
        ' и тут же создаем новую открытую строку. Даты/номер приказа закрытия
        ' зеркально используются как даты/номер приказа открытия новой строки.
        If Not private_TryBuildMovementClosingValues(sourceTable, context, closingOrderNo, closingOnFoodDate, closingArrivalDate) Then GoTo CleanFail

        If Not private_TryGetRequiredSourceText(sourceTable, sourceTable.Rows.Item(1), MOVEMENT_TARGET_IPN, closingTargetIpn) Then GoTo CleanFail
        If Not private_TryFindLastRowByIpn(targetTable, closingTargetIpn, targetRowRange) Then GoTo CleanFail
        If Not private_TryWriteMovementClosingRow(targetTable, targetRowRange, closingOrderNo, closingOnFoodDate, closingArrivalDate, basisSummaryText) Then GoTo CleanFail

        ' Mirror/opening часть: те же значения, которыми закрыли старую строку,
        ' становятся начальными значениями новой строки.
        mirrorOpeningOrderNo = closingOrderNo
        mirrorOpeningDepartureDate = closingArrivalDate
        mirrorOpeningFoodFromDate = closingOnFoodDate

        If Not private_TryGetAppendRowRange(targetTable, targetRowRange, insertedRow) Then GoTo CleanFail
        If Not private_TryWriteMovementRow( _
            targetTable, targetRowRange, targetValues, _
            mirrorOpeningOrderNo, mirrorOpeningFoodFromDate, mirrorOpeningDepartureDate, _
            writeSpecialOpeningFields, specialDurationValue, specialVkNoValue, _
            shouldWriteMappedEvent, mappedEventText, basisSummaryText, _
            tvoFioText, tvoIpnText, tvoPositionText) Then GoTo CleanFail
    Else
        ' Opening: обычное выбытие. Создаем новую строку и заполняем поля выбытия:
        ' наказ вибуття, з продовольчого, вибуття, плюс базовые первые 6 колонок.
        If Not private_TryBuildMovementOutgoingValues(sourceTable, context, outgoingOrderNo, outgoingFoodFromDate, outgoingDepartureDate) Then GoTo CleanFail
        If Not private_TryGetAppendRowRange(targetTable, targetRowRange, insertedRow) Then GoTo CleanFail
        If Not private_TryWriteMovementRow( _
            targetTable, targetRowRange, targetValues, _
            outgoingOrderNo, outgoingFoodFromDate, outgoingDepartureDate, _
            writeSpecialOpeningFields, specialDurationValue, specialVkNoValue, _
            shouldWriteMappedEvent, mappedEventText, basisSummaryText, _
            tvoFioText, tvoIpnText, tvoPositionText) Then GoTo CleanFail
    End If

    If Not openedByExporter And SAVE_ALREADY_OPEN_WORKBOOK Then targetWb.Save
    Export = True
    GoTo CleanExit

CleanFail:
    private_LogError "Movement export failed before completion."
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
    private_LogError "Movement export exception: [" & VBA.CStr(Err.Number) & "] " & Err.Description
    VBA.MsgBox "PrototypeNew: Movement export failed. " & Err.Description, VBA.vbExclamation, "PrototypeNew / Movement export"
    On Error Resume Next
    If openedByExporter Then targetWb.Close SaveChanges:=False
    If fastModeStarted Then m_Base.RestoreFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    On Error GoTo 0
End Function

Private Function private_IsMirrorTransferSectionType(ByVal sectionTypeText As String) As Boolean
    private_IsMirrorTransferSectionType = m_Data.IsMovementMirrorTransferSectionType(sectionTypeText)
End Function

Private Function private_GetSectionTypeTextFromContext( _
    ByVal context As Object, _
    ByVal sourceTable As obj_TableDynamic _
) As String
    Dim sectionTypeText As String

    sectionTypeText = private_GetContextText(context, MOVEMENT_CONTEXT_SECTION_TYPE)
    If VBA.Len(sectionTypeText) > 0 Then
        private_GetSectionTypeTextFromContext = sectionTypeText
        Exit Function
    End If
    If sourceTable Is Nothing Then Exit Function
    sectionTypeText = VBA.Trim$(sourceTable.SectionTitle)
    If VBA.Len(sectionTypeText) > 0 Then
        private_GetSectionTypeTextFromContext = sectionTypeText
        Exit Function
    End If
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

Private Function private_TryValidateManualOrderNoSpecified(ByVal context As Object) As Boolean
    Dim manualOrderNoText As String

    manualOrderNoText = private_GetContextText(context, MOVEMENT_CONTEXT_MANUAL_ORDER_NO)
    If VBA.Len(VBA.Trim$(manualOrderNoText)) = 0 Then
        private_LogError "Movement export blocked: manual order number is empty."
        VBA.MsgBox "PrototypeNew: order number is required for Movement export. Fill the order number field before exporting.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    private_TryValidateManualOrderNoSpecified = True
End Function

Private Function private_TryBuildSpecialOpeningValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal context As Object, _
    ByVal sectionTypeText As String, _
    ByRef outShouldWrite As Boolean, _
    ByRef outDurationValue As Variant, _
    ByRef outVkNoValue As Variant _
) As Boolean
    Dim sourceRow As obj_Row

    outShouldWrite = False
    outDurationValue = VBA.vbNullString
    outVkNoValue = VBA.vbNullString

    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    If Not private_ShouldWriteSpecialOpeningFieldsForSectionType(sectionTypeText) Then
        private_TryBuildSpecialOpeningValues = True
        Exit Function
    End If

    outShouldWrite = True
    outDurationValue = private_GetOptionalSourceTextByAnyColumn(sourceTable, sourceRow, MOVEMENT_SOURCE_DURATION_DAYS)
    outVkNoValue = m_DataProvider.CommonData.FormatVacationTicketNoForExport( _
        private_GetOptionalSourceTextByAnyColumn(sourceTable, sourceRow, MOVEMENT_SOURCE_VK_NO), _
        private_GetContextText(context, MOVEMENT_CONTEXT_MANUAL_ORDER_NO))
    private_TryBuildSpecialOpeningValues = True
End Function

Private Function private_TryBuildMovementTvoValues( _
    ByVal sourceTables As Collection, _
    ByRef outFioText As String, _
    ByRef outIpnText As String, _
    ByRef outPositionText As String _
) As Boolean
    Dim tableIndex As Long
    Dim rowIndex As Long
    Dim tvoTable As obj_TableDynamic
    Dim tvoRow As obj_Row
    Dim fioText As String
    Dim ipnText As String
    Dim positionText As String
    Dim separatorText As String

    outFioText = VBA.vbNullString
    outIpnText = VBA.vbNullString
    outPositionText = VBA.vbNullString
    separatorText = VBA.vbLf & VBA.vbLf
    If sourceTables Is Nothing Then Exit Function

    For tableIndex = 2 To sourceTables.Count
        Set tvoTable = Nothing
        Set tvoTable = sourceTables.Item(tableIndex)
        If tvoTable Is Nothing Then GoTo ContinueTable
        If VBA.StrComp(private_NormalizeText(tvoTable.SectionTitle), _
            private_NormalizeText(META_SECTION_TYPE_TVO), VBA.vbTextCompare) <> 0 Then GoTo ContinueTable

        For rowIndex = 1 To tvoTable.RowCount
            Set tvoRow = tvoTable.Rows.Item(rowIndex)
            If tvoRow Is Nothing Then Exit Function
            If Not private_TryGetSourceTextByAnyColumn(tvoTable, tvoRow, fioText, SOURCE_ALIAS_FIO, MOVEMENT_TARGET_TVO_FIO) Then GoTo InvalidTvoRow
            If Not private_TryGetSourceTextByAnyColumn(tvoTable, tvoRow, ipnText, SOURCE_ALIAS_IPN, MOVEMENT_TARGET_TVO_IPN) Then GoTo InvalidTvoRow
            If Not private_TryGetSourceTextByAnyColumn(tvoTable, tvoRow, positionText, SOURCE_ALIAS_POSITION_CODE, MOVEMENT_TARGET_TVO_POSITION) Then GoTo InvalidTvoRow

            If VBA.Len(outFioText) > 0 Then
                outFioText = outFioText & separatorText
                outIpnText = outIpnText & separatorText
                outPositionText = outPositionText & separatorText
            End If
            outFioText = outFioText & fioText
            outIpnText = outIpnText & ipnText
            outPositionText = outPositionText & positionText
        Next rowIndex
ContinueTable:
    Next tableIndex

    private_TryBuildMovementTvoValues = True
    Exit Function

InvalidTvoRow:
    VBA.MsgBox "PrototypeNew: meta-ТВО строка должна содержать ПІБ, ІПН и код посади.", _
        VBA.vbExclamation, "PrototypeNew / Movement export"
End Function

Private Function private_ShouldWriteSpecialOpeningFieldsForSectionType(ByVal rawSectionType As String) As Boolean
    private_ShouldWriteSpecialOpeningFieldsForSectionType = m_Data.ShouldWriteMovementSpecialOpeningFields(rawSectionType)
End Function

Private Function private_GetOptionalSourceTextByAnyColumn( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ParamArray columnNames() As Variant _
) As Variant
    Dim columnName As Variant
    Dim valueObj As Variant

    If sourceTable Is Nothing Then Exit Function
    If sourceRow Is Nothing Then Exit Function

    For Each columnName In columnNames
        valueObj = private_GetOptionalSourceText(sourceTable, sourceRow, VBA.CStr(columnName))
        If VBA.Len(VBA.Trim$(VBA.CStr(valueObj))) > 0 Then
            private_GetOptionalSourceTextByAnyColumn = valueObj
            Exit Function
        End If
    Next columnName
End Function

Private Function private_IsClosingSectionType(ByVal normalizedSectionType As String) As Boolean
    private_IsClosingSectionType = m_Data.IsMovementClosingSectionType(normalizedSectionType)
End Function

Private Function private_TryBuildMovementClosingValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal context As Object, _
    ByRef outOrderNo As Variant, _
    ByRef outOnFoodDate As Variant, _
    ByRef outArrivalDate As Variant _
) As Boolean
    Dim sourceRow As obj_Row
    Dim manualOrderNo As Variant
    Dim arrivalDateRaw As Variant
    Dim arrivalDateFromForm As Date
    Dim orderDateByNo As Date
    Dim foodSupportDate As Date
    Dim hasArrivalDateFromForm As Boolean
    Dim hasOrderDateByNo As Boolean

    outOrderNo = VBA.vbNullString
    outOnFoodDate = VBA.vbNullString
    outArrivalDate = VBA.vbNullString

    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    manualOrderNo = private_GetContextText(context, MOVEMENT_CONTEXT_MANUAL_ORDER_NO)
    outOrderNo = manualOrderNo
    If VBA.Len(VBA.Trim$(VBA.CStr(outOrderNo))) = 0 Then
        outOrderNo = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_INCOMING_NO)
    End If

    arrivalDateRaw = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_DEPARTURE_DATE)
    hasArrivalDateFromForm = private_TryParseIncomingDateWithOrderContext(arrivalDateRaw, outOrderNo, arrivalDateFromForm)
    If VBA.Len(VBA.Trim$(VBA.CStr(arrivalDateRaw))) > 0 And Not hasArrivalDateFromForm Then
        private_LogError "Movement failed to resolve arrival date raw='" & private_EscapeForLog(VBA.CStr(arrivalDateRaw)) & "' by orderNo='" & private_EscapeForLog(VBA.CStr(outOrderNo)) & "'."
        VBA.MsgBox "PrototypeNew: failed to resolve full date for '" & MOVEMENT_SOURCE_DEPARTURE_DATE & "' from value '" & VBA.CStr(arrivalDateRaw) & "' and order number '" & VBA.CStr(outOrderNo) & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    hasOrderDateByNo = private_TryResolveOrderDateFromCommonData(outOrderNo, orderDateByNo)

    If hasArrivalDateFromForm Then
        outArrivalDate = arrivalDateFromForm
    ElseIf hasOrderDateByNo Then
        outArrivalDate = orderDateByNo
    End If

    If m_DataProvider.CommonData.TryCalculateFoodSupportDate( _
        hasArrivalDateFromForm, arrivalDateFromForm, foodSupportDate) Then
        outOnFoodDate = foodSupportDate
    End If

    private_TryBuildMovementClosingValues = True
End Function

Private Function private_TryFindLastRowByIpn( _
    ByVal targetTable As ListObject, _
    ByVal ipnValue As String, _
    ByRef outRowRange As Range _
) As Boolean
    Dim ipnColumnIndex As Long
    Dim rowIndex As Long
    Dim candidateValue As String
    Dim expectedValue As String

    Set outRowRange = Nothing
    If targetTable Is Nothing Then Exit Function

    expectedValue = private_NormalizeComparableToken(ipnValue)
    If VBA.Len(expectedValue) = 0 Then
        private_LogError "Movement closing failed because source IПН is empty."
        VBA.MsgBox "PrototypeNew: Movement closing requires source value '" & MOVEMENT_TARGET_IPN & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    ipnColumnIndex = private_FindTargetColumnIndex(targetTable, MOVEMENT_TARGET_IPN)
    If ipnColumnIndex <= 0 Then
        private_LogError "Movement closing failed because target column '" & private_EscapeForLog(MOVEMENT_TARGET_IPN) & "' was not found."
        VBA.MsgBox "PrototypeNew: Movement target column '" & MOVEMENT_TARGET_IPN & "' was not found.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    For rowIndex = targetTable.ListRows.Count To 1 Step -1
        candidateValue = private_NormalizeComparableToken(targetTable.ListRows.Item(rowIndex).Range.Cells(1, ipnColumnIndex).Value2)
        If VBA.StrComp(candidateValue, expectedValue, VBA.vbTextCompare) = 0 Then
            Set outRowRange = targetTable.ListRows.Item(rowIndex).Range
            private_TryFindLastRowByIpn = True
            Exit Function
        End If
    Next rowIndex

    private_LogError "Movement closing row was not found by ІПН='" & private_EscapeForLog(expectedValue) & "'."
    VBA.MsgBox "PrototypeNew: failed to find latest Movement row by '" & MOVEMENT_TARGET_IPN & "' = '" & ipnValue & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
End Function

Private Function private_NormalizeComparableToken(ByVal rawValue As Variant) As String
    Dim valueText As String

    valueText = VBA.Trim$(VBA.CStr(rawValue))
    valueText = VBA.Replace(valueText, " ", VBA.vbNullString)
    private_NormalizeComparableToken = valueText
End Function

Private Function private_TryWriteMovementClosingRow( _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByVal arrivalOrderNo As Variant, _
    ByVal onFoodDate As Variant, _
    ByVal arrivalDate As Variant, _
    ByVal basisSummaryText As String _
) As Boolean
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_ARRIVAL_ORDER_NO, arrivalOrderNo) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_ON_FOOD, onFoodDate) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_ARRIVAL, arrivalDate) Then Exit Function
    If VBA.Len(VBA.Trim$(basisSummaryText)) > 0 Then
        If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_RETURN_REASON, basisSummaryText) Then Exit Function
    End If

    private_TryWriteMovementClosingRow = True
End Function

' //
' // Internal
' //
Private Function private_TryBuildMovementRowValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sectionTypeText As String, _
    ByRef outValues As Variant _
) As Boolean
    Dim sourceRow As obj_Row
    Dim requiredValue As Variant
    Dim destinationValue As Variant
    Dim mappedEventText As String

    Set sourceRow = Nothing
    If VBA.IsArray(outValues) Then Erase outValues
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    ReDim outValues(1 To 1, 1 To MOVEMENT_TARGET_COLUMN_COUNT)

    If Not private_TryGetRequiredSourceText(sourceTable, sourceRow, "Звання", requiredValue) Then Exit Function
    outValues(1, 1) = requiredValue
    If Not private_TryGetRequiredSourceText(sourceTable, sourceRow, "ПІБ", requiredValue) Then Exit Function
    outValues(1, 2) = requiredValue
    If Not private_TryGetRequiredSourceText(sourceTable, sourceRow, "ІПН", requiredValue) Then Exit Function
    outValues(1, 3) = requiredValue
    If Not private_TryGetRequiredSourceText(sourceTable, sourceRow, "Код посади", requiredValue) Then Exit Function
    outValues(1, 4) = private_NormalizePositionCodeForMovement(requiredValue)

    If private_TryMapSectionTypeToEventText(sectionTypeText, mappedEventText) Then
        outValues(1, 5) = mappedEventText
    Else
        outValues(1, 5) = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_EVENT)
    End If

    destinationValue = private_ResolveMovementDestinationValue(sourceTable, sourceRow, sectionTypeText)
    outValues(1, 6) = destinationValue

    private_TryBuildMovementRowValues = True
End Function

Private Function private_NormalizePositionCodeForMovement(ByVal sourcePositionCodeValue As Variant) As String
    Dim sourcePositionCodeText As String
    Dim normalizedCodeText As String

    If VBA.IsError(sourcePositionCodeValue) Then Exit Function
    If VBA.IsNull(sourcePositionCodeValue) Then Exit Function
    If VBA.IsEmpty(sourcePositionCodeValue) Then Exit Function

    sourcePositionCodeText = VBA.CStr(sourcePositionCodeValue)
    normalizedCodeText = private_NormalizeSpecialPositionPrefix(sourcePositionCodeText)

    If VBA.Left$(normalizedCodeText, VBA.Len(SPECIAL_POSITION_PREFIX_ROZP)) = SPECIAL_POSITION_PREFIX_ROZP Then
        private_NormalizePositionCodeForMovement = SPECIAL_POSITION_CODE_ROZP
        Exit Function
    End If
    If VBA.Left$(normalizedCodeText, VBA.Len(SPECIAL_POSITION_PREFIX_SPIS)) = SPECIAL_POSITION_PREFIX_SPIS Then
        private_NormalizePositionCodeForMovement = SPECIAL_POSITION_CODE_SPIS
        Exit Function
    End If

    private_NormalizePositionCodeForMovement = sourcePositionCodeText
End Function

Private Function private_NormalizeSpecialPositionPrefix(ByVal sourcePositionCodeText As String) As String
    Dim normalizedCodeText As String

    normalizedCodeText = VBA.UCase$(VBA.Trim$(sourcePositionCodeText))
    normalizedCodeText = VBA.Replace(normalizedCodeText, "А", "A")
    normalizedCodeText = VBA.Replace(normalizedCodeText, "В", "B")
    normalizedCodeText = VBA.Replace(normalizedCodeText, " ", VBA.vbNullString)
    private_NormalizeSpecialPositionPrefix = normalizedCodeText
End Function

Private Function private_ResolveMovementDestinationValue( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal sectionTypeText As String _
) As Variant
    ' В профиле выбытия в медицинскую роту поле больницы намеренно скрыто,
    ' поэтому направление задаётся семантикой самой секции.
    If VBA.StrComp( _
        private_NormalizeText(sectionTypeText), _
        private_NormalizeText(m_Data.SectionTypeToTreatmentMedicalCompany), _
        VBA.vbTextCompare) = 0 Then
        private_ResolveMovementDestinationValue = MEDICAL_COMPANY_DESTINATION
        Exit Function
    End If

    If m_Data.UsesMovementVacationDestination(sectionTypeText) Then
        private_ResolveMovementDestinationValue = private_GetOptionalSourceText(sourceTable, sourceRow, "Відпустка")
        Exit Function
    End If

    private_ResolveMovementDestinationValue = private_GetOptionalSourceText(sourceTable, sourceRow, "Лікарня скорочена назва")
End Function

Private Function private_TryBuildMovementBasisSummary( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal context As Object, _
    ByRef outBasisSummary As String _
) As Boolean
    Dim sourceRow As obj_Row
    Dim reportRankText As String
    Dim reportPersonText As String
    Dim reportPositionCodeText As String
    Dim incomingNoText As String
    Dim incomingDateText As String
    Dim orderNoText As String
    Dim reportRankGenitive As String
    Dim reportPositionGenitive As String
    Dim reportTvoPositionGenitive As String
    Dim reportPositionShortGenitive As String
    Dim reportPersonInitialsGenitive As String
    Dim personRankText As String
    Dim personIpnText As String
    Dim personRankGenitive As String
    Dim personInitialsGenitive As String
    Dim reporterCoreText As String
    Dim reporterText As String
    Dim isReporterTvo As Boolean
    Dim incomingDateValue As Date
    Dim incomingDateResolvedText As String
    Dim basisDetailsText As String

    outBasisSummary = VBA.vbNullString
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    If m_DataProvider Is Nothing Then Exit Function
    If m_DataProvider.CommonData Is Nothing Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, reportRankText, MOVEMENT_SOURCE_REPORT_RANK, "Звання (рапорт)", "Рапорт ТВО") Then reportRankText = VBA.vbNullString
    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, reportPersonText, MOVEMENT_SOURCE_REPORT_PERSON, "ФІО (рапорт)", "Рапорт кого") Then reportPersonText = VBA.vbNullString
    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, reportPositionCodeText, MOVEMENT_SOURCE_REPORT_POSITION_CODE, "Код посади (рапорт)") Then reportPositionCodeText = VBA.vbNullString
    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, incomingNoText, MOVEMENT_SOURCE_INCOMING_NO) Then incomingNoText = VBA.vbNullString
    If Not private_TryGetSourceTextByAnyColumn(sourceTable, sourceRow, incomingDateText, MOVEMENT_SOURCE_INCOMING_DATE) Then incomingDateText = VBA.vbNullString

    If private_IsSelfReportText(reportPersonText) Then
        ' Пустое поле рапортующего и значение "сам" означают, что рапорт
        ' подал основной военнослужащий. В основании указываем его звание и
        ' фамилию с инициалами в родительном падеже, но не должность.
        If Not private_TryGetSourceTextByAnyColumn( _
            sourceTable, sourceRow, personRankText, MOVEMENT_SOURCE_PERSON_RANK, "Звання") Then
            VBA.MsgBox "PrototypeNew: self-report basis requires the person's rank.", _
                VBA.vbExclamation, "PrototypeNew / Movement export"
            Exit Function
        End If
        If Not private_TryGetSourceTextByAnyColumn( _
            sourceTable, sourceRow, personIpnText, MOVEMENT_SOURCE_PERSON_IPN, "ІПН") Then
            VBA.MsgBox "PrototypeNew: self-report basis requires the person's IPN.", _
                VBA.vbExclamation, "PrototypeNew / Movement export"
            Exit Function
        End If
        If VBA.Len(VBA.Trim$(personRankText)) = 0 Or VBA.Len(VBA.Trim$(personIpnText)) = 0 Then
            VBA.MsgBox "PrototypeNew: self-report basis requires non-empty rank and IPN.", _
                VBA.vbExclamation, "PrototypeNew / Movement export"
            Exit Function
        End If
        If Not m_DataProvider.CommonData.TryResolveRankGenitive( _
            personRankText, personRankGenitive) Then Exit Function
        If Not m_DataProvider.CommonData.TryResolveFioInitialsGenitive( _
            personIpnText, personInitialsGenitive) Then Exit Function

        reporterText = private_JoinNonEmptyParts(personRankGenitive, personInitialsGenitive)
        If VBA.Len(VBA.Trim$(reporterText)) = 0 Then
            VBA.MsgBox "PrototypeNew: failed to build rank and initials for self-report basis.", _
                VBA.vbExclamation, "PrototypeNew / Movement export"
            Exit Function
        End If
    Else
        If Not m_DataProvider.CommonData.TryResolveRankGenitive(reportRankText, reportRankGenitive) Then Exit Function
        If Not m_DataProvider.TryResolveReporterTvoPositionGenitive(reportPersonText, reportTvoPositionGenitive, isReporterTvo) Then Exit Function
        If isReporterTvo Then
            reportPositionGenitive = reportTvoPositionGenitive
        Else
            If Not m_DataProvider.CommonData.TryResolvePositionGenitive(reportPositionCodeText, reportPositionGenitive) Then Exit Function
        End If
        If Not m_DataProvider.CommonData.TryResolveFioInitialsGenitiveByName(reportPersonText, reportPersonInitialsGenitive) Then Exit Function

        ' Movement использует короткую формулировку основания:
        ' "рапорт ком. 1 мб майора РУБАНА І.І.".
        ' Поэтому звание не ставим перед должностью, а саму должность
        ' приводим к нижнему регистру первой буквы и сокращаем локальными
        ' правилами, которые раньше жили в формуле целевой таблицы.
        If isReporterTvo Then
            reportPositionShortGenitive = "ТВО " & private_LowerFirstLetter(reportPositionGenitive)
        Else
            reportPositionShortGenitive = private_AbbreviateMovementBasisText( _
                private_LowerFirstLetter(reportPositionGenitive))
        End If
        reporterCoreText = private_JoinNonEmptyParts( _
            private_JoinNonEmptyParts(reportPositionShortGenitive, reportRankGenitive), _
            reportPersonInitialsGenitive)
        reporterText = reporterCoreText
        If VBA.Len(reporterText) = 0 Then reporterText = reportPersonText
    End If

    If VBA.Len(reporterText) = 0 Then reporterText = "військовослужбовця"
    If Not isReporterTvo Then reporterText = private_LowerFirstLetter(reporterText)

    incomingNoText = m_DataProvider.NormalizeIncomingNoForExport(incomingNoText)
    orderNoText = private_GetContextText(context, MOVEMENT_CONTEXT_MANUAL_ORDER_NO)
    If VBA.Len(orderNoText) = 0 Then orderNoText = incomingNoText

    ' Блок входящего документа добавляем, если указан хотя бы один реквизит.
    ' Номер и дата формируются независимо; весь блок пропускается только тогда,
    ' когда оба значения пустые.
    If VBA.Len(incomingDateText) > 0 Then
        If Not private_TryParseIncomingDateWithOrderContext(incomingDateText, orderNoText, incomingDateValue) Then
            VBA.MsgBox "PrototypeNew: failed to resolve full date for '" & MOVEMENT_SOURCE_INCOMING_DATE & _
                "' from value '" & incomingDateText & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
            Exit Function
        End If
        incomingDateResolvedText = VBA.Format$(incomingDateValue, "dd.mm.yyyy")
    End If

    If VBA.Len(incomingNoText) > 0 Then basisDetailsText = "вх. № " & incomingNoText
    If VBA.Len(incomingDateResolvedText) > 0 Then
        If VBA.Len(basisDetailsText) > 0 Then basisDetailsText = basisDetailsText & " "
        basisDetailsText = basisDetailsText & "від " & incomingDateResolvedText
    End If

    outBasisSummary = "рапорт " & reporterText
    If VBA.Len(basisDetailsText) > 0 Then outBasisSummary = outBasisSummary & " (" & basisDetailsText & ")"
    outBasisSummary = outBasisSummary & "."
    private_TryBuildMovementBasisSummary = True
End Function

Private Function private_TryGetRequiredSourceText( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal columnName As String, _
    ByRef outValue As Variant _
) As Boolean
    Dim columnIndex As Long

    outValue = VBA.vbNullString
    columnIndex = private_GetSourceColumnIndex(sourceTable, columnName)
    If columnIndex <= 0 Then
        private_LogError "Movement source column is missing: '" & private_EscapeForLog(columnName) & "'."
        VBA.MsgBox "PrototypeNew: Movement export requires source column '" & columnName & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    outValue = sourceRow.GetCellValue(columnIndex)
    private_TryGetRequiredSourceText = True
End Function

Private Function private_GetOptionalSourceText( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal columnName As String _
) As Variant
    Dim columnIndex As Long

    columnIndex = private_GetSourceColumnIndex(sourceTable, columnName)
    If columnIndex <= 0 Then
        private_GetOptionalSourceText = VBA.vbNullString
        Exit Function
    End If

    private_GetOptionalSourceText = sourceRow.GetCellValue(columnIndex)
End Function

Private Function private_TryGetSourceTextByAnyColumn( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByRef outValue As String, _
    ParamArray columnNames() As Variant _
) As Boolean
    Dim columnName As Variant
    Dim valueText As String

    outValue = VBA.vbNullString
    If sourceTable Is Nothing Then Exit Function
    If sourceRow Is Nothing Then Exit Function

    For Each columnName In columnNames
        valueText = VBA.Trim$(VBA.CStr(private_GetOptionalSourceText(sourceTable, sourceRow, VBA.CStr(columnName))))
        If VBA.Len(valueText) > 0 Then
            outValue = valueText
            private_TryGetSourceTextByAnyColumn = True
            Exit Function
        End If
    Next columnName
End Function

Private Function private_GetSourceColumnIndex(ByVal sourceTable As obj_TableDynamic, ByVal columnName As String) As Long
    Dim sourceColIndex As Long
    Dim sourceColumn As obj_Column
    Dim expectedName As String

    If sourceTable Is Nothing Then Exit Function
    If sourceTable.TryGetColumnIndexByAlias(columnName, private_GetSourceColumnIndex) Then Exit Function
    If sourceTable.TryGetColumnIndexByName(columnName, private_GetSourceColumnIndex) Then Exit Function

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
    private_IsSelfReportText = (VBA.Len(valueText) = 0 Or _
        VBA.StrComp(valueText, "сам", VBA.vbTextCompare) = 0)
End Function

Private Function private_LowerFirstLetter(ByVal valueText As String) As String
    If VBA.Len(valueText) = 0 Then Exit Function
    private_LowerFirstLetter = VBA.LCase$(VBA.Left$(valueText, 1)) & VBA.Mid$(valueText, 2)
End Function

Private Function private_AbbreviateMovementBasisText(ByVal valueText As String) As String
    valueText = private_CutTextBeforeMarker(valueText, "; відпускний квиток")
    valueText = private_CutTextBeforeMarker(valueText, "; посвідчення про відрядження")

    valueText = VBA.Replace(valueText, "тимчасово виконуючого обов'язки", "ТВО", 1, -1, VBA.vbTextCompare)
    valueText = VBA.Replace(valueText, "начальника медичного пункту", "НМП", 1, -1, VBA.vbTextCompare)
    valueText = VBA.Replace(valueText, "командира ", "ком. ", 1, -1, VBA.vbTextCompare)
    valueText = VBA.Replace(valueText, " механізованого батальйону", " мб", 1, -1, VBA.vbTextCompare)
    valueText = VBA.Replace(valueText, " танкового батальйону", " тб", 1, -1, VBA.vbTextCompare)
    valueText = VBA.Replace(valueText, " стрілецького батальйону", " сб", 1, -1, VBA.vbTextCompare)
    valueText = VBA.Replace(valueText, "довідка військово-лікарської комісії", "довідка ВЛК", 1, -1, VBA.vbTextCompare)
    valueText = VBA.Replace(valueText, "медичної карти стаціонарного хворого", "МКСХ", 1, -1, VBA.vbTextCompare)
    valueText = VBA.Replace(valueText, "медична карта стаціонарного хворого", "МКСХ", 1, -1, VBA.vbTextCompare)
    valueText = VBA.Replace(valueText, "військової частини", "вч.", 1, -1, VBA.vbTextCompare)

    valueText = private_NormalizeInlineText(valueText)
    If VBA.Len(valueText) > 0 Then
        If VBA.Right$(valueText, 1) = "." Then valueText = VBA.Left$(valueText, VBA.Len(valueText) - 1)
    End If
    private_AbbreviateMovementBasisText = valueText
End Function

Private Function private_CutTextBeforeMarker(ByVal valueText As String, ByVal markerText As String) As String
    Dim markerPos As Long

    markerPos = VBA.InStr(1, valueText, markerText, VBA.vbTextCompare)
    If markerPos > 0 Then valueText = VBA.Left$(valueText, markerPos - 1)
    private_CutTextBeforeMarker = VBA.Trim$(valueText)
End Function

Private Function private_NormalizeInlineText(ByVal valueText As String) As String
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    valueText = VBA.Replace(valueText, VBA.ChrW(160), " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    private_NormalizeInlineText = VBA.Trim$(valueText)
End Function

Private Function private_TryGetAppendRowRange( _
    ByVal targetTable As ListObject, _
    ByRef outRowRange As Range, _
    ByRef outInsertedRow As ListRow _
) As Boolean
    Dim reusableRowIndex As Long

    Set outRowRange = Nothing
    Set outInsertedRow = Nothing
    If targetTable Is Nothing Then Exit Function

    If targetTable.ListColumns.Count < MOVEMENT_TARGET_COLUMN_COUNT Then
        private_LogError "Movement target table has fewer than 6 columns. Actual=" & VBA.CStr(targetTable.ListColumns.Count)
        VBA.MsgBox "PrototypeNew: Movement target table has fewer than 6 columns.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    reusableRowIndex = private_FindTrailingEmptyReuseRowIndex(targetTable)
    If reusableRowIndex > 0 Then
        Set outRowRange = targetTable.ListRows.Item(reusableRowIndex).Range
        private_TryGetAppendRowRange = Not outRowRange Is Nothing
        Exit Function
    End If

    Set outInsertedRow = targetTable.ListRows.Add
    If outInsertedRow Is Nothing Then
        private_LogError "Movement append row returned Nothing."
        VBA.MsgBox "PrototypeNew: failed to append Movement row to target table.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    Set outRowRange = outInsertedRow.Range
    private_TryGetAppendRowRange = Not outRowRange Is Nothing
End Function

Private Function private_FindTrailingEmptyReuseRowIndex(ByVal targetTable As ListObject) As Long
    Dim firstScanRowIndex As Long
    Dim rowIndex As Long
    Dim firstTrailingEmptyRowIndex As Long

    If targetTable Is Nothing Then Exit Function
    If targetTable.ListRows.Count <= 0 Then Exit Function

    firstScanRowIndex = targetTable.ListRows.Count - MOVEMENT_TRAILING_EMPTY_LOOKBACK_ROWS + 1
    If firstScanRowIndex < 1 Then firstScanRowIndex = 1

    For rowIndex = targetTable.ListRows.Count To firstScanRowIndex Step -1
        If private_IsRowEmpty(targetTable.ListRows.Item(rowIndex).Range) Then
            firstTrailingEmptyRowIndex = rowIndex
        Else
            Exit For
        End If
    Next rowIndex

    private_FindTrailingEmptyReuseRowIndex = firstTrailingEmptyRowIndex
End Function

Private Function private_IsRowEmpty(ByVal rowRange As Range) As Boolean
    Dim cellObj As Range
    Dim valueText As String

    If rowRange Is Nothing Then Exit Function

    For Each cellObj In rowRange.Cells
        If cellObj.HasFormula Then GoTo ContinueCell
        valueText = VBA.Trim$(VBA.CStr(cellObj.Value2))
        If VBA.Len(valueText) > 0 Then Exit Function

ContinueCell:
    Next cellObj

    private_IsRowEmpty = True
End Function

Private Function private_TryWriteMovementRow( _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByRef targetValues As Variant, _
    ByVal outgoingOrderNo As Variant, _
    ByVal outgoingFoodFromDate As Variant, _
    ByVal outgoingDepartureDate As Variant, _
    ByVal writeSpecialFields As Boolean, _
    ByVal specialDurationValue As Variant, _
    ByVal specialVkNoValue As Variant, _
    ByVal writeMappedEvent As Boolean, _
    ByVal mappedEventText As String, _
    ByVal basisSummaryText As String, _
    ByVal tvoFioText As String, _
    ByVal tvoIpnText As String, _
    ByVal tvoPositionText As String _
) As Boolean
    Dim columnIndex As Long

    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function
    If rowRange.Columns.Count < MOVEMENT_TARGET_COLUMN_COUNT Then
        private_LogError "Movement target row has fewer than 6 cells. Actual=" & VBA.CStr(rowRange.Columns.Count)
        VBA.MsgBox "PrototypeNew: Movement target row has fewer than 6 cells.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    For columnIndex = 1 To MOVEMENT_TARGET_COLUMN_COUNT
        If Not private_TryWriteCellValueWithFormulaPolicy(rowRange.Cells(1, columnIndex), targetValues(1, columnIndex), "movement:write-cell col=" & VBA.CStr(columnIndex)) Then Exit Function
    Next columnIndex

    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_ORDER_NO, outgoingOrderNo) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_FOOD_FROM, outgoingFoodFromDate) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_DEPARTURE, outgoingDepartureDate) Then Exit Function

    If writeSpecialFields Then
        If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_DURATION_TERM, specialDurationValue) Then Exit Function
        If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_ESCORT_DOCUMENT, specialVkNoValue) Then Exit Function
    End If

    If writeMappedEvent Then
        If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_EVENT, mappedEventText) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(basisSummaryText)) > 0 Then
        If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_OUT_REASON, basisSummaryText) Then Exit Function
    End If
    ' Пишем колонки даже при пустой цепочке, чтобы новая строка имела полный
    ' и предсказуемый набор значений ТВО.
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_TVO_FIO, tvoFioText) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_TVO_IPN, tvoIpnText) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_TVO_POSITION, tvoPositionText) Then Exit Function

    private_TryWriteMovementRow = True
End Function

Private Function private_TryMapSectionTypeToEventText( _
    ByVal rawSectionType As String, _
    ByRef outEventText As String _
) As Boolean
    private_TryMapSectionTypeToEventText = m_Data.TryMapMovementSectionTypeToEventText(rawSectionType, outEventText)
End Function

Private Function private_TryBuildMovementOutgoingValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal context As Object, _
    ByRef outOrderNo As Variant, _
    ByRef outFoodFromDate As Variant, _
    ByRef outDepartureDate As Variant _
) As Boolean
    Dim sourceRow As obj_Row
    Dim incomingDate As Date
    Dim departureDateFromForm As Date
    Dim orderDateByNo As Date
    Dim manualOrderNo As Variant
    Dim incomingDateRaw As Variant
    Dim departureDateRaw As Variant
    Dim hasIncomingDate As Boolean
    Dim hasDepartureDateFromForm As Boolean
    Dim hasOrderDateByNo As Boolean
    Dim foodSupportDate As Date

    outOrderNo = VBA.vbNullString
    outFoodFromDate = VBA.vbNullString
    outDepartureDate = VBA.vbNullString

    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    manualOrderNo = private_GetContextText(context, MOVEMENT_CONTEXT_MANUAL_ORDER_NO)
    outOrderNo = manualOrderNo
    If VBA.Len(VBA.Trim$(VBA.CStr(outOrderNo))) = 0 Then
        outOrderNo = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_INCOMING_NO)
    End If
    incomingDateRaw = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_INCOMING_DATE)
    departureDateRaw = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_DEPARTURE_DATE)

    hasIncomingDate = private_TryParseIncomingDateWithOrderContext(incomingDateRaw, outOrderNo, incomingDate)
    If VBA.Len(VBA.Trim$(VBA.CStr(incomingDateRaw))) > 0 And Not hasIncomingDate Then
        private_LogError "Movement failed to resolve incoming date raw='" & private_EscapeForLog(VBA.CStr(incomingDateRaw)) & "' by orderNo='" & private_EscapeForLog(VBA.CStr(outOrderNo)) & "'."
        VBA.MsgBox "PrototypeNew: failed to resolve full date for '" & MOVEMENT_SOURCE_INCOMING_DATE & "' from value '" & VBA.CStr(incomingDateRaw) & "' and order number '" & VBA.CStr(outOrderNo) & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    hasDepartureDateFromForm = private_TryParseIncomingDateWithOrderContext(departureDateRaw, outOrderNo, departureDateFromForm)
    hasOrderDateByNo = private_TryResolveOrderDateFromCommonData(outOrderNo, orderDateByNo)

    If hasDepartureDateFromForm Then
        outDepartureDate = departureDateFromForm
    ElseIf hasIncomingDate Then
        outDepartureDate = incomingDate
    End If

    If m_DataProvider.CommonData.TryCalculateFoodSupportDate( _
        hasDepartureDateFromForm, departureDateFromForm, foodSupportDate) Then
        outFoodFromDate = foodSupportDate
    End If

    private_TryBuildMovementOutgoingValues = True
End Function

Private Function private_TryParseIncomingDateWithOrderContext( _
    ByVal rawIncomingDate As Variant, _
    ByVal rawOrderNo As Variant, _
    ByRef outDateValue As Date _
) As Boolean
    Dim orderDate As Date

    If private_TryResolveOrderDateFromCommonData(rawOrderNo, orderDate) Then
        private_TryParseIncomingDateWithOrderContext = ex_Helpers.fn_TryResolveDateWithContext(rawIncomingDate, orderDate, outDateValue)
    ElseIf ex_Helpers.fn_IsShortDateValue(rawIncomingDate) Then
        outDateValue = SENTINEL_SHORT_DATE
        private_TryParseIncomingDateWithOrderContext = True
    Else
        private_TryParseIncomingDateWithOrderContext = ex_Helpers.fn_TryResolveDateWithContext(rawIncomingDate, SENTINEL_SHORT_DATE, outDateValue)
    End If
End Function

Private Function private_TryResolveOrderDateFromCommonData( _
    ByVal rawOrderNo As Variant, _
    ByRef outOrderDate As Date _
) As Boolean
    ' Movement не читает карту приказов напрямую. Единый источник даты приказа
    ' живет в obj_PEB_ExptrCommonDataPrvdr, чтобы WORD/DailyScope/Movement
    ' одинаково резолвили сокращенные даты от одного OrderDate.
    If m_DataProvider Is Nothing Then Exit Function
    If m_DataProvider.CommonData Is Nothing Then Exit Function
    If Not m_DataProvider.CommonData.SetOrderNo(rawOrderNo) Then Exit Function
    If Not m_DataProvider.CommonData.HasOrderDate Then Exit Function

    outOrderDate = m_DataProvider.CommonData.OrderDate
    private_TryResolveOrderDateFromCommonData = True
End Function

Private Function private_TryWriteNamedColumnValue( _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByVal targetColumnName As String, _
    ByVal incomingValue As Variant _
) As Boolean
    Dim targetColumnIndex As Long

    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    targetColumnIndex = private_FindTargetColumnIndex(targetTable, targetColumnName)
    If targetColumnIndex <= 0 Then
        private_TryWriteNamedColumnValue = True
        Exit Function
    End If

    private_TryWriteNamedColumnValue = private_TryWriteCellValueWithFormulaPolicy( _
        rowRange.Cells(1, targetColumnIndex), _
        incomingValue, _
        "movement:write-named col='" & private_EscapeForLog(targetColumnName) & "'")
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
    ByVal incomingValue As Variant, _
    ByVal logPrefix As String _
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

Private Function private_FormatLogDateValue(ByVal valueObj As Variant) As String
    If VBA.IsDate(valueObj) Then
        private_FormatLogDateValue = VBA.Format$(VBA.CDate(valueObj), "dd.mm.yyyy")
    Else
        private_FormatLogDateValue = VBA.CStr(valueObj)
    End If
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

Private Sub private_LogInfo(ByVal message As String)
    ex_Core.fn_Diagnostic_LogInfo "movement-export: " & message
End Sub

Private Sub private_LogError(ByVal message As String)
    ex_Core.fn_Diagnostic_LogError "Movement export: " & message
End Sub

Private Function private_EscapeForLog(ByVal valueText As String) As String
    private_EscapeForLog = VBA.Replace(VBA.CStr(valueText), "'", "''")
End Function

Private Function private_BoolText(ByVal value As Boolean) As String
    If value Then
        private_BoolText = "True"
    Else
        private_BoolText = "False"
    End If
End Function
