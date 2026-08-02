VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExptrCfgDataPrvdr"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Private m_CommonData As obj_PEB_ExptrCommonDataPrvdr
Private m_PersonnelWorkbookPath As String
Private m_PersonnelTableRef As String
Private m_MovementWorkbookPath As String
Private m_MovementSheetName As String
Private m_MovementRangeStartMarker As String
Private m_MovementRangeEndMarker As String
Private m_MovementSnapshotPath As String
Private m_MovementSnapshotSourceModifiedAt As Date
Private m_HasMovementSnapshotSourceModifiedAt As Boolean
Private m_TemporaryFilePaths As Collection
' Provider хранит только единый engine. Он сам переключается между SQL для
' закрытого Movement и чтением Worksheet, если источник уже открыт пользователем.
Private m_QueryEngine As obj_ExtWorkbookQueryEngine

Private Const CONFIG_PERSONNEL_FILE_PATH_KEY As String = "Source.Personnel.FilePath"
Private Const CONFIG_PERSONNEL_STATE_RANGE_KEY As String = "Personnel.Sheet[StateMain].SheetName"
Private Const CONFIG_MOVEMENT_FILE_PATH_KEY As String = "Export.Movement.FilePath"
Private Const CONFIG_MOVEMENT_SHEET_NAME_KEY As String = "Export.Movement.SheetName"
Private Const CONFIG_MOVEMENT_RANGE_START_KEY As String = "Export.Movement.RangeStartMarker"
Private Const CONFIG_MOVEMENT_RANGE_END_KEY As String = "Export.Movement.RangeEndMarker"
Private Const MOVEMENT_IPN_HEADER As String = "ІПН"
Private Const MOVEMENT_EVENT_HEADER As String = "Подія"
Private Const MOVEMENT_DEPARTURE_DATE_HEADER As String = "Вибуття"
Private Const MOVEMENT_DEPARTURE_ORDER_HEADER As String = "Наказ вибуття"
Private Const MOVEMENT_ARRIVAL_DATE_HEADER As String = "Прибуття"
Private Const MOVEMENT_ARRIVAL_ORDER_HEADER As String = "Наказ прибуття"
Private Const MOVEMENT_ON_FOOD_HEADER As String = "На продовольче"
Private Const MOVEMENT_TVO_FIO_HEADER As String = "ТВО ПІБ"
Private Const MOVEMENT_TVO_IPN_HEADER As String = "ТВО ІПН"
Private Const MOVEMENT_TVO_POSITION_HEADER As String = "ТВО Посада"
Private Const MOVEMENT_ESCORT_DOCUMENT_HEADER As String = "Супровідний документ"
Private Const PERSONNEL_TVO_HEADER As String = "ТВО"
Private Const PERSONNEL_UNIT_HEADER As String = "#"
Private Const PERSONNEL_POSITION_CODE_HEADER As String = "Код посади"
Private Const PERSONNEL_POSITION_NAME_HEADER As String = "Повна назва посади"
Private Const REPORT_POSITION_CODE_ALIAS As String = "_ReportPositionCode"
Private Const EXCEL_MAX_ROW As Long = 20000

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
' // API
' //
Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    m_IsDisposed = False
    Set m_CommonData = New obj_PEB_ExptrCommonDataPrvdr
    m_PersonnelWorkbookPath = VBA.vbNullString
    m_PersonnelTableRef = VBA.vbNullString
    m_MovementWorkbookPath = VBA.vbNullString
    m_MovementSheetName = VBA.vbNullString
    m_MovementRangeStartMarker = VBA.vbNullString
    m_MovementRangeEndMarker = VBA.vbNullString
    m_MovementSnapshotPath = VBA.vbNullString
    m_MovementSnapshotSourceModifiedAt = 0
    m_HasMovementSnapshotSourceModifiedAt = False
    Set m_TemporaryFilePaths = New Collection
    Set m_QueryEngine = New obj_ExtWorkbookQueryEngine

    If Not m_CommonData.Initialize() Then Exit Function
    If Not m_QueryEngine.Initialize Then Exit Function
    If Not private_TryLoadConfig(configTable) Then Exit Function

    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_CommonData Is Nothing Then m_CommonData.Dispose
    Set m_CommonData = Nothing
    If Not m_QueryEngine Is Nothing Then m_QueryEngine.Dispose
    Set m_QueryEngine = Nothing
    ' ADO должен освободить handle раньше удаления snapshot. Удаляем только
    ' файлы, созданные этим экземпляром provider-а, а не произвольные *.tmp.xlsx
    ' в каталоге пользователя.
    private_DeleteTemporaryFiles
    m_PersonnelWorkbookPath = VBA.vbNullString
    m_PersonnelTableRef = VBA.vbNullString
    m_MovementWorkbookPath = VBA.vbNullString
    m_MovementSheetName = VBA.vbNullString
    m_MovementRangeStartMarker = VBA.vbNullString
    m_MovementRangeEndMarker = VBA.vbNullString
    m_MovementSnapshotPath = VBA.vbNullString
    m_MovementSnapshotSourceModifiedAt = 0
    m_HasMovementSnapshotSourceModifiedAt = False
    Set m_TemporaryFilePaths = Nothing
    On Error GoTo 0
End Sub

Public Property Get CommonData() As obj_PEB_ExptrCommonDataPrvdr
    Set CommonData = m_CommonData
End Property

' Настройки журнала движения сохраняются в общем provider-е, чтобы будущие
' проверки могли прочитать последнюю запись человека и определить его
' предыдущий статус до запуска конкретного экспортера.
Public Property Get MovementWorkbookPath() As String
    MovementWorkbookPath = m_MovementWorkbookPath
End Property

Public Property Get MovementSheetName() As String
    MovementSheetName = m_MovementSheetName
End Property

' Возвращает все строки и все фактические колонки Movement для одного ИПН.
' Этот read-only API предназначен для встроенного preview на PEB-странице;
' export validation продолжает использовать узкий запрос последней строки.
Public Function TryGetMovementHistoryByIpn( _
    ByVal ipnText As String, _
    ByRef outTable As obj_TableDynamic, _
    Optional ByVal maxRows As Long = 0 _
) As Boolean
    Dim resolvedPath As String
    Dim snapshotPath As String
    Dim movementTableRef As String
    Dim query As obj_ExtWorkbookQuery

    Set outTable = Nothing
    If m_IsDisposed Then Exit Function

    ipnText = private_NormalizeLookupKey(ipnText)
    If VBA.Len(ipnText) = 0 Then
        VBA.MsgBox "Для перегляду історії руху заповніть поле 'ІПН'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement"
        Exit Function
    End If

    If Not private_TryResolveMovementQueryContext( _
        resolvedPath, movementTableRef) Then Exit Function
    If m_QueryEngine Is Nothing Then Exit Function
    ' История является read-only preview сохранённого состояния Movement.
    ' Даже если оригинал открыт и содержит live-изменения, snapshot копируется
    ' с диска и намеренно не включает значения до сохранения пользователем.
    If Not private_TryGetMovementSnapshotPath( _
        resolvedPath, snapshotPath) Then Exit Function

    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = snapshotPath
    query.TableRef = movementTableRef
    query.SelectAllColumns = True
    If maxRows > 0 Then
        ' Сначала читаем совпадения в естественном порядке. После запроса
        ' оставляем последние N строк: так UI получает последние события,
        ' но показывает их хронологически, а не в обратном порядке.
        query.MaxRows = 0
    End If
    If Not query.AddCondition( _
        MOVEMENT_IPN_HEADER, _
        en_ExtWorkbookQueryOp.ExtQueryOpEquals, _
        ipnText, _
        True) Then Exit Function

    If Not m_QueryEngine.TryExecute(query, outTable) Then Exit Function
    If outTable Is Nothing Then Exit Function
    If maxRows > 0 Then
        If Not private_TryKeepLastRows(outTable, maxRows) Then Exit Function
    End If
    outTable.SectionTitle = private_BuildMovementHistorySectionTitle(resolvedPath)
    TryGetMovementHistoryByIpn = True
End Function

Private Function private_TryKeepLastRows( _
    ByRef tableToTrim As obj_TableDynamic, _
    ByVal maxRows As Long _
) As Boolean
    Dim trimmedTable As obj_TableDynamic
    Dim tableColumn As obj_Column
    Dim sourceRow As obj_Row
    Dim clonedRow As obj_Row
    Dim clonedRowObj As Object
    Dim firstRowIndex As Long
    Dim columnIndex As Long
    Dim rowIndex As Long

    If tableToTrim Is Nothing Then Exit Function
    If maxRows <= 0 Or tableToTrim.RowCount <= maxRows Then
        private_TryKeepLastRows = True
        Exit Function
    End If

    Set trimmedTable = New obj_TableDynamic
    If Not trimmedTable.Initialize Then Exit Function
    For columnIndex = 1 To tableToTrim.ColumnCount
        Set tableColumn = tableToTrim.Columns.Item(columnIndex)
        If tableColumn Is Nothing Then Exit Function
        If Not trimmedTable.PushColumn(tableColumn) Then Exit Function
    Next columnIndex

    firstRowIndex = tableToTrim.RowCount - maxRows + 1
    For rowIndex = firstRowIndex To tableToTrim.RowCount
        Set sourceRow = tableToTrim.Rows.Item(rowIndex)
        If sourceRow Is Nothing Then Exit Function
        Set clonedRowObj = sourceRow.Clone(tableToTrim.ColumnCount)
        If clonedRowObj Is Nothing Then Exit Function
        If Not TypeOf clonedRowObj Is obj_Row Then Exit Function
        Set clonedRow = clonedRowObj
        If Not trimmedTable.PushRow(clonedRow) Then Exit Function
    Next rowIndex

    Set tableToTrim = trimmedTable
    private_TryKeepLastRows = True
End Function

' Ищет не последнюю физическую строку, а последнюю запись человека, которая
' действительно содержит обе части данных предыдущего отпускного билета.
' Запрос выполняется по snapshot сохранённого Movement, поэтому live-строка,
' добавленная перед WORD export и ещё не сохранённая, не скрывает старый билет.
Public Function TryGetLatestMovementVacationTicket( _
    ByVal ipnText As String, _
    ByRef outFound As Boolean, _
    ByRef outEscortDocumentText As String, _
    ByRef outDepartureOrderText As String _
) As Boolean
    Dim resolvedPath As String
    Dim snapshotPath As String
    Dim movementTableRef As String
    Dim query As obj_ExtWorkbookQuery
    Dim resultTable As obj_TableDynamic
    Dim resultRow As obj_Row
    Dim rowIndex As Long
    Dim escortDocumentText As String
    Dim departureOrderText As String

    outFound = False
    outEscortDocumentText = VBA.vbNullString
    outDepartureOrderText = VBA.vbNullString
    If m_IsDisposed Then Exit Function

    ipnText = private_NormalizeLookupKey(ipnText)
    If VBA.Len(ipnText) = 0 Then Exit Function
    If Not private_TryResolveMovementQueryContext( _
        resolvedPath, movementTableRef) Then Exit Function
    If m_QueryEngine Is Nothing Then Exit Function
    If Not private_TryGetMovementSnapshotPath( _
        resolvedPath, snapshotPath) Then Exit Function

    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = snapshotPath
    query.TableRef = movementTableRef
    query.ReverseOrder = True
    If Not query.AddCondition( _
        MOVEMENT_IPN_HEADER, _
        en_ExtWorkbookQueryOp.ExtQueryOpEquals, _
        ipnText, _
        True) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_ESCORT_DOCUMENT_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_DEPARTURE_ORDER_HEADER) Then Exit Function
    If Not m_QueryEngine.TryExecute(query, resultTable) Then Exit Function
    If resultTable Is Nothing Then Exit Function

    ' Engine уже возвращает строки в обратном физическом порядке. Проверяем
    ' обе колонки вместе, чтобы номер и дата не были взяты из разных событий.
    For rowIndex = 1 To resultTable.RowCount
        Set resultRow = resultTable.Rows.Item(rowIndex)
        If resultRow Is Nothing Then GoTo ContinueRow
        escortDocumentText = VBA.vbNullString
        departureOrderText = VBA.vbNullString
        If Not resultRow.TryGetCellValueByColumn( _
            MOVEMENT_ESCORT_DOCUMENT_HEADER, escortDocumentText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn( _
            MOVEMENT_DEPARTURE_ORDER_HEADER, departureOrderText) Then Exit Function
        escortDocumentText = VBA.Trim$(escortDocumentText)
        departureOrderText = VBA.Trim$(departureOrderText)
        If VBA.Len(escortDocumentText) > 0 And _
            VBA.Len(departureOrderText) > 0 Then
            outEscortDocumentText = escortDocumentText
            outDepartureOrderText = departureOrderText
            outFound = True
            Exit For
        End If
ContinueRow:
    Next rowIndex

    TryGetLatestMovementVacationTicket = True
End Function

' Совместимый узкий API для callers, которым данные ТВО не нужны.
' Основной export-flow вызывает объединённый helper напрямую и переиспользует
' outTvoChain; эта обёртка сохранена только для прежнего публичного контракта.
Public Function TryGetLatestMovementEvent( _
    ByVal ipnText As String, _
    ByRef outFound As Boolean, _
    ByRef outEventText As String, _
    ByRef outIsClosed As Boolean, _
    ByRef outDepartureDateText As String, _
    ByRef outArrivalDateText As String _
) As Boolean
    Dim ignoredTvoChain As Collection
    Dim ignoredLatestRecord As Object

    TryGetLatestMovementEvent = TryGetLatestMovementEventAndTvoChain( _
        ipnText, _
        outFound, _
        outEventText, _
        outIsClosed, _
        outDepartureDateText, _
        outArrivalDateText, _
        ignoredTvoChain, _
        ignoredLatestRecord)
End Function

' Читает одним запросом последнюю физическую строку Movement-листа и цепочку ТВО.
' Отсутствие строки не является ошибкой: в этом случае outFound=False.
' Критерий закрытия совпадает с obj_PEB_ExptrMovement: заполнены приказ
' прибытия, дата постановки на продовольствие и дата прибытия.
Public Function TryGetLatestMovementEventAndTvoChain( _
    ByVal ipnText As String, _
    ByRef outFound As Boolean, _
    ByRef outEventText As String, _
    ByRef outIsClosed As Boolean, _
    ByRef outDepartureDateText As String, _
    ByRef outArrivalDateText As String, _
    ByRef outTvoChain As Collection, _
    ByRef outLatestRecord As Object _
) As Boolean
    Dim resolvedPath As String
    Dim movementTableRef As String
    Dim query As obj_ExtWorkbookQuery
    Dim resultTable As obj_TableDynamic
    Dim resultRow As obj_Row
    Dim arrivalOrderText As String
    Dim onFoodDateText As String
    Dim tvoFioText As String
    Dim tvoIpnText As String
    Dim tvoPositionText As String
    Dim departureOrderText As String
    Dim escortDocumentText As String

    outFound = False
    outEventText = VBA.vbNullString
    outIsClosed = False
    outDepartureDateText = VBA.vbNullString
    outArrivalDateText = VBA.vbNullString
    Set outTvoChain = New Collection
    Set outLatestRecord = VBA.CreateObject("Scripting.Dictionary")
    outLatestRecord.CompareMode = 1

    If m_IsDisposed Then Exit Function
    ipnText = private_NormalizeLookupKey(ipnText)
    If VBA.Len(ipnText) = 0 Then
        VBA.MsgBox "PrototypeNew: Movement history lookup requires a non-empty IPN.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_MovementWorkbookPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Movement.FilePath' is empty.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_MovementSheetName)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Movement.SheetName' is empty.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_MovementRangeStartMarker)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Movement.RangeStartMarker' is empty.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_MovementRangeEndMarker)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Movement.RangeEndMarker' is empty.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    If Not private_TryResolveMovementQueryContext(resolvedPath, movementTableRef) Then Exit Function
    If m_QueryEngine Is Nothing Then Exit Function
    ' Описание запроса не зависит от текущего состояния Movement-книги.
    ' ReverseOrder + MaxRows=1 означает последнюю физическую запись этого ИПН
    ' как для SQL recordset, так и для массива открытого Worksheet.
    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = resolvedPath
    query.TableRef = movementTableRef
    If Not query.AddCondition(MOVEMENT_IPN_HEADER, en_ExtWorkbookQueryOp.ExtQueryOpEquals, ipnText, True) Then Exit Function
    query.ReverseOrder = True
    query.MaxRows = 1
    If Not query.AddSelectColumn(MOVEMENT_EVENT_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_DEPARTURE_DATE_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_DEPARTURE_ORDER_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_ARRIVAL_DATE_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_ARRIVAL_ORDER_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_ON_FOOD_HEADER) Then Exit Function
    ' Валидация и восстановление ТВО используют одну и ту же последнюю строку.
    ' Читаем TVO-колонки тем же запросом, чтобы exporter не сканировал Movement
    ' повторно после успешной проверки IsExportAllowed.
    If Not query.AddSelectColumn(MOVEMENT_TVO_FIO_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_TVO_IPN_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_TVO_POSITION_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_ESCORT_DOCUMENT_HEADER) Then Exit Function
    If Not m_QueryEngine.TryExecute(query, resultTable) Then Exit Function

    If Not resultTable Is Nothing Then outFound = (resultTable.RowCount > 0)
    If outFound Then
        Set resultRow = resultTable.Rows.Item(1)
        If resultRow Is Nothing Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_EVENT_HEADER, outEventText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_DEPARTURE_DATE_HEADER, outDepartureDateText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_DEPARTURE_ORDER_HEADER, departureOrderText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_ARRIVAL_DATE_HEADER, outArrivalDateText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_ARRIVAL_ORDER_HEADER, arrivalOrderText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_ON_FOOD_HEADER, onFoodDateText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_TVO_FIO_HEADER, tvoFioText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_TVO_IPN_HEADER, tvoIpnText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_TVO_POSITION_HEADER, tvoPositionText) Then Exit Function
        If Not resultRow.TryGetCellValueByColumn(MOVEMENT_ESCORT_DOCUMENT_HEADER, escortDocumentText) Then Exit Function
        outLatestRecord(MOVEMENT_DEPARTURE_ORDER_HEADER) = departureOrderText
        outLatestRecord(MOVEMENT_ESCORT_DOCUMENT_HEADER) = escortDocumentText
        If Not private_TryBuildTvoChain( _
            tvoFioText, _
            tvoIpnText, _
            tvoPositionText, _
            ipnText, _
            outTvoChain) Then Exit Function
    End If
    If outFound Then
        outIsClosed = (VBA.Len(VBA.Trim$(arrivalOrderText)) > 0) _
            And (VBA.Len(VBA.Trim$(onFoodDateText)) > 0) _
            And (VBA.Len(VBA.Trim$(outArrivalDateText)) > 0)
    End If
    TryGetLatestMovementEventAndTvoChain = True
End Function

' Единая предварительная проверка для всех PEB exporters. Обычное выбытие
' разрешено, если прошлой записи нет либо она закрыта. Прибытие, наоборот,
' может закрывать только существующую открытую запись. Смены статуса временно
' пропускаются без проверки до появления отдельных правил.
Public Function IsExportAllowed( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal exportSectionType As String, _
    ByRef outErrorMessage As String, _
    ByRef outLatestTvoChain As Collection, _
    ByRef outLatestMovementRecord As Object, _
    Optional ByVal applyValidationRules As Boolean = True _
) As Boolean
    Dim data As obj_PrsnlEvntBuilderData
    Dim ipnText As String
    Dim found As Boolean
    Dim previousEventText As String
    Dim previousIsClosed As Boolean
    Dim previousDepartureDateText As String
    Dim previousArrivalDateText As String
    Dim requiredPreviousEventText As String

    outErrorMessage = VBA.vbNullString
    Set outLatestTvoChain = New Collection
    Set outLatestMovementRecord = VBA.CreateObject("Scripting.Dictionary")
    outLatestMovementRecord.CompareMode = 1
    If m_IsDisposed Then
        outErrorMessage = "Exporter data provider is disposed."
        Exit Function
    End If
    If sourceTable Is Nothing Then
        outErrorMessage = "Export source table is empty."
        Exit Function
    End If
    If sourceTable.RowCount <= 0 Then
        outErrorMessage = "Export source table is empty."
        Exit Function
    End If

    exportSectionType = VBA.Trim$(exportSectionType)
    If VBA.Len(exportSectionType) = 0 Then
        outErrorMessage = "Export SectionType is empty."
        Exit Function
    End If

    If Not private_TryGetFirstRowText(sourceTable, MOVEMENT_IPN_HEADER, ipnText) Then
        outErrorMessage = "Export source does not contain a valid IPN."
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(ipnText)) = 0 Then
        outErrorMessage = "Export source IPN is empty."
        Exit Function
    End If
    If Not TryGetLatestMovementEventAndTvoChain( _
        ipnText, found, previousEventText, previousIsClosed, _
        previousDepartureDateText, previousArrivalDateText, outLatestTvoChain, outLatestMovementRecord) Then
        outErrorMessage = "Failed to read the latest Movement event for IPN '" & ipnText & "'."
        Exit Function
    End If

    ' Даже при отключённой блокирующей валидации snapshot последней Movement-
    ' записи читается выше: WORD и DailyScope используют его для ТВО и данных
    ' предыдущего события. Флаг отключает только правила opening/closing.
    If Not applyValidationRules Then
        IsExportAllowed = True
        Exit Function
    End If

    Set data = New obj_PrsnlEvntBuilderData
    ' Большинство mirror-переходов допускает общий pipeline закрытия. Для смены
    ' вида отпуска этого недостаточно: новая запись обязана закрыть открытое
    ' событие именно исходного вида, иначе можно завершить несвязанный статус.
    If data.TryGetRequiredPreviousMovementEvent(exportSectionType, requiredPreviousEventText) Then
        If Not found Then
            outErrorMessage = "Export was stopped because there is no Movement event to close." & _
                VBA.vbCrLf & "IPN: " & ipnText & _
                VBA.vbCrLf & "Required previous event: " & requiredPreviousEventText
            Exit Function
        End If
        If previousIsClosed Then
            outErrorMessage = "Export was stopped because the latest Movement event is already closed." & _
                VBA.vbCrLf & "IPN: " & ipnText & _
                VBA.vbCrLf & "Previous event: " & previousEventText & _
                VBA.vbCrLf & "Required previous event: " & requiredPreviousEventText
            Exit Function
        End If
        If VBA.StrComp( _
            private_NormalizeLookupKey(previousEventText), _
            private_NormalizeLookupKey(requiredPreviousEventText), _
            VBA.vbTextCompare) <> 0 Then
            outErrorMessage = "Export was stopped because the latest Movement event has an unexpected type." & _
                VBA.vbCrLf & "IPN: " & ipnText & _
                VBA.vbCrLf & "Previous event: " & previousEventText & _
                VBA.vbCrLf & "Required previous event: " & requiredPreviousEventText
            Exit Function
        End If
    End If

    ' Mirror transfer по-прежнему не блокируется правилами opening/closing,
    ' но последняя строка уже прочитана и доступна экспортерам как snapshot.
    If data.IsMovementMirrorTransferSectionType(exportSectionType) Then
        IsExportAllowed = True
        Exit Function
    End If

    If data.IsMovementClosingSectionType(exportSectionType) Then
        If Not found Then
            outErrorMessage = "Export was stopped because there is no Movement event to close." & _
                VBA.vbCrLf & "IPN: " & ipnText
            Exit Function
        End If
        If previousIsClosed Then
            outErrorMessage = "Export was stopped because the latest Movement event is already closed." & _
                VBA.vbCrLf & "IPN: " & ipnText & _
                VBA.vbCrLf & "Previous event: " & previousEventText & _
                VBA.vbCrLf & "Departure date: " & previousDepartureDateText & _
                VBA.vbCrLf & "Arrival date: " & previousArrivalDateText & _
                VBA.vbCrLf & "Existing Movement values were not changed."
            Exit Function
        End If

        IsExportAllowed = True
        Exit Function
    End If

    If found And Not previousIsClosed Then
        outErrorMessage = "Export was stopped because the latest Movement event is not closed." & _
            VBA.vbCrLf & "IPN: " & ipnText & _
            VBA.vbCrLf & "Previous event: " & previousEventText & _
            VBA.vbCrLf & "Departure date: " & previousDepartureDateText & _
            VBA.vbCrLf & "Arrival date: " & previousArrivalDateText
        Exit Function
    End If

    IsExportAllowed = True
End Function

Public Function NormalizeIncomingNoForExport(ByVal incomingNoText As String) As String
    Dim explicitNumberText As String
    Dim trimmedIncomingNoText As String

    trimmedIncomingNoText = VBA.Trim$(incomingNoText)
    ' Ведущий знак № является явным указанием пользователя не применять
    ' служебные 1656/ и -в. Сам знак удаляем: WORD-шаблон уже выводит "вх. №"
    ' отдельно от значения placeholder.
    If VBA.Left$(trimmedIncomingNoText, 1) = "№" Then
        explicitNumberText = VBA.Trim$(VBA.Mid$(trimmedIncomingNoText, 2))
        NormalizeIncomingNoForExport = explicitNumberText
        Exit Function
    End If

    ' Только чистое значение из 1–5 ASCII-цифр получает служебное обрамление.
    ' Пробелы, уже существующие префиксы/суффиксы и более длинные номера
    ' считаются самостоятельным форматом и возвращаются без изменений.
    NormalizeIncomingNoForExport = incomingNoText
    If VBA.Len(incomingNoText) < 1 Or VBA.Len(incomingNoText) > 5 Then Exit Function
    If incomingNoText Like "*[!0-9]*" Then Exit Function

    NormalizeIncomingNoForExport = "1656/" & incomingNoText & "-в"
End Function

' Возвращает сохранённую в последней Movement-строке цепочку ТВО.
' Три TVO-колонки содержат синхронные многострочные списки: элемент с одним
' индексом во всех списках описывает один уровень замещения.
Public Function TryGetLatestMovementTvoChain( _
    ByVal ipnText As String, _
    ByRef outFound As Boolean, _
    ByRef outChain As Collection _
) As Boolean
    Dim resolvedPath As String
    Dim movementTableRef As String
    Dim query As obj_ExtWorkbookQuery
    Dim resultTable As obj_TableDynamic
    Dim resultRow As obj_Row
    Dim tvoFioText As String
    Dim tvoIpnText As String
    Dim tvoPositionText As String

    outFound = False
    Set outChain = New Collection
    If m_IsDisposed Then Exit Function

    ipnText = private_NormalizeLookupKey(ipnText)
    If VBA.Len(ipnText) = 0 Then
        VBA.MsgBox "PrototypeNew: Movement TVO lookup requires a non-empty IPN.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If Not private_TryResolveMovementQueryContext(resolvedPath, movementTableRef) Then Exit Function
    If m_QueryEngine Is Nothing Then Exit Function
    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = resolvedPath
    query.TableRef = movementTableRef
    If Not query.AddCondition(MOVEMENT_IPN_HEADER, en_ExtWorkbookQueryOp.ExtQueryOpEquals, ipnText, True) Then Exit Function
    query.ReverseOrder = True
    query.MaxRows = 1
    If Not query.AddSelectColumn(MOVEMENT_TVO_FIO_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_TVO_IPN_HEADER) Then Exit Function
    If Not query.AddSelectColumn(MOVEMENT_TVO_POSITION_HEADER) Then Exit Function
    If Not m_QueryEngine.TryExecute(query, resultTable) Then Exit Function

    If Not resultTable Is Nothing Then outFound = (resultTable.RowCount > 0)

    If Not outFound Then
        TryGetLatestMovementTvoChain = True
        Exit Function
    End If
    Set resultRow = resultTable.Rows.Item(1)
    If resultRow Is Nothing Then Exit Function
    If Not resultRow.TryGetCellValueByColumn(MOVEMENT_TVO_FIO_HEADER, tvoFioText) Then Exit Function
    If Not resultRow.TryGetCellValueByColumn(MOVEMENT_TVO_IPN_HEADER, tvoIpnText) Then Exit Function
    If Not resultRow.TryGetCellValueByColumn(MOVEMENT_TVO_POSITION_HEADER, tvoPositionText) Then Exit Function

    If Not private_TryBuildTvoChain(tvoFioText, tvoIpnText, tvoPositionText, ipnText, outChain) Then Exit Function

    TryGetLatestMovementTvoChain = True
End Function

' Преобразует три синхронных многострочных Movement-поля в один snapshot цепочки.
' Helper используется как отдельным TVO lookup, так и объединённым запросом
' validation + TVO, поэтому правила проверки структуры не дублируются.
Private Function private_TryBuildTvoChain( _
    ByVal tvoFioText As String, _
    ByVal tvoIpnText As String, _
    ByVal tvoPositionText As String, _
    ByVal ownerIpn As String, _
    ByRef outChain As Collection _
) As Boolean
    Dim fioItems As Collection
    Dim ipnItems As Collection
    Dim positionItems As Collection
    Dim chainItem As Object
    Dim itemIndex As Long

    Set outChain = New Collection
    Set fioItems = private_SplitNonEmptyLines(tvoFioText)
    Set ipnItems = private_SplitNonEmptyLines(tvoIpnText)
    Set positionItems = private_SplitNonEmptyLines(tvoPositionText)

    If fioItems.Count <> ipnItems.Count Or fioItems.Count <> positionItems.Count Then
        VBA.MsgBox "PrototypeNew: Movement TVO chain columns have different item counts." & _
            VBA.vbCrLf & "IPN: " & ownerIpn & _
            VBA.vbCrLf & MOVEMENT_TVO_FIO_HEADER & ": " & VBA.CStr(fioItems.Count) & _
            VBA.vbCrLf & MOVEMENT_TVO_IPN_HEADER & ": " & VBA.CStr(ipnItems.Count) & _
            VBA.vbCrLf & MOVEMENT_TVO_POSITION_HEADER & ": " & VBA.CStr(positionItems.Count), _
            VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    ' Movement хранит готовую линейную цепочку произвольной длины. Каждый
    ' следующий элемент возвращается на собственную должность PositionCode.
    For itemIndex = 1 To fioItems.Count
        If VBA.Len(private_NormalizeLookupKey(VBA.CStr(ipnItems.Item(itemIndex)))) = 0 Then
            VBA.MsgBox "PrototypeNew: Movement TVO chain contains an empty IPN at depth " & VBA.CStr(itemIndex) & ".", VBA.vbExclamation, "PrototypeNew / exporter data provider"
            Exit Function
        End If

        Set chainItem = VBA.CreateObject("Scripting.Dictionary")
        chainItem.CompareMode = 1
        chainItem("Depth") = itemIndex
        chainItem("FIO") = VBA.CStr(fioItems.Item(itemIndex))
        chainItem("IPN") = VBA.CStr(ipnItems.Item(itemIndex))
        chainItem("PositionCode") = VBA.CStr(positionItems.Item(itemIndex))
        outChain.Add chainItem
    Next itemIndex

    private_TryBuildTvoChain = True
End Function

Public Function TryResolveReporterTvoPositionGenitive( _
    ByVal reporterFioText As String, _
    ByRef outPositionGenitive As String, _
    ByRef outIsTvo As Boolean _
) As Boolean
    Dim tvoPositionCode As String

    If m_IsDisposed Then Exit Function
    If m_CommonData Is Nothing Then Exit Function

    reporterFioText = private_NormalizeLookupKey(reporterFioText)
    outPositionGenitive = VBA.vbNullString
    outIsTvo = False

    If VBA.Len(reporterFioText) = 0 Or private_IsSelfReportText(reporterFioText) Then
        TryResolveReporterTvoPositionGenitive = True
        Exit Function
    End If

    ' ШПО остается основным источником склонений. Personnel используется здесь
    ' только для отсутствующей в ШПО связи: в чьей строке колонка "ТВО"
    ' содержит ФИО рапортующего. Код должности найденной строки затем склоняется
    ' обычным CommonData provider-ом по справочнику должностей ШПО.
    If Not private_TryLookupPersonnelTvoPositionCode( _
        reporterFioText, _
        tvoPositionCode, _
        outIsTvo) Then Exit Function
    If Not outIsTvo Then
        TryResolveReporterTvoPositionGenitive = True
        Exit Function
    End If

    If Not m_CommonData.TryResolvePositionGenitive( _
        tvoPositionCode, _
        outPositionGenitive) Then Exit Function

    TryResolveReporterTvoPositionGenitive = True
End Function

' Возвращает текущую должность рапортующего и все предыдущие физические
' должности того же подразделения. Граница определяется по колонке "#".
' Запрос выполняется только по явному Ctrl+. пользователя.
Public Function TryGetReporterTvoPositionCandidates( _
    ByVal currentPositionCode As String, _
    ByRef outCandidates As obj_TableDynamic _
) As Boolean
    Dim resolvedPath As String
    Dim query As obj_ExtWorkbookQuery
    Dim lookupTable As obj_TableDynamic
    Dim lookupRow As obj_Row
    Dim personnelTable As obj_TableDynamic
    Dim personnelRow As obj_Row
    Dim candidateTable As obj_TableDynamic
    Dim candidateColumn As obj_Column
    Dim candidateRow As obj_Row
    Dim positionCodeText As String
    Dim positionNameText As String
    Dim unitCodeText As String
    Dim currentUnitCode As String
    Dim currentRowIndex As Long
    Dim sourceRowIndex As Long

    Set outCandidates = Nothing
    currentPositionCode = private_NormalizeLookupKey(currentPositionCode)
    If VBA.Len(currentPositionCode) = 0 Then
        VBA.MsgBox "Поле 'Код посади (рапорт)' порожнє. Спочатку виберіть рапортуючого.", _
            VBA.vbExclamation, "PrsnlEventBuilder / ТВО"
        Exit Function
    End If

    resolvedPath = private_ResolveWorkbookPath(m_PersonnelWorkbookPath)
    If VBA.Len(resolvedPath) = 0 Or VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: personnel source workbook was not found: " & _
            m_PersonnelWorkbookPath, VBA.vbExclamation, _
            "PrototypeNew / exporter data provider"
        Exit Function
    End If

    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = resolvedPath
    query.TableRef = m_PersonnelTableRef
    query.MaxRows = 2
    If Not query.AddSelectColumn(PERSONNEL_POSITION_CODE_HEADER) Then Exit Function
    If Not query.AddSelectColumn(PERSONNEL_UNIT_HEADER) Then Exit Function
    If Not query.AddCondition( _
        PERSONNEL_POSITION_CODE_HEADER, en_ExtWorkbookQueryOp.ExtQueryOpEquals, _
        currentPositionCode, True) Then Exit Function
    If Not m_QueryEngine.TryExecute(query, lookupTable) Then Exit Function
    If lookupTable Is Nothing Then Exit Function
    If lookupTable.RowCount = 0 Then
        VBA.MsgBox "Код посади '" & currentPositionCode & "' не знайдено у ШПС.", _
            VBA.vbExclamation, "PrsnlEventBuilder / ТВО"
        Exit Function
    End If
    If lookupTable.RowCount > 1 Then
        VBA.MsgBox "У ШПС знайдено кілька рядків із кодом посади '" & _
            currentPositionCode & "'.", VBA.vbExclamation, _
            "PrsnlEventBuilder / ТВО"
        Exit Function
    End If
    Set lookupRow = lookupTable.Rows.Item(1)
    If lookupRow Is Nothing Then Exit Function
    If Not lookupRow.TryGetCellValueByColumn( _
        PERSONNEL_UNIT_HEADER, currentUnitCode) Then Exit Function
    currentUnitCode = private_NormalizeLookupKey(currentUnitCode)
    If VBA.Len(currentUnitCode) = 0 Then
        VBA.MsgBox "Для коду посади '" & currentPositionCode & _
            "' у ШПС не заповнено колонку '#'. Неможливо визначити межі підрозділу.", _
            VBA.vbExclamation, "PrsnlEventBuilder / ТВО"
        Exit Function
    End If

    ' Второй запрос читает только текущее подразделение, а не весь диапазон
    ' ШПС. Физический порядок строк сохраняется и используется ниже.
    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = resolvedPath
    query.TableRef = m_PersonnelTableRef
    query.MaxRows = 0
    If Not query.AddSelectColumn(PERSONNEL_POSITION_CODE_HEADER) Then Exit Function
    If Not query.AddSelectColumn(PERSONNEL_POSITION_NAME_HEADER) Then Exit Function
    If Not query.AddSelectColumn(PERSONNEL_UNIT_HEADER) Then Exit Function
    If Not query.AddCondition( _
        PERSONNEL_UNIT_HEADER, en_ExtWorkbookQueryOp.ExtQueryOpEquals, _
        currentUnitCode, True) Then Exit Function
    If Not m_QueryEngine.TryExecute(query, personnelTable) Then Exit Function
    If personnelTable Is Nothing Then Exit Function

    For sourceRowIndex = 1 To personnelTable.RowCount
        Set personnelRow = personnelTable.Rows.Item(sourceRowIndex)
        If personnelRow Is Nothing Then Exit Function
        If Not personnelRow.TryGetCellValueByColumn( _
            PERSONNEL_POSITION_CODE_HEADER, positionCodeText) Then Exit Function
        If VBA.StrComp( _
            private_NormalizeLookupKey(positionCodeText), currentPositionCode, _
            VBA.vbTextCompare) = 0 Then
            currentRowIndex = sourceRowIndex
        End If
    Next sourceRowIndex
    If currentRowIndex = 0 Then
        VBA.MsgBox "Код посади '" & currentPositionCode & "' не знайдено у ШПС.", _
            VBA.vbExclamation, "PrsnlEventBuilder / ТВО"
        Exit Function
    End If

    Set candidateTable = New obj_TableDynamic
    If Not candidateTable.Initialize() Then Exit Function
    Set candidateColumn = New obj_Column
    candidateColumn.Name = "Код посади (рапорт)"
    If Not candidateColumn.AddAlias(REPORT_POSITION_CODE_ALIAS) Then Exit Function
    If Not candidateTable.PushColumn(candidateColumn) Then Exit Function
    Set candidateColumn = New obj_Column
    candidateColumn.Name = "Посада ТВО"
    If Not candidateTable.PushColumn(candidateColumn) Then Exit Function

    sourceRowIndex = currentRowIndex
    Do While sourceRowIndex >= 1
        Set personnelRow = personnelTable.Rows.Item(sourceRowIndex)
        If personnelRow Is Nothing Then Exit Function
        If Not personnelRow.TryGetCellValueByColumn( _
            PERSONNEL_UNIT_HEADER, unitCodeText) Then Exit Function
        unitCodeText = private_NormalizeLookupKey(unitCodeText)
        If VBA.Len(unitCodeText) > 0 Then
            If VBA.StrComp(unitCodeText, currentUnitCode, _
                VBA.vbTextCompare) <> 0 Then Exit Do
        End If
        If Not personnelRow.TryGetCellValueByColumn( _
            PERSONNEL_POSITION_CODE_HEADER, positionCodeText) Then Exit Function
        positionCodeText = VBA.Trim$(positionCodeText)
        If VBA.Len(positionCodeText) > 0 Then
            If Not personnelRow.TryGetCellValueByColumn( _
                PERSONNEL_POSITION_NAME_HEADER, positionNameText) Then Exit Function
            Set candidateRow = New obj_Row
            candidateRow.PushCellRaw positionCodeText
            candidateRow.PushCellRaw positionNameText
            If Not candidateTable.PushRow(candidateRow) Then Exit Function
        End If
        sourceRowIndex = sourceRowIndex - 1
    Loop

    Set outCandidates = candidateTable
    TryGetReporterTvoPositionCandidates = True
End Function

' //
' // Internal
' //
Private Function private_TryGetFirstRowText( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal columnAlias As String, _
    ByRef outValue As String _
) As Boolean
    Dim columnIndex As Long
    Dim sourceRow As obj_Row

    outValue = VBA.vbNullString
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    columnAlias = VBA.Trim$(columnAlias)
    If VBA.Len(columnAlias) = 0 Then Exit Function

    If Not sourceTable.TryGetColumnIndexByAlias(columnAlias, columnIndex) Then
        If Not sourceTable.TryGetColumnIndexByName(columnAlias, columnIndex) Then Exit Function
    End If
    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    outValue = VBA.Trim$(VBA.CStr(sourceRow.GetCellValue(columnIndex)))
    private_TryGetFirstRowText = True
End Function

Private Function private_TryResolveMovementQueryContext( _
    ByRef outResolvedPath As String, _
    ByRef outTableRef As String _
) As Boolean
    outResolvedPath = VBA.vbNullString
    outTableRef = VBA.vbNullString
    If VBA.Len(VBA.Trim$(m_MovementWorkbookPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Movement.FilePath' is empty.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_MovementSheetName)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Movement.SheetName' is empty.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_MovementRangeStartMarker)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Movement.RangeStartMarker' is empty.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_MovementRangeEndMarker)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Movement.RangeEndMarker' is empty.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    outResolvedPath = private_ResolveWorkbookPath(m_MovementWorkbookPath)
    If VBA.Len(outResolvedPath) = 0 Or VBA.Len(VBA.Dir$(outResolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: Movement workbook was not found: " & m_MovementWorkbookPath, VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    outTableRef = private_BuildMovementSheetTableRef( _
        m_MovementSheetName, m_MovementRangeStartMarker, m_MovementRangeEndMarker, EXCEL_MAX_ROW)
    If VBA.Len(outTableRef) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to build Movement table range from profile configuration.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        ' Здесь больше нет локальных ADO-объектов, требующих cleanup: соединением
        ' владеет obj_ExtWorkbookQueryEngine. При ошибке построения range просто
        ' завершаем текущую операцию.
        Exit Function
    End If

    private_TryResolveMovementQueryContext = True
End Function

Private Function private_TryGetMovementSnapshotPath( _
    ByVal sourcePath As String, _
    ByRef outSnapshotPath As String _
) As Boolean
    Dim snapshotPath As String
    Dim errorDescription As String
    Dim sourceFileChanged As Boolean

    outSnapshotPath = VBA.vbNullString
    sourcePath = VBA.Trim$(sourcePath)
    If VBA.Len(sourcePath) = 0 Then Exit Function

    ' Snapshot создаётся один раз на время жизни provider-а. Благодаря
    ' неизменному пути obj_ExtWorkbookQueryEngine переиспользует одно и то же
    ' ADO-соединение между запросами истории для разных военнослужащих.
    If VBA.Len(m_MovementSnapshotPath) > 0 Then
        sourceFileChanged = private_HasMovementSourceFileChanged(sourcePath)
        If VBA.Len(VBA.Dir$(m_MovementSnapshotPath)) > 0 And _
            Not sourceFileChanged Then
            outSnapshotPath = m_MovementSnapshotPath
            private_TryGetMovementSnapshotPath = True
            Exit Function
        End If

        ' При сохранении Movement или удалении snapshot извне сначала закрываем
        ' ADO handle. Только после этого Windows разрешит заменить временный
        ' файл по тому же пути актуальной сохранённой копией.
        If Not m_QueryEngine Is Nothing Then m_QueryEngine.Dispose
    Set m_QueryEngine = New obj_ExtWorkbookQueryEngine
    If Not m_QueryEngine.Initialize Then Exit Function
        On Error Resume Next
        If VBA.Len(VBA.Dir$(m_MovementSnapshotPath)) > 0 Then
            VBA.Kill m_MovementSnapshotPath
        End If
        On Error GoTo 0
        m_MovementSnapshotPath = VBA.vbNullString
        m_MovementSnapshotSourceModifiedAt = 0
        m_HasMovementSnapshotSourceModifiedAt = False
    End If

    snapshotPath = private_BuildMovementSnapshotPath(sourcePath)
    If VBA.Len(snapshotPath) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to build the Movement snapshot path." & _
            VBA.vbCrLf & "Source: " & sourcePath, _
            VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    On Error GoTo EH
    ' Удаляем только ожидаемую копию этого Movement. Она могла остаться после
    ' аварийного завершения Excel, когда Dispose не получил управление.
    If VBA.Len(VBA.Dir$(snapshotPath)) > 0 Then VBA.Kill snapshotPath

    ' Файловая копия читает последнее сохранённое состояние оригинала. Поэтому
    ' snapshot остаётся стабильным, даже когда открытый Movement продолжает
    ' получать несохранённые live-изменения через ListObject.
    If Not private_TryCopySnapshotFile(sourcePath, snapshotPath) Then
        Err.Raise VBA.vbObjectError + 2101, _
            TypeName(Me), _
            "The saved Movement file could not be copied."
    End If

    m_MovementSnapshotPath = snapshotPath
    ' Timestamp фиксируется в момент создания snapshot. Если пользователь
    ' позже сохранит Movement, заголовок старой копии не станет ошибочно
    ' показывать уже новую дату оригинала.
    On Error Resume Next
    m_MovementSnapshotSourceModifiedAt = VBA.FileDateTime(sourcePath)
    m_HasMovementSnapshotSourceModifiedAt = (Err.Number = 0)
    Err.Clear
    On Error GoTo EH
    If m_TemporaryFilePaths Is Nothing Then Set m_TemporaryFilePaths = New Collection
    private_TrackTemporaryFilePath snapshotPath
    outSnapshotPath = snapshotPath
    private_TryGetMovementSnapshotPath = True
    Exit Function

EH:
    errorDescription = Err.Description
    On Error Resume Next
    If VBA.Len(snapshotPath) > 0 Then
        If VBA.Len(VBA.Dir$(snapshotPath)) > 0 Then VBA.Kill snapshotPath
    End If
    On Error GoTo 0
    VBA.MsgBox "PrototypeNew: failed to create the Movement snapshot." & _
        VBA.vbCrLf & "Source: " & sourcePath & _
        VBA.vbCrLf & "Snapshot: " & snapshotPath & _
        VBA.vbCrLf & "Error: " & errorDescription, _
        VBA.vbExclamation, "PrototypeNew / exporter data provider"
End Function

Private Sub private_TrackTemporaryFilePath(ByVal temporaryFilePath As String)
    Dim trackedFilePath As Variant

    temporaryFilePath = VBA.Trim$(temporaryFilePath)
    If VBA.Len(temporaryFilePath) = 0 Then Exit Sub
    If m_TemporaryFilePaths Is Nothing Then Set m_TemporaryFilePaths = New Collection

    ' При обновлении snapshot используется тот же путь. Не добавляем его
    ' повторно, чтобы Dispose выполнял одну операцию удаления на файл.
    For Each trackedFilePath In m_TemporaryFilePaths
        If VBA.StrComp( _
            VBA.Replace$(VBA.CStr(trackedFilePath), "/", "\"), _
            VBA.Replace$(temporaryFilePath, "/", "\"), _
            VBA.vbTextCompare) = 0 Then Exit Sub
    Next trackedFilePath
    m_TemporaryFilePaths.Add temporaryFilePath
End Sub

Private Function private_HasMovementSourceFileChanged( _
    ByVal sourcePath As String _
) As Boolean
    Dim currentSourceModifiedAt As Date

    ' Если timestamp прочитать нельзя, сохраняем рабочий snapshot: ошибка
    ' проверки свежести не должна ломать повторный запрос истории.
    If Not m_HasMovementSnapshotSourceModifiedAt Then Exit Function
    On Error GoTo DateUnavailable
    currentSourceModifiedAt = VBA.FileDateTime(sourcePath)
    private_HasMovementSourceFileChanged = _
        (VBA.CDbl(currentSourceModifiedAt) <> _
            VBA.CDbl(m_MovementSnapshotSourceModifiedAt))

DateUnavailable:
End Function

Private Function private_TryCopySnapshotFile( _
    ByVal sourcePath As String, _
    ByVal snapshotPath As String _
) As Boolean
    Dim fileSystemObject As Object

    On Error GoTo EH
    Set fileSystemObject = VBA.CreateObject("Scripting.FileSystemObject")
    ' В отличие от VBA.FileCopy, FSO не отклоняет файл только потому, что книга
    ' открыта в Excel. Копируется именно содержимое на диске, а не Workbook DOM.
    fileSystemObject.CopyFile sourcePath, snapshotPath, True
    private_TryCopySnapshotFile = True
EH:
    Set fileSystemObject = Nothing
End Function

Private Function private_BuildMovementHistorySectionTitle( _
    ByVal sourcePath As String _
) As String
    Dim sourceFileDate As Date

    On Error GoTo DateUnavailable
    If m_HasMovementSnapshotSourceModifiedAt Then
        sourceFileDate = m_MovementSnapshotSourceModifiedAt
    Else
        sourceFileDate = VBA.FileDateTime(sourcePath)
    End If
    private_BuildMovementHistorySectionTitle = _
        "Історія руху — snapshot файлу від " & _
        VBA.Format$(sourceFileDate, "dd.mm.yyyy hh:nn:ss")
    Exit Function

DateUnavailable:
    ' Ошибка чтения timestamp не должна отменять уже выполненный запрос.
    private_BuildMovementHistorySectionTitle = "Історія руху — snapshot файлу"
End Function

Private Function private_BuildMovementSnapshotPath(ByVal sourcePath As String) As String
    Dim extensionPos As Long
    Dim separatorPos As Long

    sourcePath = VBA.Trim$(sourcePath)
    extensionPos = VBA.InStrRev(sourcePath, ".")
    separatorPos = VBA.InStrRev(VBA.Replace$(sourcePath, "/", "\"), "\")
    If extensionPos <= separatorPos Then Exit Function

    private_BuildMovementSnapshotPath = _
        VBA.Left$(sourcePath, extensionPos - 1) & _
        " (snapshot).tmp" & VBA.Mid$(sourcePath, extensionPos)
End Function

Private Sub private_DeleteTemporaryFiles()
    Dim fileIndex As Long
    Dim temporaryFilePath As String

    If m_TemporaryFilePaths Is Nothing Then Exit Sub

    On Error Resume Next
    For fileIndex = m_TemporaryFilePaths.Count To 1 Step -1
        temporaryFilePath = VBA.Trim$(VBA.CStr(m_TemporaryFilePaths.Item(fileIndex)))
        If VBA.Len(temporaryFilePath) > 0 Then
            If VBA.Len(VBA.Dir$(temporaryFilePath)) > 0 Then VBA.Kill temporaryFilePath
        End If
        m_TemporaryFilePaths.Remove fileIndex
    Next fileIndex
    On Error GoTo 0
End Sub

Private Function private_SplitNonEmptyLines(ByVal valueText As String) As Collection
    Dim result As Collection
    Dim normalizedText As String
    Dim parts() As String
    Dim partIndex As Long
    Dim partText As String

    Set result = New Collection
    normalizedText = VBA.Replace(VBA.CStr(valueText), VBA.vbCrLf, VBA.vbLf)
    normalizedText = VBA.Replace(normalizedText, VBA.vbCr, VBA.vbLf)
    parts = VBA.Split(normalizedText, VBA.vbLf)
    For partIndex = LBound(parts) To UBound(parts)
        partText = VBA.Trim$(VBA.CStr(parts(partIndex)))
        If VBA.Len(partText) > 0 Then result.Add partText
    Next partIndex
    Set private_SplitNonEmptyLines = result
End Function

Private Function private_TryLoadConfig(ByVal configTable As obj_ConfigTable) As Boolean
    Dim prsnlEvntBuilderCfgParser As obj_PrsnlEvntBuilderCfgParser
    Dim rawPersonnelTableRef As String

    private_TryLoadConfig = True
    If configTable Is Nothing Then Exit Function

    ' Используем единый parser режима как resolverDataContext. При этом ключи
    ' Personnel/Movement и их семантика остаются в config-dependent provider,
    ' а parser только возвращает разрешённые значения по переданному ключу.
    Set prsnlEvntBuilderCfgParser = New obj_PrsnlEvntBuilderCfgParser
    If Not prsnlEvntBuilderCfgParser.Initialize(configTable) Then
        private_TryLoadConfig = False
        GoTo CleanExit
    End If

    If Not prsnlEvntBuilderCfgParser.TryGetRequiredValue( _
        CONFIG_PERSONNEL_FILE_PATH_KEY, m_PersonnelWorkbookPath) Then
        VBA.MsgBox "PrototypeNew: required profile key '" & _
            CONFIG_PERSONNEL_FILE_PATH_KEY & "' is missing or could not be resolved.", _
            VBA.vbExclamation, "PrototypeNew / exporter config data provider"
        private_TryLoadConfig = False
        GoTo CleanExit
    End If
    If Not prsnlEvntBuilderCfgParser.TryGetRequiredValue( _
        CONFIG_PERSONNEL_STATE_RANGE_KEY, rawPersonnelTableRef) Then
        VBA.MsgBox "PrototypeNew: required profile key '" & _
            CONFIG_PERSONNEL_STATE_RANGE_KEY & "' is missing.", _
            VBA.vbExclamation, "PrototypeNew / exporter config data provider"
        private_TryLoadConfig = False
        GoTo CleanExit
    End If
    m_PersonnelTableRef = private_BuildConfiguredAdoRangeRef(rawPersonnelTableRef)
    ' Movement optional: профили без соответствующего exporter-а не обязаны
    ' объявлять его источник, но при наличии resolver он применяется тем же parser.
    m_MovementWorkbookPath = prsnlEvntBuilderCfgParser.GetOptionalValue(CONFIG_MOVEMENT_FILE_PATH_KEY)
    m_MovementSheetName = prsnlEvntBuilderCfgParser.GetOptionalValue(CONFIG_MOVEMENT_SHEET_NAME_KEY)
    m_MovementRangeStartMarker = prsnlEvntBuilderCfgParser.GetOptionalValue(CONFIG_MOVEMENT_RANGE_START_KEY)
    m_MovementRangeEndMarker = prsnlEvntBuilderCfgParser.GetOptionalValue(CONFIG_MOVEMENT_RANGE_END_KEY)

CleanExit:
    On Error Resume Next
    If Not prsnlEvntBuilderCfgParser Is Nothing Then prsnlEvntBuilderCfgParser.Dispose
    Set prsnlEvntBuilderCfgParser = Nothing
    On Error GoTo 0
End Function


' Ищет должность, обязанности по которой временно исполняет рапортующий.
' QueryEngine обеспечивает одинаковый контракт для открытого Personnel
' (чтение Range) и закрытого Personnel (ACE/ADO SQL с HDR=YES).
Private Function private_TryLookupPersonnelTvoPositionCode( _
    ByVal reporterFioText As String, _
    ByRef outPositionCode As String, _
    ByRef outFound As Boolean _
) As Boolean
    Dim resolvedPath As String
    Dim reporterSurnameText As String
    Dim query As obj_ExtWorkbookQuery
    Dim resultTable As obj_TableDynamic
    Dim resultRow As obj_Row

    outPositionCode = VBA.vbNullString
    outFound = False

    reporterSurnameText = private_GetFirstLookupWord(reporterFioText)
    If VBA.Len(reporterSurnameText) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to extract the reporter surname from '" & reporterFioText & "'.", _
            VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    If m_QueryEngine Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(m_PersonnelWorkbookPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key '" & CONFIG_PERSONNEL_FILE_PATH_KEY & "' is empty.", _
            VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_PersonnelTableRef)) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key '" & CONFIG_PERSONNEL_STATE_RANGE_KEY & "' is empty.", _
            VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    resolvedPath = private_ResolveWorkbookPath(m_PersonnelWorkbookPath)
    If VBA.Len(resolvedPath) = 0 Or VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: personnel source workbook was not found: " & m_PersonnelWorkbookPath, _
            VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = resolvedPath
    query.TableRef = m_PersonnelTableRef
    query.MaxRows = 2
    If Not query.AddSelectColumn(PERSONNEL_POSITION_CODE_HEADER) Then Exit Function
    If Not query.AddCondition( _
        PERSONNEL_TVO_HEADER, _
        en_ExtWorkbookQueryOp.ExtQueryOpContains, _
        reporterSurnameText, _
        True) Then Exit Function

    If Not m_QueryEngine.TryExecute(query, resultTable) Then Exit Function
    If resultTable Is Nothing Then Exit Function
    If resultTable.RowCount = 0 Then
        private_TryLookupPersonnelTvoPositionCode = True
        Exit Function
    End If
    If resultTable.RowCount > 1 Then
        VBA.MsgBox "PrototypeNew: reporter '" & reporterFioText & _
            "' is referenced as TVO in more than one Personnel row. Export was stopped because the TVO position is ambiguous.", _
            VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    Set resultRow = resultTable.Rows.Item(1)
    If resultRow Is Nothing Then Exit Function
    If Not resultRow.TryGetCellValueByColumn( _
        PERSONNEL_POSITION_CODE_HEADER, _
        outPositionCode) Then Exit Function

    outPositionCode = VBA.Trim$(outPositionCode)
    If VBA.Len(outPositionCode) = 0 Then
        VBA.MsgBox "PrototypeNew: Personnel row referencing reporter '" & reporterFioText & _
            "' as TVO has an empty '" & PERSONNEL_POSITION_CODE_HEADER & "' value.", _
            VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    outFound = True
    private_TryLookupPersonnelTvoPositionCode = True
End Function


' Получает ИПН рапортующего из Personnel только для резервной проверки
' Movement. Отсутствие строки не означает ошибку: рапортующий может не быть
' военнослужащим из текущего снимка ШПС.
Private Function private_GetFirstLookupWord(ByVal valueText As String) As String
    Dim separatorPos As Long

    valueText = private_NormalizeLookupKey(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    separatorPos = VBA.InStr(1, valueText, " ", VBA.vbBinaryCompare)
    If separatorPos > 1 Then
        private_GetFirstLookupWord = VBA.Left$(valueText, separatorPos - 1)
    Else
        private_GetFirstLookupWord = valueText
    End If
End Function


Private Function private_BuildAdoConnectionString(ByVal sourcePath As String) As String
    Dim ext As String
    Dim props As String

    sourcePath = VBA.Trim$(sourcePath)
    ext = VBA.LCase$(VBA.Mid$(sourcePath, VBA.InStrRev(sourcePath, ".") + 1))
    Select Case ext
        Case "xls"
            props = "Excel 8.0"
        Case "xlsx"
            props = "Excel 12.0 Xml"
        Case "xlsm"
            props = "Excel 12.0 Macro"
        Case "xlsb"
            props = "Excel 12.0"
        Case Else
            props = "Excel 12.0 Xml"
    End Select
    props = props & ";HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"

    private_BuildAdoConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePath & _
        ";Extended Properties=""" & props & """;"
End Function

Private Function private_QuoteSqlIdentifier(ByVal identifierText As String) As String
    identifierText = VBA.Replace(VBA.Trim$(identifierText), "]", "]]")
    private_QuoteSqlIdentifier = "[" & identifierText & "]"
End Function

Private Function private_BuildNormalizedSqlTextExpression(ByVal identifierText As String) As String
    Dim quotedIdentifier As String
    Dim safeValueExpression As String

    quotedIdentifier = private_QuoteSqlIdentifier(identifierText)
    safeValueExpression = "IIf(IsNull(" & quotedIdentifier & "), '', " & quotedIdentifier & ")"

    ' Excel/ACE SQL падает с "Invalid use of Null", если вызвать CStr(Null).
    ' Поэтому сначала заменяем Null на пустую строку, а уже потом чистим
    ' переносы/неразрывные пробелы для сравнения ФИО из колонки ТВО.
    private_BuildNormalizedSqlTextExpression = _
        "LCase(Trim(Replace(Replace(Replace(Replace(CStr(" & safeValueExpression & _
        "), Chr(160), ' '), Chr(13), ' '), Chr(10), ' '), Chr(9), ' ')))"
End Function

Private Function private_BuildConfiguredAdoRangeRef(ByVal rawRangeRef As String) As String
    rawRangeRef = VBA.Trim$(rawRangeRef)
    If VBA.Len(rawRangeRef) = 0 Then Exit Function

    ' Экранирование закрывающей скобки относится к ADO-контракту provider,
    ' поэтому намеренно не переносится в общий parser строковых значений.
    rawRangeRef = VBA.Replace(rawRangeRef, "]", "]]")
    private_BuildConfiguredAdoRangeRef = "[" & rawRangeRef & "]"
End Function

Private Function private_BuildMovementSheetTableRef( _
    ByVal sheetName As String, _
    ByVal rangeStartMarker As String, _
    ByVal rangeEndMarker As String, _
    ByVal lastUsedRow As Long _
) As String
    Dim endColumnText As String
    Dim rx As Object

    sheetName = VBA.Trim$(sheetName)
    If VBA.Len(sheetName) = 0 Then Exit Function
    rangeStartMarker = VBA.Replace(VBA.Trim$(rangeStartMarker), "$", VBA.vbNullString)
    rangeEndMarker = VBA.Replace(VBA.Trim$(rangeEndMarker), "$", VBA.vbNullString)
    If VBA.Len(rangeStartMarker) = 0 Or VBA.Len(rangeEndMarker) = 0 Or lastUsedRow <= 0 Then Exit Function

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "[^A-Z]"
    endColumnText = VBA.UCase$(rx.Replace(rangeEndMarker, VBA.vbNullString))
    If VBA.Len(endColumnText) = 0 Then Exit Function

    sheetName = VBA.Replace(sheetName, "]", "]]")
    private_BuildMovementSheetTableRef = "[" & sheetName & "$" & rangeStartMarker & ":" & _
        endColumnText & VBA.CStr(lastUsedRow) & "]"
End Function

Private Function private_TryGetMovementLastRowThroughAdo( _
    ByVal conn As Object, _
    ByRef outLastRow As Long _
) As Boolean
    Dim rs As Object
    Dim movementTableRef As String
    Dim headerRow As Long

    outLastRow = 0
    If conn Is Nothing Then Exit Function
    ' Начинаем range с настроенной строки заголовков ($B6). При запросе всего
    ' sheet HDR=YES использует строку 1, поэтому поле "ІПН" не существует.
    movementTableRef = private_BuildMovementSheetTableRef( _
        m_MovementSheetName, _
        m_MovementRangeStartMarker, _
        m_MovementRangeEndMarker, _
        EXCEL_MAX_ROW)
    If VBA.Len(movementTableRef) = 0 Then Exit Function
    headerRow = private_ExtractRowNumber(m_MovementRangeStartMarker)
    If headerRow <= 0 Then Exit Function

    On Error GoTo EH
    Set rs = VBA.CreateObject("ADODB.Recordset")
    ' Для RecordCount достаточно одной стабильной колонки. SELECT * заставлял
    ' ACE материализовать всю широкую Movement-таблицу перед каждой проверкой.
    rs.Open "SELECT " & private_QuoteSqlIdentifier(MOVEMENT_IPN_HEADER) & _
        " FROM " & movementTableRef, conn, 3, 1

    ' RecordCount не включает строку заголовков диапазона. Переводим число
    ' data rows обратно в физический номер строки исходного листа.
    If rs.RecordCount >= 0 Then outLastRow = headerRow + rs.RecordCount
    private_TryGetMovementLastRowThroughAdo = (outLastRow > 0)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
    Exit Function

EH:
    VBA.MsgBox "PrototypeNew: failed to resolve the last Movement row through SQL." & _
        VBA.vbCrLf & "Sheet: " & m_MovementSheetName & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, "PrototypeNew / exporter data provider"
    Resume CleanExit
End Function

Private Function private_ExtractRowNumber(ByVal cellMarker As String) As Long
    Dim rx As Object
    Dim matches As Object

    cellMarker = VBA.Replace(VBA.Trim$(cellMarker), "$", VBA.vbNullString)
    If VBA.Len(cellMarker) = 0 Then Exit Function
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "([0-9]+)$"
    Set matches = rx.Execute(cellMarker)
    If matches.Count <= 0 Then Exit Function
    private_ExtractRowNumber = VBA.CLng(matches.Item(0).SubMatches(0))
End Function

Private Function private_AdoSqlTextLiteral(ByVal valueText As String) As String
    private_AdoSqlTextLiteral = "'" & VBA.Replace(private_NormalizeLookupKey(valueText), "'", "''") & "'"
End Function

Private Function private_AdoSqlLikeContainsLiteral(ByVal valueText As String, ByVal wildcardText As String) As String
    valueText = private_NormalizeLookupKey(valueText)
    valueText = VBA.Replace(valueText, "'", "''")
    valueText = VBA.Replace(valueText, "%", "[%]")
    valueText = VBA.Replace(valueText, "_", "[_]")
    If VBA.Len(wildcardText) = 0 Then wildcardText = "%"
    private_AdoSqlLikeContainsLiteral = "'" & wildcardText & valueText & wildcardText & "'"
End Function

Private Function private_RecordsetFieldText(ByVal valueIn As Variant) As String
    If VBA.IsNull(valueIn) Or VBA.IsEmpty(valueIn) Then Exit Function
    private_RecordsetFieldText = VBA.Trim$(VBA.CStr(valueIn))
End Function

Private Function private_ResolveWorkbookPath(ByVal workbookPath As String) As String
    workbookPath = VBA.Trim$(workbookPath)
    If VBA.Len(workbookPath) = 0 Then Exit Function

    If VBA.InStr(1, workbookPath, ":", VBA.vbBinaryCompare) > 0 _
        Or VBA.Left$(workbookPath, 2) = "\\" Then
        private_ResolveWorkbookPath = workbookPath
    Else
        private_ResolveWorkbookPath = ex_XmlCore.fn_CombineBasePath(ThisWorkbook, workbookPath)
    End If
End Function


Private Function private_NormalizeLookupKey(ByVal valueText As String) As String
    valueText = VBA.Trim$(VBA.CStr(valueText))
    valueText = VBA.Replace(valueText, VBA.ChrW$(160), " ")
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    private_NormalizeLookupKey = VBA.LCase$(VBA.Trim$(valueText))
End Function

Private Function private_IsSelfReportText(ByVal valueText As String) As Boolean
    valueText = private_NormalizeLookupKey(valueText)
    private_IsSelfReportText = (VBA.StrComp(valueText, "сам", VBA.vbTextCompare) = 0)
End Function
