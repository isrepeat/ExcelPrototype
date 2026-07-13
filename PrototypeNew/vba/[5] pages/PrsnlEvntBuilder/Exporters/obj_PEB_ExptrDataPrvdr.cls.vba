VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExptrDataPrvdr"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False
#Const LOGGING_VERBOSE_ENABLED = False
' Проверяет допустимость нового события по последней Movement-записи человека.
' Смены статуса и закрывающие события пока пропускаются отдельными правилами.
#Const MOVEMENT_EXPORT_VALIDATION_ENABLED = True

Private m_IsDisposed As Boolean
Private m_CommonData As obj_PEB_ExptrCommonDataPrvdr
Private m_MovementWorkbookPath As String
Private m_MovementSheetName As String
Private m_MovementRangeStartMarker As String
Private m_MovementRangeEndMarker As String
Private m_MovementConnection As Object

Private Const CONFIG_MOVEMENT_FILE_PATH_KEY As String = "Export.Movement.FilePath"
Private Const CONFIG_MOVEMENT_SHEET_NAME_KEY As String = "Export.Movement.SheetName"
Private Const CONFIG_MOVEMENT_RANGE_START_KEY As String = "Export.Movement.RangeStartMarker"
Private Const CONFIG_MOVEMENT_RANGE_END_KEY As String = "Export.Movement.RangeEndMarker"
Private Const MOVEMENT_IPN_HEADER As String = "ІПН"
Private Const MOVEMENT_EVENT_HEADER As String = "Подія"
Private Const MOVEMENT_DEPARTURE_DATE_HEADER As String = "Вибуття"
Private Const MOVEMENT_ARRIVAL_DATE_HEADER As String = "Прибуття"
Private Const MOVEMENT_ARRIVAL_ORDER_HEADER As String = "Наказ прибуття"
Private Const MOVEMENT_ON_FOOD_HEADER As String = "На продовольче"
Private Const MOVEMENT_TVO_FIO_HEADER As String = "ТВО ПІБ"
Private Const MOVEMENT_TVO_IPN_HEADER As String = "ТВО ІПН"
Private Const MOVEMENT_TVO_POSITION_HEADER As String = "ТВО Посада"
Private Const EXCEL_MAX_ROW As Long = 1048576

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
    m_MovementWorkbookPath = VBA.vbNullString
    m_MovementSheetName = VBA.vbNullString
    m_MovementRangeStartMarker = VBA.vbNullString
    m_MovementRangeEndMarker = VBA.vbNullString
    Set m_MovementConnection = Nothing

    If Not m_CommonData.Initialize() Then Exit Function
    If Not private_TryLoadConfig(configTable) Then Exit Function

    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_CommonData Is Nothing Then m_CommonData.Dispose
    Set m_CommonData = Nothing
    m_MovementWorkbookPath = VBA.vbNullString
    m_MovementSheetName = VBA.vbNullString
    m_MovementRangeStartMarker = VBA.vbNullString
    m_MovementRangeEndMarker = VBA.vbNullString
    If Not m_MovementConnection Is Nothing Then
        If m_MovementConnection.State <> 0 Then m_MovementConnection.Close
    End If
    Set m_MovementConnection = Nothing
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

' Читает последнюю физическую строку Movement-листа для указанного ИПН.
' Отсутствие строки не является ошибкой: в этом случае outFound=False.
' Критерий закрытия совпадает с obj_PEB_ExptrMovement: заполнены приказ
' прибытия, дата постановки на продовольствие и дата прибытия.
Public Function TryGetLatestMovementEvent( _
    ByVal ipnText As String, _
    ByRef outFound As Boolean, _
    ByRef outEventText As String, _
    ByRef outIsClosed As Boolean, _
    ByRef outDepartureDateText As String, _
    ByRef outArrivalDateText As String _
) As Boolean
    Dim resolvedPath As String
    Dim movementTableRef As String
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim arrivalOrderText As String
    Dim onFoodDateText As String
    outFound = False
    outEventText = VBA.vbNullString
    outIsClosed = False
    outDepartureDateText = VBA.vbNullString
    outArrivalDateText = VBA.vbNullString

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
    On Error GoTo LookupFail
    If Not private_TryGetMovementConnection(resolvedPath, conn) Then Exit Function

    ' ACE возвращает строки Excel-листа в физическом порядке. Проходим весь
    ' результат и перезаписываем outputs, поэтому после EOF остаётся последняя
    ' строка человека, аналогично обратному обходу ListObject в Movement exporter.
    sql = "SELECT " & _
        private_QuoteSqlIdentifier(MOVEMENT_EVENT_HEADER) & ", " & _
        private_QuoteSqlIdentifier(MOVEMENT_DEPARTURE_DATE_HEADER) & ", " & _
        private_QuoteSqlIdentifier(MOVEMENT_ARRIVAL_DATE_HEADER) & ", " & _
        private_QuoteSqlIdentifier(MOVEMENT_ARRIVAL_ORDER_HEADER) & ", " & _
        private_QuoteSqlIdentifier(MOVEMENT_ON_FOOD_HEADER) & _
        " FROM " & movementTableRef & _
        " WHERE " & private_BuildNormalizedSqlTextExpression(MOVEMENT_IPN_HEADER) & _
        " = " & private_AdoSqlTextLiteral(ipnText)

    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1
    Do While Not rs.EOF
        outFound = True
        outEventText = private_RecordsetFieldText(rs.Fields(0).Value)
        outDepartureDateText = private_RecordsetFieldText(rs.Fields(1).Value)
        outArrivalDateText = private_RecordsetFieldText(rs.Fields(2).Value)
        arrivalOrderText = private_RecordsetFieldText(rs.Fields(3).Value)
        onFoodDateText = private_RecordsetFieldText(rs.Fields(4).Value)
        rs.MoveNext
    Loop
    If outFound Then
        outIsClosed = (VBA.Len(VBA.Trim$(arrivalOrderText)) > 0) _
            And (VBA.Len(VBA.Trim$(onFoodDateText)) > 0) _
            And (VBA.Len(VBA.Trim$(outArrivalDateText)) > 0)
    End If
    TryGetLatestMovementEvent = True

CleanupDone:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    Set rs = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Exit Function

LookupFail:
    VBA.MsgBox "PrototypeNew: failed to read latest Movement event." & _
        VBA.vbCrLf & "Workbook: " & m_MovementWorkbookPath & _
        VBA.vbCrLf & "Sheet: " & m_MovementSheetName & _
        VBA.vbCrLf & "IPN: " & ipnText & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, "PrototypeNew / exporter data provider"
    Resume CleanupDone
End Function

' Единая предварительная проверка для всех PEB exporters. Обычное выбытие
' нельзя экспортировать, пока последняя Movement-запись человека не закрыта.
' Прибытия разрешаются: именно они закрывают предыдущую запись. Смены статуса
' временно пропускаются без проверки до появления отдельных правил.
Public Function IsExportAllowed( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal exportSectionType As String, _
    ByRef outErrorMessage As String _
) As Boolean
    Dim data As obj_PrsnlEvntBuilderData
    Dim ipnText As String
    Dim found As Boolean
    Dim previousEventText As String
    Dim previousIsClosed As Boolean
    Dim previousDepartureDateText As String
    Dim previousArrivalDateText As String
    outErrorMessage = VBA.vbNullString
#If Not MOVEMENT_EXPORT_VALIDATION_ENABLED Then
    ' Пока правила переходов между событиями не завершены, все exporters
    ' получают единое разрешение без обращения к Movement.
    IsExportAllowed = True
    Exit Function
#End If

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

    Set data = New obj_PrsnlEvntBuilderData
    If data.IsMovementMirrorTransferSectionType(exportSectionType) Then
        IsExportAllowed = True
        Exit Function
    End If
    If data.IsMovementClosingSectionType(exportSectionType) Then
        IsExportAllowed = True
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
    If Not TryGetLatestMovementEvent( _
        ipnText, found, previousEventText, previousIsClosed, _
        previousDepartureDateText, previousArrivalDateText) Then
        outErrorMessage = "Failed to read the latest Movement event for IPN '" & ipnText & "'."
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
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim tvoFioText As String
    Dim tvoIpnText As String
    Dim tvoPositionText As String
    Dim fioItems As Collection
    Dim ipnItems As Collection
    Dim positionItems As Collection
    Dim chainItem As Object
    Dim itemIndex As Long
    Dim lookupErrorNumber As Long
    Dim lookupErrorSource As String
    Dim lookupErrorDescription As String
    outFound = False
    Set outChain = New Collection
    If m_IsDisposed Then Exit Function

    ipnText = private_NormalizeLookupKey(ipnText)
    If VBA.Len(ipnText) = 0 Then
        VBA.MsgBox "PrototypeNew: Movement TVO lookup requires a non-empty IPN.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If
    If Not private_TryResolveMovementQueryContext(resolvedPath, movementTableRef) Then Exit Function

    On Error GoTo LookupFail
    If Not private_TryGetMovementConnection(resolvedPath, conn) Then Exit Function
    sql = "SELECT " & _
        private_QuoteSqlIdentifier(MOVEMENT_TVO_FIO_HEADER) & ", " & _
        private_QuoteSqlIdentifier(MOVEMENT_TVO_IPN_HEADER) & ", " & _
        private_QuoteSqlIdentifier(MOVEMENT_TVO_POSITION_HEADER) & _
        " FROM " & movementTableRef & _
        " WHERE " & private_BuildNormalizedSqlTextExpression(MOVEMENT_IPN_HEADER) & _
        " = " & private_AdoSqlTextLiteral(ipnText)

    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1
    Do While Not rs.EOF
        outFound = True
        ' SELECT задаёт стабильный порядок полей. Чтение по индексу не зависит
        ' от того, как ACE нормализует кириллические имена заголовков recordset.
        tvoFioText = private_RecordsetFieldText(rs.Fields(0).Value)
        tvoIpnText = private_RecordsetFieldText(rs.Fields(1).Value)
        tvoPositionText = private_RecordsetFieldText(rs.Fields(2).Value)
        rs.MoveNext
    Loop

    If Not outFound Then
        TryGetLatestMovementTvoChain = True
        GoTo CleanupDone
    End If

    Set fioItems = private_SplitNonEmptyLines(tvoFioText)
    Set ipnItems = private_SplitNonEmptyLines(tvoIpnText)
    Set positionItems = private_SplitNonEmptyLines(tvoPositionText)
    If fioItems.Count <> ipnItems.Count Or fioItems.Count <> positionItems.Count Then
        VBA.MsgBox "PrototypeNew: Movement TVO chain columns have different item counts." & _
            VBA.vbCrLf & "IPN: " & ipnText & _
            VBA.vbCrLf & MOVEMENT_TVO_FIO_HEADER & ": " & VBA.CStr(fioItems.Count) & _
            VBA.vbCrLf & MOVEMENT_TVO_IPN_HEADER & ": " & VBA.CStr(ipnItems.Count) & _
            VBA.vbCrLf & MOVEMENT_TVO_POSITION_HEADER & ": " & VBA.CStr(positionItems.Count), _
            VBA.vbExclamation, "PrototypeNew / exporter data provider"
        GoTo CleanupDone
    End If
    ' Movement уже хранит готовую линейную цепочку произвольной длины. Каждый
    ' следующий элемент замещает предыдущего участника на его штатной должности,
    ' указанной в той же строке ТВО Посада. Здесь не выполняется рекурсивный
    ' поиск по другим Movement-записям и не поддерживается замена ранее
    ' назначенного ТВО другим человеком: читаем только сохранённый snapshot.
    For itemIndex = 1 To fioItems.Count
        If VBA.Len(private_NormalizeLookupKey(VBA.CStr(ipnItems.Item(itemIndex)))) = 0 Then
            VBA.MsgBox "PrototypeNew: Movement TVO chain contains an empty IPN at depth " & VBA.CStr(itemIndex) & ".", VBA.vbExclamation, "PrototypeNew / exporter data provider"
            GoTo CleanupDone
        End If

        Set chainItem = VBA.CreateObject("Scripting.Dictionary")
        chainItem.CompareMode = 1
        chainItem("Depth") = itemIndex
        chainItem("FIO") = VBA.CStr(fioItems.Item(itemIndex))
        chainItem("IPN") = VBA.CStr(ipnItems.Item(itemIndex))
        chainItem("PositionCode") = VBA.CStr(positionItems.Item(itemIndex))
        outChain.Add chainItem
    Next itemIndex

    TryGetLatestMovementTvoChain = True

CleanupDone:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    Set rs = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Exit Function

LookupFail:
    ' Err нужно сохранить сразу: обращения к ADO/VBA в диагностике или cleanup
    ' могут очистить исходную ошибку и оставить в MsgBox пустое поле Error.
    lookupErrorNumber = Err.Number
    lookupErrorSource = Err.Source
    lookupErrorDescription = Err.Description
    VBA.MsgBox "PrototypeNew: failed to read Movement TVO chain." & _
        VBA.vbCrLf & "Workbook: " & m_MovementWorkbookPath & _
        VBA.vbCrLf & "Sheet: " & m_MovementSheetName & _
        VBA.vbCrLf & "IPN: " & ipnText & _
        VBA.vbCrLf & "SQL: " & sql & _
        VBA.vbCrLf & "Error " & VBA.CStr(lookupErrorNumber) & _
        " (" & lookupErrorSource & "): " & lookupErrorDescription, _
        VBA.vbExclamation, "PrototypeNew / exporter data provider"
    Resume CleanupDone
End Function

Public Function TryResolveReporterTvoPositionGenitive( _
    ByVal reporterFioText As String, _
    ByRef outPositionGenitive As String, _
    ByRef outIsTvo As Boolean _
) As Boolean
    If m_IsDisposed Then Exit Function
    If m_CommonData Is Nothing Then Exit Function

    reporterFioText = private_NormalizeLookupKey(reporterFioText)
    outPositionGenitive = VBA.vbNullString
    outIsTvo = False

    If VBA.Len(reporterFioText) = 0 Or private_IsSelfReportText(reporterFioText) Then
        TryResolveReporterTvoPositionGenitive = True
        Exit Function
    End If

    ' Внешняя таблица особового состава больше не является источником PEB. Определение того,
    ' что рапортующий сам является ТВО, будет восстановлено отдельным правилом
    ' по Movement; до этого момента используем его штатную должность формы.
    TryResolveReporterTvoPositionGenitive = True
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
    Dim conn As Object
    Dim movementLastRow As Long

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
    On Error GoTo EH
    If Not private_TryGetMovementConnection(outResolvedPath, conn) Then GoTo CleanExit
    If Not private_TryGetMovementLastRowThroughAdo(conn, movementLastRow) Then GoTo CleanExit

    outTableRef = private_BuildMovementSheetTableRef( _
        m_MovementSheetName, m_MovementRangeStartMarker, m_MovementRangeEndMarker, movementLastRow)
    If VBA.Len(outTableRef) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to build Movement table range from profile configuration.", VBA.vbExclamation, "PrototypeNew / exporter data provider"
        GoTo CleanExit
    End If

    private_TryResolveMovementQueryContext = True

CleanExit:
    On Error Resume Next
    Set conn = Nothing
    On Error GoTo 0
    Exit Function

EH:
    VBA.MsgBox "PrototypeNew: failed to resolve Movement SQL range." & _
        VBA.vbCrLf & "Workbook: " & outResolvedPath & _
        VBA.vbCrLf & "Sheet: " & m_MovementSheetName & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, "PrototypeNew / exporter data provider"
    Resume CleanExit
End Function

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
    Dim cfgParserBase As obj_CfgParserBase
    Dim configEntries As Collection
    Dim cfgMap As Object

    private_TryLoadConfig = True
    If configTable Is Nothing Then Exit Function

    Set cfgParserBase = New obj_CfgParserBase
    If Not cfgParserBase.Initialize(configTable) Then
        private_TryLoadConfig = False
        GoTo CleanExit
    End If
    If Not cfgParserBase.TryGetConfigEntries(configEntries) Then
        private_TryLoadConfig = False
        GoTo CleanExit
    End If
    If Not cfgParserBase.BuildConfigDictionary(configEntries, cfgMap) Then
        private_TryLoadConfig = False
        GoTo CleanExit
    End If

    ' Эти ключи пока только экспонируются будущим helper-ам определения
    ' предыдущего статуса. Они optional, чтобы существующие профили без
    ' Movement export продолжали инициализировать provider.
    m_MovementWorkbookPath = cfgParserBase.GetOptionalConfigValue( _
        cfgMap, _
        CONFIG_MOVEMENT_FILE_PATH_KEY, _
        VBA.vbNullString)
    m_MovementSheetName = cfgParserBase.GetOptionalConfigValue( _
        cfgMap, _
        CONFIG_MOVEMENT_SHEET_NAME_KEY, _
        VBA.vbNullString)
    m_MovementRangeStartMarker = cfgParserBase.GetOptionalConfigValue( _
        cfgMap, _
        CONFIG_MOVEMENT_RANGE_START_KEY, _
        VBA.vbNullString)
    m_MovementRangeEndMarker = cfgParserBase.GetOptionalConfigValue( _
        cfgMap, _
        CONFIG_MOVEMENT_RANGE_END_KEY, _
        VBA.vbNullString)

CleanExit:
    On Error Resume Next
    If Not cfgParserBase Is Nothing Then cfgParserBase.Dispose
    Set cfgParserBase = Nothing
    On Error GoTo 0
End Function


Private Function private_TryGetMovementConnection( _
    ByVal resolvedPath As String, _
    ByRef outConnection As Object _
) As Boolean
    Set outConnection = Nothing
    If VBA.Len(VBA.Trim$(resolvedPath)) = 0 Then Exit Function

    On Error GoTo ConnectionFail
    If m_MovementConnection Is Nothing Then
        Set m_MovementConnection = VBA.CreateObject("ADODB.Connection")
        m_MovementConnection.Open private_BuildAdoConnectionString(resolvedPath)
    ElseIf m_MovementConnection.State = 0 Then
        m_MovementConnection.Open private_BuildAdoConnectionString(resolvedPath)
    End If

    Set outConnection = m_MovementConnection
    private_TryGetMovementConnection = True
    Exit Function

ConnectionFail:
    VBA.MsgBox "PrototypeNew: failed to open Movement data source." & _
        VBA.vbCrLf & "Workbook: " & resolvedPath & _
        VBA.vbCrLf & "Error: " & Err.Description, _
        VBA.vbExclamation, "PrototypeNew / exporter data provider"
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
