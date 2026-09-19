Option Explicit
Attribute VB_Name = "module_PeriodsGeneration"

#Const ENABLE_LOGGING = False
#Const ENABLE_DEBUG_LOGGING = False

' ============================================================
' Configuration
' ============================================================
' ----------------------------
' Source
' ----------------------------
Private Const SOURCE_SHEET_NAME As String = "Відсутні"
Private Const SOURCE_TABLE_NAME As String = "ТВ"

Private Const SOURCE_COL_RANK As String = "Звання"
Private Const SOURCE_COL_NAME As String = "ПІБ"
Private Const SOURCE_COL_TAX_ID As String = "ІПН"
Private Const SOURCE_COL_POSITION As String = "Посада"
Private Const SOURCE_COL_EVENT As String = "Подія"

Private Const SOURCE_COL_PERIOD_FROM As String = "Вибуття"
Private Const SOURCE_COL_PERIOD_TO As String = "Прибуття"

Private Const SOURCE_COL_DEPARTURE_ORDER As String = "Вибуття.Наказ"
Private Const SOURCE_COL_ARRIVAL_ORDER As String = "Прибуття.Наказ"

' ----------------------------
' Target
' ----------------------------
Private Const TARGET_SHEET_NAME As String = "Періоди вибуття"
Private Const TARGET_TABLE_NAME As String = "ПеріодиВибуття"

Private Const TARGET_COL_RANK As String = "Звання"
Private Const TARGET_COL_NAME As String = "ПІБ"
Private Const TARGET_COL_TAX_ID As String = "ІПН"
Private Const TARGET_COL_POSITION As String = "Посада"
Private Const TARGET_COL_EVENT As String = "Подія"

Private Const TARGET_COL_DEPARTURE_ORDER As String = "Наказ.Вибуття"
Private Const TARGET_COL_PERIOD_FROM As String = "Період.Вибуття"
Private Const TARGET_COL_PERIOD_TO As String = "Період.Прибуття"
Private Const TARGET_COL_ARRIVAL_ORDER As String = "Наказ.Прибуття"
Private Const TARGET_COL_PERIOD_COUNT As String = "Період.Кількість"

' ----------------------------
' Parameters
' ----------------------------
Private Const PARAM_PERIOD_FROM_CELL As String = "D3"
Private Const PARAM_PERIOD_TO_CELL As String = "E3"

' ----------------------------
' Merge options
' ----------------------------
' Только эти типы событий можно объединять в один непрерывный период.
' Любое другое событие создает границу периода.
Private Const MERGEABLE_EVENTS As String = _
    "Відпустка для лікування" & _
    "|Стаціонарне лікування" & _
    "|ВЛК за межами"

' Только периоды с этими событиями отображаются в результате.
Private Const DISPLAYABLE_EVENTS As String = _
    "Відпустка для лікування" & _
    "|Стаціонарне лікування" & _
    "|ВЛК за межами"

Private Const EVENTS_SEPARATOR As String = " | "
Private Const OPEN_PERIOD_TO_TEXT As String = "по теперішній час"

' ----------------------------
' UI
' ----------------------------
Private Const UI_YIELD_INTERVAL As Long = 100

' ----------------------------
' Logging
' ----------------------------
Private Const LOG_FILE_SUFFIX As String = "_logs.txt"

' ----------------------------
' Messages
' ----------------------------
Private Const MSG_SOURCE_NOT_FOUND As String ="Source workbook was not found. Open the RUKH workbook and try again."
Private Const MSG_TARGET_NOT_FOUND As String ="Target table was not found."
Private Const MSG_INVALID_PERIOD As String ="Invalid calculation period."
Private Const MSG_SOURCE_COLUMN_NOT_FOUND As String ="A required source column was not found."
Private Const MSG_TARGET_COLUMN_NOT_FOUND As String ="A required target column was not found."
Private Const MSG_NO_PERIODS As String ="No periods intersect the requested range."
Private Const MSG_DONE_PREFIX As String ="Calculation completed. Period count: "
Private Const MSG_CANCELLED As String ="Calculation cancelled."
Private Const MSG_UNEXPECTED_ERROR As String ="Unexpected error occurred."
Private Const MSG_CALCULATION_ALREADY_RUNNING As String = _
    "Period calculation is already running. Wait for it to complete."

' ============================================================
' Module state
' ============================================================
Private gCancelRequested As Boolean
Private gCalculationRunning As Boolean

' ============================================================
' Data structures
' ============================================================
Private Type PeriodInfo
    Rank As String
    PersonName As String
    taxId As String
    Position As String
    EventName As String
    LastEventName As String
    HasDisplayableEvent As Boolean
    IsOpen As Boolean

    DepartureOrder As String
    dateFrom As Date

    dateTo As Date
    ArrivalOrder As String
End Type

' ============================================================
' Entry points
' ============================================================
Public Sub CalculatePeriods()
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceTable As ListObject

    Dim targetSheet As Worksheet
    Dim targetTable As ListObject

    Dim filterDateFrom As Date
    Dim filterDateTo As Date

    Dim result As Boolean

    If gCalculationRunning Then
        MsgBox MSG_CALCULATION_ALREADY_RUNNING, vbExclamation
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    gCalculationRunning = True
    gCancelRequested = False
    Application.StatusBar = "Preparing calculation..."
    ClearLog
    LogDebug "Calculation started"
    ' --------------------------------------------------------
    ' Поиск исходной таблицы
    ' --------------------------------------------------------
    If Not FindSourceTable(sourceWorkbook, sourceSheet, sourceTable) Then
        LogError "Source table was not found"
        MsgBox MSG_SOURCE_NOT_FOUND, vbExclamation
        GoTo ExitPoint
    End If
    ClearSourceFilters sourceTable
    If IsCancellationRequested(True) Then
        GoTo Cancelled
    End If
    ' --------------------------------------------------------
    ' Поиск целевой таблицы
    ' --------------------------------------------------------
    If Not FindTargetTable(targetSheet, targetTable) Then
        LogError "Target table was not found"
        MsgBox MSG_TARGET_NOT_FOUND, vbExclamation
        GoTo ExitPoint
    End If
    ' --------------------------------------------------------
    ' Проверка обязательных данных
    ' --------------------------------------------------------
    If Not ValidateSourceColumns(sourceTable) Then
        MsgBox MSG_SOURCE_COLUMN_NOT_FOUND, vbExclamation
        GoTo ExitPoint
    End If
    If Not ValidateTargetColumns(targetTable) Then
        MsgBox MSG_TARGET_COLUMN_NOT_FOUND, vbExclamation
        GoTo ExitPoint
    End If
    ' --------------------------------------------------------
    ' Чтение параметров расчета
    ' --------------------------------------------------------
    If Not TryReadDateParameters(targetSheet, filterDateFrom, filterDateTo) Then
        LogError "Invalid calculation range"
        MsgBox MSG_INVALID_PERIOD, vbExclamation
        GoTo ExitPoint
    End If
    LogDebug "Calculation range: " & Format$(filterDateFrom, "dd.mm.yyyy") & " - " & Format$(filterDateTo, "dd.mm.yyyy")
    ' --------------------------------------------------------
    ' Выполнение расчета
    ' --------------------------------------------------------
    result = BuildPeriods(sourceTable, targetTable, filterDateFrom, filterDateTo)
    If gCancelRequested Then
        GoTo Cancelled
    End If
    GoTo ExitPoint
Cancelled:
    LogWarning "Calculation cancelled by user"
    MsgBox MSG_CANCELLED, vbInformation
ExitPoint:
    gCalculationRunning = False
    Application.StatusBar = False
    Exit Sub
ErrorHandler:
    LogError "Unexpected error" & " | Number=" & Err.Number & " | Description=" & Err.Description
    gCalculationRunning = False
    Application.StatusBar = False
    MsgBox MSG_UNEXPECTED_ERROR, vbCritical
End Sub

Public Sub CancelCalculation()
    gCancelRequested = True
    Application.StatusBar = "Cancelling calculation..."
End Sub

' ============================================================
' Main workflow
' ============================================================
Private Function BuildPeriods( _
    ByVal sourceTable As ListObject, _
    ByVal targetTable As ListObject, _
    ByVal filterDateFrom As Date, _
    ByVal filterDateTo As Date _
) As Boolean
    Dim sourcePeriods() As PeriodInfo
    Dim mergedPeriods() As PeriodInfo

    Dim sourceCount As Long
    Dim mergedCount As Long
    Dim finalCount As Long

    BuildPeriods = False
    ' --------------------------------------------------------
    ' Чтение закрытых и открытых событий
    ' --------------------------------------------------------
    Application.StatusBar = "Reading source data..."
    sourceCount = ReadPeriods(sourceTable, sourcePeriods)
    If gCancelRequested Then
        Exit Function
    End If
    LogDebug "Source event count: " & sourceCount
    If sourceCount = 0 Then
        ClearAllTargetRows targetTable
        MsgBox MSG_NO_PERIODS, vbInformation
        BuildPeriods = True
        Exit Function
    End If
    ' --------------------------------------------------------
    ' Сортировка событий
    ' --------------------------------------------------------
    Application.StatusBar = "Sorting source periods..."
    SortPeriods sourcePeriods, sourceCount
    If gCancelRequested Then
        Exit Function
    End If
    ' --------------------------------------------------------
    ' Объединение непрерывных периодов
    ' --------------------------------------------------------
    Application.StatusBar = "Merging periods..."
    mergedCount = MergePeriodsByPerson(sourcePeriods, sourceCount, mergedPeriods)
    If gCancelRequested Then
        Exit Function
    End If
    LogDebug "Merged period count: " & mergedCount
    ' --------------------------------------------------------
    ' Фильтрация по заданному диапазону
    '
    ' Исходные даты периодов сохраняются без обрезки.
    ' --------------------------------------------------------
    Application.StatusBar = "Filtering result..."
    finalCount = FilterPeriodsByRange(mergedPeriods, mergedCount, filterDateFrom, filterDateTo)
    If gCancelRequested Then
        Exit Function
    End If
    LogDebug "Result period count: " & finalCount
    ' --------------------------------------------------------
    ' Запись результата
    ' --------------------------------------------------------
    Application.StatusBar = "Writing result..."
    If Not WritePeriods(targetTable, mergedPeriods, finalCount) Then
        Exit Function
    End If
    If gCancelRequested Then
        Exit Function
    End If
    Application.StatusBar = False
    ' If finalCount = 0 Then
    '     MsgBox MSG_NO_PERIODS, vbInformation
    ' Else
    '     MsgBox MSG_DONE_PREFIX & finalCount, vbInformation
    ' End If
    LogDebug "Calculation completed"
    BuildPeriods = True
End Function

' ============================================================
' Read events
' ============================================================
Private Function ReadPeriods(ByVal sourceTable As ListObject, ByRef result() As PeriodInfo) As Long
    Dim data As Variant

    Dim rowCount As Long
    Dim rowIndex As Long
    Dim count As Long

    Dim rankIndex As Long
    Dim nameIndex As Long
    Dim taxIdIndex As Long
    Dim positionIndex As Long
    Dim eventIndex As Long

    Dim fromIndex As Long
    Dim toIndex As Long

    Dim departureOrderIndex As Long
    Dim arrivalOrderIndex As Long

    Dim dateFrom As Date
    Dim dateTo As Date
    Dim isOpen As Boolean

    Dim taxId As String

    Dim invalidPeriodCount As Long
    Dim missingTaxIdCount As Long
    Dim openEventCount As Long

    ReadPeriods = 0
    If sourceTable.DataBodyRange Is Nothing Then
        Exit Function
    End If
    rankIndex = sourceTable.ListColumns(SOURCE_COL_RANK).Index
    nameIndex = sourceTable.ListColumns(SOURCE_COL_NAME).Index
    taxIdIndex = sourceTable.ListColumns(SOURCE_COL_TAX_ID).Index
    positionIndex = sourceTable.ListColumns(SOURCE_COL_POSITION).Index
    eventIndex = sourceTable.ListColumns(SOURCE_COL_EVENT).Index
    fromIndex = sourceTable.ListColumns(SOURCE_COL_PERIOD_FROM).Index
    toIndex = sourceTable.ListColumns(SOURCE_COL_PERIOD_TO).Index
    departureOrderIndex = sourceTable.ListColumns(SOURCE_COL_DEPARTURE_ORDER).Index
    arrivalOrderIndex = sourceTable.ListColumns(SOURCE_COL_ARRIVAL_ORDER).Index
    data = sourceTable.DataBodyRange.Value2
    rowCount = UBound(data, 1)
    ReDim result(1 To rowCount)
    count = 0
    For rowIndex = 1 To rowCount
        If rowIndex Mod UI_YIELD_INTERVAL = 0 Then
            Application.StatusBar = "Reading source data: " & rowIndex & " / " & rowCount
            If IsCancellationRequested(True) Then
                ReadPeriods = count
                Exit Function
            End If
        End If
        If Not TryConvertToDate(data(rowIndex, fromIndex), dateFrom) Then
            invalidPeriodCount = invalidPeriodCount + 1
            GoTo NextRow
        End If
        isOpen = False
        If Not TryConvertToDate(data(rowIndex, toIndex), dateTo) Then
            If IsError(data(rowIndex, toIndex)) Then
                invalidPeriodCount = invalidPeriodCount + 1
                GoTo NextRow
            End If
            If Len(Trim$(SafeString(data(rowIndex, toIndex)))) > 0 Then
                invalidPeriodCount = invalidPeriodCount + 1
                GoTo NextRow
            End If
            isOpen = True
            openEventCount = openEventCount + 1
            dateTo = DateSerial(9999, 12, 31)
        End If
        If dateTo < dateFrom Then
            invalidPeriodCount = invalidPeriodCount + 1
            GoTo NextRow
        End If
        taxId = SafeString(data(rowIndex, taxIdIndex))
        If Len(taxId) = 0 Then
            missingTaxIdCount = missingTaxIdCount + 1
            GoTo NextRow
        End If
        count = count + 1
        With result(count)
            .Rank = SafeString(data(rowIndex, rankIndex))
            .PersonName = SafeString(data(rowIndex, nameIndex))
            .taxId = taxId
            .Position = SafeString(data(rowIndex, positionIndex))
            .EventName = SafeString(data(rowIndex, eventIndex))
            .LastEventName = .EventName
            .HasDisplayableEvent = IsDisplayableEvent(.EventName)
            .IsOpen = isOpen
            .DepartureOrder = SafeString(data(rowIndex, departureOrderIndex))
            .dateFrom = dateFrom
            .dateTo = dateTo
            .ArrivalOrder = SafeString(data(rowIndex, arrivalOrderIndex))
        End With
NextRow:
    Next rowIndex
    If count > 0 Then
        ReDim Preserve result(1 To count)
    End If
    If invalidPeriodCount > 0 Then
        LogWarning "Invalid source periods skipped: " & invalidPeriodCount
    End If
    If missingTaxIdCount > 0 Then
        LogWarning "Source rows with empty TaxId skipped: " & missingTaxIdCount
    End If
    LogDebug "Open events included: " & openEventCount
    ReadPeriods = count
End Function

' ============================================================
' Date conversion
' ============================================================
Private Function TryConvertToDate(ByVal value As Variant, ByRef result As Date) As Boolean
    TryConvertToDate = False
    If IsError(value) Then
        Exit Function
    End If
    If IsEmpty(value) Then
        Exit Function
    End If
    If IsNull(value) Then
        Exit Function
    End If
    On Error GoTo ConversionFailed
    If IsDate(value) Or IsNumeric(value) Then
        result = CDate(value)
        TryConvertToDate = True
    End If
    Exit Function
ConversionFailed:
    TryConvertToDate = False
End Function

' ============================================================
' Sort
' ============================================================
Private Sub SortPeriods(ByRef periods() As PeriodInfo, ByVal count As Long)
    If count <= 1 Then
        Exit Sub
    End If
    QuickSortPeriods periods, 1, count
End Sub

Private Sub QuickSortPeriods(ByRef periods() As PeriodInfo, ByVal low As Long, ByVal high As Long)
    Dim i As Long
    Dim j As Long

    Dim pivot As PeriodInfo
    Dim temp As PeriodInfo

    If gCancelRequested Then
        Exit Sub
    End If
    i = low
    j = high
    pivot = periods((low + high) \ 2)
    Do While i <= j
        Do While ComparePeriods(periods(i), pivot) < 0
            i = i + 1
        Loop
        Do While ComparePeriods(periods(j), pivot) > 0
            j = j - 1
        Loop
        If i <= j Then
            temp = periods(i)
            periods(i) = periods(j)
            periods(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If IsCancellationRequested(True) Then
        Exit Sub
    End If
    If low < j Then
        QuickSortPeriods periods, low, j
    End If
    If gCancelRequested Then
        Exit Sub
    End If
    If i < high Then
        QuickSortPeriods periods, i, high
    End If
End Sub

Private Function ComparePeriods(ByRef leftPeriod As PeriodInfo, ByRef rightPeriod As PeriodInfo) As Long
    Dim taxIdCompare As Long

    taxIdCompare = StrComp(leftPeriod.taxId, rightPeriod.taxId, vbTextCompare)
    If taxIdCompare < 0 Then
        ComparePeriods = -1
        Exit Function
    End If
    If taxIdCompare > 0 Then
        ComparePeriods = 1
        Exit Function
    End If
If leftPeriod.dateFrom <rightPeriod.dateFrom Then
        ComparePeriods = -1
ElseIf leftPeriod.dateFrom >rightPeriod.dateFrom Then
        ComparePeriods = 1
ElseIf leftPeriod.dateTo <rightPeriod.dateTo Then
        ComparePeriods = -1
ElseIf leftPeriod.dateTo >rightPeriod.dateTo Then
        ComparePeriods = 1
    Else
        ComparePeriods = 0
    End If
End Function

' ============================================================
' Merge by person
' ============================================================
Private Function MergePeriodsByPerson( _
    ByRef source() As PeriodInfo, _
    ByVal sourceCount As Long, _
    ByRef result() As PeriodInfo _
) As Long
    Dim sourceIndex As Long
    Dim resultCount As Long

    If sourceCount = 0 Then
        MergePeriodsByPerson = 0
        Exit Function
    End If
    ReDim result(1 To sourceCount)
    resultCount = 1
    result(resultCount) = source(1)
    For sourceIndex = 2 To sourceCount
        If sourceIndex Mod UI_YIELD_INTERVAL = 0 Then
            Application.StatusBar = "Merging periods: " & sourceIndex & " / " & sourceCount
            If IsCancellationRequested(True) Then
                MergePeriodsByPerson = resultCount
                Exit Function
            End If
        End If
        If IsSamePerson(result(resultCount), source(sourceIndex)) Then
            If IsContinuous(result(resultCount), source(sourceIndex)) Then
                MergeIntoPeriod result(resultCount), source(sourceIndex)
            Else
                resultCount = resultCount + 1
                result(resultCount) = source(sourceIndex)
            End If
        Else
            resultCount = resultCount + 1
            result(resultCount) = source(sourceIndex)
        End If
    Next sourceIndex
    ReDim Preserve result(1 To resultCount)
    MergePeriodsByPerson = resultCount
End Function

' ============================================================
' Merge two periods
' ============================================================
Private Sub MergeIntoPeriod(ByRef targetPeriod As PeriodInfo, ByRef sourcePeriod As PeriodInfo)
    If Len(targetPeriod.EventName) = 0 Then
        targetPeriod.EventName = sourcePeriod.EventName
    ElseIf Len(sourcePeriod.EventName) > 0 Then
        targetPeriod.EventName = targetPeriod.EventName & EVENTS_SEPARATOR & sourcePeriod.EventName
    End If
    targetPeriod.LastEventName = sourcePeriod.LastEventName
    targetPeriod.HasDisplayableEvent = _
        targetPeriod.HasDisplayableEvent Or sourcePeriod.HasDisplayableEvent
    targetPeriod.IsOpen = targetPeriod.IsOpen Or sourcePeriod.IsOpen
If sourcePeriod.dateTo >targetPeriod.dateTo Then
        targetPeriod.dateTo = sourcePeriod.dateTo
        targetPeriod.ArrivalOrder = sourcePeriod.ArrivalOrder
    End If
End Sub

' ============================================================
' Person comparison
' ============================================================
Private Function IsSamePerson(ByRef leftPeriod As PeriodInfo, ByRef rightPeriod As PeriodInfo) As Boolean
    IsSamePerson = StrComp(leftPeriod.taxId, rightPeriod.taxId, vbTextCompare) = 0
End Function

' ============================================================
' Continuity
' ============================================================
Private Function IsContinuous(ByRef currentPeriod As PeriodInfo, ByRef NextPeriod As PeriodInfo) As Boolean
    IsContinuous = False
    If Not IsMergeableEvent(currentPeriod.LastEventName) Then
        Exit Function
    End If
    If Not IsMergeableEvent(NextPeriod.EventName) Then
        Exit Function
    End If
    ' Периоды пересекаются или имеют общую граничную дату.
If NextPeriod.dateFrom <=currentPeriod.dateTo Then
        IsContinuous = True
        Exit Function
    End If
    ' Правило объединения соседних календарных дней временно отключено.
    '
    ' 01.06 - 10.06
    ' 11.06 - 20.06
    '
    ' Соседние периоды объединяются только при совпадении номеров приказов.
    ' If NextPeriod.dateFrom = DateAdd("d", 1, currentPeriod.dateTo) Then
    '     If OrdersMatch(currentPeriod.ArrivalOrder, NextPeriod.DepartureOrder) Then
    '         IsContinuous = True
    '     End If
    ' End If
End Function

Private Function IsMergeableEvent(ByVal eventName As String) As Boolean
    IsMergeableEvent = IsEventInList(eventName, MERGEABLE_EVENTS)
End Function

Private Function IsDisplayableEvent(ByVal eventName As String) As Boolean
    IsDisplayableEvent = IsEventInList(eventName, DISPLAYABLE_EVENTS)
End Function

Private Function IsEventInList(ByVal eventName As String, ByVal eventList As String) As Boolean
    eventName = Trim$(eventName)
    If Len(eventName) = 0 Then
        IsEventInList = False
        Exit Function
    End If
    IsEventInList = InStr( _
        1, _
        "|" & eventList & "|", _
        "|" & eventName & "|", _
        vbTextCompare _
    ) > 0
End Function

' ============================================================
' Order comparison
' ============================================================
Private Function OrdersMatch(ByVal closingOrder As String, ByVal openingOrder As String) As Boolean
    closingOrder = Trim$(closingOrder)
    openingOrder = Trim$(openingOrder)
    OrdersMatch = False
    If Len(closingOrder) = 0 Then
        Exit Function
    End If
    If Len(openingOrder) = 0 Then
        Exit Function
    End If
    OrdersMatch = StrComp(closingOrder, openingOrder, vbTextCompare) = 0
End Function

' ============================================================
' Filter by requested range
'
' Исходные даты периодов сохраняются без обрезки.
' В результат попадают только периоды с событиями из DISPLAYABLE_EVENTS,
' которые пересекаются с заданным диапазоном.
' ============================================================
Private Function FilterPeriodsByRange( _
    ByRef periods() As PeriodInfo, _
    ByVal periodCount As Long, _
    ByVal filterDateFrom As Date, _
    ByVal filterDateTo As Date _
) As Long
    Dim readIndex As Long
    Dim writeIndex As Long

    writeIndex = 0
    For readIndex = 1 To periodCount
        If readIndex Mod UI_YIELD_INTERVAL = 0 Then
            If IsCancellationRequested(True) Then
                FilterPeriodsByRange = writeIndex
                Exit Function
            End If
        End If
        If Not periods(readIndex).HasDisplayableEvent Then
            GoTo NextPeriod
        End If
        If periods(readIndex).dateTo < filterDateFrom Then
            GoTo NextPeriod
        End If
        If periods(readIndex).dateFrom > filterDateTo Then
            GoTo NextPeriod
        End If
        writeIndex = writeIndex + 1
        If writeIndex <> readIndex Then
            periods(writeIndex) = periods(readIndex)
        End If
NextPeriod:
    Next readIndex
    FilterPeriodsByRange = writeIndex
End Function

' ============================================================
' Fast bulk target write
'
' Лишние строки ListObject удаляются одним вызовом Range.Delete.
' Для VBA это аналог выделения строк таблицы и нажатия Ctrl + Minus.
' ============================================================
Private Function WritePeriods( _
    ByVal targetTable As ListObject, _
    ByRef periods() As PeriodInfo, _
    ByVal count As Long _
) As Boolean
    Dim targetRange As Range

    Dim data() As Variant

    Dim rowIndex As Long
    Dim columnCount As Long

    Dim rankIndex As Long
    Dim nameIndex As Long
    Dim taxIdIndex As Long
    Dim positionIndex As Long
    Dim eventIndex As Long

    Dim departureOrderIndex As Long
    Dim fromIndex As Long
    Dim toIndex As Long
    Dim arrivalOrderIndex As Long
    Dim periodCountIndex As Long

    WritePeriods = False
    If IsCancellationRequested(True) Then
        Exit Function
    End If
    ' --------------------------------------------------------
    ' Быстрый аналог Ctrl + Minus
    ' --------------------------------------------------------
    DeleteExcessTargetRows targetTable, count
    If IsCancellationRequested(True) Then
        Exit Function
    End If
    ' --------------------------------------------------------
    ' Пустой результат
    ' --------------------------------------------------------
    If count = 0 Then
        WritePeriods = True
        Exit Function
    End If
    columnCount = targetTable.ListColumns.count
    ' --------------------------------------------------------
    ' Определение индексов целевых колонок
    ' --------------------------------------------------------
    rankIndex = targetTable.ListColumns(TARGET_COL_RANK).Index
    nameIndex = targetTable.ListColumns(TARGET_COL_NAME).Index
    taxIdIndex = targetTable.ListColumns(TARGET_COL_TAX_ID).Index
    positionIndex = targetTable.ListColumns(TARGET_COL_POSITION).Index
    eventIndex = targetTable.ListColumns(TARGET_COL_EVENT).Index
    departureOrderIndex = targetTable.ListColumns(TARGET_COL_DEPARTURE_ORDER).Index
    fromIndex = targetTable.ListColumns(TARGET_COL_PERIOD_FROM).Index
    toIndex = targetTable.ListColumns(TARGET_COL_PERIOD_TO).Index
    arrivalOrderIndex = targetTable.ListColumns(TARGET_COL_ARRIVAL_ORDER).Index
    periodCountIndex = targetTable.ListColumns(TARGET_COL_PERIOD_COUNT).Index
    ' --------------------------------------------------------
    ' Однократное расширение таблицы, если строк недостаточно
    '
    ' Если строк было больше, чем требуется для нового результата,
    ' после DeleteExcessTargetRows таблица уже имеет нужный размер.
    ' --------------------------------------------------------
    If targetTable.ListRows.count <> count Then
        Set targetRange = targetTable.HeaderRowRange.Resize(count + 1, columnCount)
        targetTable.Resize targetRange
    End If
    If IsCancellationRequested(True) Then
        Exit Function
    End If
    ' --------------------------------------------------------
    ' Формирование результата в памяти
    ' --------------------------------------------------------
    ReDim data(1 To count, 1 To columnCount)
    For rowIndex = 1 To count
        data(rowIndex, rankIndex) = periods(rowIndex).Rank
        data(rowIndex, nameIndex) = periods(rowIndex).PersonName
        data(rowIndex, taxIdIndex) = periods(rowIndex).taxId
        data(rowIndex, positionIndex) = periods(rowIndex).Position
        data(rowIndex, eventIndex) = periods(rowIndex).EventName
        data(rowIndex, departureOrderIndex) = periods(rowIndex).DepartureOrder
        data(rowIndex, fromIndex) = periods(rowIndex).dateFrom
        If periods(rowIndex).IsOpen Then
            data(rowIndex, toIndex) = OPEN_PERIOD_TO_TEXT
            data(rowIndex, arrivalOrderIndex) = vbNullString
            data(rowIndex, periodCountIndex) = CStr(DateDiff( _
                "d", _
                periods(rowIndex).dateFrom, _
                Date _
            ))
        Else
            data(rowIndex, toIndex) = periods(rowIndex).dateTo
            data(rowIndex, arrivalOrderIndex) = periods(rowIndex).ArrivalOrder
            data(rowIndex, periodCountIndex) = CStr(DateDiff( _
                "d", _
                periods(rowIndex).dateFrom, _
                periods(rowIndex).dateTo _
            ))
        End If
        If rowIndex Mod UI_YIELD_INTERVAL = 0 Then
            Application.StatusBar = "Preparing output: " & rowIndex & " / " & count
            If IsCancellationRequested(True) Then
                Exit Function
            End If
        End If
    Next rowIndex
    ' --------------------------------------------------------
    ' Единая пакетная запись в Excel
    ' --------------------------------------------------------
    Application.StatusBar = "Writing result to worksheet..."
    DoEvents
    targetTable.ListColumns(TARGET_COL_PERIOD_COUNT).DataBodyRange.NumberFormat = "@"
    targetTable.DataBodyRange.value = data
    ' --------------------------------------------------------
    ' Форматирование дат
    ' --------------------------------------------------------
    targetTable.ListColumns(TARGET_COL_PERIOD_FROM).DataBodyRange.NumberFormat = "dd.mm.yyyy"
    targetTable.ListColumns(TARGET_COL_PERIOD_TO).DataBodyRange.NumberFormat = "dd.mm.yyyy"
    DoEvents
    WritePeriods = True
End Function

' ============================================================
' Fast Ctrl-minus equivalent
'
' Пример:
'
' Старый результат = 3480 строк
' Новый результат = 900 строк
'
' Строки 901..3480 выделяются одним Range и удаляются
' одной операцией Delete.
'
' Без ClearFormats.
' Без EntireRow.Delete.
' Без цикла по ListRows.
' ============================================================
Private Sub DeleteExcessTargetRows(ByVal targetTable As ListObject, ByVal newRowCount As Long)
    Dim oldRowCount As Long
    Dim deleteCount As Long

    Dim staleRange As Range

    oldRowCount = targetTable.ListRows.count
    If oldRowCount = 0 Then
        Exit Sub
    End If
    ' --------------------------------------------------------
    ' Удаление всех строк таблицы
    ' --------------------------------------------------------
    If newRowCount <= 0 Then
        If Not targetTable.DataBodyRange Is Nothing Then
            targetTable.DataBodyRange.Delete Shift:=xlShiftUp
        End If
        Exit Sub
    End If
    ' --------------------------------------------------------
    ' Лишних строк нет
    ' --------------------------------------------------------
    If oldRowCount <= newRowCount Then
        Exit Sub
    End If
    deleteCount = oldRowCount - newRowCount
    ' --------------------------------------------------------
    ' Выделение только устаревшего хвоста ListObject
    ' --------------------------------------------------------
Set staleRange =targetTable.DataBodyRange.Rows(newRowCount + 1).Resize(deleteCount,targetTable.ListColumns.count)
    ' --------------------------------------------------------
    ' Одна операция удаления.
    '
    ' Это ключевой участок оптимизации.
    ' --------------------------------------------------------
    staleRange.Delete Shift:=xlShiftUp
End Sub

' ============================================================
' Clear all target rows
' ============================================================
Private Sub ClearAllTargetRows(ByVal targetTable As ListObject)
    If targetTable.DataBodyRange Is Nothing Then
        Exit Sub
    End If
    targetTable.DataBodyRange.Delete Shift:=xlShiftUp
End Sub

' ============================================================
' Cancellation
' ============================================================
Private Function IsCancellationRequested(Optional ByVal yieldToUi As Boolean = False) As Boolean
    If yieldToUi Then
        DoEvents
    End If
    IsCancellationRequested = gCancelRequested
End Function

' ============================================================
' Find source
' ============================================================
Private Function FindSourceTable( _
    ByRef resultWorkbook As Workbook, _
    ByRef resultSheet As Worksheet, _
    ByRef resultTable As ListObject _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim table As ListObject

    FindSourceTable = False
    For Each wb In Application.Workbooks
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(SOURCE_SHEET_NAME)
        On Error GoTo 0
        If Not ws Is Nothing Then
            For Each table In ws.ListObjects
                If StrComp(table.Name, SOURCE_TABLE_NAME, vbBinaryCompare) = 0 Then
                    Set resultWorkbook = wb
                    Set resultSheet = ws
                    Set resultTable = table
                    FindSourceTable = True
                    Exit Function
                End If
            Next table
        End If
    Next wb
End Function

' ============================================================
' Find target
' ============================================================
Private Function FindTargetTable(ByRef resultSheet As Worksheet, ByRef resultTable As ListObject) As Boolean
    FindTargetTable = False
    Set resultSheet = Nothing
    Set resultTable = Nothing
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    If Not resultSheet Is Nothing Then
        Set resultTable = resultSheet.ListObjects(TARGET_TABLE_NAME)
    End If
    On Error GoTo 0
    FindTargetTable = Not resultTable Is Nothing
End Function

' ============================================================
' Clear source filters
' ============================================================
Private Sub ClearSourceFilters(ByVal sourceTable As ListObject)
    On Error Resume Next
    If sourceTable.AutoFilter.FilterMode Then
        sourceTable.AutoFilter.ShowAllData
    End If
    On Error GoTo 0
End Sub

' ============================================================
' Parameters
' ============================================================
Private Function TryReadDateParameters( _
    ByVal targetSheet As Worksheet, _
    ByRef resultFrom As Date, _
    ByRef resultTo As Date _
) As Boolean
    Dim valueFrom As Variant
    Dim valueTo As Variant

    TryReadDateParameters = False
    valueFrom = targetSheet.Range(PARAM_PERIOD_FROM_CELL).Value2
    valueTo = targetSheet.Range(PARAM_PERIOD_TO_CELL).Value2
    If Not TryConvertToDate(valueFrom, resultFrom) Then
        Exit Function
    End If
    If Not TryConvertToDate(valueTo, resultTo) Then
        Exit Function
    End If
    If resultFrom > resultTo Then
        Exit Function
    End If
    TryReadDateParameters = True
End Function

' ============================================================
' Validation
' ============================================================
Private Function ValidateSourceColumns(ByVal sourceTable As ListObject) As Boolean
    ValidateSourceColumns = False
    If Not RequireColumn(sourceTable, SOURCE_COL_RANK, "source Rank") Then Exit Function
    If Not RequireColumn(sourceTable, SOURCE_COL_NAME, "source Name") Then Exit Function
    If Not RequireColumn(sourceTable, SOURCE_COL_TAX_ID, "source TaxId") Then Exit Function
    If Not RequireColumn(sourceTable, SOURCE_COL_POSITION, "source Position") Then Exit Function
    If Not RequireColumn(sourceTable, SOURCE_COL_EVENT, "source Event") Then Exit Function
    If Not RequireColumn(sourceTable, SOURCE_COL_PERIOD_FROM, "source PeriodFrom") Then Exit Function
    If Not RequireColumn(sourceTable, SOURCE_COL_PERIOD_TO, "source PeriodTo") Then Exit Function
    If Not RequireColumn(sourceTable, SOURCE_COL_DEPARTURE_ORDER, "source DepartureOrder") Then Exit Function
    If Not RequireColumn(sourceTable, SOURCE_COL_ARRIVAL_ORDER, "source ArrivalOrder") Then Exit Function
    ValidateSourceColumns = True
End Function

Private Function ValidateTargetColumns(ByVal targetTable As ListObject) As Boolean
    ValidateTargetColumns = False
    If Not RequireColumn(targetTable, TARGET_COL_RANK, "target Rank") Then Exit Function
    If Not RequireColumn(targetTable, TARGET_COL_NAME, "target Name") Then Exit Function
    If Not RequireColumn(targetTable, TARGET_COL_TAX_ID, "target TaxId") Then Exit Function
    If Not RequireColumn(targetTable, TARGET_COL_POSITION, "target Position") Then Exit Function
    If Not RequireColumn(targetTable, TARGET_COL_EVENT, "target Event") Then Exit Function
    If Not RequireColumn(targetTable, TARGET_COL_DEPARTURE_ORDER, "target DepartureOrder") Then Exit Function
    If Not RequireColumn(targetTable, TARGET_COL_PERIOD_FROM, "target PeriodFrom") Then Exit Function
    If Not RequireColumn(targetTable, TARGET_COL_PERIOD_TO, "target PeriodTo") Then Exit Function
    If Not RequireColumn(targetTable, TARGET_COL_ARRIVAL_ORDER, "target ArrivalOrder") Then Exit Function
    If Not RequireColumn(targetTable, TARGET_COL_PERIOD_COUNT, "target PeriodCount") Then Exit Function
    ValidateTargetColumns = True
End Function

Private Function RequireColumn( _
    ByVal table As ListObject, _
    ByVal columnName As String, _
    ByVal diagnosticName As String _
) As Boolean
    Dim column As ListColumn

    Set column = Nothing
    On Error Resume Next
    Set column = table.ListColumns(columnName)
    On Error GoTo 0
    RequireColumn = Not column Is Nothing
    If Not RequireColumn Then
        LogError "Required column missing: " & diagnosticName
    End If
End Function

' ============================================================
' Helpers
' ============================================================
Private Function SafeString(ByVal value As Variant) As String
    If IsError(value) Then
        SafeString = vbNullString
    ElseIf IsNull(value) Then
        SafeString = vbNullString
    ElseIf IsEmpty(value) Then
        SafeString = vbNullString
    Else
        SafeString = CStr(value)
    End If
End Function

' ============================================================
' Logging
' ============================================================
Private Sub ClearLog()
#If ENABLE_LOGGING Then
    Dim fileNumber As Integer

    fileNumber = FreeFile
    Open GetLogFilePath() For Output As #fileNumber
    Close #fileNumber
#End If
End Sub

Private Sub LogError(ByVal message As String)
#If ENABLE_LOGGING Then
    WriteLog "ERROR: " & message
#End If
End Sub

Private Sub LogWarning(ByVal message As String)
#If ENABLE_LOGGING Then
    WriteLog "WARNING: " & message
#End If
End Sub

Private Sub LogDebug(ByVal message As String)
#If ENABLE_LOGGING Then
#If ENABLE_DEBUG_LOGGING Then
    WriteLog "DEBUG: " & message
#End If
#End If
End Sub

Private Sub WriteLog(ByVal message As String)
#If ENABLE_LOGGING Then
    Dim fileNumber As Integer

    fileNumber = FreeFile
    Open GetLogFilePath() For Append As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #fileNumber
#End If
End Sub

Private Function GetLogFilePath() As String
    Dim workbookName As String
    Dim baseName As String

    Dim dotPosition As Long

    workbookName = ThisWorkbook.Name
    dotPosition = InStrRev(workbookName, ".")
    If dotPosition > 0 Then
        baseName = Left$(workbookName, dotPosition - 1)
    Else
        baseName = workbookName
    End If
    GetLogFilePath = ThisWorkbook.Path & Application.PathSeparator & baseName & LOG_FILE_SUFFIX
End Function
