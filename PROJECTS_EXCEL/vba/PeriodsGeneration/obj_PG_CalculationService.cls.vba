VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PG_CalculationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_configuration As obj_PG_Configuration
Private m_runContext As obj_PG_RunContext
Private m_interval As Long


' //
' // Жизненный цикл
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize( _
    ByVal configuration As obj_PG_Configuration, _
    ByVal runContext As obj_PG_RunContext _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If configuration Is Nothing Or runContext Is Nothing Then
        Exit Function
    End If
    Set m_configuration = configuration
    Set m_runContext = runContext
    m_interval = VBA.CLng(m_configuration.GetText("legacy.UI_YIELD_INTERVAL"))
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_configuration = Nothing
    Set m_runContext = Nothing
End Sub

Public Function Calculate( _
    ByVal sourceTable As ListObject, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    ByRef values As Variant _
) As Long
    Dim sourcePeriods() As obj_PG_Period
    Dim mergedPeriods() As obj_PG_Period
    Dim sourceCount As Long
    Dim mergedCount As Long
    Dim count As Long
    Dim data() As Variant
    Dim rowIndex As Long
    Dim countTo As Date
    Dim today As Date

    private_EnsureReady
    values = Empty
    today = VBA.Date
    m_runContext.CheckCancel 0
    Application.StatusBar = m_configuration.GetText("status.Reading")
    sourceCount = private_ReadPeriods(sourceTable, sourcePeriods)
    If sourceCount = 0 Then
        Exit Function
    End If
    Application.StatusBar = m_configuration.GetText("status.Sorting")
    private_SortPeriods sourcePeriods, sourceCount
    Application.StatusBar = m_configuration.GetText("status.Merging")
    mergedCount = private_MergePeriodsByPerson(sourcePeriods, sourceCount, mergedPeriods)
    Application.StatusBar = m_configuration.GetText("status.Filtering")
    count = private_FilterPeriodsByRange(mergedPeriods, mergedCount, startDate, endDate)
    m_runContext.CheckCancel 0
    If count = 0 Then
        Exit Function
    End If
    Application.StatusBar = m_configuration.GetText("status.Preparing")
    ReDim data(1 To count, 1 To 10)
    For rowIndex = 1 To count
        m_runContext.CheckCancel rowIndex
        With mergedPeriods(rowIndex)
            data(rowIndex, 1) = .Rank
            data(rowIndex, 2) = .PersonName
            data(rowIndex, 3) = .taxId
            data(rowIndex, 4) = .Position
            data(rowIndex, 5) = .EventName
            data(rowIndex, 6) = .DepartureOrder
            data(rowIndex, 7) = VBA.Format$(.dateFrom, m_configuration.GetText("format.date"))
            If .IsOpen Then
                data(rowIndex, 8) = m_configuration.GetText("legacy.OPEN_PERIOD_TO_TEXT")
                data(rowIndex, 9) = VBA.vbNullString
                countTo = today
            Else
                data(rowIndex, 8) = VBA.Format$(.dateTo, m_configuration.GetText("format.date"))
                data(rowIndex, 9) = .ArrivalOrder
                countTo = .dateTo
            End If
            data(rowIndex, 10) = VBA.CStr(VBA.DateDiff("d", .dateFrom, countTo))
        End With
    Next rowIndex
    values = data
    Calculate = count
End Function

' //
' // Вспомогательные методы
' //
Private Function private_ReadPeriods( _
    ByVal sourceTable As ListObject, _
    ByRef result() As obj_PG_Period _
) As Long
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

    private_ReadPeriods = 0
    If sourceTable.DataBodyRange Is Nothing Then
        Exit Function
    End If
    rankIndex = sourceTable.ListColumns(m_configuration.GetText("legacy.SOURCE_COL_RANK")).Index
    nameIndex = sourceTable.ListColumns(m_configuration.GetText("legacy.SOURCE_COL_NAME")).Index
    taxIdIndex = sourceTable.ListColumns(m_configuration.GetText("legacy.SOURCE_COL_TAX_ID")).Index
    positionIndex = sourceTable.ListColumns(m_configuration.GetText("legacy.SOURCE_COL_POSITION")).Index
    eventIndex = sourceTable.ListColumns(m_configuration.GetText("legacy.SOURCE_COL_EVENT")).Index
    fromIndex = sourceTable.ListColumns(m_configuration.GetText("legacy.SOURCE_COL_PERIOD_FROM")).Index
    toIndex = sourceTable.ListColumns(m_configuration.GetText("legacy.SOURCE_COL_PERIOD_TO")).Index
    departureOrderIndex = sourceTable.ListColumns(m_configuration.GetText("legacy.SOURCE_COL_DEPARTURE_ORDER")).Index
    arrivalOrderIndex = sourceTable.ListColumns(m_configuration.GetText("legacy.SOURCE_COL_ARRIVAL_ORDER")).Index
    data = sourceTable.DataBodyRange.Value2
    rowCount = UBound(data, 1)
    ReDim result(1 To rowCount)
    count = 0
    For rowIndex = 1 To rowCount
        If rowIndex Mod m_interval = 0 Then
            Application.StatusBar = VBA.Replace( _
                VBA.Replace(m_configuration.GetText("status.ReadingProgress"), "{current}", VBA.CStr(rowIndex)), _
                "{total}", VBA.CStr(rowCount) _
            )
            If private_IsCancellationRequested(True) Then
                private_ReadPeriods = count
                Exit Function
            End If
        End If
        If Not private_TryConvertToDate(data(rowIndex, fromIndex), dateFrom) Then
            invalidPeriodCount = invalidPeriodCount + 1
            GoTo NextRow
        End If
        isOpen = False
        If Not private_TryConvertToDate(data(rowIndex, toIndex), dateTo) Then
            If IsError(data(rowIndex, toIndex)) Then
                invalidPeriodCount = invalidPeriodCount + 1
                GoTo NextRow
            End If
            If Len(Trim$(private_SafeString(data(rowIndex, toIndex)))) > 0 Then
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
        taxId = private_SafeString(data(rowIndex, taxIdIndex))
        If Len(taxId) = 0 Then
            missingTaxIdCount = missingTaxIdCount + 1
            GoTo NextRow
        End If
        count = count + 1
        Set result(count) = New obj_PG_Period
        If Not result(count).Initialize() Then
            m_configuration.Fail m_configuration.GetText("message.PublishFailed")
        End If
        With result(count)
            .Rank = private_SafeString(data(rowIndex, rankIndex))
            .PersonName = private_SafeString(data(rowIndex, nameIndex))
            .taxId = taxId
            .Position = private_SafeString(data(rowIndex, positionIndex))
            .EventName = private_SafeString(data(rowIndex, eventIndex))
            .LastEventName = .EventName
            .HasDisplayableEvent = private_IsDisplayableEvent(.EventName)
            .IsOpen = isOpen
            .DepartureOrder = private_SafeString(data(rowIndex, departureOrderIndex))
            .dateFrom = dateFrom
            .dateTo = dateTo
            .ArrivalOrder = private_SafeString(data(rowIndex, arrivalOrderIndex))
        End With
NextRow:
    Next rowIndex
    If count > 0 Then
        ReDim Preserve result(1 To count)
    End If
    If invalidPeriodCount > 0 Then
        m_runContext.LogWarning "Invalid source periods skipped: " & invalidPeriodCount
    End If
    If missingTaxIdCount > 0 Then
        m_runContext.LogWarning "Source rows with empty TaxId skipped: " & missingTaxIdCount
    End If
    m_runContext.LogDebug "Open events included: " & openEventCount
    private_ReadPeriods = count
End Function

' //
' // Преобразование дат
' //
Private Function private_TryConvertToDate( _
    ByVal value As Variant, _
    ByRef result As Date _
) As Boolean
    private_TryConvertToDate = False
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
        private_TryConvertToDate = True
    End If
    Exit Function
ConversionFailed:
    private_TryConvertToDate = False
End Function

' //
' // Сортировка
' //
Private Sub private_SortPeriods( _
    ByRef periods() As obj_PG_Period, _
    ByVal count As Long _
)
    If count <= 1 Then
        Exit Sub
    End If
    private_QuickSortPeriods periods, 1, count
End Sub

Private Sub private_QuickSortPeriods( _
    ByRef periods() As obj_PG_Period, _
    ByVal low As Long, _
    ByVal high As Long _
)
    Dim i As Long
    Dim j As Long

    Dim pivot As obj_PG_Period
    Dim temp As obj_PG_Period

    If private_IsCancellationRequested() Then
        Exit Sub
    End If
    i = low
    j = high
    Set pivot = periods((low + high) \ 2)
    Do While i <= j
        Do While private_ComparePeriods(periods(i), pivot) < 0
            i = i + 1
        Loop
        Do While private_ComparePeriods(periods(j), pivot) > 0
            j = j - 1
        Loop
        If i <= j Then
            Set temp = periods(i)
            Set periods(i) = periods(j)
            Set periods(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If private_IsCancellationRequested(True) Then
        Exit Sub
    End If
    If low < j Then
        private_QuickSortPeriods periods, low, j
    End If
    If private_IsCancellationRequested() Then
        Exit Sub
    End If
    If i < high Then
        private_QuickSortPeriods periods, i, high
    End If
End Sub

Private Function private_ComparePeriods( _
    ByRef leftPeriod As obj_PG_Period, _
    ByRef rightPeriod As obj_PG_Period _
) As Long
    Dim taxIdCompare As Long

    taxIdCompare = StrComp(leftPeriod.taxId, rightPeriod.taxId, vbTextCompare)
    If taxIdCompare < 0 Then
        private_ComparePeriods = -1
        Exit Function
    End If
    If taxIdCompare > 0 Then
        private_ComparePeriods = 1
        Exit Function
    End If
    If leftPeriod.dateFrom < rightPeriod.dateFrom Then
        private_ComparePeriods = -1
    ElseIf leftPeriod.dateFrom > rightPeriod.dateFrom Then
        private_ComparePeriods = 1
    ElseIf leftPeriod.dateTo < rightPeriod.dateTo Then
        private_ComparePeriods = -1
    ElseIf leftPeriod.dateTo > rightPeriod.dateTo Then
        private_ComparePeriods = 1
    Else
        private_ComparePeriods = 0
    End If
End Function

' //
' // Объединение по человеку
' //
Private Function private_MergePeriodsByPerson( _
    ByRef source() As obj_PG_Period, _
    ByVal sourceCount As Long, _
    ByRef result() As obj_PG_Period _
) As Long
    Dim sourceIndex As Long
    Dim resultCount As Long

    If sourceCount = 0 Then
        private_MergePeriodsByPerson = 0
        Exit Function
    End If
    ReDim result(1 To sourceCount)
    resultCount = 1
    Set result(resultCount) = source(1)
    For sourceIndex = 2 To sourceCount
        If sourceIndex Mod m_interval = 0 Then
            Application.StatusBar = VBA.Replace( _
                VBA.Replace(m_configuration.GetText("status.MergingProgress"), "{current}", VBA.CStr(sourceIndex)), _
                "{total}", VBA.CStr(sourceCount) _
            )
            If private_IsCancellationRequested(True) Then
                private_MergePeriodsByPerson = resultCount
                Exit Function
            End If
        End If
        If private_IsSamePerson(result(resultCount), source(sourceIndex)) Then
            If private_IsContinuous(result(resultCount), source(sourceIndex)) Then
                private_MergeIntoPeriod result(resultCount), source(sourceIndex)
            Else
                resultCount = resultCount + 1
                Set result(resultCount) = source(sourceIndex)
            End If
        Else
            resultCount = resultCount + 1
            Set result(resultCount) = source(sourceIndex)
        End If
    Next sourceIndex
    ReDim Preserve result(1 To resultCount)
    private_MergePeriodsByPerson = resultCount
End Function

' //
' // Merge two periods
' //
Private Sub private_MergeIntoPeriod( _
    ByRef targetPeriod As obj_PG_Period, _
    ByRef sourcePeriod As obj_PG_Period _
)
    If Len(targetPeriod.EventName) = 0 Then
        targetPeriod.EventName = sourcePeriod.EventName
    ElseIf Len(sourcePeriod.EventName) > 0 Then
        targetPeriod.EventName = VBA.Replace( _
            VBA.Replace(m_configuration.GetText("legacy.EVENTS_SEPARATOR"), "{left}", targetPeriod.EventName), _
            "{right}", sourcePeriod.EventName _
        )
    End If
    targetPeriod.LastEventName = sourcePeriod.LastEventName
    targetPeriod.HasDisplayableEvent = _
        targetPeriod.HasDisplayableEvent Or sourcePeriod.HasDisplayableEvent
    targetPeriod.IsOpen = targetPeriod.IsOpen Or sourcePeriod.IsOpen
    If sourcePeriod.dateTo > targetPeriod.dateTo Then
        targetPeriod.dateTo = sourcePeriod.dateTo
        targetPeriod.ArrivalOrder = sourcePeriod.ArrivalOrder
    End If
End Sub

' //
' // Сравнение людей
' //
Private Function private_IsSamePerson( _
    ByRef leftPeriod As obj_PG_Period, _
    ByRef rightPeriod As obj_PG_Period _
) As Boolean
    private_IsSamePerson = StrComp(leftPeriod.taxId, rightPeriod.taxId, vbTextCompare) = 0
End Function

' //
' // Непрерывность
' //
Private Function private_IsContinuous( _
    ByRef currentPeriod As obj_PG_Period, _
    ByRef NextPeriod As obj_PG_Period _
) As Boolean
    private_IsContinuous = False
    If Not private_IsMergeableEvent(currentPeriod.LastEventName) Then
        Exit Function
    End If
    If Not private_IsMergeableEvent(NextPeriod.EventName) Then
        Exit Function
    End If
    ' Периоды пересекаются или имеют общую граничную дату.
    If NextPeriod.dateFrom <= currentPeriod.dateTo Then
        private_IsContinuous = True
        Exit Function
    End If
    ' Правило объединения соседних календарных дней временно отключено.
    '
    ' 01.06 - 10.06
    ' 11.06 - 20.06
    '
    ' Соседние периоды объединяются только при совпадении номеров приказов.
    ' If NextPeriod.dateFrom = DateAdd("d", 1, currentPeriod.dateTo) Then
    '     If private_OrdersMatch(currentPeriod.ArrivalOrder, NextPeriod.DepartureOrder) Then
    '         private_IsContinuous = True
    '     End If
    ' End If
End Function

Private Function private_IsMergeableEvent(ByVal eventName As String) As Boolean
    private_IsMergeableEvent = private_IsEventInList(eventName, m_configuration.GetText("legacy.MERGEABLE_EVENTS"))
End Function

Private Function private_IsDisplayableEvent(ByVal eventName As String) As Boolean
    private_IsDisplayableEvent = private_IsEventInList(eventName, m_configuration.GetText("legacy.DISPLAYABLE_EVENTS"))
End Function

Private Function private_IsEventInList( _
    ByVal eventName As String, _
    ByVal eventList As String _
) As Boolean
    eventName = Trim$(eventName)
    If Len(eventName) = 0 Then
        private_IsEventInList = False
        Exit Function
    End If
    private_IsEventInList = InStr( _
        1, _
        "|" & eventList & "|", _
        "|" & eventName & "|", _
        vbTextCompare _
    ) > 0
End Function

' //
' // Сравнение приказов
' //
Private Function private_OrdersMatch( _
    ByVal closingOrder As String, _
    ByVal openingOrder As String _
) As Boolean
    closingOrder = Trim$(closingOrder)
    openingOrder = Trim$(openingOrder)
    private_OrdersMatch = False
    If Len(closingOrder) = 0 Then
        Exit Function
    End If
    If Len(openingOrder) = 0 Then
        Exit Function
    End If
    private_OrdersMatch = StrComp(closingOrder, openingOrder, vbTextCompare) = 0
End Function

' ============================================================
' Фильтрация по диапазону
'
' Исходные даты периодов сохраняются без обрезки.
' В результат попадают только периоды с событиями из m_configuration.GetText("legacy.DISPLAYABLE_EVENTS"),
' которые пересекаются с заданным диапазоном.
' ============================================================
Private Function private_FilterPeriodsByRange( _
    ByRef periods() As obj_PG_Period, _
    ByVal periodCount As Long, _
    ByVal filterDateFrom As Date, _
    ByVal filterDateTo As Date _
) As Long
    Dim readIndex As Long
    Dim writeIndex As Long

    writeIndex = 0
    For readIndex = 1 To periodCount
        If readIndex Mod m_interval = 0 Then
            If private_IsCancellationRequested(True) Then
                private_FilterPeriodsByRange = writeIndex
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
            Set periods(writeIndex) = periods(readIndex)
        End If
NextPeriod:
    Next readIndex
    private_FilterPeriodsByRange = writeIndex
End Function



Private Function private_SafeString(ByVal value As Variant) As String
    If IsError(value) Then
        private_SafeString = vbNullString
    ElseIf IsNull(value) Then
        private_SafeString = vbNullString
    ElseIf IsEmpty(value) Then
        private_SafeString = vbNullString
    Else
        private_SafeString = CStr(value)
    End If
End Function

Private Function private_IsCancellationRequested( _
    Optional ByVal yieldToUi As Boolean = False _
) As Boolean
    m_runContext.CheckCancel 1
    If yieldToUi Then
        m_runContext.CheckCancel 0
    End If
    ' Отмена прерывает весь запуск исключением из CheckCancel.
    private_IsCancellationRequested = False
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2300, "obj_PG_CalculationService", "Service is not initialized."
    End If
End Sub