Attribute VB_Name = "module_PersonnelAtDisposalDays"
Option Explicit

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
Private Const SOURCE_COL_NAME As String = "ПІБ"
Private Const ROSTER_SHEET_NAME As String = "Список"
Private Const ROSTER_TABLE_NAME As String = "Список"
Private Const ROSTER_COL_NAME As String = "ПІБ"
Private Const ROSTER_COL_TAX_ID As String = "ІПН"
Private Const MSG_PERSON_UNVERIFIED As String = "Person not found in events or the personnel roster: "
Private Const MSG_AMBIGUOUS_NAME As String = "Full name matches multiple tax IDs: "
Private Const SOURCE_COL_TAX_ID As String = "ІПН"
Private Const SOURCE_COL_EVENT As String = "Подія"
Private Const SOURCE_COL_PERIOD_FROM As String = "Вибуття"
Private Const SOURCE_COL_PERIOD_TO As String = "Прибуття"

' ----------------------------
' Target
' ----------------------------
Private Const TARGET_SHEET_NAME As String = "Розрахунок днів РОЗП"
Private Const TARGET_TABLE_NAME As String = "РозрахунокДнівРОЗП"
Private Const TARGET_COL_NAME As String = "ПІБ"
Private Const TARGET_COL_TAX_ID As String = "ІПН"
Private Const TARGET_COL_START As String = "Дата виведення у розпорядження"
Private Const TARGET_COL_DAYS As String = "Кількість днів у розпорядженні"
Private Const ERROR_REPORT_NAME As String = "DisposalDaysErrorReport"
Private Const ERROR_REPORT_GAP As Long = 2
Private Const ERROR_COL_NAME As String = "Full name"
Private Const ERROR_COL_TAX_ID As String = "Tax ID"
Private Const ERROR_COL_DESCRIPTION As String = "Error"
Private Const MSG_SKIPPED_PEOPLE As String = "; skipped: "
Private Const TARGET_COL_TRIPS As String = "Враховувати відрядження"
Private Const TARGET_COL_PERIODS As String = "Періоди"

' ----------------------------
' Parameters
' ----------------------------
Private Const INPUT_REFERENCE_SEPARATOR As String = "|"
Private Const INPUT_SHEET_SEPARATOR As String = "!"
Private Const PARAM_INPUT_PATH_CELL As String = "I2"
Private Const PARAM_END_DATE_CELL As String = "I3"

Private Const PARAM_COL_NAME As String = "ПІБ"
Private Const PARAM_COL_TAX_ID As String = "ІПН"
Private Const PARAM_COL_START As String = "ДЗС"
Private Const PARAM_COL_TRIPS As String = "ВІДРЯДЖЕННЯ"

' ----------------------------
' Events
' ----------------------------
' Эти события всегда исключаются из зачтённых дней.
' Для расширения списка добавьте название через |.
Private Const EXCLUDED_EVENTS As String = _
    "Стаціонарне лікування" & _
    "|Відпустка для лікування" & _
    "|ВЛК за межами"

' Эти события исключаются только при выключенном Враховувати відрядження.
Private Const OPTIONAL_EXCLUDED_EVENTS As String = _
    "Відрядження"

Private Const PRESENT_EVENT_NAME As String = "В строю"
Private Const EVENT_NAMES_SEPARATOR As String = " / "
Private Const PERIOD_LABEL_OPEN As String = " ("
Private Const PERIOD_LABEL_CLOSE As String = ")"
Private Const PERIODS_SEPARATOR As String = " | "
Private Const PERIOD_RANGE_SEPARATOR As String = "–"

' ----------------------------
' UI
' ----------------------------
Private Const UI_YIELD_INTERVAL As Long = 100
Private Const UI_YIELD_SECONDS As Single = 0.1

' ----------------------------
' Formats and dates
' ----------------------------
Private Const FORMAT_DATE As String = "dd.mm.yyyy"
Private Const FORMAT_TEXT As String = "@"
Private Const DATE_INTERVAL_DAY As String = "d"
Private Const DATE_TEXT_PATTERN As String = "##.##.####"
Private Const FORMAT_INTEGER As String = "0"
Private Const DATE_1904_OFFSET As Long = 1462
Private Const MIN_DATE_SERIAL As Long = 61
Private Const MAX_DATE_SERIAL As Long = 2958465
Private Const NBSP_CODE As Long = 160

' ----------------------------
' Logging
' ----------------------------
Private Const LOG_FILE_SUFFIX As String = "_logs.txt"
Private Const LOG_ERROR_PREFIX As String = "ERROR: "
Private Const LOG_WARNING_PREFIX As String = "WARNING: "
Private Const LOG_DEBUG_PREFIX As String = "DEBUG: "
Private Const FORMAT_LOG_TIMESTAMP As String = "yyyy-mm-dd hh:nn:ss"
Private Const LOG_ERROR_SOURCE As String = "WriteLog"

' ----------------------------
' Messages
' ----------------------------
Private Const MSG_ALREADY_RUNNING As String = "An operation is already running. Wait for completion or click Cancel."
Private Const MSG_PARAMETER_LOADING_STARTED As String = "Parameter loading started"
Private Const MSG_SELECT_TOOL_SHEET As String = "Select the tool worksheet containing the input path and end date."
Private Const MSG_WRONG_WORKBOOK As String = "Run this command from a worksheet in the tool workbook."
Private Const MSG_PARAMETER_WORKSHEET As String = "Parameter worksheet: "
Private Const MSG_PARAMETER_LOADING_COMPLETED As String = "Parameter loading completed"
Private Const MSG_INPUT_FILE As String = "Input file: "
Private Const MSG_END_DATE As String = "End date: "
Private Const MSG_OPERATION_CANCELLED As String = "Operation cancelled."
Private Const MSG_OPERATION_STOPPED As String = "Operation stopped: "
Private Const MSG_CANCELLING_OPERATION As String = "Cancelling operation..."
Private Const MSG_RUKH_SOURCE As String = "RUKH source: "
Private Const MSG_PARAMETER_CLOSE_FAILED As String = "Unable to close the parameter workbook: "
Private Const MSG_PARAMETER_SHEET_MISSING As String = "Parameter worksheet not found: "
Private Const MSG_PARAMETERS_EMPTY As String = "The parameter table is empty."
Private Const MSG_SOURCE_FILTER_FAILED As String = "Unable to clear filters in source table "
Private Const MSG_SOURCE_EMPTY As String = "The source RUKH table is empty."
Private Const MSG_PEOPLE As String = "People: "
Private Const MSG_RUKH_EVENTS As String = "; RUKH events: "
Private Const MSG_WORKSHEET_ROW As String = ", worksheet row "
Private Const MSG_TAX_ID As String = ": Tax ID"
Private Const MSG_DUPLICATE_TAX_ID As String = ": duplicate tax ID "
Private Const MSG_TIME_POINT As String = ": Start date"
Private Const MSG_START_AFTER_END As String = ": the start date is later than the end date."
Private Const MSG_FULL_NAME As String = ": Full name"
Private Const MSG_EVENT As String = ": Event"
Private Const MSG_DEPARTURE As String = ": Departure"
Private Const MSG_ARRIVAL_CELL_ERROR As String = ": the arrival cell contains an error."
Private Const MSG_ARRIVAL As String = ": Arrival"
Private Const MSG_ARRIVAL_BEFORE_DEPARTURE As String = ": arrival is earlier than departure."
Private Const MSG_TAX_ID_PREFIX As String = "Tax ID "
Private Const MSG_PERSON_NOT_FOUND As String = " was not found in RUKH. Check the ID and source workbook."
Private Const MSG_WRITING_RESULT As String = "Writing result: "
Private Const MSG_ROWS As String = " rows"
Private Const MSG_TARGET_BODY_MISSING As String = "The result table has no data rows after preparation."
Private Const MSG_TARGET_PREPARE_FAILED As String = "Unable to prepare the result table: "
Private Const MSG_TARGET_ROWS_READY As String = "Result table rows prepared: "
Private Const MSG_WRITING_RESULTS As String = "Writing results..."
Private Const MSG_CALCULATION_COMPLETED_PEOPLE As String = "Calculation completed. People: "
Private Const MSG_NO_EXCLUSIONS As String = "no exclusions"
Private Const MSG_OVERLAPS_DUPLICATES_ADJACENCY_AND_GAP As String = "overlaps, duplicates, adjacency and gap"
Private Const MSG_BUSINESS_TRIP_EXCLUDED As String = "business trip excluded"
Private Const MSG_BUSINESS_TRIP_INCLUDED As String = "business trip included"
Private Const MSG_TREATMENT_EXCLUDED As String = "treatment excluded"
Private Const MSG_TEST As String = "test"
Private Const MSG_TEXT_DATE As String = "text date"
Private Const MSG_1904_DATE_SYSTEM As String = "1904 date system"
Private Const MSG_EXAMPLE_EXCLUSIONS As String = "example exclusions"
Private Const MSG_EXAMPLE_RESULT As String = "example result"
Private Const MSG_ALGORITHM_CHECKS_PASSED As String = "Algorithm checks passed."
Private Const MSG_CHECK_FAILED As String = "Check failed: "
Private Const MSG_RELATIVE_PATH_BASE As String = "Save the tool workbook in a local or network folder before using a relative input path."
Private Const MSG_AMBIGUOUS_PATH As String = "Use a full drive or UNC path, or a path relative to the tool workbook: "
Private Const MSG_FILE_NOT_FOUND_OR_INACCESSIBLE As String = ": file not found or inaccessible: "
Private Const MSG_CANCELLED As String = "Cancelled"
Private Const MSG_CALCULATING_DAYS As String = "Calculating days: "
Private Const MSG_MULTIPLE_SOURCES As String = "Multiple RUKH tables are open. Keep only one source workbook open."
Private Const MSG_OPEN_SOURCE As String = "Open the RUKH workbook: worksheet '"
Private Const MSG_TABLE As String = "', table '"
Private Const MSG_TABLE_NOT_FOUND As String = "Table not found: '"
Private Const MSG_ON_WORKSHEET As String = "' on worksheet '"
Private Const MSG_CHECK_TABLE_CONFIGURATION As String = "'. Check the table and worksheet names in Configuration."
Private Const MSG_TABLE_TEXT As String = "Table '"
Private Const MSG_IS_MISSING_COLUMN As String = "' is missing column '"
Private Const MSG_RESULT_COLUMN_COUNT As String = "The result table has an unexpected column count."
Private Const MSG_DISABLE_TOTALS As String = "Disable the totals row in the result table."
Private Const MSG_CELL_ERROR As String = ": the cell contains an error."
Private Const MSG_REQUIRED_VALUE As String = ": a required value is missing."
Private Const MSG_SUBTRACT_BUSINESS_TRIPS As String = ": Count business trip days"
Private Const MSG_INVALID_TRIP_FLAG As String = ": the business trip flag must be TRUE / FALSE or 1 / 0 (Ukrainian and Russian equivalents are accepted)."
Private Const MSG_INVALID_DATE As String = ": enter a valid Excel date or dd.mm.yyyy text (not earlier than 01.03.1900)."
Private Const MSG_EXPECTED As String = ": expected "
Private Const MSG_ACTUAL As String = ", actual "
Private Const MSG_SAVE_BEFORE_LOGGING As String = "Save the tool workbook before enabling logging."
Private Const MSG_OPERATION_CANCELLED_BY_THE_USER As String = "Operation cancelled by the user."
Private Const MSG_NUMBER As String = "Number="
Private Const MSG_DESCRIPTION As String = " | Description="
Private Const MSG_LOG_WRITE_FAILED As String = "Unable to write the log: "
Private Const MSG_INVALID_REFERENCE As String = "Use file.xlsx|Worksheet!B2, where B2 is the first table header cell. "
Private Const MSG_PARAMETER_TABLE_MISSING As String = "No table starts at the specified header cell: "

' ----------------------------
' Validation and runtime
' ----------------------------
Private Const CANCEL_ERROR As Long = vbObjectError + 2101
Private Const DICTIONARY_PROG_ID As String = "Scripting.Dictionary"
Private Const FILE_SYSTEM_PROG_ID As String = "Scripting.FileSystemObject"
Private Const ERROR_SOURCE As String = "PersonnelAtDisposalDays"
Private Const OPERATION_ERROR As Long = vbObjectError + 2100

' ----------------------------
' Test data
' ----------------------------
Private Const TEST_START_DATE_TEXT As String = "25.05.2026"

' ============================================================
' Состояние выполнения и отмены
' ============================================================
Private mRunning As Boolean
Private mCancel As Boolean
Private mLastUiYield As Single

' ============================================================
' Снимок параметров запуска
' ============================================================
Private Type CalculationParameters
    InputPath As String
    SheetName As String
    HeaderAddress As String
    EndDay As Long
End Type

' ============================================================
' Запуск с текущего интерфейса
' ============================================================
' Читает I2/I3, открывает параметры и записывает рассчитанные дни в целевую таблицу.
Public Sub CalculateMovementDays()
    Dim parameterSheet As Worksheet
    Dim calculationParameters As CalculationParameters
    Dim parameterWorkbook As Workbook, params As ListObject
    Dim openedHere As Boolean
    Dim oldStatus As Variant, errorText As String, errorNumber As Long
    If mRunning Then
        VBA.MsgBox MSG_ALREADY_RUNNING, vbExclamation
        Exit Sub
    End If
    oldStatus = Application.StatusBar
    On Error GoTo Failed
    mRunning = True
    mCancel = False
    ClearLog
    LogDebug MSG_PARAMETER_LOADING_STARTED
    ' Сохраняем ссылку до DoEvents: пользователь может переключить лист.
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Fail MSG_SELECT_TOOL_SHEET
    Set parameterSheet = Application.ActiveSheet
    If Not parameterSheet.Parent Is ThisWorkbook Then Fail MSG_WRONG_WORKBOOK
    CheckCancel 0
    LogDebug MSG_PARAMETER_WORKSHEET & parameterSheet.Name
    ReadCalculationParameters parameterSheet, calculationParameters
    CheckCancel 0
    LogDebug MSG_PARAMETER_LOADING_COMPLETED
    Set parameterWorkbook = OpenParameterWorkbook(calculationParameters.InputPath, openedHere)
    CheckCancel 0
    Set params = ReadParameterTable(parameterWorkbook, calculationParameters)
    BuildMovementDays params, calculationParameters.EndDay
Cleanup:
    If openedHere Then CloseParameterWorkbook parameterWorkbook
    Application.StatusBar = oldStatus
    mRunning = False
    mCancel = False
    Exit Sub
Failed:
    errorText = VBA.Err.Description
    errorNumber = VBA.Err.Number
    LogFailure errorNumber, errorText
    If errorNumber = CANCEL_ERROR Then
        VBA.MsgBox MSG_OPERATION_CANCELLED, vbInformation
    Else
        VBA.MsgBox MSG_OPERATION_STOPPED & errorText, vbExclamation
    End If
    Resume Cleanup
End Sub

' ============================================================
' Запрос отмены
' ============================================================
' Кнопка Скасувати только устанавливает флаг. Обработчик CheckCancel
' завершает операцию в ближайшей точке передачи управления через DoEvents.
Public Sub CancelMovementDays()
    If mRunning Then
        mCancel = True
        Application.StatusBar = MSG_CANCELLING_OPERATION
    End If
End Sub

' ============================================================
' Проверка алгоритма
' ============================================================
' Проверяет интервалы и преобразование параметров без изменения листов.
' Запускается вручную в Excel.
Public Sub TestMovementDays()
    Dim rows As Collection, namedRows As Collection
    Dim item As Variant
    Dim periodsText As String
    On Error GoTo Failed
    Set rows = New Collection
    AssertDays UnionDays(rows), 0, MSG_NO_EXCLUSIONS
    AssertDays BuildCountedPeriods(rows, VBA.CLng(VBA.DateSerial(2026, 8, 1)), _
        VBA.CLng(VBA.DateSerial(2026, 8, 31)), periodsText), 30, MSG_NO_EXCLUSIONS
    rows.Add VBA.Array(10, 20)
    rows.Add VBA.Array(15, 25)
    rows.Add VBA.Array(10, 20)
    rows.Add VBA.Array(25, 30)
    rows.Add VBA.Array(35, 40)
    AssertDays UnionDays(rows), 25, MSG_OVERLAPS_DUPLICATES_ADJACENCY_AND_GAP
    AssertDays VBA.CLng(IsExcluded("Відрядження", True)), -1, MSG_BUSINESS_TRIP_EXCLUDED
    AssertDays VBA.CLng(IsExcluded("Відрядження", False)), 0, MSG_BUSINESS_TRIP_INCLUDED
    AssertDays VBA.CLng(IsExcluded("Стаціонарне лікування", False)), -1, MSG_TREATMENT_EXCLUDED
    AssertDays ReadDay(TEST_START_DATE_TEXT, False, MSG_TEST), VBA.CLng(VBA.DateSerial(2026, 5, 25)), MSG_TEXT_DATE
    AssertDays ReadDay(44705, True, MSG_TEST), 46167, MSG_1904_DATE_SYSTEM
    ' Пример со скриншота после обрезки событий границами расчёта.
    Set rows = New Collection
    rows.Add VBA.Array(VBA.CLng(VBA.DateSerial(2026, 5, 25)), VBA.CLng(VBA.DateSerial(2026, 6, 2)))
    rows.Add VBA.Array(VBA.CLng(VBA.DateSerial(2026, 6, 2)), VBA.CLng(VBA.DateSerial(2026, 7, 2)))
    rows.Add VBA.Array(VBA.CLng(VBA.DateSerial(2026, 7, 2)), VBA.CLng(VBA.DateSerial(2026, 7, 11)))
    rows.Add VBA.Array(VBA.CLng(VBA.DateSerial(2026, 7, 11)), VBA.CLng(VBA.DateSerial(2026, 8, 9)))
    rows.Add VBA.Array(VBA.CLng(VBA.DateSerial(2026, 8, 11)), VBA.CLng(VBA.DateSerial(2026, 8, 31)))
    AssertDays UnionDays(rows), 96, MSG_EXAMPLE_EXCLUSIONS
    AssertDays VBA.DateDiff(DATE_INTERVAL_DAY, VBA.DateSerial(2026, 5, 25), VBA.DateSerial(2026, 8, 31)) - UnionDays(rows), 2, MSG_EXAMPLE_RESULT
    Set namedRows = New Collection
    For Each item In rows
        namedRows.Add VBA.Array(item(0), item(1), "Test exclusion", True)
    Next item
    AssertDays BuildCountedPeriods(namedRows, VBA.CLng(VBA.DateSerial(2026, 5, 25)), _
        VBA.CLng(VBA.DateSerial(2026, 8, 31)), periodsText), 2, MSG_EXAMPLE_RESULT
    VBA.MsgBox MSG_ALGORITHM_CHECKS_PASSED, vbInformation
    Exit Sub
Failed:
    VBA.MsgBox MSG_CHECK_FAILED & VBA.Err.Description, vbCritical
End Sub

' ============================================================
' Расчёт по таблице параметров и открытому РУХ
' ============================================================
Private Sub BuildMovementDays(ByVal params As ListObject, ByVal lastDay As Long)
    Dim source As ListObject, target As ListObject, roster As ListObject
    Dim people As Object, intervals As Object
    Dim eventTaxIndex As Object, eventNameIndex As Object, eventNameIds As Object
    Dim rosterTaxIndex As Object, rosterNameIndex As Object, rosterNameIds As Object
    Dim rows As Collection
    Dim p As Variant, s As Variant, rosterData As Variant
    Dim matches As Collection, rosterMatches As Collection, rowNumber As Variant
    Dim sourceNameCol As Long, rosterTaxCol As Long, rosterNameCol As Long
    Dim matchError As String, resolvedTaxId As String
    Dim personErrors() As String
    Dim resultNames() As Variant, resultTaxIds() As Variant
    Dim resultStarts() As Variant, resultDays() As Variant
    Dim resultPeriods() As Variant, periodsText As String
    Dim resultTrips() As Variant
    Dim starts() As Long, trips() As Boolean
    Dim taxCol As Long, eventCol As Long, fromCol As Long, toCol As Long
    Dim pTax As Long, pName As Long, pStart As Long, pTrip As Long
    Dim i As Long, person As Long, count As Long
    Dim parameterRows() As Long
    Dim failures As Collection, validCount As Long
    Dim firstDay As Long, arrival As Long
    Dim taxId As String, eventName As String, context As String
    ' --------------------------------------------------------
    ' Поиск таблиц и проверка структуры до чтения данных.
    ' --------------------------------------------------------
    Set source = FindSource()
    Set roster = RequireTable(ROSTER_SHEET_NAME, ROSTER_TABLE_NAME, source.Parent.Parent)
    rosterTaxCol = ColumnIndex(roster, ROSTER_COL_TAX_ID)
    rosterNameCol = ColumnIndex(roster, ROSTER_COL_NAME)
    If Not roster.DataBodyRange Is Nothing Then rosterData = roster.DataBodyRange.Value2
    ClearSourceFilters source
    CheckCancel 0
    Set target = RequireTable(TARGET_SHEET_NAME, TARGET_TABLE_NAME)
    LogDebug MSG_RUKH_SOURCE & source.Parent.Parent.Name & " / " & source.Parent.Name & " / " & source.Name
    ValidateOutput target
    pTax = ColumnIndex(params, PARAM_COL_TAX_ID)
    pName = ColumnIndex(params, PARAM_COL_NAME)
    pStart = ColumnIndex(params, PARAM_COL_START)
    pTrip = ColumnIndex(params, PARAM_COL_TRIPS)
    sourceNameCol = ColumnIndex(source, SOURCE_COL_NAME)
    taxCol = ColumnIndex(source, SOURCE_COL_TAX_ID)
    eventCol = ColumnIndex(source, SOURCE_COL_EVENT)
    fromCol = ColumnIndex(source, SOURCE_COL_PERIOD_FROM)
    toCol = ColumnIndex(source, SOURCE_COL_PERIOD_TO)
    ' --------------------------------------------------------
    ' Пакетное чтение таблиц и подготовка индексов людей в памяти.
    ' --------------------------------------------------------
    If params.DataBodyRange Is Nothing Then Fail MSG_PARAMETERS_EMPTY
    p = params.DataBodyRange.Value2
    If Not source.DataBodyRange Is Nothing Then s = source.DataBodyRange.Value2
    BuildPersonIndex s, source.ListRows.Count, taxCol, sourceNameCol, _
        eventTaxIndex, eventNameIndex, eventNameIds
    BuildPersonIndex rosterData, roster.ListRows.Count, rosterTaxCol, rosterNameCol, _
        rosterTaxIndex, rosterNameIndex, rosterNameIds
    count = CompactParameterRows(p, parameterRows)
    If count = 0 Then Fail MSG_PARAMETERS_EMPTY
    LogDebug MSG_PEOPLE & count & MSG_RUKH_EVENTS & source.ListRows.Count
    ReDim personErrors(1 To count)
    ReDim starts(1 To count)
    ReDim trips(1 To count)
    ReDim resultNames(1 To count, 1 To 1)
    ReDim resultTaxIds(1 To count, 1 To 1)
    ReDim resultStarts(1 To count, 1 To 1)
    ReDim resultDays(1 To count, 1 To 1)
    ReDim resultPeriods(1 To count, 1 To 1)
    ReDim resultTrips(1 To count, 1 To 1)
    Set failures = New Collection
    Set people = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set intervals = VBA.CreateObject(DICTIONARY_PROG_ID)
    ' --------------------------------------------------------
    ' Проверка индивидуальных дат и флагов; подготовка строк результата.
    ' --------------------------------------------------------
    For i = 1 To count
        context = params.Parent.Name & MSG_WORKSHEET_ROW & params.DataBodyRange.Row + parameterRows(i) - 1
        taxId = RequiredText(p(i, pTax), context & MSG_TAX_ID)
        If people.Exists(taxId) Then Fail context & MSG_DUPLICATE_TAX_ID & taxId
        people.Add taxId, i
        starts(i) = ReadDay(p(i, pStart), params.Parent.Parent.Date1904, context & MSG_TIME_POINT)
        If starts(i) > lastDay Then Fail context & MSG_START_AFTER_END
        ' Включённый флаг сохраняет дни командировок в зачтённом времени.
        trips(i) = ReadFlag(p(i, pTrip), context)
        ' Сохраняем исходное значение флага, включая пустую ячейку.
        resultTrips(i, 1) = p(i, pTrip)
        resultTaxIds(i, 1) = taxId
        resultNames(i, 1) = RequiredText(p(i, pName), context & MSG_FULL_NAME)
        resultStarts(i, 1) = VBA.CDate(starts(i))
        Set rows = New Collection
        intervals.Add taxId, rows
        CheckCancel i
    Next i
    ' После сброса фильтров обрабатываются все строки таблицы РУХ.
    ' --------------------------------------------------------
    ' Сбор всех событий выбранных людей с обрезкой по границам расчёта.
    ' --------------------------------------------------------
    For person = 1 To count
        CheckCancel 0
        taxId = resultTaxIds(person, 1)
        Set matches = FindPersonRows(eventTaxIndex, eventNameIndex, eventNameIds, _
            taxId, VBA.CStr(resultNames(person, 1)), matchError)
        If VBA.Len(matchError) > 0 Then
            personErrors(person) = matchError
            GoTo NextMatchedPerson
        End If
        If matches.Count = 0 Then
            Set rosterMatches = FindPersonRows(rosterTaxIndex, rosterNameIndex, rosterNameIds, _
                taxId, VBA.CStr(resultNames(person, 1)), matchError)
            If VBA.Len(matchError) > 0 Then
                personErrors(person) = matchError
                GoTo NextMatchedPerson
            End If
            If rosterMatches.Count = 0 Then
                personErrors(person) = MSG_PERSON_UNVERIFIED & resultNames(person, 1) & " / " & taxId
                GoTo NextMatchedPerson
            End If
            ' Найденный по полному имени ІПН уточняет поиск событий для случая ДРАЙ.
            resolvedTaxId = MatchText(rosterData(rosterMatches(1), rosterTaxCol))
            If VBA.Len(resolvedTaxId) > 0 Then
                Set matches = FindPersonRows(eventTaxIndex, eventNameIndex, eventNameIds, _
                    resolvedTaxId, VBA.CStr(resultNames(person, 1)), matchError)
                If VBA.Len(matchError) > 0 Then
                    personErrors(person) = matchError
                    GoTo NextMatchedPerson
                End If
            End If
        End If
        ' Наличие в Списке подтверждено, даже если коллекция событий пуста.
        For Each rowNumber In matches
            CheckCancel VBA.CLng(rowNumber)
            context = source.Parent.Parent.Name & " / " & SOURCE_SHEET_NAME & _
                MSG_WORKSHEET_ROW & source.DataBodyRange.Row + VBA.CLng(rowNumber) - 1
            eventName = RequiredText(s(VBA.CLng(rowNumber), eventCol), context & MSG_EVENT)
            firstDay = ReadDay(s(VBA.CLng(rowNumber), fromCol), source.Parent.Parent.Date1904, context & MSG_DEPARTURE)
            If VBA.IsError(s(VBA.CLng(rowNumber), toCol)) Then Fail context & MSG_ARRIVAL_CELL_ERROR
            If VBA.Len(VBA.Trim$(VBA.CStr(s(VBA.CLng(rowNumber), toCol)))) = 0 Then
                arrival = lastDay
            Else
                arrival = ReadDay(s(VBA.CLng(rowNumber), toCol), source.Parent.Parent.Date1904, context & MSG_ARRIVAL)
                If arrival < firstDay Then Fail context & MSG_ARRIVAL_BEFORE_DEPARTURE
            End If
            If firstDay < starts(person) Then firstDay = starts(person)
            If arrival > lastDay Then arrival = lastDay
            If firstDay < arrival Then
                Set rows = intervals(VBA.CStr(resultTaxIds(person, 1)))
                rows.Add VBA.Array(firstDay, arrival, eventName, IsExcluded(eventName, Not trips(person)))
            End If
        Next rowNumber
NextMatchedPerson:
    Next person
    ' --------------------------------------------------------
    ' Вычитание объединённых интервалов из общего количества дней.
    ' --------------------------------------------------------
    For i = 1 To count
        CheckCancel i
        taxId = resultTaxIds(i, 1)
        If VBA.Len(personErrors(i)) > 0 Then
            failures.Add VBA.Array(resultNames(i, 1), taxId, personErrors(i))
            LogWarning personErrors(i)
            GoTo NextResultPerson
        End If
        ' Без событий период В строю разрешён только после проверки Списка.
        Set rows = intervals(taxId)
        validCount = validCount + 1
        resultDays(validCount, 1) = BuildCountedPeriods(rows, starts(i), lastDay, periodsText)
        resultPeriods(validCount, 1) = periodsText
        resultNames(validCount, 1) = resultNames(i, 1)
        resultTaxIds(validCount, 1) = taxId
        resultStarts(validCount, 1) = resultStarts(i, 1)
        resultTrips(validCount, 1) = resultTrips(i, 1)
NextResultPerson:
    Next i
    count = validCount
    ' Ошибки данных и отмена до этого места сохраняют предыдущий результат.
    CheckCancel 0
    LogDebug MSG_WRITING_RESULT & count & MSG_ROWS
    Application.StatusBar = MSG_WRITING_RESULTS
    ClearErrorReport
    If count = 0 Then
        If Not target.DataBodyRange Is Nothing Then target.DataBodyRange.Delete
        GoTo WriteErrors
    End If
    If Not target.DataBodyRange Is Nothing Then target.DataBodyRange.ClearContents
    PrepareTargetRows target, count, VBA.CStr(resultNames(1, 1))
    target.ListColumns(TARGET_COL_TAX_ID).DataBodyRange.NumberFormat = FORMAT_TEXT
    target.ListColumns(TARGET_COL_START).DataBodyRange.NumberFormat = FORMAT_DATE
    target.ListColumns(TARGET_COL_DAYS).DataBodyRange.NumberFormat = FORMAT_INTEGER
    target.ListColumns(TARGET_COL_PERIODS).DataBodyRange.NumberFormat = FORMAT_TEXT
    ' Пакетная запись в именованные колонки умной таблицы.
    WriteResultColumn target.ListColumns(TARGET_COL_NAME), resultNames, count
    WriteResultColumn target.ListColumns(TARGET_COL_TAX_ID), resultTaxIds, count
    WriteResultColumn target.ListColumns(TARGET_COL_START), resultStarts, count
    WriteResultColumn target.ListColumns(TARGET_COL_DAYS), resultDays, count
    WriteResultColumn target.ListColumns(TARGET_COL_PERIODS), resultPeriods, count
    WriteResultColumn target.ListColumns(TARGET_COL_TRIPS), resultTrips, count
WriteErrors:
    WriteErrorReport target, failures
    LogDebug MSG_CALCULATION_COMPLETED_PEOPLE & count & MSG_SKIPPED_PEOPLE & failures.Count
    VBA.MsgBox MSG_CALCULATION_COMPLETED_PEOPLE & count & MSG_SKIPPED_PEOPLE & failures.Count, vbInformation
End Sub





' ============================================================
' Сопоставление человека: сначала ІПН, затем полное имя
' ============================================================
' Проверяем индекс ІПН до перехода к индексу имени. Повторные события одного
' человека допустимы; разные ІПН у одинакового имени считаются неоднозначностью.
Private Function FindPersonRows(ByVal taxIndex As Object, ByVal nameIndex As Object, _
    ByVal nameIds As Object, ByVal taxId As String, ByVal fullName As String, _
    ByRef matchError As String) As Collection
    Dim ids As Object
    matchError = vbNullString
    If taxIndex.Exists(taxId) Then
        Set FindPersonRows = taxIndex(taxId)
    ElseIf nameIndex.Exists(fullName) Then
        Set FindPersonRows = nameIndex(fullName)
        Set ids = nameIds(fullName)
        If ids.Count > 1 Then matchError = MSG_AMBIGUOUS_NAME & fullName
    Else
        Set FindPersonRows = New Collection
    End If
End Function

' ============================================================
' Однократное индексирование людей в исходной таблице
' ============================================================
' Сохраняем номера строк по ІПН и полному имени. Набор ІПН каждого имени
' позволяет проверить неоднозначность без повторного перебора событий.
Private Sub BuildPersonIndex(ByRef data As Variant, ByVal count As Long, _
    ByVal taxColumn As Long, ByVal nameColumn As Long, ByRef taxIndex As Object, _
    ByRef nameIndex As Object, ByRef nameIds As Object)
    Dim i As Long, taxId As String, fullName As String
    Dim rows As Collection, ids As Object
    Set taxIndex = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set nameIndex = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set nameIds = VBA.CreateObject(DICTIONARY_PROG_ID)
    nameIndex.CompareMode = vbTextCompare
    nameIds.CompareMode = vbTextCompare
    For i = 1 To count
        CheckCancel i
        taxId = MatchText(data(i, taxColumn))
        fullName = MatchText(data(i, nameColumn))
        If VBA.Len(taxId) > 0 Then
            If Not taxIndex.Exists(taxId) Then
                Set rows = New Collection
                taxIndex.Add taxId, rows
            End If
            Set rows = taxIndex(taxId)
            rows.Add i
        End If
        If VBA.Len(fullName) > 0 Then
            If Not nameIndex.Exists(fullName) Then
                Set rows = New Collection
                nameIndex.Add fullName, rows
                Set ids = VBA.CreateObject(DICTIONARY_PROG_ID)
                nameIds.Add fullName, ids
            End If
            Set rows = nameIndex(fullName)
            rows.Add i
            Set ids = nameIds(fullName)
            If VBA.Len(taxId) > 0 Then ids(taxId) = True
        End If
    Next i
End Sub

' ============================================================
' Текст для поиска в исходных данных
' ============================================================
' Пустые и ошибочные ячейки не дают совпадения. Внутренние пробелы имени
' не схлопываются, сокращения и нечёткий поиск не применяются.
Private Function MatchText(ByVal value As Variant) As String
    If VBA.IsError(value) Or VBA.IsNull(value) Then Exit Function
    MatchText = VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))
End Function

' ============================================================
' Запись только успешно рассчитанных строк колонки
' ============================================================
Private Sub WriteResultColumn(ByVal column As ListColumn, ByRef values() As Variant, ByVal count As Long)
    Dim data() As Variant, i As Long
    ReDim data(1 To count, 1 To 1)
    For i = 1 To count
        data(i, 1) = values(i, 1)
    Next i
    column.DataBodyRange.Value = data
End Sub

' ============================================================
' Очистка предыдущего списка ошибок
' ============================================================
' Именованный диапазон отмечает только созданный инструментом отчёт.
Private Sub ClearErrorReport()
    Dim reportName As Name
    For Each reportName In ThisWorkbook.Names
        If VBA.StrComp(reportName.Name, ERROR_REPORT_NAME, vbTextCompare) = 0 Then
            reportName.RefersToRange.ClearContents
            reportName.Delete
            Exit Sub
        End If
    Next reportName
End Sub

' ============================================================
' Список пропущенных людей под основной таблицей
' ============================================================
Private Sub WriteErrorReport(ByVal target As ListObject, ByVal failures As Collection)
    Dim report As Range, data() As Variant, item As Variant, i As Long
    If failures.Count = 0 Then Exit Sub
    ReDim data(1 To failures.Count + 1, 1 To 3)
    data(1, 1) = ERROR_COL_NAME
    data(1, 2) = ERROR_COL_TAX_ID
    data(1, 3) = ERROR_COL_DESCRIPTION
    For i = 1 To failures.Count
        item = failures(i)
        data(i + 1, 1) = item(0)
        data(i + 1, 2) = item(1)
        data(i + 1, 3) = item(2)
    Next i
    Set report = target.Parent.Cells(target.Range.Row + target.Range.Rows.Count + _
        ERROR_REPORT_GAP, target.Range.Column).Resize(failures.Count + 1, 3)
    ThisWorkbook.Names.Add Name:=ERROR_REPORT_NAME, RefersTo:="=" & report.Address(External:=True)
    report.NumberFormat = FORMAT_TEXT
    report.Value = data
End Sub

' ============================================================
' Подготовка строк пустой целевой таблицы
' ============================================================
' После ClearContents и Resize Excel может не создать DataBodyRange
' для единственной пустой строки. Заполняем первую строку до Resize.
Private Sub PrepareTargetRows(ByVal target As ListObject, ByVal count As Long, ByVal firstName As String)
    Dim firstRow As ListRow
    On Error GoTo Failed
    If target.DataBodyRange Is Nothing Then
        Set firstRow = target.ListRows.Add
    End If
    If target.DataBodyRange Is Nothing Then Fail MSG_TARGET_BODY_MISSING
    target.ListColumns(TARGET_COL_NAME).DataBodyRange.Cells(1, 1).Value = firstName
    target.Resize target.HeaderRowRange.Resize(count + 1, target.ListColumns.Count)
    If target.DataBodyRange Is Nothing Then Fail MSG_TARGET_BODY_MISSING
    LogDebug MSG_TARGET_ROWS_READY & target.ListRows.Count
    Exit Sub
Failed:
    Fail MSG_TARGET_PREPARE_FAILED & VBA.Err.Description
End Sub

' ============================================================
' Открытие внешней книги параметров
' ============================================================
' Уже открытую книгу используем без закрытия. Новую открываем только для чтения.
Private Function OpenParameterWorkbook(ByVal filePath As String, ByRef openedHere As Boolean) As Workbook
    Dim workbook As Workbook
    Dim oldSecurity As Long, errorNumber As Long, errorText As String
    For Each workbook In Application.Workbooks
        If VBA.StrComp(workbook.FullName, filePath, vbTextCompare) = 0 Then
            Set OpenParameterWorkbook = workbook
            Exit Function
        End If
    Next workbook
    oldSecurity = Application.AutomationSecurity
    On Error GoTo Failed
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set OpenParameterWorkbook = Application.Workbooks.Open(Filename:=filePath, _
        UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    openedHere = True
    Application.AutomationSecurity = oldSecurity
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Application.AutomationSecurity = oldSecurity
    VBA.Err.Raise errorNumber, ERROR_SOURCE, errorText
End Function

' ============================================================
' Закрытие книги, открытой инструментом
' ============================================================
Private Sub CloseParameterWorkbook(ByVal workbook As Workbook)
    On Error GoTo Failed
    workbook.Close SaveChanges:=False
    Exit Sub
Failed:
    VBA.MsgBox MSG_PARAMETER_CLOSE_FAILED & VBA.Err.Description, vbExclamation
End Sub


' ============================================================
' Поиск таблицы параметров по первой ячейке заголовка
' ============================================================
' Лист и адрес берутся из I2. Имя таблицы и позиции колонок не фиксируются.
Private Function ReadParameterTable(ByVal workbook As Workbook, _
    ByRef parameters As CalculationParameters) As ListObject
    Dim ws As Worksheet, parameterSheet As Worksheet
    Dim headerCell As Range, table As ListObject
    For Each ws In workbook.Worksheets
        If VBA.StrComp(ws.Name, parameters.SheetName, vbTextCompare) = 0 Then
            Set parameterSheet = ws
            Exit For
        End If
    Next ws
    If parameterSheet Is Nothing Then Fail MSG_PARAMETER_SHEET_MISSING & parameters.SheetName
    Set headerCell = parameterSheet.Range(parameters.HeaderAddress)
    For Each table In parameterSheet.ListObjects
        If Not table.HeaderRowRange Is Nothing Then
            If table.HeaderRowRange.Cells(1, 1).Address = headerCell.Address Then
                Set ReadParameterTable = table
                Exit Function
            End If
        End If
    Next table
    Fail MSG_PARAMETER_TABLE_MISSING & parameters.SheetName & INPUT_SHEET_SEPARATOR & parameters.HeaderAddress
End Function

' ============================================================
' Пропуск полностью пустых строк таблицы параметров
' ============================================================
' Частично заполненные строки сохраняются для проверки обязательных полей.
Private Function CompactParameterRows(ByRef data As Variant, ByRef sourceRows() As Long) As Long
    Dim i As Long, j As Long, count As Long, hasValue As Boolean
    ReDim sourceRows(1 To UBound(data, 1))
    For i = 1 To UBound(data, 1)
        CheckCancel i
        hasValue = False
        For j = 1 To UBound(data, 2)
            If VBA.IsError(data(i, j)) Then
                hasValue = True
            ElseIf VBA.Len(VBA.Trim$(VBA.CStr(data(i, j)))) > 0 Then
                hasValue = True
            End If
        Next j
        If hasValue Then
            count = count + 1
            sourceRows(count) = i
            If count <> i Then
                For j = 1 To UBound(data, 2)
                    data(count, j) = data(i, j)
                Next j
            End If
        End If
    Next i
    CompactParameterRows = count
End Function

' ============================================================
' Чтение и проверка параметров запуска
' ============================================================
' Создаёт снимок пути и даты, затем проверяет доступность файла.
' Между чтением двух ячеек управление интерфейсу не передаётся.
Private Sub ReadCalculationParameters(ByVal ws As Worksheet, ByRef parameters As CalculationParameters)
    Dim fileSystem As Object
    ' Оба значения читаются подряд, без передачи управления между чтениями.
    parameters.InputPath = RequiredText(ws.Range(PARAM_INPUT_PATH_CELL).Value2, ws.Name & "!" & PARAM_INPUT_PATH_CELL)
    parameters.EndDay = ReadDay(ws.Range(PARAM_END_DATE_CELL).Value2, ws.Parent.Date1904, ws.Name & "!" & PARAM_END_DATE_CELL)
    ParseInputReference parameters
    parameters.InputPath = ResolveInputPath(parameters.InputPath)
    LogDebug MSG_INPUT_FILE & parameters.InputPath
    LogDebug MSG_END_DATE & VBA.Format$(VBA.CDate(parameters.EndDay), FORMAT_DATE)
    CheckCancel 0
    Set fileSystem = VBA.CreateObject(FILE_SYSTEM_PROG_ID)
    If Not fileSystem.FileExists(parameters.InputPath) Then _
        Fail ws.Name & "!" & PARAM_INPUT_PATH_CELL & MSG_FILE_NOT_FOUND_OR_INACCESSIBLE & parameters.InputPath
End Sub



' ============================================================
' Разбор ссылки на файл, лист и начало таблицы
' ============================================================
' Формат: путь|лист!B2. Лист может содержать пробелы и быть заключён в апострофы.
' Адрес принимается только как одна ячейка A1, без именованных диапазонов.
Private Sub ParseInputReference(ByRef parameters As CalculationParameters)
    Dim parts As Variant, location As String, separatorPosition As Long
    Dim address As String, i As Long, character As String, digitsStarted As Boolean
    parts = VBA.Split(parameters.InputPath, INPUT_REFERENCE_SEPARATOR)
    If UBound(parts) <> 1 Then Fail MSG_INVALID_REFERENCE
    parameters.InputPath = VBA.Trim$(parts(0))
    location = VBA.Trim$(parts(1))
    separatorPosition = VBA.InStrRev(location, INPUT_SHEET_SEPARATOR)
    If separatorPosition <= 1 Then Fail MSG_INVALID_REFERENCE
    parameters.SheetName = VBA.Trim$(VBA.Left$(location, separatorPosition - 1))
    If VBA.Left$(parameters.SheetName, 1) = "'" And VBA.Right$(parameters.SheetName, 1) = "'" Then
        parameters.SheetName = VBA.Replace(VBA.Mid$(parameters.SheetName, 2, _
            VBA.Len(parameters.SheetName) - 2), "''", "'")
    End If
    address = VBA.UCase$(VBA.Replace(VBA.Trim$(VBA.Mid$(location, separatorPosition + 1)), "$", ""))
    If VBA.Len(parameters.InputPath) = 0 Or VBA.Len(parameters.SheetName) = 0 Then Fail MSG_INVALID_REFERENCE
    If Not VBA.Left$(address, 1) Like "[A-Z]" Then Fail MSG_INVALID_REFERENCE
    For i = 1 To VBA.Len(address)
        character = VBA.Mid$(address, i, 1)
        If character Like "[0-9]" Then
            digitsStarted = True
        ElseIf Not character Like "[A-Z]" Or digitsStarted Then
            Fail MSG_INVALID_REFERENCE
        End If
    Next i
    If Not digitsStarted Then Fail MSG_INVALID_REFERENCE
    parameters.HeaderAddress = address
End Sub

' ============================================================
' Разрешение пути относительно книги инструмента
' ============================================================
' .\, ..\ и имя файла без папки отсчитываются от ThisWorkbook.Path.
' Полные пути диска и UNC не зависят от текущей папки Excel.
Private Function ResolveInputPath(ByVal inputPath As String) As String
    Dim fileSystem As Object
    Dim basePath As String
    Set fileSystem = VBA.CreateObject(FILE_SYSTEM_PROG_ID)
    inputPath = VBA.Replace(inputPath, "/", "\")
    If VBA.Left$(inputPath, 2) = "\\" Or inputPath Like "[A-Za-z]:\*" Then
        ResolveInputPath = fileSystem.GetAbsolutePathName(inputPath)
        Exit Function
    End If
    ' Пути вида C:file.xlsx и \file.xlsx зависят от текущего диска Excel.
    If VBA.Left$(inputPath, 1) = "\" Or VBA.InStr(1, inputPath, ":", vbBinaryCompare) > 0 Then
        Fail MSG_AMBIGUOUS_PATH & inputPath
    End If
    basePath = ThisWorkbook.Path
    If VBA.Len(basePath) = 0 Or VBA.InStr(1, basePath, "://", vbBinaryCompare) > 0 Then
        Fail MSG_RELATIVE_PATH_BASE
    End If
    ResolveInputPath = fileSystem.GetAbsolutePathName(fileSystem.BuildPath(basePath, inputPath))
End Function

' ============================================================
' Отмена и отзывчивость интерфейса
' ============================================================
' Проверяет запрос отмены и каждые 100 шагов обрабатывает события UI.
' Индекс 0 используется для обязательной проверки между этапами.
' Исключение CANCEL_ERROR обрабатывается публичным методом запуска.
Private Sub CheckCancel(ByVal index As Long)
    Dim currentTime As Single
    If Not mRunning Then Exit Sub
    If mCancel Then VBA.Err.Raise CANCEL_ERROR, ERROR_SOURCE, MSG_CANCELLED
    If index Mod UI_YIELD_INTERVAL <> 0 Then Exit Sub
    currentTime = VBA.Timer
    If index <> 0 And currentTime >= mLastUiYield Then
        If currentTime - mLastUiYield < UI_YIELD_SECONDS Then Exit Sub
    End If
    mLastUiYield = currentTime
    Application.StatusBar = MSG_CALCULATING_DAYS & index
    VBA.DoEvents
    If mCancel Then VBA.Err.Raise CANCEL_ERROR, ERROR_SOURCE, MSG_CANCELLED
End Sub

' ============================================================
' Правила исключения событий
' ============================================================
' Медицинские события исключаются всегда, командировки — по флагу человека.
Private Function IsExcluded(ByVal eventName As String, ByVal subtractTrips As Boolean) As Boolean
    IsExcluded = IsEventInList(eventName, EXCLUDED_EVENTS)
    If Not IsExcluded And subtractTrips Then
        IsExcluded = IsEventInList(eventName, OPTIONAL_EXCLUDED_EVENTS)
    End If
End Function

' ============================================================
' Поиск события в конфигурационном списке
' ============================================================
' Как в module_PeriodsGeneration: границы | исключают совпадение по подстроке.
Private Function IsEventInList(ByVal eventName As String, ByVal eventList As String) As Boolean
    eventName = VBA.Trim$(eventName)
    If VBA.Len(eventName) = 0 Then Exit Function
    IsEventInList = VBA.InStr(1, "|" & eventList & "|", "|" & eventName & "|", vbTextCompare) > 0
End Function


' ============================================================
' Зачтённые периоды и их общая длительность
' ============================================================
' Разбиваем шкалу по границам всех событий. Исключения имеют приоритет.
' Для зачтённых отрезков берём названия активных событий или В строю.
' Число дней и текст формируются из одних и тех же зачтённых интервалов.
Private Function BuildCountedPeriods(ByVal rows As Collection, ByVal firstDay As Long, _
    ByVal lastDay As Long, ByRef periodsText As String) As Long
    Dim boundaries() As Long, unused() As Long
    Dim i As Long, j As Long, count As Long, item As Variant
    Dim segmentStart As Long, segmentEnd As Long, pendingStart As Long, pendingEnd As Long
    Dim label As String, pendingLabel As String, excluded As Boolean
    Dim names As Object, key As Variant
    periodsText = vbNullString
    If firstDay >= lastDay Then Exit Function
    count = rows.Count * 2 + 2
    ReDim boundaries(1 To count)
    ReDim unused(1 To count)
    boundaries(1) = firstDay
    boundaries(2) = lastDay
    For i = 1 To rows.Count
        CheckCancel i
        item = rows(i)
        boundaries(i * 2 + 1) = item(0)
        boundaries(i * 2 + 2) = item(1)
    Next i
    SortIntervals boundaries, unused, 1, count
    For i = 1 To count - 1
        CheckCancel i
        segmentStart = boundaries(i)
        segmentEnd = boundaries(i + 1)
        If segmentStart >= segmentEnd Then GoTo NextSegment
        excluded = False
        Set names = VBA.CreateObject(DICTIONARY_PROG_ID)
        names.CompareMode = vbTextCompare
        For j = 1 To rows.Count
            CheckCancel j
            item = rows(j)
            If item(0) < segmentEnd And item(1) > segmentStart Then
                If item(3) Then
                    excluded = True
                    Exit For
                End If
                names(item(2)) = True
            End If
        Next j
        If Not excluded Then
            label = vbNullString
            For Each key In names.Keys
                If VBA.Len(label) > 0 Then label = label & EVENT_NAMES_SEPARATOR
                label = label & VBA.CStr(key)
            Next key
            If names.Count = 0 Then label = PRESENT_EVENT_NAME
            BuildCountedPeriods = BuildCountedPeriods + segmentEnd - segmentStart
            If pendingEnd = segmentStart And pendingLabel = label Then
                pendingEnd = segmentEnd
            Else
                If VBA.Len(pendingLabel) > 0 Then _
                    AppendCountedPeriod periodsText, pendingStart, pendingEnd, pendingLabel
                pendingStart = segmentStart
                pendingEnd = segmentEnd
                pendingLabel = label
            End If
        ElseIf VBA.Len(pendingLabel) > 0 Then
            AppendCountedPeriod periodsText, pendingStart, pendingEnd, pendingLabel
            pendingLabel = vbNullString
        End If
NextSegment:
    Next i
    If VBA.Len(pendingLabel) > 0 Then AppendCountedPeriod periodsText, pendingStart, pendingEnd, pendingLabel
End Function

' ============================================================
' Представление зачтённого периода
' ============================================================
' Для чтения в Excel обе отображаемые даты включаются в период.
' Внутренний исключающий конец уменьшается на день только при форматировании.
Private Sub AppendCountedPeriod(ByRef periodsText As String, ByVal firstDay As Long, _
    ByVal endExclusive As Long, ByVal eventName As String)
    If VBA.Len(periodsText) > 0 Then periodsText = periodsText & PERIODS_SEPARATOR
    periodsText = periodsText & eventName & PERIOD_LABEL_OPEN & _
        VBA.Format$(VBA.CDate(firstDay), FORMAT_DATE) & PERIOD_RANGE_SEPARATOR & _
        VBA.Format$(VBA.CDate(endExclusive - 1), FORMAT_DATE) & PERIOD_LABEL_CLOSE
End Sub

' ============================================================
' Объединение исключаемых интервалов
' ============================================================
' Сортирует границы и объединяет пересечения и смежные интервалы.
' Каждый день вычитается один раз; разрывы остаются в зачтённых днях.
Private Function UnionDays(ByVal rows As Collection) As Long
    Dim a() As Long, b() As Long, item As Variant
    Dim i As Long, leftEdge As Long, rightEdge As Long
    If rows.Count = 0 Then Exit Function
    ReDim a(1 To rows.Count)
    ReDim b(1 To rows.Count)
    For i = 1 To rows.Count
        CheckCancel i
        item = rows(i)
        a(i) = item(0)
        b(i) = item(1)
    Next i
    SortIntervals a, b, 1, rows.Count
    leftEdge = a(1)
    rightEdge = b(1)
    For i = 2 To rows.Count
        CheckCancel i
        If a(i) <= rightEdge Then
            If b(i) > rightEdge Then rightEdge = b(i)
        Else
            UnionDays = UnionDays + rightEdge - leftEdge
            leftEdge = a(i)
            rightEdge = b(i)
        End If
    Next i
    UnionDays = UnionDays + rightEdge - leftEdge
End Function

' ============================================================
' Сортировка интервалов
' ============================================================
' Быстрая сортировка по началу интервала. Массивы начал и концов
' переставляются синхронно, чтобы сохранить пары границ.
Private Sub SortIntervals(ByRef a() As Long, ByRef b() As Long, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Long, temp As Long
    i = low
    j = high
    pivot = a(low + (high - low) \ 2)
    Do While i <= j
        CheckCancel i
        Do While a(i) < pivot
            i = i + 1
            CheckCancel i
        Loop
        Do While a(j) > pivot
            j = j - 1
            CheckCancel j
        Loop
        If i <= j Then
            temp = a(i): a(i) = a(j): a(j) = temp
            temp = b(i): b(i) = b(j): b(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then SortIntervals a, b, low, j
    If i < high Then SortIntervals a, b, i, high
End Sub

' ============================================================
' Сброс фильтров источника РУХ
' ============================================================
' Снимает фильтры именно исходной умной таблицы перед чтением событий.
' Ошибка передаётся обработчику запуска: продолжать с несброшенным фильтром нельзя.
Private Sub ClearSourceFilters(ByVal sourceTable As ListObject)
    Dim sourceFilter As AutoFilter
    On Error GoTo Failed
    Set sourceFilter = sourceTable.AutoFilter
    ' Если объект фильтра отсутствует, сбрасывать в этой таблице нечего.
    If sourceFilter Is Nothing Then Exit Sub
    If sourceFilter.FilterMode Then
        sourceFilter.ShowAllData
    End If
    Exit Sub
Failed:
    Fail MSG_SOURCE_FILTER_FAILED & sourceTable.Name & ": " & VBA.Err.Description
End Sub

' ============================================================
' Поиск источника РУХ
' ============================================================
' Ищет единственную исходную таблицу среди открытых книг.
' Отсутствие источника и несколько совпадений считаются ошибкой.
Private Function FindSource() As ListObject
    Dim wb As Workbook, ws As Worksheet, lo As ListObject
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            If ws.Name = SOURCE_SHEET_NAME Then
                For Each lo In ws.ListObjects
                    If lo.Name = SOURCE_TABLE_NAME Then
                        If Not FindSource Is Nothing Then Fail MSG_MULTIPLE_SOURCES
                        Set FindSource = lo
                    End If
                Next lo
            End If
        Next ws
    Next wb
    If FindSource Is Nothing Then Fail MSG_OPEN_SOURCE & SOURCE_SHEET_NAME & MSG_TABLE & SOURCE_TABLE_NAME & "'."
End Function

' ============================================================
' Поиск обязательной таблицы инструмента
' ============================================================
' Ищет таблицу по точным именам листа и таблицы в ThisWorkbook.
Private Function RequireTable(ByVal sheetName As String, ByVal tableName As String, _
    Optional ByVal workbook As Workbook = Nothing) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If workbook Is Nothing Then Set workbook = ThisWorkbook
    For Each ws In workbook.Worksheets
        If ws.Name = sheetName Then
            For Each lo In ws.ListObjects
                If lo.Name = tableName Then
                    Set RequireTable = lo
                    Exit Function
                End If
            Next lo
        End If
    Next ws
    Fail MSG_TABLE_NOT_FOUND & tableName & MSG_ON_WORKSHEET & sheetName & MSG_CHECK_TABLE_CONFIGURATION
End Function

' ============================================================
' Проверка обязательных колонок
' ============================================================
' Возвращает индекс колонки по имени; отсутствие колонки останавливает операцию.
Private Function ColumnIndex(ByVal lo As ListObject, ByVal header As String) As Long
    Dim col As ListColumn
    For Each col In lo.ListColumns
        If col.Name = header Then
            ColumnIndex = col.Index
            Exit Function
        End If
    Next col
    Fail MSG_TABLE_TEXT & lo.Name & MSG_IS_MISSING_COLUMN & header & "'."
End Function

' ============================================================
' Проверка структуры результата
' ============================================================
' Проверяет состав и число колонок; порядок определяется по именам.
Private Sub ValidateOutput(ByVal lo As ListObject)
    Dim headers As Variant, i As Long, columnNumber As Long
    headers = VBA.Array(TARGET_COL_NAME, TARGET_COL_TAX_ID, TARGET_COL_START, _
        TARGET_COL_TRIPS, TARGET_COL_DAYS, TARGET_COL_PERIODS)
    If lo.ListColumns.Count <> UBound(headers) - LBound(headers) + 1 Then Fail MSG_RESULT_COLUMN_COUNT
    For i = LBound(headers) To UBound(headers)
        columnNumber = ColumnIndex(lo, VBA.CStr(headers(i)))
    Next i
    If lo.ShowTotals Then Fail MSG_DISABLE_TOTALS
End Sub

' ============================================================
' Преобразование обязательного текста
' ============================================================
' Удаляет пробелы по краям и нормализует неразрывные пробелы.
' Пустое значение и ошибка ячейки не подменяются значением по умолчанию.
Private Function RequiredText(ByVal value As Variant, ByVal context As String) As String
    If VBA.IsError(value) Or VBA.IsNull(value) Then Fail context & MSG_CELL_ERROR
    RequiredText = VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))
    If VBA.Len(RequiredText) = 0 Then Fail context & MSG_REQUIRED_VALUE
End Function

' ============================================================
' Чтение флага командировок
' ============================================================
' Принимает только явно перечисленные варианты логического значения.
Private Function ReadFlag(ByVal value As Variant, ByVal context As String) As Boolean
    ' В новом формате пустой флаг явно означает Ні: командировки вычитаются.
    If Not VBA.IsError(value) And Not VBA.IsNull(value) Then
        If VBA.Len(VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))) = 0 Then
            ReadFlag = False
            Exit Function
        End If
    End If
    Select Case VBA.LCase$(RequiredText(value, context & MSG_SUBTRACT_BUSINESS_TRIPS))
        Case "так", "да", "true", "1": ReadFlag = True
        Case "ні", "нет", "false", "0": ReadFlag = False
        Case Else: Fail context & MSG_INVALID_TRIP_FLAG
    End Select
End Function

' ============================================================
' Чтение календарной даты
' ============================================================
' Текстовые даты строго dd.mm.yyyy; числовые учитывают систему дат книги.
' Время отбрасывается: инструмент считает календарные дни.
Private Function ReadDay(ByVal value As Variant, ByVal date1904 As Boolean, ByVal context As String) As Long
    Dim text As String, parsed As Date, serial As Double
    On Error GoTo InvalidDate
    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsEmpty(value) Then GoTo InvalidDate
    If VBA.VarType(value) = vbString Then
        text = VBA.Trim$(VBA.CStr(value))
        If Not text Like DATE_TEXT_PATTERN Then GoTo InvalidDate
        parsed = VBA.DateSerial(VBA.CInt(VBA.Right$(text, 4)), VBA.CInt(VBA.Mid$(text, 4, 2)), VBA.CInt(VBA.Left$(text, 2)))
        If VBA.Format$(parsed, FORMAT_DATE) <> text Then GoTo InvalidDate
        serial = VBA.CDbl(parsed)
    Else
        If VBA.VarType(value) = vbBoolean Or Not VBA.IsNumeric(value) Then GoTo InvalidDate
        serial = VBA.Int(VBA.CDbl(value))
        If date1904 Then serial = serial + DATE_1904_OFFSET
    End If
    If serial < MIN_DATE_SERIAL Or serial > MAX_DATE_SERIAL Then GoTo InvalidDate
    ReadDay = VBA.CLng(serial)
    Exit Function
InvalidDate:
    Fail context & MSG_INVALID_DATE
End Function

' ============================================================
' Обработка ошибок
' ============================================================
' Передаёт конкретное описание ошибки обработчику публичного метода,
' который показывает MsgBox и завершает текущую операцию.
Private Sub Fail(ByVal message As String)
    VBA.Err.Raise OPERATION_ERROR, ERROR_SOURCE, message
End Sub

' ============================================================
' Проверка ожидаемого результата
' ============================================================
' Вспомогательная проверка для TestMovementDays с описанием расхождения.
Private Sub AssertDays(ByVal actual As Long, ByVal expected As Long, ByVal context As String)
    If actual <> expected Then Fail context & MSG_EXPECTED & expected & MSG_ACTUAL & actual
End Sub

' ============================================================
' Логирование
' ============================================================

' Лог создаётся рядом с книгой: <имя книги>_logs.txt.
' ClearLog очищает предыдущий лог при запуске операции.
' ERROR и WARNING включаются ENABLE_LOGGING; DEBUG требует обоих флагов.
' При отключённом ENABLE_LOGGING файловые операции не компилируются.
Private Sub ClearLog()
#If ENABLE_LOGGING Then
    Dim fileNumber As Integer

    fileNumber = VBA.FreeFile
    Open GetLogFilePath() For Output As #fileNumber
    Close #fileNumber
#End If
End Sub

Private Sub LogError(ByVal message As String)
#If ENABLE_LOGGING Then
    WriteLog LOG_ERROR_PREFIX & message
#End If
End Sub

Private Sub LogWarning(ByVal message As String)
#If ENABLE_LOGGING Then
    WriteLog LOG_WARNING_PREFIX & message
#End If
End Sub

Private Sub LogDebug(ByVal message As String)
#If ENABLE_LOGGING Then
#If ENABLE_DEBUG_LOGGING Then
    WriteLog LOG_DEBUG_PREFIX & message
#End If
#End If
End Sub

Private Sub WriteLog(ByVal message As String)
#If ENABLE_LOGGING Then
    Dim fileNumber As Integer

    Dim fileOpened As Boolean
    Dim errorNumber As Long, errorText As String

    On Error GoTo Failed
    fileNumber = VBA.FreeFile
    Open GetLogFilePath() For Append As #fileNumber
    fileOpened = True
    Print #fileNumber, VBA.Format$(VBA.Now, FORMAT_LOG_TIMESTAMP) & " | " & message
    Close #fileNumber
    Exit Sub
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    ' Ошибка закрытия не должна подменять исходную ошибку записи.
    On Error Resume Next
    If fileOpened Then Close #fileNumber
    On Error GoTo 0
    VBA.Err.Raise errorNumber, LOG_ERROR_SOURCE, errorText
#End If
End Sub

Private Function GetLogFilePath() As String
    Dim workbookName As String
    Dim baseName As String

    Dim dotPosition As Long

    If VBA.Len(ThisWorkbook.Path) = 0 Then
        Fail MSG_SAVE_BEFORE_LOGGING
    End If
    workbookName = ThisWorkbook.Name
    dotPosition = VBA.InStrRev(workbookName, ".")
    If dotPosition > 0 Then
        baseName = VBA.Left$(workbookName, dotPosition - 1)
    Else
        baseName = workbookName
    End If
    GetLogFilePath = ThisWorkbook.Path & Application.PathSeparator & baseName & LOG_FILE_SUFFIX
End Function

' ============================================================
' Логирование при завершении с ошибкой или отменой
' ============================================================

' Отдельный обработчик защищает основной Cleanup от повторной ошибки логгера.
' Недоступность лога сообщается явно; исходная ошибка остаётся у вызывающего кода.
Private Sub LogFailure(ByVal errorNumber As Long, ByVal errorText As String)
#If ENABLE_LOGGING Then
    On Error GoTo Failed
    If errorNumber = CANCEL_ERROR Then
        LogWarning MSG_OPERATION_CANCELLED_BY_THE_USER
    Else
        LogError MSG_NUMBER & errorNumber & MSG_DESCRIPTION & errorText
    End If
    Exit Sub
Failed:
    VBA.MsgBox MSG_LOG_WRITE_FAILED & VBA.Err.Description, vbExclamation
#End If
End Sub
