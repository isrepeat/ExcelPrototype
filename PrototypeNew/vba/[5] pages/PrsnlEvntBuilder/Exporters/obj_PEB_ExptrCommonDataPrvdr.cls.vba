VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExptrCommonDataPrvdr"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False
#Const LOGGING_VERBOSE_ENABLED = False

Private m_IsDisposed As Boolean
Private m_OrderNo As Variant
Private m_OrderDate As Date
Private m_HasOrderDate As Boolean
Private m_OrderYear As Long
Private m_WorkbookConnections As Object
' Все новые обращения к ШПО/справочникам описываются одинаковым query object.
' Engine читает закрытый источник через ADO, а открытый — из живого Worksheet,
' включая несохраненные изменения пользователя.
Private m_QueryEngine As obj_ExtWorkbookQueryEngine

Private Const DEFAULT_SHPO_REL_PATH As String = "modes\PrsnlEvntBuilder\ШПО.xlsx"
Private Const ALF_SHEET_NAME As String = "АЛФ"
' АЛФ сохраняет обычную однострочную шапку и узкий диапазон A:J.
Private Const ALF_RANGE_START As String = "A1"
Private Const ALF_RANGE_END_COLUMN As String = "J"
Private Const DEFAULT_INSTITUTIONS_REL_PATH As String = "modes\PrsnlEvntBuilder\Установи.xlsx"
Private Const INSTITUTIONS_SHEET_NAME As String = "Лікувальні Заклади"
Private Const INSTITUTIONS_RANGE_START As String = "A3"
Private Const INSTITUTIONS_RANGE_END_COLUMN As String = "F"
Private Const INSTITUTIONS_RANGE_END_ROW As Long = 10000
Private Const RANKS_SHEET_NAME As String = "Звання"
Private Const RANKS_RANGE_START As String = "A1"
Private Const RANKS_RANGE_END_COLUMN As String = "E"
Private Const POSITIONS_SHEET_NAME As String = "Посади"
Private Const POSITIONS_RANGE_START As String = "A1"
Private Const POSITIONS_RANGE_END_COLUMN As String = "E"
Private Const DEFAULT_ORDER_MAP_REL_PATH As String = "modes\PrsnlEvntBuilder\Накази.xlsx"
Private Const ORDER_MAP_SHEET_NAME As String = "Накази"
Private Const ORDER_MAP_2026_RANGE_START As String = "D2"
Private Const ORDER_MAP_2026_RANGE_END_COLUMN As String = "E"
Private Const ORDER_MAP_2025_RANGE_START As String = "A2"
Private Const ORDER_MAP_2025_RANGE_END_COLUMN As String = "B"
' Справочники фактически помещаются в 12 000 строк; это ниже общего лимита
' 20 000 и уменьшает область каждого точечного ACE-запроса.
Private Const EXCEL_MAX_ROW As Long = 12000

Private Const ALF_KEY_HEADER As String = "ІПН"
Private Const ALF_FIO_KEY_HEADER As String = "ПІБ"
Private Const ALF_GENITIVE_HEADER As String = "Родовий"
Private Const ALF_ACCUSATIVE_HEADER As String = "Знахідний"
Private Const ALF_DATIVE_HEADER As String = "Давальний"
Private Const ALF_INITIALS_GENITIVE_HEADER As String = "ПІП (Родовий)"
Private Const INSTITUTIONS_KEY_HEADER As String = "Позначення"
Private Const INSTITUTIONS_NAME_HEADER As String = "Назва"
Private Const INSTITUTIONS_GENITIVE_HEADER As String = "Родовий"
Private Const INSTITUTIONS_ACCUSATIVE_HEADER As String = "Знахідний"
Private Const INSTITUTIONS_DATIVE_HEADER As String = "Давальний"
Private Const RANKS_KEY_HEADER As String = "Звання"
Private Const RANKS_GENITIVE_HEADER As String = "Родовий"
Private Const RANKS_ACCUSATIVE_HEADER As String = "Знахідний"
Private Const RANKS_DATIVE_HEADER As String = "Давальний"
Private Const POSITIONS_KEY_HEADER As String = "Код"
Private Const POSITIONS_GENITIVE_HEADER As String = "Родовий"
Private Const POSITIONS_DATIVE_HEADER As String = "Давальний"
Private Const POSITIONS_DEFAULT_HEADER As String = "Назва"
Private Const OS_SHEET_NAME As String = "ОС"
' В обновлённой ШПО лист ОС использует ту же трёхстрочную шапку. Основной ІПН
' остаётся в пределах прежнего узкого диапазона A:AB.
Private Const OS_RANGE_START As String = "A3"
Private Const OS_RANGE_END_COLUMN As String = "AB"
Private Const OS_IPN_HEADER As String = "ІПН"
Private Const OS_RANK_HEADER As String = "Військове звання фактично"
Private Const SPECIAL_POSITION_PREFIX_ROZP As String = "A1A"
Private Const SPECIAL_POSITION_PREFIX_SPIS As String = "A1B"
Private Const SPECIAL_POSITION_CODE_ROZP As String = "РОЗП"
Private Const SPECIAL_POSITION_CODE_SPIS As String = "СПИС"
Private Const SPECIAL_POSITION_NAME_ROZP_OFFICER As String = "який перебуває у розпорядженні командира військової частини А3369"
Private Const SPECIAL_POSITION_NAME_ROZP_OTHER As String = "який перебуває у розпорядженні командира військової частини А7383"
Private Const SPECIAL_POSITION_UNIT_ROZP_OFFICER As String = "А3369"
Private Const SPECIAL_POSITION_UNIT_ROZP_OTHER As String = "А7383"
Private Const ORDER_DATE_COLUMN_NAME As String = "Дата наказу"
Private Const ORDER_NO_COLUMN_NAME As String = "Номер наказу"

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize(Optional ByVal configTable As obj_ConfigTable = Nothing) As Boolean
    m_IsDisposed = False
    m_OrderNo = VBA.vbNullString
    m_OrderDate = 0
    m_HasOrderDate = False
    m_OrderYear = 0
    Set m_WorkbookConnections = VBA.CreateObject("Scripting.Dictionary")
    m_WorkbookConnections.CompareMode = 1
    Set m_QueryEngine = New obj_ExtWorkbookQueryEngine
    If Not m_QueryEngine.Initialize Then Exit Function
    Initialize = True
End Function

Public Function TryResolveFioDative( _
    ByVal ipnText As String, _
    ByRef outFioDative As String _
) As Boolean
    ' АЛФ связывает человека по ИПН, поэтому одинаковые ФИО не создают
    ' неоднозначность при получении полного имени в дательном падеже.
    If m_IsDisposed Then Exit Function
    ipnText = private_NormalizeLookupKey(ipnText)
    outFioDative = VBA.vbNullString
    If VBA.Len(ipnText) = 0 Then
        TryResolveFioDative = True
        Exit Function
    End If

    TryResolveFioDative = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(ALF_SHEET_NAME, ALF_RANGE_START, ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_KEY_HEADER, ALF_DATIVE_HEADER, ipnText, "ШПО / АЛФ", outFioDative)
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    m_OrderNo = VBA.vbNullString
    m_OrderDate = 0
    m_HasOrderDate = False
    m_OrderYear = 0
    On Error Resume Next
    If Not m_QueryEngine Is Nothing Then m_QueryEngine.Dispose
    Set m_QueryEngine = Nothing
    private_CloseWorkbookConnections
    Set m_WorkbookConnections = Nothing
    On Error GoTo 0
End Sub

' Статический provider общих данных PrsnlEvntBuilder.
' Здесь остаются только стабильные справочники, не завязанные на профиль:
' ШПО (АЛФ/Посади/Звання), Установи, Накази.
' Динамические источники вроде ежедневной ШПС держит obj_PEB_ExptrCfgDataPrvdr.
Public Function SetOrderNo(ByVal orderNo As Variant) As Boolean
    Dim normalizedOrderNo As String

    If m_IsDisposed Then Exit Function

    normalizedOrderNo = private_NormalizeOrderNumberToken(orderNo)
    ' После controller-resolve пара уже содержит правильный, возможно не
    ' текущий год. Повторные внутренние обращения с тем же номером не должны
    ' снова переключать lookup на системный год.
    If m_HasOrderDate And VBA.StrComp( _
        normalizedOrderNo, private_NormalizeOrderNumberToken(m_OrderNo), _
        VBA.vbTextCompare) = 0 Then
        SetOrderNo = True
        Exit Function
    End If

    ' Номер приказа задается один раз перед export/render. Если дату удалось
    ' найти в отдельной "Мапі наказів", она становится базовой датой для всех
    ' сокращенных дат текущего экспорта.
    m_OrderNo = orderNo
    m_OrderDate = 0
    m_HasOrderDate = False

    m_OrderYear = VBA.Year(VBA.Date)
    If VBA.Len(normalizedOrderNo) > 0 Then _
        m_HasOrderDate = TryResolveOrderDateByNumber(orderNo, m_OrderDate, False, , m_OrderYear)

    SetOrderNo = True
End Function

' Преобразует оба допустимых UI-ввода в единственный внутренний контракт:
' канонический номер приказа и его полную дату.
Public Function TryResolveOrderReference( _
    ByVal numberOrDate As Variant, _
    ByVal explicitYear As Variant, _
    ByRef outOrderNo As String, _
    ByRef outOrderDate As Date _
) As Boolean
    Dim inputText As String
    Dim orderYear As Long
    Dim fullDate As Date

    outOrderNo = VBA.vbNullString
    outOrderDate = 0
    ClearOrderReference
    inputText = VBA.Trim$(VBA.CStr(numberOrDate))
    If VBA.Len(inputText) = 0 Then
        VBA.MsgBox "Вкажіть номер або повну дату наказу.", VBA.vbExclamation, "PrsnlEventBuilder / Наказ"
        Exit Function
    End If

    If private_LooksLikeDateInput(inputText) Then
        If Not private_TryParseFullOrderDate(inputText, fullDate) Then
            VBA.MsgBox "Дата наказу має бути повною та коректною у форматі ДД.ММ.РРРР.", _
                VBA.vbExclamation, "PrsnlEventBuilder / Наказ"
            Exit Function
        End If
        orderYear = VBA.Year(fullDate)
        If Not private_TryResolveOrderNoByDate(fullDate, orderYear, outOrderNo) Then Exit Function
        outOrderDate = fullDate
    Else
        If Not private_TryResolveExplicitOrderYear(explicitYear, orderYear) Then Exit Function
        outOrderNo = private_NormalizeOrderNumberToken(inputText)
        If VBA.Len(outOrderNo) = 0 Then Exit Function
        If Not TryResolveOrderDateByNumber(outOrderNo, outOrderDate, True, , orderYear) Then Exit Function
        If outOrderDate = 0 Then
            VBA.MsgBox "Наказ № " & outOrderNo & " за " & VBA.CStr(orderYear) & _
                " рік не знайдено у довіднику «Накази».", VBA.vbExclamation, "PrsnlEventBuilder / Наказ"
            Exit Function
        End If
    End If

    m_OrderNo = outOrderNo
    m_OrderDate = outOrderDate
    m_OrderYear = orderYear
    m_HasOrderDate = True
    TryResolveOrderReference = True
End Function

Public Sub ClearOrderReference()
    m_OrderNo = VBA.vbNullString
    m_OrderDate = 0
    m_OrderYear = 0
    m_HasOrderDate = False
End Sub

Public Function SetResolvedOrderPair( _
    ByVal orderNo As Variant, _
    ByVal orderDate As Date _
) As Boolean
    Dim normalizedOrderNo As String

    If m_IsDisposed Then Exit Function
    normalizedOrderNo = private_NormalizeOrderNumberToken(orderNo)
    If VBA.Len(normalizedOrderNo) = 0 Or orderDate = 0 Then Exit Function
    m_OrderNo = normalizedOrderNo
    m_OrderDate = VBA.DateValue(orderDate)
    m_OrderYear = VBA.Year(orderDate)
    m_HasOrderDate = True
    SetResolvedOrderPair = True
End Function

Public Property Get OrderNo() As Variant
    OrderNo = m_OrderNo
End Property

Public Property Get HasOrderDate() As Boolean
    HasOrderDate = m_HasOrderDate
End Property

Public Property Get OrderDate() As Date
    OrderDate = m_OrderDate
End Property

Public Property Get OrderYear() As Long
    OrderYear = m_OrderYear
End Property

Public Function FormatVacationTicketNoForExport( _
    ByVal rawTicketNo As Variant, _
    ByVal orderNo As Variant _
) As String
    Dim formattedValue As String

    If TryFormatVacationTicketNoForExport(rawTicketNo, orderNo, formattedValue) Then
        FormatVacationTicketNoForExport = formattedValue
    End If
End Function

Public Function TryFormatVacationTicketNoForExport( _
    ByVal rawTicketNo As Variant, _
    ByVal orderNo As Variant, _
    ByRef outTicketNo As String _
) As Boolean
    Dim ticketNoText As String
    Dim orderNoText As String
    Dim rx As Object
    Dim parts As Variant
    Dim partIndex As Long
    Dim comparableTicketNo As String

    outTicketNo = VBA.vbNullString
    ticketNoText = VBA.Trim$(VBA.CStr(rawTicketNo))
    orderNoText = private_NormalizeOrderNumberToken(orderNo)
    If VBA.Len(ticketNoText) = 0 Then
        TryFormatVacationTicketNoForExport = True
        Exit Function
    End If

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.Pattern = "^\d+$"

    parts = VBA.Split(ticketNoText, "/")
    If UBound(parts) > 2 Then GoTo InvalidFormat
    For partIndex = LBound(parts) To UBound(parts)
        If Not rx.Test(VBA.CStr(parts(partIndex))) Then GoTo InvalidFormat
    Next partIndex

    Select Case UBound(parts) - LBound(parts) + 1
        Case 1
            comparableTicketNo = ticketNoText
            Do While VBA.Len(comparableTicketNo) > 1 And VBA.Left$(comparableTicketNo, 1) = "0"
                comparableTicketNo = VBA.Mid$(comparableTicketNo, 2)
            Loop
            If VBA.Len(comparableTicketNo) > 4 Or _
               (VBA.Len(comparableTicketNo) = 4 And VBA.StrComp(comparableTicketNo, "3000", VBA.vbBinaryCompare) > 0) Then
                outTicketNo = "3/71/" & ticketNoText
            Else
                If VBA.Len(orderNoText) = 0 Or Not rx.Test(orderNoText) Then
                    VBA.MsgBox _
                        "PrototypeNew: отпускной билет '" & ticketNoText & _
                        "' требует числовой номер текущего приказа.", _
                        VBA.vbExclamation, "PrototypeNew / WORD export"
                    Exit Function
                End If
                outTicketNo = VBA.CStr(VBA.Year(VBA.Date)) & "/" & orderNoText & "/" & ticketNoText
            End If

        Case 2
            outTicketNo = VBA.CStr(VBA.Year(VBA.Date)) & "/" & ticketNoText

        Case 3
            outTicketNo = ticketNoText

        Case Else
            GoTo InvalidFormat
    End Select

    TryFormatVacationTicketNoForExport = True
    Exit Function

InvalidFormat:
    VBA.MsgBox _
        "PrototypeNew: некорректный номер отпускного билета '" & ticketNoText & "'." & VBA.vbCrLf & _
        "Допустимые формы: <номер>, <приказ>/<билет> или <год>/<приказ>/<билет>." & VBA.vbCrLf & _
        "Каждая часть должна содержать только цифры.", _
        VBA.vbExclamation, "PrototypeNew / WORD export"
End Function

Public Function TryCalculateFoodSupportDate( _
    ByVal hasExplicitEventDate As Boolean, _
    ByVal explicitEventDate As Date, _
    ByRef outFoodSupportDate As Date _
) As Boolean
    ' Единое правило для снятия и зачисления на продовольственное обеспечение:
    ' минимально допустима дата OrderDate + 1; более поздняя явная дата события
    ' имеет приоритет. Если дата приказа неизвестна, сохраняем прежнее поведение
    ' Movement и используем явную дату, когда она доступна.
    outFoodSupportDate = 0
    If m_HasOrderDate Then
        outFoodSupportDate = VBA.DateAdd("d", 1, m_OrderDate)
        If hasExplicitEventDate Then
            If explicitEventDate > outFoodSupportDate Then outFoodSupportDate = explicitEventDate
        End If
        TryCalculateFoodSupportDate = True
        Exit Function
    End If

    If hasExplicitEventDate Then
        outFoodSupportDate = explicitEventDate
        TryCalculateFoodSupportDate = True
    End If
End Function

Public Function TryResolveOrderDateByNumber( _
    ByVal orderNo As Variant, _
    ByRef outOrderDate As Date, _
    Optional ByVal allowMissing As Boolean = False, _
    Optional ByRef outFound As Boolean = False, _
    Optional ByVal orderYear As Long = 0 _
) As Boolean
    Dim orderNoToken As String
    Dim orderMapPath As String
    Dim orderMapRangeStart As String
    Dim orderMapRangeEndColumn As String
    Dim currentYear As Long

    If m_IsDisposed Then Exit Function
    outFound = False
    orderNoToken = private_NormalizeOrderNumberToken(orderNo)
    If VBA.Len(orderNoToken) = 0 Then Exit Function
    If Not private_TryResolveOrderMapWorkbookPath(orderMapPath) Then Exit Function

    ' «Накази» разложены горизонтальными блоками по годам. Номер ищем только
    ' в блоке текущего системного года: одинаковый номер прошлого года не
    ' должен ошибочно считаться текущим приказом.
    currentYear = orderYear
    If currentYear = 0 Then currentYear = VBA.Year(VBA.Date)
    If Not private_TryGetOrderMapRange( _
        currentYear, orderMapRangeStart, orderMapRangeEndColumn) Then Exit Function

    TryResolveOrderDateByNumber = private_TryLookupWorkbookDate( _
        orderMapPath, _
        private_BuildAdoRangeRef( _
            ORDER_MAP_SHEET_NAME, _
            orderMapRangeStart, _
            orderMapRangeEndColumn & VBA.CStr(EXCEL_MAX_ROW)), _
        ORDER_NO_COLUMN_NAME, _
        ORDER_DATE_COLUMN_NAME, _
        orderNoToken, _
        "Накази / " & VBA.CStr(currentYear), _
        outOrderDate, _
        allowMissing, _
        outFound)
End Function

Public Function TryResolveFioGenitive( _
    ByVal ipnText As String, _
    ByRef outFioGenitive As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    ipnText = private_NormalizeLookupKey(ipnText)
    outFioGenitive = VBA.vbNullString
    If VBA.Len(ipnText) = 0 Then
        TryResolveFioGenitive = True
        Exit Function
    End If

    TryResolveFioGenitive = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_KEY_HEADER, _
        ALF_GENITIVE_HEADER, _
        ipnText, _
        "ШПО / АЛФ", _
        outFioGenitive)
End Function

Public Function TryResolveFioAccusative( _
    ByVal ipnText As String, _
    ByRef outFioAccusative As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    ipnText = private_NormalizeLookupKey(ipnText)
    outFioAccusative = VBA.vbNullString
    If VBA.Len(ipnText) = 0 Then
        TryResolveFioAccusative = True
        Exit Function
    End If

    TryResolveFioAccusative = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_KEY_HEADER, _
        ALF_ACCUSATIVE_HEADER, _
        ipnText, _
        "ШПО / АЛФ", _
        outFioAccusative)
End Function

Public Function TryResolveFioInitialsGenitive( _
    ByVal ipnText As String, _
    ByRef outFioInitialsGenitive As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    ipnText = private_NormalizeLookupKey(ipnText)
    outFioInitialsGenitive = VBA.vbNullString
    If VBA.Len(ipnText) = 0 Then
        TryResolveFioInitialsGenitive = True
        Exit Function
    End If

    TryResolveFioInitialsGenitive = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_KEY_HEADER, _
        ALF_INITIALS_GENITIVE_HEADER, _
        ipnText, _
        "ШПО / АЛФ", _
        outFioInitialsGenitive)
End Function

Public Function TryResolveFioInitialsGenitiveOptional( _
    ByVal ipnText As String, _
    ByRef outFioInitialsGenitive As String, _
    ByRef outFound As Boolean _
) As Boolean
    If m_IsDisposed Then Exit Function
    ipnText = private_NormalizeLookupKey(ipnText)
    outFioInitialsGenitive = VBA.vbNullString
    outFound = False
    If VBA.Len(ipnText) = 0 Then
        TryResolveFioInitialsGenitiveOptional = True
        Exit Function
    End If

    TryResolveFioInitialsGenitiveOptional = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_KEY_HEADER, _
        ALF_INITIALS_GENITIVE_HEADER, _
        ipnText, _
        "ШПО / АЛФ", _
        outFioInitialsGenitive, _
        allowMissingRow:=True, _
        outFound:=outFound, _
        requireUniqueMatch:=False, _
        logMissingRow:=False)
End Function

Public Function TryResolveFioFormsOptional( _
    ByVal ipnText As String, _
    ByRef outFioGenitive As String, _
    ByRef outFioAccusative As String, _
    ByRef outFioInitialsGenitive As String, _
    ByRef outFound As Boolean _
) As Boolean
    ipnText = private_NormalizeLookupKey(ipnText)
    outFioGenitive = VBA.vbNullString
    outFioAccusative = VBA.vbNullString
    outFioInitialsGenitive = VBA.vbNullString
    outFound = False
    If m_IsDisposed Then Exit Function
    If VBA.Len(ipnText) = 0 Then
        TryResolveFioFormsOptional = True
        Exit Function
    End If

    ' Для preview отсутствие человека в АЛФ является предупреждением, а не
    ' ошибкой pipeline. Сначала выполняем один неблокирующий запрос. Остальные
    ' формы читаем только при наличии строки с таким ИПН.
    If Not private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_KEY_HEADER, _
        ALF_GENITIVE_HEADER, _
        ipnText, _
        "ШПО / АЛФ", _
        outFioGenitive, _
        allowMissingRow:=True, _
        outFound:=outFound, _
        requireUniqueMatch:=False, _
        logMissingRow:=False) Then Exit Function
    If Not outFound Then
        TryResolveFioFormsOptional = True
        Exit Function
    End If

    If Not TryResolveFioAccusative( _
        ipnText, outFioAccusative) Then Exit Function
    If Not TryResolveFioInitialsGenitive( _
        ipnText, outFioInitialsGenitive) Then Exit Function
    TryResolveFioFormsOptional = True
End Function

Public Function TryResolveFioGenitiveByName( _
    ByVal fioText As String, _
    ByRef outFioGenitive As String _
) As Boolean
    TryResolveFioGenitiveByName = private_TryResolveAlfByFio( _
        fioText, _
        ALF_GENITIVE_HEADER, _
        "full genitive name", _
        outFioGenitive)
End Function

Public Function TryResolveFioFormsByNameOptional( _
    ByVal fioText As String, _
    ByRef outFioGenitive As String, _
    ByRef outFioAccusative As String, _
    ByRef outFioInitialsGenitive As String, _
    ByRef outFound As Boolean _
) As Boolean
    fioText = private_NormalizeLookupKey(fioText)
    outFioGenitive = VBA.vbNullString
    outFioAccusative = VBA.vbNullString
    outFioInitialsGenitive = VBA.vbNullString
    outFound = False
    If m_IsDisposed Then Exit Function
    If VBA.Len(fioText) = 0 Then
        TryResolveFioFormsByNameOptional = True
        Exit Function
    End If

    If Not private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(ALF_SHEET_NAME, ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_FIO_KEY_HEADER, ALF_GENITIVE_HEADER, fioText, _
        "ШПО / АЛФ / ПІБ", outFioGenitive, True, outFound, _
        False, False) Then Exit Function
    If Not outFound Then
        TryResolveFioFormsByNameOptional = True
        Exit Function
    End If
    If Not private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(ALF_SHEET_NAME, ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_FIO_KEY_HEADER, ALF_ACCUSATIVE_HEADER, fioText, _
        "ШПО / АЛФ / ПІБ", outFioAccusative) Then Exit Function
    If Not private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(ALF_SHEET_NAME, ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_FIO_KEY_HEADER, ALF_INITIALS_GENITIVE_HEADER, fioText, _
        "ШПО / АЛФ / ПІБ", outFioInitialsGenitive) Then Exit Function
    TryResolveFioFormsByNameOptional = True
End Function

Public Function TryResolveFioDefaultByGenitive( _
    ByVal fioGenitive As String, _
    ByRef outFioDefault As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    fioGenitive = private_NormalizeLookupKey(fioGenitive)
    outFioDefault = VBA.vbNullString
    If VBA.Len(fioGenitive) = 0 Then
        TryResolveFioDefaultByGenitive = True
        Exit Function
    End If

    ' Обратный поиск нужен extractor-у приказов: склонённое ФИО из текста
    ' связывается с несклонённой формой в той же строке АЛФ.
    TryResolveFioDefaultByGenitive = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_GENITIVE_HEADER, _
        ALF_FIO_KEY_HEADER, _
        fioGenitive, _
        "ШПО / АЛФ / ФИО в родовом падеже", _
        outFioDefault, _
        allowMissingRow:=False, _
        requireUniqueMatch:=True)
End Function

Public Function TryFindFioDefaultByGenitive( _
    ByVal fioGenitive As String, _
    ByRef outFound As Boolean, _
    ByRef outFioDefault As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    fioGenitive = private_NormalizeLookupKey(fioGenitive)
    outFound = False
    outFioDefault = VBA.vbNullString
    If VBA.Len(fioGenitive) = 0 Then
        TryFindFioDefaultByGenitive = True
        Exit Function
    End If

    ' Неблокирующий вариант для нормализации: отсутствие человека в АЛФ
    ' является штатным результатом, а не ошибкой всего pipeline.
    TryFindFioDefaultByGenitive = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_GENITIVE_HEADER, _
        ALF_FIO_KEY_HEADER, _
        fioGenitive, _
        "ШПО / АЛФ / ФИО в родовом падеже", _
        outFioDefault, _
        allowMissingRow:=True, _
        outFound:=outFound, _
        requireUniqueMatch:=False)
End Function

Public Function TryFindFioDefaultByDeclinedForm( _
    ByVal fioDeclined As String, _
    ByRef outFound As Boolean, _
    ByRef outFioDefault As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    fioDeclined = private_NormalizeLookupKey(fioDeclined)
    outFound = False
    outFioDefault = VBA.vbNullString
    If VBA.Len(fioDeclined) = 0 Then
        TryFindFioDefaultByDeclinedForm = True
        Exit Function
    End If

    ' В приказах ФИО не обязано быть в родовом падеже: формулировка события
    ' может использовать также дательный или винительный. Ищем точную строку
    ' по всем поддерживаемым формам и из найденной строки возвращаем ПІБ.
    ' Промах отдельной колонки не логируется: ошибкой является только итоговое
    ' отсутствие человека, которое уже обрабатывает вызывающий transformer.
    If Not private_TryFindFioDefaultByDeclensionHeader( _
        fioDeclined, ALF_GENITIVE_HEADER, outFound, outFioDefault) Then Exit Function
    If outFound Then
        TryFindFioDefaultByDeclinedForm = True
        Exit Function
    End If
    If Not private_TryFindFioDefaultByDeclensionHeader( _
        fioDeclined, ALF_DATIVE_HEADER, outFound, outFioDefault) Then Exit Function
    If outFound Then
        TryFindFioDefaultByDeclinedForm = True
        Exit Function
    End If
    If Not private_TryFindFioDefaultByDeclensionHeader( _
        fioDeclined, ALF_ACCUSATIVE_HEADER, outFound, outFioDefault) Then Exit Function

    TryFindFioDefaultByDeclinedForm = True
End Function

Private Function private_TryFindFioDefaultByDeclensionHeader( _
    ByVal fioDeclined As String, _
    ByVal declensionHeader As String, _
    ByRef outFound As Boolean, _
    ByRef outFioDefault As String _
) As Boolean
    ' requireUniqueMatch=False намеренно: старый АЛФ может содержать повторные
    ' склонённые формы. Для нормализации достаточно первой подходящей строки;
    ' строгая проверка уникальности остаётся в публичных resolve-методах.
    private_TryFindFioDefaultByDeclensionHeader = _
        private_TryLookupWorkbookValue( _
            DEFAULT_SHPO_REL_PATH, _
            private_BuildAdoRangeRef( _
                ALF_SHEET_NAME, _
                ALF_RANGE_START, _
                ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
            declensionHeader, _
            ALF_FIO_KEY_HEADER, _
            fioDeclined, _
            "ШПО / АЛФ / " & declensionHeader, _
            outFioDefault, _
            allowMissingRow:=True, _
            outFound:=outFound, _
            requireUniqueMatch:=False, _
            logMissingRow:=False)
End Function

Public Function TryResolveFioInitialsGenitiveByName( _
    ByVal fioText As String, _
    ByRef outFioInitialsGenitive As String _
) As Boolean
    TryResolveFioInitialsGenitiveByName = private_TryResolveAlfByFio( _
        fioText, _
        ALF_INITIALS_GENITIVE_HEADER, _
        "short genitive name", _
        outFioInitialsGenitive)
End Function

Public Function TryResolveHospitalGenitive( _
    ByVal hospitalShortText As String, _
    ByRef outHospitalGenitive As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    hospitalShortText = private_NormalizeLookupKey(hospitalShortText)
    outHospitalGenitive = VBA.vbNullString
    If VBA.Len(hospitalShortText) = 0 Then
        TryResolveHospitalGenitive = True
        Exit Function
    End If

    TryResolveHospitalGenitive = private_TryLookupWorkbookValue( _
        DEFAULT_INSTITUTIONS_REL_PATH, _
        private_BuildAdoRangeRef( _
            INSTITUTIONS_SHEET_NAME, _
            INSTITUTIONS_RANGE_START, _
            INSTITUTIONS_RANGE_END_COLUMN & VBA.CStr(INSTITUTIONS_RANGE_END_ROW)), _
        INSTITUTIONS_KEY_HEADER, _
        INSTITUTIONS_GENITIVE_HEADER, _
        hospitalShortText, _
        "Установи", _
        outHospitalGenitive)
End Function

Public Function TryResolveHospitalName( _
    ByVal hospitalShortText As String, _
    ByRef outHospitalName As String, _
    Optional ByVal allowMissing As Boolean = False, _
    Optional ByRef outFound As Boolean = False _
) As Boolean
    If m_IsDisposed Then Exit Function
    hospitalShortText = private_NormalizeLookupKey(hospitalShortText)
    outHospitalName = VBA.vbNullString
    outFound = False
    If VBA.Len(hospitalShortText) = 0 Then
        TryResolveHospitalName = True
        Exit Function
    End If

    TryResolveHospitalName = private_TryLookupWorkbookValue( _
        DEFAULT_INSTITUTIONS_REL_PATH, _
        private_BuildAdoRangeRef( _
            INSTITUTIONS_SHEET_NAME, _
            INSTITUTIONS_RANGE_START, _
            INSTITUTIONS_RANGE_END_COLUMN & VBA.CStr(INSTITUTIONS_RANGE_END_ROW)), _
        INSTITUTIONS_KEY_HEADER, _
        INSTITUTIONS_NAME_HEADER, _
        hospitalShortText, _
        "Установи", _
        outHospitalName, _
        allowMissingRow:=allowMissing, _
        outFound:=outFound, _
        logMissingRow:=Not allowMissing)
End Function

Public Function TryResolveHospitalAccusative( _
    ByVal hospitalShortText As String, _
    ByRef outHospitalAccusative As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    hospitalShortText = private_NormalizeLookupKey(hospitalShortText)
    outHospitalAccusative = VBA.vbNullString
    If VBA.Len(hospitalShortText) = 0 Then
        TryResolveHospitalAccusative = True
        Exit Function
    End If

    TryResolveHospitalAccusative = private_TryLookupWorkbookValue( _
        DEFAULT_INSTITUTIONS_REL_PATH, _
        private_BuildAdoRangeRef( _
            INSTITUTIONS_SHEET_NAME, _
            INSTITUTIONS_RANGE_START, _
            INSTITUTIONS_RANGE_END_COLUMN & VBA.CStr(INSTITUTIONS_RANGE_END_ROW)), _
        INSTITUTIONS_KEY_HEADER, _
        INSTITUTIONS_ACCUSATIVE_HEADER, _
        hospitalShortText, _
        "Установи", _
        outHospitalAccusative)
End Function

Public Function TryResolveHospitalDative( _
    ByVal hospitalShortText As String, _
    ByRef outHospitalDative As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    hospitalShortText = private_NormalizeLookupKey(hospitalShortText)
    outHospitalDative = VBA.vbNullString
    If VBA.Len(hospitalShortText) = 0 Then
        TryResolveHospitalDative = True
        Exit Function
    End If

    TryResolveHospitalDative = private_TryLookupWorkbookValue( _
        DEFAULT_INSTITUTIONS_REL_PATH, _
        private_BuildAdoRangeRef( _
            INSTITUTIONS_SHEET_NAME, _
            INSTITUTIONS_RANGE_START, _
            INSTITUTIONS_RANGE_END_COLUMN & VBA.CStr(INSTITUTIONS_RANGE_END_ROW)), _
        INSTITUTIONS_KEY_HEADER, _
        INSTITUTIONS_DATIVE_HEADER, _
        hospitalShortText, _
        "Установи", _
        outHospitalDative)
End Function

Public Function TryResolveRankGenitive( _
    ByVal rankText As String, _
    ByRef outRankGenitive As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    rankText = private_NormalizeLookupKey(rankText)
    outRankGenitive = VBA.vbNullString
    If VBA.Len(rankText) = 0 Then
        TryResolveRankGenitive = True
        Exit Function
    End If

    TryResolveRankGenitive = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            RANKS_SHEET_NAME, _
            RANKS_RANGE_START, _
            RANKS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        RANKS_KEY_HEADER, _
        RANKS_GENITIVE_HEADER, _
        rankText, _
        "ШПО / Звання", _
        outRankGenitive)
End Function

Public Function TryResolveRankGenitiveOptional( _
    ByVal rankText As String, _
    ByRef outRankGenitive As String, _
    ByRef outFound As Boolean _
) As Boolean
    If m_IsDisposed Then Exit Function
    rankText = private_NormalizeLookupKey(rankText)
    outRankGenitive = VBA.vbNullString
    outFound = False
    If VBA.Len(rankText) = 0 Then
        TryResolveRankGenitiveOptional = True
        Exit Function
    End If

    ' Отсутствующая строка звания допускает fallback исходной формой в preview.
    ' Ошибки открытия книги и выполнения запроса по-прежнему возвращают False.
    TryResolveRankGenitiveOptional = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            RANKS_SHEET_NAME, _
            RANKS_RANGE_START, _
            RANKS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        RANKS_KEY_HEADER, _
        RANKS_GENITIVE_HEADER, _
        rankText, _
        "ШПО / Звання", _
        outRankGenitive, _
        allowMissingRow:=True, _
        outFound:=outFound, _
        requireUniqueMatch:=False, _
        logMissingRow:=False)
End Function

Public Function TryFindRankDefaultByDeclinedForm( _
    ByVal rankDeclined As String, _
    ByRef outFound As Boolean, _
    ByRef outRankDefault As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    rankDeclined = private_NormalizeLookupKey(rankDeclined)
    outFound = False
    outRankDefault = VBA.vbNullString
    If VBA.Len(rankDeclined) = 0 Then
        TryFindRankDefaultByDeclinedForm = True
        Exit Function
    End If

    If Not private_TryFindRankDefaultByDeclensionHeader( _
        rankDeclined, RANKS_GENITIVE_HEADER, outFound, _
        outRankDefault) Then Exit Function
    If outFound Then
        TryFindRankDefaultByDeclinedForm = True
        Exit Function
    End If
    If Not private_TryFindRankDefaultByDeclensionHeader( _
        rankDeclined, RANKS_ACCUSATIVE_HEADER, outFound, _
        outRankDefault) Then Exit Function
    If outFound Then
        TryFindRankDefaultByDeclinedForm = True
        Exit Function
    End If
    If Not private_TryFindRankDefaultByDeclensionHeader( _
        rankDeclined, RANKS_DATIVE_HEADER, outFound, _
        outRankDefault) Then Exit Function

    TryFindRankDefaultByDeclinedForm = True
End Function

Private Function private_TryFindRankDefaultByDeclensionHeader( _
    ByVal rankDeclined As String, _
    ByVal declensionHeader As String, _
    ByRef outFound As Boolean, _
    ByRef outRankDefault As String _
) As Boolean
    private_TryFindRankDefaultByDeclensionHeader = _
        private_TryLookupWorkbookValue( _
            DEFAULT_SHPO_REL_PATH, _
            private_BuildAdoRangeRef( _
                RANKS_SHEET_NAME, _
                RANKS_RANGE_START, _
                RANKS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
            declensionHeader, _
            RANKS_KEY_HEADER, _
            rankDeclined, _
            "ШПО / Звання / " & declensionHeader, _
            outRankDefault, _
            allowMissingRow:=True, _
            outFound:=outFound, _
            requireUniqueMatch:=False, _
            logMissingRow:=False)
End Function

Public Function TryResolveRankDative( _
    ByVal rankText As String, _
    ByRef outRankDative As String _
) As Boolean
    ' Используем ту же таблицу званий, что и для Genitive, но колонку
    ' "Давальний". Пустое исходное звание является допустимым значением.
    If m_IsDisposed Then Exit Function
    rankText = private_NormalizeLookupKey(rankText)
    outRankDative = VBA.vbNullString
    If VBA.Len(rankText) = 0 Then
        TryResolveRankDative = True
        Exit Function
    End If

    TryResolveRankDative = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(RANKS_SHEET_NAME, RANKS_RANGE_START, RANKS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        RANKS_KEY_HEADER, RANKS_DATIVE_HEADER, rankText, "ШПО / Звання", outRankDative)
End Function

Public Function TryResolvePositionGenitive( _
    ByVal positionText As String, _
    ByRef outPositionGenitive As String, _
    Optional ByVal rankText As String = "" _
) As Boolean
    Dim isResolved As Boolean

    If m_IsDisposed Then Exit Function
    positionText = private_NormalizePositionCodeForLookup(positionText)
    outPositionGenitive = VBA.vbNullString
    If VBA.Len(positionText) = 0 Then
        TryResolvePositionGenitive = True
        Exit Function
    End If

    isResolved = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            POSITIONS_SHEET_NAME, _
            POSITIONS_RANGE_START, _
            POSITIONS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        POSITIONS_KEY_HEADER, _
        POSITIONS_GENITIVE_HEADER, _
        positionText, _
        "ШПО / Посади", _
        outPositionGenitive)
    If Not isResolved Then Exit Function

    private_EnsureRozpUnitSuffix positionText, rankText, outPositionGenitive
    TryResolvePositionGenitive = True
End Function

Public Function TryResolvePositionDative( _
    ByVal positionText As String, _
    ByRef outPositionDative As String _
) As Boolean
    ' Должность ищется по стабильному коду, а не по отображаемому названию.
    If m_IsDisposed Then Exit Function
    positionText = private_NormalizePositionCodeForLookup(positionText)
    outPositionDative = VBA.vbNullString
    If VBA.Len(positionText) = 0 Then
        TryResolvePositionDative = True
        Exit Function
    End If

    TryResolvePositionDative = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(POSITIONS_SHEET_NAME, POSITIONS_RANGE_START, POSITIONS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        POSITIONS_KEY_HEADER, POSITIONS_DATIVE_HEADER, positionText, "ШПО / Посади", outPositionDative)
End Function

Public Function TryResolvePositionGenitiveOptional( _
    ByVal positionText As String, _
    ByRef outPositionGenitive As String, _
    ByRef outFound As Boolean, _
    Optional ByVal rankText As String = "" _
) As Boolean
    Dim isResolved As Boolean

    ' WORD preview may use the current source position when a TVO position
    ' code has not yet been added to the declension dictionary.
    If m_IsDisposed Then Exit Function
    positionText = private_NormalizePositionCodeForLookup(positionText)
    outPositionGenitive = VBA.vbNullString
    outFound = False
    If VBA.Len(positionText) = 0 Then
        TryResolvePositionGenitiveOptional = True
        Exit Function
    End If

    isResolved = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(POSITIONS_SHEET_NAME, POSITIONS_RANGE_START, POSITIONS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        POSITIONS_KEY_HEADER, POSITIONS_GENITIVE_HEADER, positionText, "ШПО / Посади", outPositionGenitive, True, outFound)
    If Not isResolved Then Exit Function

    If outFound Then private_EnsureRozpUnitSuffix positionText, rankText, outPositionGenitive
    TryResolvePositionGenitiveOptional = True
End Function

Public Function TryResolvePositionDativeOptional( _
    ByVal positionText As String, _
    ByRef outPositionDative As String, _
    ByRef outFound As Boolean _
) As Boolean
    If m_IsDisposed Then Exit Function
    positionText = private_NormalizePositionCodeForLookup(positionText)
    outPositionDative = VBA.vbNullString
    outFound = False
    If VBA.Len(positionText) = 0 Then
        TryResolvePositionDativeOptional = True
        Exit Function
    End If

    TryResolvePositionDativeOptional = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(POSITIONS_SHEET_NAME, POSITIONS_RANGE_START, POSITIONS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        POSITIONS_KEY_HEADER, POSITIONS_DATIVE_HEADER, positionText, "ШПО / Посади", outPositionDative, True, outFound)
End Function

Public Function TryResolvePositionDefault( _
    ByVal positionCodeText As String, _
    ByRef outPositionText As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    positionCodeText = private_NormalizePositionCodeForLookup(positionCodeText)
    outPositionText = VBA.vbNullString
    If VBA.Len(positionCodeText) = 0 Then Exit Function

    TryResolvePositionDefault = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(POSITIONS_SHEET_NAME, POSITIONS_RANGE_START, POSITIONS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        POSITIONS_KEY_HEADER, POSITIONS_DEFAULT_HEADER, positionCodeText, "ШПО / Посади", outPositionText)
End Function

' Возвращает обычную, родительную и дательную формы должности одним SQL.
' Обычное название сохраняется для fallback и диагностического preview.
Public Function TryResolvePositionFormsOptional( _
    ByVal positionCodeText As String, _
    ByRef outPositionDefault As String, _
    ByRef outPositionGenitive As String, _
    ByRef outPositionDative As String, _
    ByRef outFound As Boolean, _
    Optional ByVal rankText As String = "" _
) As Boolean
    If m_IsDisposed Then Exit Function

    positionCodeText = private_NormalizePositionCodeForLookup(positionCodeText)
    outPositionDefault = VBA.vbNullString
    outPositionGenitive = VBA.vbNullString
    outPositionDative = VBA.vbNullString
    outFound = False
    If VBA.Len(positionCodeText) = 0 Then
        TryResolvePositionFormsOptional = True
        Exit Function
    End If

    TryResolvePositionFormsOptional = private_TryLookupWorkbookThreeValues( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(POSITIONS_SHEET_NAME, POSITIONS_RANGE_START, POSITIONS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        POSITIONS_KEY_HEADER, POSITIONS_DEFAULT_HEADER, POSITIONS_GENITIVE_HEADER, POSITIONS_DATIVE_HEADER, _
        positionCodeText, "ШПО / Посади", _
        outPositionDefault, outPositionGenitive, outPositionDative, outFound)
    If TryResolvePositionFormsOptional And outFound Then
        private_EnsureRozpUnitSuffix positionCodeText, rankText, outPositionGenitive
    End If
End Function

Public Function TryResolveRankByIpn( _
    ByVal ipnText As String, _
    ByRef outRankText As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    ipnText = private_NormalizeLookupKey(ipnText)
    outRankText = VBA.vbNullString
    If VBA.Len(ipnText) = 0 Then Exit Function

    TryResolveRankByIpn = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef(OS_SHEET_NAME, OS_RANGE_START, OS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        OS_IPN_HEADER, OS_RANK_HEADER, ipnText, "ШПО / ОС", outRankText)
End Function

Public Function TryResolveSpecialPositionMapping( _
    ByVal sourcePositionCodeText As String, _
    ByVal sourceRankText As String, _
    ByRef outTargetPositionCodeText As String, _
    ByRef outTargetPositionNameText As String _
) As Boolean
    Dim normalizedCodeText As String

    If m_IsDisposed Then Exit Function
    outTargetPositionCodeText = VBA.vbNullString
    outTargetPositionNameText = VBA.vbNullString
    normalizedCodeText = private_NormalizePositionCodeForLookup(sourcePositionCodeText)

    Select Case normalizedCodeText
        Case SPECIAL_POSITION_CODE_ROZP
            outTargetPositionCodeText = SPECIAL_POSITION_CODE_ROZP
            If private_IsOfficerRank(sourceRankText) Then
                outTargetPositionNameText = SPECIAL_POSITION_NAME_ROZP_OFFICER
            Else
                outTargetPositionNameText = SPECIAL_POSITION_NAME_ROZP_OTHER
            End If
            TryResolveSpecialPositionMapping = True
        Case SPECIAL_POSITION_CODE_SPIS
            outTargetPositionCodeText = SPECIAL_POSITION_CODE_SPIS
            If Not TryResolvePositionDefault( _
                SPECIAL_POSITION_CODE_SPIS, _
                outTargetPositionNameText) Then Exit Function
            TryResolveSpecialPositionMapping = True
    End Select
End Function

' //
' // Internal
' //
Private Sub private_EnsureRozpUnitSuffix( _
    ByVal normalizedPositionCode As String, _
    ByVal rankText As String, _
    ByRef positionText As String _
)
    Dim unitText As String

    If normalizedPositionCode <> SPECIAL_POSITION_CODE_ROZP Then Exit Sub
    If VBA.Len(VBA.Trim$(rankText)) = 0 Then Exit Sub

    If private_IsOfficerRank(rankText) Then
        unitText = SPECIAL_POSITION_UNIT_ROZP_OFFICER
    Else
        unitText = SPECIAL_POSITION_UNIT_ROZP_OTHER
    End If

    ' Справочник может уже содержать номер части. Сначала удаляем любой
    ' известный вариант, затем ставим номер, соответствующий званию.
    positionText = VBA.Replace( _
        positionText, SPECIAL_POSITION_UNIT_ROZP_OFFICER, _
        VBA.vbNullString, 1, -1, VBA.vbTextCompare)
    positionText = VBA.Replace( _
        positionText, SPECIAL_POSITION_UNIT_ROZP_OTHER, _
        VBA.vbNullString, 1, -1, VBA.vbTextCompare)
    positionText = VBA.Trim$(positionText)
    If VBA.Len(positionText) > 0 Then
        positionText = positionText & " " & unitText
    End If
End Sub

Private Function private_NormalizePositionCodeForLookup(ByVal positionText As String) As String
    Dim normalizedCode As String
    Dim normalizedSpecialCode As String

    normalizedCode = VBA.UCase$(private_NormalizeLookupKey(positionText))
    If normalizedCode = SPECIAL_POSITION_CODE_ROZP Or normalizedCode = SPECIAL_POSITION_CODE_SPIS Then
        private_NormalizePositionCodeForLookup = normalizedCode
        Exit Function
    End If
    ' Коды ШПС могут содержать визуально одинаковые кириллические А/В.
    normalizedSpecialCode = VBA.Replace(normalizedCode, " ", VBA.vbNullString)
    normalizedSpecialCode = VBA.Replace(normalizedSpecialCode, "А", "A")
    normalizedSpecialCode = VBA.Replace(normalizedSpecialCode, "В", "B")
    If VBA.Left$(normalizedSpecialCode, VBA.Len(SPECIAL_POSITION_PREFIX_ROZP)) = SPECIAL_POSITION_PREFIX_ROZP Then
        private_NormalizePositionCodeForLookup = SPECIAL_POSITION_CODE_ROZP
        Exit Function
    End If
    If VBA.Left$(normalizedSpecialCode, VBA.Len(SPECIAL_POSITION_PREFIX_SPIS)) = SPECIAL_POSITION_PREFIX_SPIS Then
        private_NormalizePositionCodeForLookup = SPECIAL_POSITION_CODE_SPIS
        Exit Function
    End If

    private_NormalizePositionCodeForLookup = normalizedCode
End Function

Private Function private_IsOfficerRank(ByVal rankText As String) As Boolean
    Dim normalizedRank As String

    normalizedRank = private_NormalizeLookupKey(rankText)
    ' В используемом наборе данных максимальное звание — полковник.
    Select Case normalizedRank
        Case "молодший лейтенант", "лейтенант", "старший лейтенант", _
             "капітан", "майор", "підполковник", "полковник"
            private_IsOfficerRank = True
    End Select
End Function

Private Function private_TryResolveAlfByFio( _
    ByVal fioText As String, _
    ByVal valueHeader As String, _
    ByVal sourceLabelSuffix As String, _
    ByRef outValue As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    fioText = private_NormalizeLookupKey(fioText)
    outValue = VBA.vbNullString
    If VBA.Len(fioText) = 0 Or private_IsSelfReportText(fioText) Then
        private_TryResolveAlfByFio = True
        Exit Function
    End If

    private_TryResolveAlfByFio = private_TryLookupWorkbookValue( _
        DEFAULT_SHPO_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_FIO_KEY_HEADER, _
        valueHeader, _
        fioText, _
        "ШПО / АЛФ / " & sourceLabelSuffix, _
        outValue, _
        allowMissingRow:=False, _
        requireUniqueMatch:=True)
End Function

Private Function private_TryLookupWorkbookValue( _
    ByVal workbookPath As String, _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal valueHeader As String, _
    ByVal lookupKey As String, _
    ByVal sourceLabel As String, _
    ByRef outValue As String, _
    Optional ByVal allowMissingRow As Boolean = False, _
    Optional ByRef outFound As Boolean = False, _
    Optional ByVal requireUniqueMatch As Boolean = False, _
    Optional ByVal logMissingRow As Boolean = True _
) As Boolean
    Dim resolvedPath As String
    Dim query As obj_ExtWorkbookQuery
    Dim resultTable As obj_TableDynamic
    Dim resultRow As obj_Row

    outValue = VBA.vbNullString
    outFound = False
    resolvedPath = private_ResolveWorkbookPath(workbookPath)
    If VBA.Len(resolvedPath) = 0 Or VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: declension source workbook was not found: " & workbookPath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If m_QueryEngine Is Nothing Then Exit Function
    ' Provider задает только смысл запроса: ключ и возвращаемую колонку.
    ' Конкретный способ доступа к книге инкапсулирован в m_QueryEngine.
    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = resolvedPath
    query.TableRef = tableRef
    If Not query.AddCondition(keyHeader, en_ExtWorkbookQueryOp.ExtQueryOpEquals, lookupKey, True) Then Exit Function
    query.MaxRows = IIf(requireUniqueMatch, 2, 1)
    If Not query.AddSelectColumn(valueHeader) Then Exit Function
    If Not m_QueryEngine.TryExecute(query, resultTable) Then Exit Function

    If resultTable Is Nothing Then
        Exit Function
    ElseIf resultTable.RowCount = 0 Then
        ' Неблокирующие составные lookup могут выполнить несколько пробных
        ' запросов. logMissingRow позволяет не превращать каждый такой пробный
        ' промах в ложную диагностическую ошибку.
        If logMissingRow Then
            ex_Core.fn_Diagnostic_LogError "peb-declension:lookup-miss source='" & sourceLabel & _
                "' key='" & lookupKey & "' workbook='" & resolvedPath & "' range='" & tableRef & "'"
        End If
        If allowMissingRow Then
            private_TryLookupWorkbookValue = True
        Else
            VBA.MsgBox "PrototypeNew: declension row was not found in " & sourceLabel & " for key: " & lookupKey, VBA.vbExclamation, "PrototypeNew / WORD export"
        End If
        Exit Function
    End If

    If requireUniqueMatch And resultTable.RowCount > 1 Then
        VBA.MsgBox "PrototypeNew: more than one row was found in " & sourceLabel & _
            " for FIO: " & lookupKey & VBA.vbCrLf & _
            "The reporter cannot be resolved unambiguously.", _
            VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    Set resultRow = resultTable.Rows.Item(1)
    If resultRow Is Nothing Then Exit Function
    outValue = VBA.Trim$(VBA.CStr(resultRow.GetCellValue(1)))
    outFound = True
    private_TryLookupWorkbookValue = True
End Function

Private Function private_TryLookupWorkbookThreeValues( _
    ByVal workbookPath As String, ByVal tableRef As String, _
    ByVal keyHeader As String, ByVal firstHeader As String, _
    ByVal secondHeader As String, ByVal thirdHeader As String, _
    ByVal lookupKey As String, ByVal sourceLabel As String, _
    ByRef outFirst As String, ByRef outSecond As String, ByRef outThird As String, _
    ByRef outFound As Boolean _
) As Boolean
    Dim resolvedPath As String
    Dim query As obj_ExtWorkbookQuery
    Dim resultTable As obj_TableDynamic
    Dim resultRow As obj_Row

    outFirst = VBA.vbNullString
    outSecond = VBA.vbNullString
    outThird = VBA.vbNullString
    outFound = False
    resolvedPath = private_ResolveWorkbookPath(workbookPath)
    If VBA.Len(resolvedPath) = 0 Or VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: declension source workbook was not found: " & workbookPath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If m_QueryEngine Is Nothing Then Exit Function
    ' Все три формы читаются одним запросом и возвращаются в порядке добавления
    ' колонок независимо от того, открыт справочник или закрыт.
    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = resolvedPath
    query.TableRef = tableRef
    If Not query.AddCondition(keyHeader, en_ExtWorkbookQueryOp.ExtQueryOpEquals, lookupKey, True) Then Exit Function
    query.MaxRows = 1
    If Not query.AddSelectColumn(firstHeader) Then Exit Function
    If Not query.AddSelectColumn(secondHeader) Then Exit Function
    If Not query.AddSelectColumn(thirdHeader) Then Exit Function
    If Not m_QueryEngine.TryExecute(query, resultTable) Then Exit Function
    If resultTable Is Nothing Then
        Exit Function
    ElseIf resultTable.RowCount = 0 Then
        private_TryLookupWorkbookThreeValues = True
        Exit Function
    End If

    Set resultRow = resultTable.Rows.Item(1)
    If resultRow Is Nothing Then Exit Function
    outFirst = VBA.Trim$(VBA.CStr(resultRow.GetCellValue(1)))
    outSecond = VBA.Trim$(VBA.CStr(resultRow.GetCellValue(2)))
    outThird = VBA.Trim$(VBA.CStr(resultRow.GetCellValue(3)))
    outFound = True
    private_TryLookupWorkbookThreeValues = True
End Function

Private Function private_IsDigitsOnly(ByVal valueText As String) As Boolean
    Dim charIndex As Long
    Dim charText As String

    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function
    For charIndex = 1 To VBA.Len(valueText)
        charText = VBA.Mid$(valueText, charIndex, 1)
        If charText < "0" Or charText > "9" Then Exit Function
    Next charIndex
    private_IsDigitsOnly = True
End Function

Private Sub private_LogLookupMissDiagnostics( _
    ByVal conn As Object, _
    ByVal resolvedPath As String, _
    ByVal tableRef As String, _
    ByVal quotedKeyHeader As String, _
    ByVal sourceLabel As String, _
    ByVal lookupKey As String, _
    ByVal lookupSql As String _
)
    Dim probeRs As Object
    Dim probeSql As String
    Dim rowCountText As String
    Dim sampleText As String
    Dim sampleValue As String
    Dim fieldTypeText As String
    Dim numericCountText As String
    Dim quotedCountText As String
    Dim cstrCountText As String
    Dim scalarErrorText As String
    Dim scanMatchText As String
    Dim scanNearText As String
    Dim scanRows As Long
    Dim fileInfoText As String
    Dim openWorkbookText As String
    Dim connectionText As String

    On Error Resume Next
    Set probeRs = VBA.CreateObject("ADODB.Recordset")
    probeSql = "SELECT COUNT(*) AS RowCount FROM " & tableRef
    probeRs.Open probeSql, conn, 0, 1
    If Err.Number = 0 And Not probeRs.EOF Then rowCountText = VBA.CStr(probeRs.Fields(0).Value)
    If Not probeRs Is Nothing Then If probeRs.State <> 0 Then probeRs.Close
    Err.Clear

    probeSql = "SELECT TOP 5 " & quotedKeyHeader & " FROM " & tableRef & _
        " WHERE " & quotedKeyHeader & " Is Not Null"
    probeRs.Open probeSql, conn, 0, 1
    If Err.Number = 0 Then
        Do While Not probeRs.EOF
            sampleValue = private_RecordsetFieldText(probeRs.Fields(0).Value)
            If VBA.Len(sampleText) > 0 Then sampleText = sampleText & " | "
            sampleText = sampleText & sampleValue
            probeRs.MoveNext
        Loop
    End If
    If Not probeRs Is Nothing Then If probeRs.State <> 0 Then probeRs.Close

    If private_IsDigitsOnly(lookupKey) Then
        numericCountText = private_DiagnosticQueryScalar(conn, _
            "SELECT COUNT(*) FROM " & tableRef & " WHERE " & quotedKeyHeader & " = " & lookupKey, scalarErrorText)
        If VBA.Len(scalarErrorText) > 0 Then numericCountText = "error: " & scalarErrorText

        scalarErrorText = VBA.vbNullString
        quotedCountText = private_DiagnosticQueryScalar(conn, _
            "SELECT COUNT(*) FROM " & tableRef & " WHERE " & quotedKeyHeader & " = " & private_AdoSqlTextLiteral(lookupKey), scalarErrorText)
        If VBA.Len(scalarErrorText) > 0 Then quotedCountText = "error: " & scalarErrorText

        scalarErrorText = VBA.vbNullString
        cstrCountText = private_DiagnosticQueryScalar(conn, _
            "SELECT COUNT(*) FROM " & tableRef & " WHERE CStr(" & quotedKeyHeader & ") = " & private_AdoSqlTextLiteral(lookupKey), scalarErrorText)
        If VBA.Len(scalarErrorText) > 0 Then cstrCountText = "error: " & scalarErrorText
    End If

    private_DiagnosticScanLookupKey conn, tableRef, quotedKeyHeader, lookupKey, _
        fieldTypeText, scanRows, scanMatchText, scanNearText
    fileInfoText = private_DiagnosticGetFileInfo(resolvedPath)
    openWorkbookText = private_DiagnosticGetOpenWorkbookInfo(resolvedPath)
    On Error Resume Next
    connectionText = VBA.CStr(conn.ConnectionString)
    Err.Clear
    On Error GoTo 0

    ex_Core.fn_Diagnostic_LogError "peb-declension:lookup-miss source='" & sourceLabel & _
        "' key='" & lookupKey & _
        "' workbook='" & resolvedPath & _
        "' range='" & tableRef & _
        "' rows='" & rowCountText & _
        "' samples='" & VBA.Replace$(sampleText, "'", "''") & _
        "' sql='" & VBA.Replace$(lookupSql, "'", "''") & "'"
    ex_Core.fn_Diagnostic_LogError "peb-declension:lookup-probes key='" & lookupKey & _
        "' fieldType='" & fieldTypeText & _
        "' numericCount='" & numericCountText & _
        "' quotedCount='" & quotedCountText & _
        "' cstrCount='" & cstrCountText & _
        "' scannedRows='" & VBA.CStr(scanRows) & _
        "' clientMatch='" & VBA.Replace$(scanMatchText, "'", "''") & _
        "' nearValues='" & VBA.Replace$(scanNearText, "'", "''") & "'"
    ex_Core.fn_Diagnostic_LogError "peb-declension:lookup-source-state file='" & _
        VBA.Replace$(fileInfoText, "'", "''") & _
        "' openWorkbook='" & VBA.Replace$(openWorkbookText, "'", "''") & _
        "' connection='" & VBA.Replace$(connectionText, "'", "''") & "'"
    Set probeRs = Nothing
    On Error GoTo 0
End Sub

Private Function private_DiagnosticQueryScalar( _
    ByVal conn As Object, _
    ByVal sql As String, _
    ByRef outErrorText As String _
) As String
    Dim rs As Object

    outErrorText = VBA.vbNullString
    On Error GoTo QueryFail
    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1
    If Not rs.EOF Then private_DiagnosticQueryScalar = private_RecordsetFieldText(rs.Fields(0).Value)
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
    Exit Function
QueryFail:
    outErrorText = Err.Description
    Resume Cleanup
End Function

Private Sub private_DiagnosticScanLookupKey( _
    ByVal conn As Object, _
    ByVal tableRef As String, _
    ByVal quotedKeyHeader As String, _
    ByVal lookupKey As String, _
    ByRef outFieldTypeText As String, _
    ByRef outScannedRows As Long, _
    ByRef outMatchText As String, _
    ByRef outNearText As String _
)
    Dim rs As Object
    Dim rawValue As Variant
    Dim valueText As String
    Dim lookupSuffix As String
    Dim nearCount As Long

    On Error GoTo ScanDone
    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open "SELECT " & quotedKeyHeader & " FROM " & tableRef & _
        " WHERE " & quotedKeyHeader & " Is Not Null", conn, 0, 1
    outFieldTypeText = VBA.CStr(rs.Fields(0).Type)
    lookupSuffix = VBA.Right$(lookupKey, 4)

    Do While Not rs.EOF
        outScannedRows = outScannedRows + 1
        rawValue = rs.Fields(0).Value
        valueText = private_RecordsetFieldText(rawValue)
        If VBA.StrComp(valueText, lookupKey, VBA.vbBinaryCompare) = 0 Then
            outMatchText = "row=" & VBA.CStr(outScannedRows) & _
                "; value=" & valueText & _
                "; varType=" & VBA.CStr(VBA.VarType(rawValue))
            Exit Do
        End If
        If nearCount < 5 And VBA.Len(lookupSuffix) > 0 Then
            If VBA.Right$(valueText, VBA.Len(lookupSuffix)) = lookupSuffix Then
                If VBA.Len(outNearText) > 0 Then outNearText = outNearText & " | "
                outNearText = outNearText & valueText
                nearCount = nearCount + 1
            End If
        End If
        rs.MoveNext
    Loop

ScanDone:
    If Err.Number <> 0 Then
        outMatchText = "scan-error: " & Err.Description
        Err.Clear
    End If
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
End Sub

Private Function private_DiagnosticGetFileInfo(ByVal filePath As String) As String
    Dim fso As Object
    Dim fileObj As Object

    On Error GoTo InfoFail
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Set fileObj = fso.GetFile(filePath)
    private_DiagnosticGetFileInfo = "path=" & fileObj.Path & _
        "; size=" & VBA.CStr(fileObj.Size) & _
        "; modified=" & VBA.Format$(fileObj.DateLastModified, "yyyy-mm-dd hh:nn:ss")
    Exit Function
InfoFail:
    private_DiagnosticGetFileInfo = "error: " & Err.Description
End Function

Private Function private_DiagnosticGetOpenWorkbookInfo(ByVal filePath As String) As String
    Dim workbookObj As Object

    On Error GoTo InfoFail
    For Each workbookObj In Application.Workbooks
        If VBA.StrComp(VBA.CStr(workbookObj.FullName), filePath, VBA.vbTextCompare) = 0 Then
            private_DiagnosticGetOpenWorkbookInfo = "open=true; saved=" & _
                VBA.LCase$(VBA.CStr(workbookObj.Saved)) & "; name=" & VBA.CStr(workbookObj.Name)
            Exit Function
        End If
    Next workbookObj
    private_DiagnosticGetOpenWorkbookInfo = "open=false"
    Exit Function
InfoFail:
    private_DiagnosticGetOpenWorkbookInfo = "error: " & Err.Description
End Function

Private Function private_TryGetWorkbookConnection( _
    ByVal resolvedPath As String, _
    ByRef outConnection As Object _
) As Boolean
    Dim connectionKey As String

    ' Каждая внешняя книга получает одно соединение на lifetime provider-а.
    ' Значения не кешируются: resolver-ы по-прежнему выполняют SQL при вызове.
    On Error GoTo ConnectionFail
    If m_WorkbookConnections Is Nothing Then
        Set m_WorkbookConnections = VBA.CreateObject("Scripting.Dictionary")
        m_WorkbookConnections.CompareMode = 1
    End If
    connectionKey = VBA.LCase$(VBA.Trim$(resolvedPath))
    If Not m_WorkbookConnections.Exists(connectionKey) Then
        Set outConnection = VBA.CreateObject("ADODB.Connection")
        outConnection.Open private_BuildAdoConnectionString(resolvedPath)
        m_WorkbookConnections.Add connectionKey, outConnection
    Else
        Set outConnection = m_WorkbookConnections(connectionKey)
        If outConnection.State = 0 Then outConnection.Open private_BuildAdoConnectionString(resolvedPath)
    End If

    private_TryGetWorkbookConnection = True
    Exit Function

ConnectionFail:
    VBA.MsgBox "PrototypeNew: failed to open external data source." & _
        VBA.vbCrLf & "Workbook: " & resolvedPath & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, "PrototypeNew / WORD export"
End Function

Private Sub private_DropWorkbookConnection(ByVal resolvedPath As String)
    Dim connectionKey As String
    Dim conn As Object

    If m_WorkbookConnections Is Nothing Then Exit Sub
    connectionKey = VBA.LCase$(VBA.Trim$(resolvedPath))
    If Not m_WorkbookConnections.Exists(connectionKey) Then Exit Sub
    On Error Resume Next
    Set conn = m_WorkbookConnections(connectionKey)
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    m_WorkbookConnections.Remove connectionKey
    Set conn = Nothing
    On Error GoTo 0
End Sub

Private Sub private_CloseWorkbookConnections()
    Dim connectionKey As Variant
    Dim conn As Object

    If m_WorkbookConnections Is Nothing Then Exit Sub
    On Error Resume Next
    For Each connectionKey In m_WorkbookConnections.Keys
        Set conn = m_WorkbookConnections(connectionKey)
        If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
        Set conn = Nothing
    Next connectionKey
    m_WorkbookConnections.RemoveAll
    On Error GoTo 0
End Sub

Private Function private_TryResolveOrderMapWorkbookPath(ByRef outPath As String) As Boolean
    outPath = private_ResolveWorkbookPath(DEFAULT_ORDER_MAP_REL_PATH)
    If VBA.Len(outPath) > 0 And VBA.Len(VBA.Dir$(outPath)) > 0 Then
        private_TryResolveOrderMapWorkbookPath = True
        Exit Function
    End If

    VBA.MsgBox "PrototypeNew: order map workbook was not found." & _
        VBA.vbCrLf & "Expected: " & DEFAULT_ORDER_MAP_REL_PATH, VBA.vbExclamation, "PrototypeNew / WORD export"
End Function

Private Function private_LooksLikeDateInput(ByVal valueText As String) As Boolean
    private_LooksLikeDateInput = (VBA.InStr(1, valueText, ".", VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, valueText, "/", VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, valueText, "-", VBA.vbBinaryCompare) > 0)
End Function

Private Function private_TryParseFullOrderDate( _
    ByVal valueText As String, _
    ByRef outDate As Date _
) As Boolean
    Dim rx As Object
    Dim matches As Object
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim parsedDate As Date

    outDate = 0
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.Pattern = "^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$"
    If Not rx.Test(VBA.Trim$(valueText)) Then Exit Function
    Set matches = rx.Execute(VBA.Trim$(valueText))
    dayValue = VBA.CLng(matches(0).SubMatches(0))
    monthValue = VBA.CLng(matches(0).SubMatches(1))
    yearValue = VBA.CLng(matches(0).SubMatches(2))
    On Error GoTo InvalidDate
    parsedDate = VBA.DateSerial(yearValue, monthValue, dayValue)
    If VBA.Day(parsedDate) <> dayValue Or VBA.Month(parsedDate) <> monthValue Or _
       VBA.Year(parsedDate) <> yearValue Then Exit Function
    outDate = parsedDate
    private_TryParseFullOrderDate = True
InvalidDate:
End Function

Private Function private_TryResolveExplicitOrderYear( _
    ByVal rawYear As Variant, _
    ByRef outYear As Long _
) As Boolean
    Dim yearText As String

    outYear = 0
    yearText = VBA.Trim$(VBA.CStr(rawYear))
    If VBA.Len(yearText) = 0 Then
        outYear = VBA.Year(VBA.Date)
        private_TryResolveExplicitOrderYear = True
        Exit Function
    End If
    If VBA.Len(yearText) <> 4 Or Not VBA.IsNumeric(yearText) Then
        VBA.MsgBox "Рік наказу має складатися з чотирьох цифр.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Наказ"
        Exit Function
    End If
    outYear = VBA.CLng(yearText)
    private_TryResolveExplicitOrderYear = True
End Function

Private Function private_TryResolveOrderNoByDate( _
    ByVal orderDate As Date, _
    ByVal orderYear As Long, _
    ByRef outOrderNo As String _
) As Boolean
    Dim orderMapPath As String
    Dim rangeStart As String
    Dim rangeEndColumn As String
    Dim query As obj_ExtWorkbookQuery
    Dim resultTable As obj_TableDynamic
    Dim resultRow As obj_Row
    Dim rowIndex As Long
    Dim candidateDate As Date
    Dim candidateNo As String
    Dim matchCount As Long
    Dim matchedNumbers As Object

    outOrderNo = VBA.vbNullString
    Set matchedNumbers = VBA.CreateObject("Scripting.Dictionary")
    matchedNumbers.CompareMode = 1
    If Not private_TryResolveOrderMapWorkbookPath(orderMapPath) Then Exit Function
    If Not private_TryGetOrderMapRange(orderYear, rangeStart, rangeEndColumn) Then Exit Function

    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = orderMapPath
    query.TableRef = private_BuildAdoRangeRef(ORDER_MAP_SHEET_NAME, rangeStart, _
        rangeEndColumn & VBA.CStr(EXCEL_MAX_ROW))
    query.MaxRows = 0
    If Not query.AddSelectColumn(ORDER_NO_COLUMN_NAME) Then Exit Function
    If Not query.AddSelectColumn(ORDER_DATE_COLUMN_NAME) Then Exit Function
    If Not m_QueryEngine.TryExecute(query, resultTable) Then Exit Function
    If resultTable Is Nothing Then Exit Function

    For rowIndex = 1 To resultTable.RowCount
        Set resultRow = resultTable.Rows.Item(rowIndex)
        If Not resultRow Is Nothing Then
            candidateDate = 0
            If ex_Helpers.fn_TryResolveDateWithContext( _
                resultRow.GetCellValue(2), VBA.DateSerial(1900, 1, 1), candidateDate) Then
                If VBA.DateValue(candidateDate) = VBA.DateValue(orderDate) Then
                    candidateNo = private_NormalizeOrderNumberToken(resultRow.GetCellValue(1))
                    If VBA.Len(candidateNo) > 0 Then
                        If Not matchedNumbers.Exists(candidateNo) Then
                            matchedNumbers.Add candidateNo, True
                            matchCount = matchCount + 1
                            outOrderNo = candidateNo
                        End If
                    End If
                End If
            End If
        End If
    Next rowIndex

    If matchCount = 0 Then
        VBA.MsgBox "За датою " & VBA.Format$(orderDate, "dd.mm.yyyy") & _
            " наказ у довіднику «Накази» не знайдено.", VBA.vbExclamation, "PrsnlEventBuilder / Наказ"
        Exit Function
    End If
    If matchCount > 1 Then
        outOrderNo = VBA.vbNullString
        VBA.MsgBox "За датою " & VBA.Format$(orderDate, "dd.mm.yyyy") & _
            " знайдено декілька наказів. Вкажіть номер наказу та, за потреби, рік.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Наказ"
        Exit Function
    End If
    private_TryResolveOrderNoByDate = True
End Function

Private Function private_TryGetOrderMapRange( _
    ByVal orderYear As Long, _
    ByRef outRangeStart As String, _
    ByRef outRangeEndColumn As String _
) As Boolean
    outRangeStart = VBA.vbNullString
    outRangeEndColumn = VBA.vbNullString
    Select Case orderYear
        Case 2025
            outRangeStart = ORDER_MAP_2025_RANGE_START
            outRangeEndColumn = ORDER_MAP_2025_RANGE_END_COLUMN
        Case 2026
            outRangeStart = ORDER_MAP_2026_RANGE_START
            outRangeEndColumn = ORDER_MAP_2026_RANGE_END_COLUMN
        Case Else
            VBA.MsgBox "У довіднику «Накази» не налаштовано таблицю для " & _
                VBA.CStr(orderYear) & " року.", VBA.vbExclamation, "PrototypeNew / Накази"
            Exit Function
    End Select
    private_TryGetOrderMapRange = True
End Function

Private Function private_TryLookupWorkbookDate( _
    ByVal workbookPath As String, _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal valueHeader As String, _
    ByVal lookupKey As String, _
    ByVal sourceLabel As String, _
    ByRef outDate As Date, _
    Optional ByVal allowMissing As Boolean = False, _
    Optional ByRef outFound As Boolean = False _
) As Boolean
    Dim resolvedPath As String
    Dim query As obj_ExtWorkbookQuery
    Dim resultTable As obj_TableDynamic
    Dim resultRow As obj_Row

    outDate = 0
    outFound = False
    resolvedPath = private_ResolveWorkbookPath(workbookPath)
    If VBA.Len(resolvedPath) = 0 Or VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: order map workbook was not found." & _
            VBA.vbCrLf & "Expected: " & workbookPath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    If m_QueryEngine Is Nothing Then Exit Function
    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = resolvedPath
    query.TableRef = tableRef
    If Not query.AddCondition(keyHeader, en_ExtWorkbookQueryOp.ExtQueryOpEquals, lookupKey, True) Then Exit Function
    query.MaxRows = 1
    If Not query.AddSelectColumn(valueHeader) Then Exit Function
    If Not m_QueryEngine.TryExecute(query, resultTable) Then Exit Function
    If resultTable Is Nothing Then Exit Function
    If resultTable.RowCount = 0 Then
        If allowMissing Then private_TryLookupWorkbookDate = True
        Exit Function
    End If
    Set resultRow = resultTable.Rows.Item(1)
    If resultRow Is Nothing Then Exit Function
    private_TryLookupWorkbookDate = ex_Helpers.fn_TryResolveDateWithContext( _
        resultRow.GetCellValue(1), _
        VBA.DateSerial(1900, 1, 1), _
        outDate)
    outFound = private_TryLookupWorkbookDate
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

Private Function private_BuildAdoRangeRef( _
    ByVal sheetName As String, _
    ByVal rangeStart As String, _
    ByVal rangeEnd As String _
) As String
    sheetName = VBA.Replace(VBA.Trim$(sheetName), "]", "]]")
    private_BuildAdoRangeRef = "[" & sheetName & "$" & VBA.Trim$(rangeStart) & ":" & VBA.Trim$(rangeEnd) & "]"
End Function

Private Function private_AdoSqlTextLiteral(ByVal valueText As String) As String
    private_AdoSqlTextLiteral = "'" & VBA.Replace(private_NormalizeLookupKey(valueText), "'", "''") & "'"
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
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    private_NormalizeLookupKey = VBA.LCase$(VBA.Trim$(valueText))
End Function

Private Function private_NormalizeOrderNumberToken(ByVal rawValue As Variant) As String
    Dim valueText As String
    Dim numericValue As Double

    valueText = VBA.Trim$(VBA.CStr(rawValue))
    valueText = VBA.Replace(valueText, " ", VBA.vbNullString)
    If VBA.Len(valueText) = 0 Then Exit Function

    If VBA.IsNumeric(valueText) Then
        numericValue = VBA.CDbl(valueText)
        If numericValue >= 0 Then
            private_NormalizeOrderNumberToken = VBA.CStr(VBA.CLng(numericValue))
            Exit Function
        End If
    End If

    private_NormalizeOrderNumberToken = valueText
End Function

Private Function private_IsSelfReportText(ByVal valueText As String) As Boolean
    valueText = private_NormalizeLookupKey(valueText)
    private_IsSelfReportText = (VBA.StrComp(valueText, "сам", VBA.vbTextCompare) = 0)
End Function
