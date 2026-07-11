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
Private m_ExportModes As Collection

Private Const EXPORT_MODE_DEFAULT As String = "Default"
Private Const EXPORT_MODE_REWRITE_LAST As String = "Rewrite Last"

Private Const DEFAULT_ALF_REL_PATH As String = "modes\PrsnlEvntBuilder\АЛФ.xlsx"
Private Const ALF_SHEET_NAME As String = "ОС"
Private Const ALF_RANGE_START As String = "A1"
Private Const ALF_RANGE_END_COLUMN As String = "J"
Private Const DEFAULT_INSTITUTIONS_REL_PATH As String = "modes\PrsnlEvntBuilder\Установи.xlsx"
Private Const INSTITUTIONS_SHEET_NAME As String = "Лікувальні Заклади"
Private Const INSTITUTIONS_RANGE_START As String = "A3"
Private Const INSTITUTIONS_RANGE_END_COLUMN As String = "E"
Private Const DEFAULT_RANKS_REL_PATH As String = "modes\PrsnlEvntBuilder\Переліки.xlsx"
Private Const RANKS_SHEET_NAME As String = "Звання"
Private Const RANKS_RANGE_START As String = "A1"
Private Const RANKS_RANGE_END_COLUMN As String = "E"
Private Const DEFAULT_POSITIONS_REL_PATH As String = "modes\PrsnlEvntBuilder\Посади.xlsm"
Private Const POSITIONS_SHEET_NAME As String = "Посади"
Private Const POSITIONS_RANGE_START As String = "A4"
Private Const POSITIONS_RANGE_END_COLUMN As String = "E"
Private Const DEFAULT_ORDER_MAP_REL_PATH As String = "modes\PrsnlEvntBuilder\Мапа наказів.xlsx"
Private Const ORDER_MAP_SHEET_NAME As String = "Накази"
Private Const ORDER_MAP_2026_RANGE_START As String = "D2"
Private Const ORDER_MAP_2026_RANGE_END_COLUMN As String = "E"
Private Const ORDER_MAP_2025_RANGE_START As String = "A2"
Private Const ORDER_MAP_2025_RANGE_END_COLUMN As String = "B"
Private Const EXCEL_MAX_ROW As Long = 1048576

Private Const ALF_KEY_HEADER As String = "ІПН"
Private Const ALF_FIO_KEY_HEADER As String = "ПІБ"
Private Const ALF_GENITIVE_HEADER As String = "Родовий"
Private Const ALF_INITIALS_GENITIVE_HEADER As String = "ПІП (Родовий)"
Private Const INSTITUTIONS_KEY_HEADER As String = "Позначення"
Private Const INSTITUTIONS_GENITIVE_HEADER As String = "Родовий"
Private Const INSTITUTIONS_ACCUSATIVE_HEADER As String = "Знахідний"
Private Const RANKS_KEY_HEADER As String = "Звання"
Private Const RANKS_GENITIVE_HEADER As String = "Родовий"
Private Const POSITIONS_KEY_HEADER As String = "Код"
Private Const POSITIONS_GENITIVE_HEADER As String = "Родовий"
Private Const ORDER_DATE_COLUMN_NAME As String = "Дата наказу"
Private Const ORDER_NO_COLUMN_NAME As String = "Номер наказу"

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
Public Function Initialize(Optional ByVal configTable As obj_ConfigTable = Nothing) As Boolean
    m_IsDisposed = False
    m_OrderNo = VBA.vbNullString
    m_OrderDate = 0
    m_HasOrderDate = False
    Set m_ExportModes = New Collection
    ' Порядок коллекции определяет цикл multi-toggle кнопки ExportMode.
    m_ExportModes.Add EXPORT_MODE_DEFAULT
    m_ExportModes.Add EXPORT_MODE_REWRITE_LAST

    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    m_OrderNo = VBA.vbNullString
    m_OrderDate = 0
    m_HasOrderDate = False
    Set m_ExportModes = Nothing
End Sub

Public Property Get ExportModes() As Collection
    Dim result As Collection
    Dim modeValue As Variant

    If m_IsDisposed Then Exit Property
    If m_ExportModes Is Nothing Then Exit Property

    Set result = New Collection
    For Each modeValue In m_ExportModes
        result.Add VBA.CStr(modeValue)
    Next modeValue
    Set ExportModes = result
End Property

Public Function TryGetExportModeName(ByVal zeroBasedIndex As Long, ByRef outModeName As String) As Boolean
    outModeName = VBA.vbNullString
    If m_IsDisposed Then Exit Function
    If m_ExportModes Is Nothing Then Exit Function
    If zeroBasedIndex < 0 Or zeroBasedIndex >= m_ExportModes.Count Then Exit Function

    outModeName = VBA.CStr(m_ExportModes.Item(zeroBasedIndex + 1))
    TryGetExportModeName = (VBA.Len(outModeName) > 0)
End Function

' Статический provider общих данных PrsnlEvntBuilder.
' Здесь остаются только стабильные справочники, не завязанные на профиль:
' АЛФ, Установи, Переліки, Посади, Мапа наказів.
' Динамические источники вроде ежедневной ШПС держит obj_PEB_ExptrDataPrvdr.
Public Function SetOrderNo(ByVal orderNo As Variant) As Boolean
    If m_IsDisposed Then Exit Function

    ' Номер приказа задается один раз перед export/render. Если дату удалось
    ' найти в отдельной "Мапі наказів", она становится базовой датой для всех
    ' сокращенных дат текущего экспорта.
    m_OrderNo = orderNo
    m_OrderDate = 0
    m_HasOrderDate = False

    If VBA.Len(private_NormalizeOrderNumberToken(orderNo)) > 0 Then
        m_HasOrderDate = TryResolveOrderDateByNumber(orderNo, m_OrderDate)
    End If

    SetOrderNo = True
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

Public Function TryResolveOrderDateByNumber( _
    ByVal orderNo As Variant, _
    ByRef outOrderDate As Date _
) As Boolean
    Dim orderNoToken As String
    Dim orderMapPath As String

    If m_IsDisposed Then Exit Function
    orderNoToken = private_NormalizeOrderNumberToken(orderNo)
    If VBA.Len(orderNoToken) = 0 Then Exit Function
    If Not private_TryResolveOrderMapWorkbookPath(orderMapPath) Then Exit Function

    ' "Мапа наказів" разложена горизонтальными блоками по годам.
    ' Сначала проверяем актуальный блок 2026, затем старый блок 2025. Важно:
    ' lookup идет через отдельный workbook, поэтому WORD/DailyScope больше не
    ' зависят от служебной таблицы внутри Movement workbook.
    If private_TryLookupWorkbookDate( _
        orderMapPath, _
        private_BuildAdoRangeRef( _
            ORDER_MAP_SHEET_NAME, _
            ORDER_MAP_2026_RANGE_START, _
            ORDER_MAP_2026_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ORDER_NO_COLUMN_NAME, _
        ORDER_DATE_COLUMN_NAME, _
        orderNoToken, _
        "Мапа наказів / 2026", _
        outOrderDate) Then
        TryResolveOrderDateByNumber = True
        Exit Function
    End If

    TryResolveOrderDateByNumber = private_TryLookupWorkbookDate( _
        orderMapPath, _
        private_BuildAdoRangeRef( _
            ORDER_MAP_SHEET_NAME, _
            ORDER_MAP_2025_RANGE_START, _
            ORDER_MAP_2025_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ORDER_NO_COLUMN_NAME, _
        ORDER_DATE_COLUMN_NAME, _
        orderNoToken, _
        "Мапа наказів / 2025", _
        outOrderDate)
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
        DEFAULT_ALF_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_KEY_HEADER, _
        ALF_GENITIVE_HEADER, _
        ipnText, _
        "АЛФ", _
        outFioGenitive)
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
            INSTITUTIONS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        INSTITUTIONS_KEY_HEADER, _
        INSTITUTIONS_GENITIVE_HEADER, _
        hospitalShortText, _
        "Установи", _
        outHospitalGenitive)
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
            INSTITUTIONS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        INSTITUTIONS_KEY_HEADER, _
        INSTITUTIONS_ACCUSATIVE_HEADER, _
        hospitalShortText, _
        "Установи", _
        outHospitalAccusative)
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
        DEFAULT_RANKS_REL_PATH, _
        private_BuildAdoRangeRef( _
            RANKS_SHEET_NAME, _
            RANKS_RANGE_START, _
            RANKS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        RANKS_KEY_HEADER, _
        RANKS_GENITIVE_HEADER, _
        rankText, _
        "Переліки / Звання", _
        outRankGenitive)
End Function

Public Function TryResolvePositionGenitive( _
    ByVal positionText As String, _
    ByRef outPositionGenitive As String _
) As Boolean
    If m_IsDisposed Then Exit Function
    positionText = private_NormalizeLookupKey(positionText)
    outPositionGenitive = VBA.vbNullString
    If VBA.Len(positionText) = 0 Then
        TryResolvePositionGenitive = True
        Exit Function
    End If

    TryResolvePositionGenitive = private_TryLookupWorkbookValue( _
        DEFAULT_POSITIONS_REL_PATH, _
        private_BuildAdoRangeRef( _
            POSITIONS_SHEET_NAME, _
            POSITIONS_RANGE_START, _
            POSITIONS_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        POSITIONS_KEY_HEADER, _
        POSITIONS_GENITIVE_HEADER, _
        positionText, _
        "Посади", _
        outPositionGenitive)
End Function

' //
' // Internal
' //
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
        DEFAULT_ALF_REL_PATH, _
        private_BuildAdoRangeRef( _
            ALF_SHEET_NAME, _
            ALF_RANGE_START, _
            ALF_RANGE_END_COLUMN & VBA.CStr(EXCEL_MAX_ROW)), _
        ALF_FIO_KEY_HEADER, _
        valueHeader, _
        fioText, _
        "АЛФ / " & sourceLabelSuffix, _
        outValue)
End Function

Private Function private_TryLookupWorkbookValue( _
    ByVal workbookPath As String, _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal valueHeader As String, _
    ByVal lookupKey As String, _
    ByVal sourceLabel As String, _
    ByRef outValue As String _
) As Boolean
    Dim resolvedPath As String
    Dim conn As Object
    Dim rs As Object
    Dim sql As String

    outValue = VBA.vbNullString
    resolvedPath = private_ResolveWorkbookPath(workbookPath)
    If VBA.Len(resolvedPath) = 0 Or VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: declension source workbook was not found: " & workbookPath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    On Error GoTo LookupFail
    Set conn = VBA.CreateObject("ADODB.Connection")
    conn.Open private_BuildAdoConnectionString(resolvedPath)

    sql = "SELECT TOP 1 " & private_QuoteSqlIdentifier(valueHeader) & _
        " FROM " & tableRef & _
        " WHERE LCase(Trim(CStr(" & private_QuoteSqlIdentifier(keyHeader) & "))) = " & private_AdoSqlTextLiteral(lookupKey)

    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1

    If rs.EOF Then
        VBA.MsgBox "PrototypeNew: declension row was not found in " & sourceLabel & " for key: " & lookupKey, VBA.vbExclamation, "PrototypeNew / WORD export"
        GoTo CleanupDone
    End If

    outValue = VBA.Trim$(private_RecordsetFieldText(rs.Fields(0).Value))
    private_TryLookupWorkbookValue = True

CleanupDone:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Exit Function

LookupFail:
    VBA.MsgBox "PrototypeNew: failed to query declension source " & sourceLabel & "." & _
        VBA.vbCrLf & "Workbook: " & workbookPath & _
        VBA.vbCrLf & "Range: " & tableRef & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, "PrototypeNew / WORD export"
    Resume CleanupDone
End Function

Private Function private_TryResolveOrderMapWorkbookPath(ByRef outPath As String) As Boolean
    outPath = private_ResolveWorkbookPath(DEFAULT_ORDER_MAP_REL_PATH)
    If VBA.Len(outPath) > 0 And VBA.Len(VBA.Dir$(outPath)) > 0 Then
        private_TryResolveOrderMapWorkbookPath = True
        Exit Function
    End If

    VBA.MsgBox "PrototypeNew: order map workbook was not found." & _
        VBA.vbCrLf & "Expected: " & DEFAULT_ORDER_MAP_REL_PATH, VBA.vbExclamation, "PrototypeNew / WORD export"
End Function

Private Function private_TryLookupWorkbookDate( _
    ByVal workbookPath As String, _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal valueHeader As String, _
    ByVal lookupKey As String, _
    ByVal sourceLabel As String, _
    ByRef outDate As Date _
) As Boolean
    Dim resolvedPath As String
    Dim conn As Object
    Dim rs As Object
    Dim sql As String

    outDate = 0
    resolvedPath = private_ResolveWorkbookPath(workbookPath)
    If VBA.Len(resolvedPath) = 0 Or VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: order map workbook was not found." & _
            VBA.vbCrLf & "Expected: " & workbookPath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    On Error GoTo LookupFail
    Set conn = VBA.CreateObject("ADODB.Connection")
    conn.Open private_BuildAdoConnectionString(resolvedPath)

    ' Номер приказа сравниваем как текстовый token. Диапазон начинается со
    ' строки заголовков блока года, поэтому ADO видит поля "Дата наказу" и
    ' "Номер наказу" так же, как остальные справочники склонений.
    sql = "SELECT TOP 1 " & private_QuoteSqlIdentifier(valueHeader) & _
        " FROM " & tableRef & _
        " WHERE LCase(Trim(CStr(" & private_QuoteSqlIdentifier(keyHeader) & "))) = " & private_AdoSqlTextLiteral(lookupKey)

    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1

    If rs.EOF Then GoTo CleanupDone

    private_TryLookupWorkbookDate = ex_Helpers.fn_TryResolveDateWithContext( _
        rs.Fields(0).Value, _
        VBA.DateSerial(1900, 1, 1), _
        outDate)

CleanupDone:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Exit Function

LookupFail:
    VBA.MsgBox "PrototypeNew: failed to query order map source " & sourceLabel & "." & _
        VBA.vbCrLf & "Workbook: " & workbookPath & _
        VBA.vbCrLf & "Range: " & tableRef & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, "PrototypeNew / WORD export"
    Resume CleanupDone
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
