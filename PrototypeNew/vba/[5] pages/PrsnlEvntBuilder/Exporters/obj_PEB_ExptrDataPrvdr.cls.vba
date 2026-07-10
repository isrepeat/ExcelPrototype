VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExptrDataPrvdr"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False
#Const LOGGING_VERBOSE_ENABLED = False

Private m_IsDisposed As Boolean
Private m_CommonData As obj_PEB_ExptrCommonDataPrvdr
Private m_PersonnelWorkbookPath As String
Private m_PersonnelTableRef As String

Private Const CONFIG_PERSONNEL_FILE_PATH_KEY As String = "Source.Personnel.FilePath"
Private Const CONFIG_PERSONNEL_STATE_RANGE_KEY As String = "Personnel.Sheet[StateMain].SheetName"
Private Const DEFAULT_PERSONNEL_STATE_RANGE_REF As String = "ШПС$A2:Q10000"
Private Const PERSONNEL_TVO_HEADER As String = "ТВО"
Private Const PERSONNEL_POSITION_CODE_HEADER As String = "Код посади"

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
    m_PersonnelTableRef = private_BuildConfiguredAdoRangeRef(DEFAULT_PERSONNEL_STATE_RANGE_REF)

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
    m_PersonnelWorkbookPath = VBA.vbNullString
    m_PersonnelTableRef = VBA.vbNullString
    On Error GoTo 0
End Sub

Public Property Get CommonData() As obj_PEB_ExptrCommonDataPrvdr
    Set CommonData = m_CommonData
End Property

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

    If Not private_TryLookupPersonnelTvoPositionCode(reporterFioText, tvoPositionCode, outIsTvo) Then Exit Function
    If Not outIsTvo Then
        TryResolveReporterTvoPositionGenitive = True
        Exit Function
    End If

    If Not m_CommonData.TryResolvePositionGenitive( _
        tvoPositionCode, _
        outPositionGenitive) Then Exit Function

    TryResolveReporterTvoPositionGenitive = True
End Function

' //
' // Internal
' //
Private Function private_TryLoadConfig(ByVal configTable As obj_ConfigTable) As Boolean
    Dim cfgParserBase As obj_CfgParserBase
    Dim configEntries As Collection
    Dim cfgMap As Object
    Dim personnelRangeRef As String

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

    m_PersonnelWorkbookPath = cfgParserBase.GetOptionalConfigValue( _
        cfgMap, _
        CONFIG_PERSONNEL_FILE_PATH_KEY, _
        VBA.vbNullString)
    personnelRangeRef = cfgParserBase.GetOptionalConfigValue( _
        cfgMap, _
        CONFIG_PERSONNEL_STATE_RANGE_KEY, _
        DEFAULT_PERSONNEL_STATE_RANGE_REF)
    m_PersonnelTableRef = private_BuildConfiguredAdoRangeRef(personnelRangeRef)

CleanExit:
    On Error Resume Next
    If Not cfgParserBase Is Nothing Then cfgParserBase.Dispose
    Set cfgParserBase = Nothing
    On Error GoTo 0
End Function

Private Function private_TryLookupPersonnelTvoPositionCode( _
    ByVal reporterFioText As String, _
    ByRef outPositionCode As String, _
    ByRef outFound As Boolean _
) As Boolean
    Dim resolvedPath As String
    Dim conn As Object
    Dim rs As Object
    Dim sql As String

    outPositionCode = VBA.vbNullString
    outFound = False

    If VBA.Len(m_PersonnelWorkbookPath) = 0 Or VBA.Len(m_PersonnelTableRef) = 0 Then
        private_TryLookupPersonnelTvoPositionCode = True
        Exit Function
    End If

    resolvedPath = private_ResolveWorkbookPath(m_PersonnelWorkbookPath)
    If VBA.Len(resolvedPath) = 0 Or VBA.Len(VBA.Dir$(resolvedPath)) = 0 Then
        VBA.MsgBox "PrototypeNew: personnel source workbook was not found: " & m_PersonnelWorkbookPath, VBA.vbExclamation, "PrototypeNew / exporter data provider"
        Exit Function
    End If

    On Error GoTo LookupFail
    Set conn = VBA.CreateObject("ADODB.Connection")
    conn.Open private_BuildAdoConnectionString(resolvedPath)

    ' TVO-позиция зависит от текущей ШПС из профиля, поэтому живет здесь,
    ' а не в статическом CommonData. Ищем строку, где ФИО рапортующего
    ' указано в колонке "ТВО", и берем код должности этой строки.
    ' На стороне SQL тоже схлопываем переносы строк: в ШПС ФИО в "ТВО"
    ' часто визуально/фактически разбито на несколько строк внутри ячейки.
    sql = "SELECT TOP 1 " & private_QuoteSqlIdentifier(PERSONNEL_POSITION_CODE_HEADER) & _
        " FROM " & m_PersonnelTableRef & _
        " WHERE " & private_BuildNormalizedSqlTextExpression(PERSONNEL_TVO_HEADER) & " = " & private_AdoSqlTextLiteral(reporterFioText) & _
        " OR " & private_BuildNormalizedSqlTextExpression(PERSONNEL_TVO_HEADER) & " LIKE " & private_AdoSqlLikeContainsLiteral(reporterFioText, "%") & _
        " OR " & private_BuildNormalizedSqlTextExpression(PERSONNEL_TVO_HEADER) & " LIKE " & private_AdoSqlLikeContainsLiteral(reporterFioText, "*")

    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1

    If Not rs.EOF Then
        outPositionCode = VBA.Trim$(private_RecordsetFieldText(rs.Fields(0).Value))
        outFound = (VBA.Len(outPositionCode) > 0)
    End If
    If Not outFound Then
        If Not private_TryLookupPersonnelTvoPositionCodeByScan( _
            conn, _
            reporterFioText, _
            outPositionCode, _
            outFound) Then GoTo CleanupDone
    End If
    private_TryLookupPersonnelTvoPositionCode = True

CleanupDone:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Exit Function

LookupFail:
    VBA.MsgBox "PrototypeNew: failed to query personnel TVO data." & _
        VBA.vbCrLf & "Workbook: " & m_PersonnelWorkbookPath & _
        VBA.vbCrLf & "Range: " & m_PersonnelTableRef & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, "PrototypeNew / exporter data provider"
    Resume CleanupDone
End Function

Private Function private_TryLookupPersonnelTvoPositionCodeByScan( _
    ByVal conn As Object, _
    ByVal reporterFioText As String, _
    ByRef outPositionCode As String, _
    ByRef outFound As Boolean _
) As Boolean
    Dim rs As Object
    Dim sql As String
    Dim tvoCellText As String

    outPositionCode = VBA.vbNullString
    outFound = False
    If conn Is Nothing Then Exit Function

    On Error GoTo ScanFail
    sql = "SELECT " & private_QuoteSqlIdentifier(PERSONNEL_POSITION_CODE_HEADER) & _
        ", " & private_QuoteSqlIdentifier(PERSONNEL_TVO_HEADER) & _
        " FROM " & m_PersonnelTableRef

    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1

    Do While Not rs.EOF
        tvoCellText = private_NormalizeLookupKey(private_RecordsetFieldText(rs.Fields(1).Value))
        If VBA.InStr(1, tvoCellText, reporterFioText, VBA.vbTextCompare) > 0 Then
            outPositionCode = VBA.Trim$(private_RecordsetFieldText(rs.Fields(0).Value))
            outFound = (VBA.Len(outPositionCode) > 0)
            Exit Do
        End If
        rs.MoveNext
    Loop

    private_TryLookupPersonnelTvoPositionCodeByScan = True

CleanupDone:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
    Exit Function

ScanFail:
    VBA.MsgBox "PrototypeNew: failed to scan personnel TVO data." & _
        VBA.vbCrLf & "Workbook: " & m_PersonnelWorkbookPath & _
        VBA.vbCrLf & "Range: " & m_PersonnelTableRef & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, "PrototypeNew / exporter data provider"
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
    If VBA.Len(rawRangeRef) = 0 Then rawRangeRef = DEFAULT_PERSONNEL_STATE_RANGE_REF
    rawRangeRef = VBA.Replace(rawRangeRef, "]", "]]")
    private_BuildConfiguredAdoRangeRef = "[" & rawRangeRef & "]"
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
