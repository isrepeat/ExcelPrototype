Option Explicit
#Const LOGGING_DEBUG_ENABLED = False

' Config example (Settings sheet: column A = key, column B = value):
'
' Source.Main.FilePath                      | C:\Data\People.xlsx
' Source.Main.Tables[PeopleTable].Ref       | People$A1:D100
'
' Query.TestQuery.From           | Source.Main.Tables[PeopleTable].Ref
' Query.TestQuery.Select         | [FullName], [BirthDate]
' Query.TestQuery.Where          | [FullName] LIKE '%{Input!B2}%'
' Query.TestQuery.Handler        | fn_Handler_FirstRowToText
' Query.TestQuery.Output         | Result!B2
'
' Query.PaymentsQuery.From       | Source.Payments.Tables[PaymentsTable].Ref
' Query.PaymentsQuery.Select     | [Звання] AS [Rank], [ПІБ] AS [FullName], [ІПН] AS [TaxId], [Посада] AS [PositionCode], [Вид] AS [PaymentType], [Рапорт] AS [Report], [Дата] AS [ReportDate], [Наказ] AS [OrderNo]
' Query.PaymentsQuery.Where      | [Вид] = 'ГДО'
'
' Query.StateQuery.From          | Source.State.Tables[StateTable].Ref
' Query.StateQuery.Select        | [ІПН] AS [TaxId], [Давальний] AS [DativeName], [Родовий] AS [GenitiveName]
'
' Source.Rank.FilePath           | C:\Data\Переліки.xlsx
' Source.Rank.Tables[RankTable].Ref | Звання$A1:E50000
'
' Query.RankQuery.From           | Source.Rank.Tables[RankTable].Ref
' Query.RankQuery.Select         | [Звання] AS [Rank], [Родовий] AS [RankGenitive], [Давальний] AS [RankDative]
'
' Source.Position.FilePath       | C:\Data\Посади.xlsm
' Source.Position.Tables[PositionTable].Ref | Посади$A4:E50000
'
' Query.PositionQuery.From       | Source.Position.Tables[PositionTable].Ref
' Query.PositionQuery.Select     | [Код] AS [PositionCode], [Давальний] AS [PositionDative]
'
' QueryGroup.GdoOrder.Queries    | PaymentsQuery, StateQuery, RankQuery, PositionQuery
' QueryGroup.GdoOrder.Handler    | fn_Handler_BuildGdoOrderText
' QueryGroup.GdoOrder.HandlerArgs| CheckBox="Виплати!N3"
' QueryGroup.GdoOrder.Output     | Виплати!G6
'
' QueryGroup handler signature:
' Public Function fn_Handler_BuildGdoOrderText(ByVal recordsetsDict As Object, ByVal queryGroupDict As Object, Optional ByVal handlerArgsDict As Object = Nothing) As String
'   recordsetsDict("PaymentsQuery") -> ADODB.Recordset
'   recordsetsDict("StateQuery")    -> ADODB.Recordset
'   recordsetsDict("RankQuery")     -> ADODB.Recordset
'   recordsetsDict("PositionQuery") -> ADODB.Recordset
'   handlerArgsDict("CheckBox")     -> TRUE/FALSE (из привязанной ячейки чекбокса)
'
' Button.RunPeople.ShapeRef      | Input!btnRunPeople
' Button.RunPeople.Query         | TestQuery
' Button.RunGdo.ShapeRef         | Виплати!btnRun
' Button.RunGdo.QueryGroup       | GdoOrder

Private Const SETTINGS_SHEET As String = "Settings"
Private Const NO_QUERY_RESULTS_TEXT As String = "<No query results>"
Private Const DIAGNOSTIC_LOG_FILE_REL_PATH As String = "Logs\\templateprototype.log"

Private gSettingsDict As Object
Private gSourcesDict As Object
Private gSourceAliasesDict As Object
Private gQueriesDict As Object
Private gQueryGroupsDict As Object
Private gButtonAliasesDict As Object
Private gButtonsDict As Object
' Parser storage shape:
' gSettingsDict: raw Settings rows by key.
'   gSettingsDict("Source.Main.FilePath") = "C:\Data\People.xlsx"
'   gSettingsDict("Query.TestQuery.From") = "Source.Main.Tables[PeopleTable].Ref"
'
' gSourcesDict: source dictionaries by file path.
'   gSourcesDict("C:\Data\People.xlsx")("Alias") = "Main"
'   gSourcesDict("C:\Data\People.xlsx")("FilePath") = "C:\Data\People.xlsx"
'   gSourcesDict("C:\Data\People.xlsx")("TablesDict")("PeopleTable") = "People$A1:D100"
'
' gSourceAliasesDict: same source dictionaries by source alias.
'   gSourceAliasesDict("Main") Is gSourcesDict("C:\Data\People.xlsx")
'
' gQueriesDict: validated query dictionaries by query alias. Handler/Output are required only for direct query execution.
'   gQueriesDict("TestQuery")("From") = "Source.Main.Tables[PeopleTable].Ref"
'   gQueriesDict("TestQuery")("SourceFilePath") = "C:\Data\People.xlsx"
'   gQueriesDict("TestQuery")("TableAlias") = "PeopleTable"
'   gQueriesDict("TestQuery")("TableRef") = "People$A1:D100"
'   gQueriesDict("TestQuery")("Handler") = "fn_Handler_FirstRowToText"
'   gQueriesDict("TestQuery")("Output") = "Result!B2"
'
' gQueryGroupsDict: validated query group dictionaries by group alias.
'   gQueryGroupsDict("GdoOrder")("Queries") = "PaymentsQuery, StateQuery, RankQuery, PositionQuery"
'   gQueryGroupsDict("GdoOrder")("QueriesDict")("PaymentsQuery") Is gQueriesDict("PaymentsQuery")
'   gQueryGroupsDict("GdoOrder")("QueriesDict")("StateQuery") Is gQueriesDict("StateQuery")
'   gQueryGroupsDict("GdoOrder")("QueriesDict")("RankQuery") Is gQueriesDict("RankQuery")
'   gQueryGroupsDict("GdoOrder")("QueriesDict")("PositionQuery") Is gQueriesDict("PositionQuery")
'   gQueryGroupsDict("GdoOrder")("Handler") = "fn_Handler_BuildGdoOrderText"
'   gQueryGroupsDict("GdoOrder")("Output") = "Виплати!G6"
'
' gButtonAliasesDict: button dictionaries by button alias before binding validation.
'   gButtonAliasesDict("RunPeople")("ShapeRef") = "Input!btnRunPeople"
'   gButtonAliasesDict("RunPeople")("Query") = "TestQuery"
'   gButtonAliasesDict("RunGdo")("QueryGroup") = "GdoOrder"
'
' gButtonsDict: validated button dictionaries by sheet and shape key.
'   gButtonsDict("Input!btnRunPeople")("Alias") = "RunPeople"
'   gButtonsDict("Input!btnRunPeople")("ShapeRef") = "Input!btnRunPeople"
'   gButtonsDict("Input!btnRunPeople")("Query") = "TestQuery"
'   gButtonsDict("Виплати!btnRun")("QueryGroup") = "GdoOrder"

' //
' // API
' //
' --------------------------------------
'  namespace Query {
' --------------------------------------
Public Sub fn_Query_RunQueryByAlias(ByVal queryAlias As String)
    Dim queryDict As Object

    On Error GoTo ErrHandler

    private_Config_Load
    Set queryDict = private_Dict_Require(gQueriesDict, queryAlias, "Query not found")
    private_Query_RunQueryDict queryDict
    Exit Sub

ErrHandler:
    VBA.MsgBox "Error: " & Err.Description, VBA.vbCritical
End Sub
' --------------------------------------
'  } // namespace Query
' --------------------------------------

' --------------------------------------
'  namespace QueryGroup {
' --------------------------------------
Public Sub fn_QueryGroup_RunQueryGroupByAlias(ByVal queryGroupAlias As String)
    Dim queryGroupDict As Object

    On Error GoTo ErrHandler

    private_Config_Load
    Set queryGroupDict = private_Dict_Require(gQueryGroupsDict, queryGroupAlias, "Query group not found")
    private_QueryGroup_RunQueryGroupDict queryGroupDict
    Exit Sub

ErrHandler:
    VBA.MsgBox "Error: " & Err.Description, VBA.vbCritical
End Sub
' --------------------------------------
'  } // namespace QueryGroup
' --------------------------------------

' --------------------------------------
'  namespace Button {
' --------------------------------------
Public Sub fn_Button_RunAssignedQuery()
    Dim buttonDict As Object
    Dim queryDict As Object
    Dim queryGroupDict As Object

    On Error GoTo ErrHandler

    private_Config_Load
    Set buttonDict = private_Button_ResolveCaller()

    If buttonDict.Exists("QueryGroup") Then
        Set queryGroupDict = private_Dict_Require(gQueryGroupsDict, buttonDict("QueryGroup"), "Button query group not found")
        private_QueryGroup_RunQueryGroupDict queryGroupDict
    Else
        Set queryDict = private_Dict_Require(gQueriesDict, buttonDict("Query"), "Button query not found")
        private_Query_RunQueryDict queryDict
    End If
    Exit Sub

ErrHandler:
    VBA.MsgBox "Error: " & Err.Description, VBA.vbCritical
End Sub
' --------------------------------------
'  } // namespace Button
' --------------------------------------

' --------------------------------------
'  namespace Handler {
' --------------------------------------
Public Function fn_Handler_FirstRowToText(ByVal rs As Object) As String
    Dim i As Long
    Dim resultText As String

    If rs.EOF Then
        fn_Handler_FirstRowToText = NO_QUERY_RESULTS_TEXT
        Exit Function
    End If

    For i = 0 To rs.Fields.Count - 1
        If i > 0 Then resultText = resultText & "; "
        resultText = resultText & private_Text_NullToEmptyString(rs.Fields(i).Value)
    Next i

    fn_Handler_FirstRowToText = resultText
End Function


Public Function fn_Handler_BuildGdoOrderText(ByVal recordsetsDict As Object, ByVal queryGroupDict As Object, Optional ByVal handlerArgsDict As Object = Nothing) As String
    Dim paymentsRs As Object
    Dim stateRs As Object
    Dim rankRs As Object
    Dim positionRs As Object
    Dim stateDict As Object
    Dim rankDict As Object
    Dim positionDict As Object
    Dim stateRowDict As Object
    Dim rankRowDict As Object
    Dim taxId As String
    Dim rankText As String
    Dim positionCode As String
    Dim resultText As String
    Dim includeTaxId As Boolean

    Set paymentsRs = private_Dict_Require(recordsetsDict, "PaymentsQuery", "QueryGroup recordset not found")
    Set stateRs = private_Dict_Require(recordsetsDict, "StateQuery", "QueryGroup recordset not found")
    Set rankRs = private_Dict_Require(recordsetsDict, "RankQuery", "QueryGroup recordset not found")
    Set positionRs = private_Dict_Require(recordsetsDict, "PositionQuery", "QueryGroup recordset not found")

    private_Recordset_RequireFieldCount paymentsRs, "PaymentsQuery", 8
    private_Recordset_RequireFieldCount stateRs, "StateQuery", 3
    private_Recordset_RequireFieldCount rankRs, "RankQuery", 3
    private_Recordset_RequireFieldCount positionRs, "PositionQuery", 2
    includeTaxId = private_HandlerArgs_GetBoolean(handlerArgsDict, "CheckBox", False)

    If paymentsRs.EOF Then
        fn_Handler_BuildGdoOrderText = NO_QUERY_RESULTS_TEXT
        Exit Function
    End If

    Set stateDict = private_Handler_BuildStateDict(stateRs)
    Set rankDict = private_Handler_BuildRankDict(rankRs)
    Set positionDict = private_Handler_BuildPositionDict(positionRs)
    paymentsRs.MoveFirst

    Do Until paymentsRs.EOF
        taxId = private_Handler_RequireText(private_Recordset_FieldTextOrIndex(paymentsRs, "TaxId", 2), "Payments row has empty TaxId.")
        rankText = private_Text_NormalizeLookupToken(private_Handler_RequireText(private_Recordset_FieldTextOrIndex(paymentsRs, "Rank", 0), "Payments row has empty Rank."))
        positionCode = private_Handler_RequireText(private_Recordset_FieldTextOrIndex(paymentsRs, "PositionCode", 3), "Payments row has empty PositionCode.")

        If Not stateDict.Exists(taxId) Then
            Err.Raise VBA.vbObjectError + 5101, , "State row not found for TaxId: " & taxId
        End If

        If Not rankDict.Exists(rankText) Then
            Err.Raise VBA.vbObjectError + 5102, , "Rank declension row not found for Rank: " & rankText
        End If

        If Not positionDict.Exists(positionCode) Then
            Err.Raise VBA.vbObjectError + 5103, , "Position declension row not found for PositionCode: " & positionCode
        End If

        Set stateRowDict = stateDict(taxId)
        Set rankRowDict = rankDict(rankText)
        If VBA.Len(resultText) > 0 Then resultText = resultText & VBA.vbCrLf & VBA.vbCrLf
        resultText = resultText & private_Handler_BuildGdoOrderItem(paymentsRs, stateRowDict, rankRowDict, positionDict(positionCode), taxId, includeTaxId)
        paymentsRs.MoveNext
    Loop

    fn_Handler_BuildGdoOrderText = resultText
End Function
' --------------------------------------
'  } // namespace Handler
' --------------------------------------

' //
' // Internal
' //
' --------------------------------------
'  namespace Query {
' --------------------------------------
Private Sub private_Query_RunQueryDict(ByVal queryDict As Object)
    Dim resultText As String
    Dim rs As Object
    Dim recordsetsDict As Object
    Dim connectionsDict As Object

    On Error GoTo ErrHandler

    private_Query_RequireStandaloneFields queryDict
    Set recordsetsDict = private_Dict_CreateTextMap()
    Set connectionsDict = private_Dict_CreateTextMap()
    private_Query_OpenRecordset queryDict, recordsetsDict, connectionsDict, queryDict("Alias")
    Set rs = recordsetsDict(queryDict("Alias"))

    If rs.EOF Then
        resultText = NO_QUERY_RESULTS_TEXT
    Else
        resultText = Application.Run(private_Handler_ResolveName(queryDict("Handler")), rs)
    End If

    private_Output_WriteResult queryDict("Output"), resultText

CleanExit:
    private_Query_CloseRuntimeDicts recordsetsDict, connectionsDict
    Exit Sub

ErrHandler:
#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogError "query:run-failed err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
    VBA.MsgBox "Error: " & Err.Description, VBA.vbCritical
    Resume CleanExit
End Sub


Private Sub private_Query_RequireStandaloneFields(ByVal queryDict As Object)
    queryDict("Handler") = private_Dict_RequireText(queryDict, "Handler", "Missing Query." & queryDict("Alias") & ".Handler")
    queryDict("Output") = private_Dict_RequireText(queryDict, "Output", "Missing Query." & queryDict("Alias") & ".Output")
End Sub


Private Sub private_Query_OpenRecordset(ByVal queryDict As Object, ByVal recordsetsDict As Object, ByVal connectionsDict As Object, ByVal recordsetAlias As String)
    Dim filePath As String
    Dim tableRef As String
    Dim sql As String
    Dim sqlAlternative As String
    Dim cn As Object
    Dim rs As Object
    Dim primaryOpenErrNumber As Long
    Dim primaryOpenErrDescription As String
    Dim altOpenErrNumber As Long
    Dim altOpenErrDescription As String

    filePath = queryDict("SourceFilePath")
    tableRef = queryDict("TableRef")
    sql = private_Sql_Build(queryDict, tableRef)
    sqlAlternative = private_Dict_GetValue(queryDict, "SqlAlternative", "")

#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogInfo "query:open-rs alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "' source='" & VBA.Replace$(filePath, "'", "''") & "' table='" & VBA.Replace$(tableRef, "'", "''") & "'"
    fn_Diagnostic_LogInfo "query:sql alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "' sql='" & VBA.Replace$(sql, "'", "''") & "'"
#End If

    Set cn = VBA.CreateObject("ADODB.Connection")
    Set rs = VBA.CreateObject("ADODB.Recordset")

    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & filePath & ";" & _
            "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"

    On Error Resume Next
    rs.Open sql, cn, 1, 1
    primaryOpenErrNumber = Err.Number
    primaryOpenErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    If primaryOpenErrNumber <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        fn_Diagnostic_LogError "query:open-rs-failed alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "' err='" & VBA.Replace$(primaryOpenErrDescription, "'", "''") & "'"
#End If

        If VBA.Len(sqlAlternative) > 0 Then
#If LOGGING_DEBUG_ENABLED Then
            fn_Diagnostic_LogWarning "query:open-rs-primary-failed-retry-alt alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "'"
            fn_Diagnostic_LogInfo "query:sql-alt alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "' sql='" & VBA.Replace$(sqlAlternative, "'", "''") & "'"
#End If

            On Error Resume Next
            rs.Open sqlAlternative, cn, 1, 1
            altOpenErrNumber = Err.Number
            altOpenErrDescription = Err.Description
            Err.Clear
            On Error GoTo 0

            If altOpenErrNumber <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
                fn_Diagnostic_LogError "query:open-rs-alt-failed alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "' err='" & VBA.Replace$(altOpenErrDescription, "'", "''") & "'"
#End If
                Err.Raise VBA.vbObjectError + 4601, , _
                          "Recordset open failed for query '" & recordsetAlias & "'. Primary error: " & primaryOpenErrDescription & "; alternative error: " & altOpenErrDescription
            End If
#If LOGGING_DEBUG_ENABLED Then
            fn_Diagnostic_LogInfo "query:open-rs-alt-has-rows alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "'"
#End If
        Else
            Err.Raise VBA.vbObjectError + 4600, , "Recordset open failed for query '" & recordsetAlias & "': " & primaryOpenErrDescription
        End If
    End If

#If LOGGING_DEBUG_ENABLED Then
    If rs.EOF Then
        fn_Diagnostic_LogWarning "query:open-rs-empty alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "'"
    Else
        fn_Diagnostic_LogInfo "query:open-rs-has-rows alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "'"
    End If
#End If

    If rs.EOF And VBA.Len(sqlAlternative) > 0 Then
#If LOGGING_DEBUG_ENABLED Then
        fn_Diagnostic_LogWarning "query:open-rs-retry-alt alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "'"
        fn_Diagnostic_LogInfo "query:sql-alt alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "' sql='" & VBA.Replace$(sqlAlternative, "'", "''") & "'"
#End If

        rs.Close
        On Error Resume Next
        rs.Open sqlAlternative, cn, 1, 1
        altOpenErrNumber = Err.Number
        altOpenErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0

        If altOpenErrNumber <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
            fn_Diagnostic_LogWarning "query:open-rs-alt-retry-failed alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "' err='" & VBA.Replace$(altOpenErrDescription, "'", "''") & "'"
            fn_Diagnostic_LogInfo "query:open-rs-alt-retry-restore-primary alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "'"
#End If
            rs.Open sql, cn, 1, 1
        End If
#If LOGGING_DEBUG_ENABLED Then
        If altOpenErrNumber = 0 And rs.EOF Then
            fn_Diagnostic_LogWarning "query:open-rs-alt-empty alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "'"
        ElseIf altOpenErrNumber = 0 Then
            fn_Diagnostic_LogInfo "query:open-rs-alt-has-rows alias='" & VBA.Replace$(recordsetAlias, "'", "''") & "'"
        End If
#End If
    End If

    connectionsDict.Add recordsetAlias, cn
    recordsetsDict.Add recordsetAlias, rs
End Sub


Private Sub private_Query_CloseRuntimeDicts(ByVal recordsetsDict As Object, ByVal connectionsDict As Object)
    Dim k As Variant
    Dim rs As Object
    Dim cn As Object

    On Error Resume Next
    If Not recordsetsDict Is Nothing Then
        For Each k In recordsetsDict.Keys
            Set rs = recordsetsDict(k)
            If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
        Next k
    End If

    If Not connectionsDict Is Nothing Then
        For Each k In connectionsDict.Keys
            Set cn = connectionsDict(k)
            If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
        Next k
    End If
End Sub
' --------------------------------------
'  } // namespace Query
' --------------------------------------
' --------------------------------------
'  namespace QueryGroup {
' --------------------------------------
Private Sub private_QueryGroup_RunQueryGroupDict(ByVal queryGroupDict As Object)
    Dim queriesDict As Object
    Dim recordsetsDict As Object
    Dim connectionsDict As Object
    Dim handlerArgsDict As Object
    Dim queryAlias As Variant
    Dim queryDict As Object
    Dim resultText As String

    On Error GoTo ErrHandler

    Set queriesDict = queryGroupDict("QueriesDict")
    Set recordsetsDict = private_Dict_CreateTextMap()
    Set connectionsDict = private_Dict_CreateTextMap()

    For Each queryAlias In queriesDict.Keys
        Set queryDict = queriesDict(queryAlias)
        private_Query_OpenRecordset queryDict, recordsetsDict, connectionsDict, VBA.CStr(queryAlias)
    Next queryAlias

    If queryGroupDict.Exists("HandlerArgsDict") Then Set handlerArgsDict = queryGroupDict("HandlerArgsDict")

    If handlerArgsDict Is Nothing Then
        resultText = Application.Run(private_Handler_ResolveName(queryGroupDict("Handler")), recordsetsDict, queryGroupDict)
    ElseIf handlerArgsDict.Count = 0 Then
        resultText = Application.Run(private_Handler_ResolveName(queryGroupDict("Handler")), recordsetsDict, queryGroupDict)
    Else
        resultText = Application.Run(private_Handler_ResolveName(queryGroupDict("Handler")), recordsetsDict, queryGroupDict, handlerArgsDict)
    End If

    private_Output_WriteResult queryGroupDict("Output"), resultText

CleanExit:
    private_Query_CloseRuntimeDicts recordsetsDict, connectionsDict
    Exit Sub

ErrHandler:
#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogError "query-group:run-failed err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
    VBA.MsgBox "Error: " & Err.Description, VBA.vbCritical
    Resume CleanExit
End Sub
' --------------------------------------
'  } // namespace QueryGroup
' --------------------------------------
' --------------------------------------
'  namespace Config {
' --------------------------------------
Private Sub private_Config_Load()
    Set gSettingsDict = private_Dict_CreateTextMap()
    Set gSourcesDict = private_Dict_CreateTextMap()
    Set gSourceAliasesDict = private_Dict_CreateTextMap()
    Set gQueriesDict = private_Dict_CreateTextMap()
    Set gQueryGroupsDict = private_Dict_CreateTextMap()
    Set gButtonAliasesDict = private_Dict_CreateTextMap()
    Set gButtonsDict = private_Dict_CreateTextMap()

    private_Config_LoadRawSettings
    private_Config_ParseSources
    private_Config_ValidateSources
    private_Config_ParseQueries
    private_Config_ValidateQueries
    private_Config_ParseQueryGroups
    private_Config_ValidateQueryGroups
    private_Config_ParseButtons
    private_Config_ValidateButtons
End Sub


Private Sub private_Config_LoadRawSettings()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim k As String
    Dim v As String

    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        k = VBA.Trim$(ws.Cells(r, "A").Value)
        v = VBA.Trim$(ws.Cells(r, "B").Value)
        If VBA.Len(k) > 0 Then
            If gSettingsDict.Exists(k) Then
                Err.Raise VBA.vbObjectError + 4000, , "Duplicate Settings key: " & k
            End If
            gSettingsDict(k) = v
        End If
    Next r
End Sub


Private Sub private_Config_ParseSources()
    Dim k As Variant

    For Each k In gSettingsDict.Keys
        If VBA.Left$(VBA.CStr(k), 7) = "Source." Then
            private_Config_ParseSourceKey VBA.CStr(k), gSettingsDict(k)
        End If
    Next k
End Sub


Private Sub private_Config_ParseSourceKey(ByVal key As String, ByVal value As String)
    Dim sourceAlias As String
    Dim tableAlias As String

    If private_Source_TryParseFilePathKey(key, sourceAlias) Then
        private_Source_SetFilePath sourceAlias, value
    ElseIf private_Source_TryParseTableRefKey(key, sourceAlias, tableAlias) Then
        private_Source_AddTableRef sourceAlias, tableAlias, value
    Else
        Err.Raise VBA.vbObjectError + 4100, , "Invalid Source config key: " & key
    End If
End Sub


Private Sub private_Config_ValidateSources()
    Dim sourceAlias As Variant
    Dim sourceDict As Object

    For Each sourceAlias In gSourceAliasesDict.Keys
        Set sourceDict = gSourceAliasesDict(sourceAlias)
        If Not sourceDict.Exists("FilePath") Then
            Err.Raise VBA.vbObjectError + 4101, , "Missing source FilePath: Source." & VBA.CStr(sourceAlias) & ".FilePath"
        End If

        If sourceDict("TablesDict").Count = 0 Then
            Err.Raise VBA.vbObjectError + 4102, , "Missing source table refs for source: " & VBA.CStr(sourceAlias)
        End If
    Next sourceAlias
End Sub


Private Sub private_Config_ParseQueries()
    Dim k As Variant

    For Each k In gSettingsDict.Keys
        If VBA.Left$(VBA.CStr(k), 6) = "Query." Then
            private_Config_ParseQueryKey VBA.CStr(k), gSettingsDict(k)
        End If
    Next k
End Sub


Private Sub private_Config_ParseQueryKey(ByVal key As String, ByVal value As String)
    Dim queryAlias As String
    Dim propName As String
    Dim queryDict As Object

    private_Config_ParseQueryPropertyKey key, queryAlias, propName

    If Not gQueriesDict.Exists(queryAlias) Then
        Set queryDict = private_Dict_CreateTextMap()
        queryDict("Alias") = queryAlias
        gQueriesDict.Add queryAlias, queryDict
    End If

    gQueriesDict(queryAlias)(propName) = value
End Sub


Private Sub private_Config_ParseQueryPropertyKey(ByVal key As String, ByRef outQueryAlias As String, ByRef outPropName As String)
    Dim remainder As String
    Dim separatorPos As Long

    remainder = VBA.Mid$(key, VBA.Len("Query.") + 1)
    separatorPos = VBA.InStr(1, remainder, ".")
    If separatorPos <= 1 Or separatorPos = VBA.Len(remainder) Then
        Err.Raise VBA.vbObjectError + 4200, , "Invalid Query config key: " & key
    End If

    outQueryAlias = VBA.Left$(remainder, separatorPos - 1)
    outPropName = VBA.Mid$(remainder, separatorPos + 1)
End Sub


Private Sub private_Config_ValidateQueries()
    Dim queryAlias As Variant
    Dim queryDict As Object

    For Each queryAlias In gQueriesDict.Keys
        Set queryDict = gQueriesDict(queryAlias)
        private_Config_ValidateQuery queryDict
    Next queryAlias
End Sub


Private Sub private_Config_ValidateQuery(ByVal queryDict As Object)
    Dim fromRef As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim sourceDict As Object
    Dim tableRef As String

    fromRef = private_Dict_RequireText(queryDict, "From", "Missing Query." & queryDict("Alias") & ".From")

    private_Source_ResolveTableRef fromRef, sourceAlias, tableAlias, sourceDict, tableRef
    queryDict("SourceAlias") = sourceAlias
    queryDict("SourceFilePath") = sourceDict("FilePath")
    queryDict("TableAlias") = tableAlias
    queryDict("TableRef") = tableRef
End Sub


Private Sub private_Config_ParseQueryGroups()
    Dim k As Variant

    For Each k In gSettingsDict.Keys
        If VBA.Left$(VBA.CStr(k), VBA.Len("QueryGroup.")) = "QueryGroup." Then
            private_Config_ParseQueryGroupKey VBA.CStr(k), gSettingsDict(k)
        End If
    Next k
End Sub


Private Sub private_Config_ParseQueryGroupKey(ByVal key As String, ByVal value As String)
    Dim queryGroupAlias As String
    Dim propName As String
    Dim queryGroupDict As Object

    private_Config_ParseQueryGroupPropertyKey key, queryGroupAlias, propName

    If Not gQueryGroupsDict.Exists(queryGroupAlias) Then
        Set queryGroupDict = private_Dict_CreateTextMap()
        queryGroupDict("Alias") = queryGroupAlias
        gQueryGroupsDict.Add queryGroupAlias, queryGroupDict
    End If

    gQueryGroupsDict(queryGroupAlias)(propName) = value
End Sub


Private Sub private_Config_ParseQueryGroupPropertyKey(ByVal key As String, ByRef outQueryGroupAlias As String, ByRef outPropName As String)
    Dim remainder As String
    Dim separatorPos As Long

    remainder = VBA.Mid$(key, VBA.Len("QueryGroup.") + 1)
    separatorPos = VBA.InStr(1, remainder, ".")
    If separatorPos <= 1 Or separatorPos = VBA.Len(remainder) Then
        Err.Raise VBA.vbObjectError + 4250, , "Invalid QueryGroup config key: " & key
    End If

    outQueryGroupAlias = VBA.Left$(remainder, separatorPos - 1)
    outPropName = VBA.Mid$(remainder, separatorPos + 1)
End Sub


Private Sub private_Config_ValidateQueryGroups()
    Dim queryGroupAlias As Variant
    Dim queryGroupDict As Object

    For Each queryGroupAlias In gQueryGroupsDict.Keys
        Set queryGroupDict = gQueryGroupsDict(queryGroupAlias)
        private_Config_ValidateQueryGroup queryGroupDict
    Next queryGroupAlias
End Sub


Private Sub private_Config_ValidateQueryGroup(ByVal queryGroupDict As Object)
    Dim queriesText As String
    Dim handlerName As String
    Dim handlerArgsText As String
    Dim outputRef As String
    Dim queryAliases As Variant
    Dim i As Long
    Dim queryAlias As String
    Dim queriesDict As Object
    Dim handlerArgsDict As Object

    queriesText = private_Dict_RequireText(queryGroupDict, "Queries", "Missing QueryGroup." & queryGroupDict("Alias") & ".Queries")
    handlerName = private_Dict_RequireText(queryGroupDict, "Handler", "Missing QueryGroup." & queryGroupDict("Alias") & ".Handler")
    handlerArgsText = private_Dict_GetValue(queryGroupDict, "HandlerArgs", "")
    outputRef = private_Dict_RequireText(queryGroupDict, "Output", "Missing QueryGroup." & queryGroupDict("Alias") & ".Output")

    Set queriesDict = private_Dict_CreateTextMap()
    queryAliases = VBA.Split(queriesText, ",")

    For i = LBound(queryAliases) To UBound(queryAliases)
        queryAlias = VBA.Trim$(VBA.CStr(queryAliases(i)))
        If VBA.Len(queryAlias) = 0 Then
            Err.Raise VBA.vbObjectError + 4251, , "Empty query alias in QueryGroup." & queryGroupDict("Alias") & ".Queries"
        End If

        If Not gQueriesDict.Exists(queryAlias) Then
            Err.Raise VBA.vbObjectError + 4252, , "QueryGroup '" & queryGroupDict("Alias") & "' references unknown query: " & queryAlias
        End If

        If queriesDict.Exists(queryAlias) Then
            Err.Raise VBA.vbObjectError + 4253, , "Duplicate query alias in QueryGroup." & queryGroupDict("Alias") & ".Queries: " & queryAlias
        End If

        queriesDict.Add queryAlias, gQueriesDict(queryAlias)
    Next i

    queryGroupDict("Handler") = handlerName
    queryGroupDict("HandlerArgs") = handlerArgsText
    queryGroupDict("Output") = outputRef
    Set handlerArgsDict = private_Config_ParseHandlerArgsDict(handlerArgsText, queryGroupDict("Alias"))
    If queryGroupDict.Exists("QueriesDict") Then queryGroupDict.Remove "QueriesDict"
    If queryGroupDict.Exists("HandlerArgsDict") Then queryGroupDict.Remove "HandlerArgsDict"
    queryGroupDict.Add "QueriesDict", queriesDict
    queryGroupDict.Add "HandlerArgsDict", handlerArgsDict
End Sub


Private Function private_Config_ParseHandlerArgsDict(ByVal handlerArgsText As String, ByVal queryGroupAlias As String) As Object
    Dim argsDict As Object
    Dim token As String
    Dim inQuotes As Boolean
    Dim i As Long
    Dim ch As String

    Set argsDict = private_Dict_CreateTextMap()
    handlerArgsText = VBA.Trim$(handlerArgsText)
    If VBA.Len(handlerArgsText) = 0 Then
        Set private_Config_ParseHandlerArgsDict = argsDict
        Exit Function
    End If

    For i = 1 To VBA.Len(handlerArgsText)
        ch = VBA.Mid$(handlerArgsText, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
            token = token & ch
        ElseIf ch = ";" And Not inQuotes Then
            private_Config_AddHandlerArgToken argsDict, token, queryGroupAlias
            token = ""
        Else
            token = token & ch
        End If
    Next i

    If inQuotes Then
        Err.Raise VBA.vbObjectError + 4254, , "Unclosed quote in QueryGroup." & queryGroupAlias & ".HandlerArgs"
    End If

    private_Config_AddHandlerArgToken argsDict, token, queryGroupAlias
    Set private_Config_ParseHandlerArgsDict = argsDict
End Function


Private Sub private_Config_AddHandlerArgToken(ByVal argsDict As Object, ByVal rawToken As String, ByVal queryGroupAlias As String)
    Dim separatorPos As Long
    Dim argName As String
    Dim argValue As String

    rawToken = VBA.Trim$(rawToken)
    If VBA.Len(rawToken) = 0 Then Exit Sub

    separatorPos = VBA.InStr(1, rawToken, "=")
    If separatorPos <= 1 Or separatorPos = VBA.Len(rawToken) Then
        Err.Raise VBA.vbObjectError + 4255, , "Invalid token in QueryGroup." & queryGroupAlias & ".HandlerArgs: " & rawToken
    End If

    argName = VBA.Trim$(VBA.Left$(rawToken, separatorPos - 1))
    argValue = VBA.Trim$(VBA.Mid$(rawToken, separatorPos + 1))
    argValue = private_Config_UnquoteHandlerArg(argValue)
    If private_CellRef_LooksLikeRef(argValue) Then
        argValue = private_CellRef_Read(argValue)
    End If

    If VBA.Len(argName) = 0 Then
        Err.Raise VBA.vbObjectError + 4256, , "Empty arg name in QueryGroup." & queryGroupAlias & ".HandlerArgs"
    End If

    If argsDict.Exists(argName) Then
        Err.Raise VBA.vbObjectError + 4257, , "Duplicate arg name in QueryGroup." & queryGroupAlias & ".HandlerArgs: " & argName
    End If

    argsDict.Add argName, argValue
End Sub


Private Function private_Config_UnquoteHandlerArg(ByVal valueText As String) As String
    valueText = VBA.Trim$(valueText)
    Do While VBA.Len(valueText) >= 2 And VBA.Left$(valueText, 1) = """" And VBA.Right$(valueText, 1) = """"
        valueText = VBA.Mid$(valueText, 2, VBA.Len(valueText) - 2)
    Loop
    valueText = VBA.Replace(valueText, """""", """")
    private_Config_UnquoteHandlerArg = valueText
End Function


Private Sub private_Config_ParseButtons()
    Dim k As Variant

    For Each k In gSettingsDict.Keys
        If VBA.Left$(VBA.CStr(k), 7) = "Button." Then
            private_Config_ParseButtonKey VBA.CStr(k), gSettingsDict(k)
        End If
    Next k
End Sub


Private Sub private_Config_ParseButtonKey(ByVal key As String, ByVal value As String)
    Dim buttonAlias As String
    Dim propName As String
    Dim buttonDict As Object

    private_Config_ParseButtonPropertyKey key, buttonAlias, propName

    If Not gButtonAliasesDict.Exists(buttonAlias) Then
        Set buttonDict = private_Dict_CreateTextMap()
        buttonDict("Alias") = buttonAlias
        gButtonAliasesDict.Add buttonAlias, buttonDict
    End If

    gButtonAliasesDict(buttonAlias)(propName) = value
End Sub


Private Sub private_Config_ParseButtonPropertyKey(ByVal key As String, ByRef outButtonAlias As String, ByRef outPropName As String)
    Dim remainder As String
    Dim separatorPos As Long

    remainder = VBA.Mid$(key, VBA.Len("Button.") + 1)
    separatorPos = VBA.InStr(1, remainder, ".")
    If separatorPos <= 1 Or separatorPos = VBA.Len(remainder) Then
        Err.Raise VBA.vbObjectError + 4300, , "Invalid Button config key: " & key
    End If

    outButtonAlias = VBA.Left$(remainder, separatorPos - 1)
    outPropName = VBA.Mid$(remainder, separatorPos + 1)
End Sub


Private Sub private_Config_ValidateButtons()
    Dim buttonAlias As Variant
    Dim buttonDict As Object
    Dim buttonKey As String
    Dim hasQuery As Boolean
    Dim hasQueryGroup As Boolean
    Dim queryDict As Object

    For Each buttonAlias In gButtonAliasesDict.Keys
        Set buttonDict = gButtonAliasesDict(buttonAlias)
        buttonDict("ShapeRef") = private_Dict_RequireText(buttonDict, "ShapeRef", "Missing Button." & buttonDict("Alias") & ".ShapeRef")
        buttonDict("ShapeRef") = private_Button_NormalizeShapeRef(buttonDict("ShapeRef"))
        hasQuery = private_Dict_HasText(buttonDict, "Query")
        hasQueryGroup = private_Dict_HasText(buttonDict, "QueryGroup")

        If hasQuery And hasQueryGroup Then
            Err.Raise VBA.vbObjectError + 4303, , "Button '" & buttonDict("Alias") & "' must reference either Query or QueryGroup, not both."
        End If

        If Not hasQuery And Not hasQueryGroup Then
            Err.Raise VBA.vbObjectError + 4304, , "Button '" & buttonDict("Alias") & "' must reference Query or QueryGroup."
        End If

        If hasQuery Then
            buttonDict("Query") = private_Dict_RequireText(buttonDict, "Query", "Missing Button." & buttonDict("Alias") & ".Query")
            If Not gQueriesDict.Exists(buttonDict("Query")) Then
                Err.Raise VBA.vbObjectError + 4301, , "Button '" & buttonDict("Alias") & "' references unknown query: " & buttonDict("Query")
            End If
            Set queryDict = gQueriesDict(buttonDict("Query"))
            private_Query_RequireStandaloneFields queryDict
        Else
            buttonDict("QueryGroup") = private_Dict_RequireText(buttonDict, "QueryGroup", "Missing Button." & buttonDict("Alias") & ".QueryGroup")
            If Not gQueryGroupsDict.Exists(buttonDict("QueryGroup")) Then
                Err.Raise VBA.vbObjectError + 4305, , "Button '" & buttonDict("Alias") & "' references unknown query group: " & buttonDict("QueryGroup")
            End If
        End If

        buttonKey = buttonDict("ShapeRef")
        If gButtonsDict.Exists(buttonKey) Then
            Err.Raise VBA.vbObjectError + 4302, , "Duplicate button binding for: " & buttonKey
        End If

        gButtonsDict.Add buttonKey, buttonDict
    Next buttonAlias
End Sub
' --------------------------------------
'  } // namespace Config
' --------------------------------------

' --------------------------------------
'  namespace Button {
' --------------------------------------
Private Function private_Button_ResolveCaller() As Object
    Dim callerValue As Variant
    Dim callerShapeName As String
    Dim buttonKey As String

    callerValue = Application.Caller
    If VBA.IsError(callerValue) Then
        Err.Raise VBA.vbObjectError + 4310, , "Button macro must be called by a worksheet shape."
    End If

    callerShapeName = VBA.Trim$(VBA.CStr(callerValue))
    If VBA.Len(callerShapeName) = 0 Then
        Err.Raise VBA.vbObjectError + 4310, , "Button macro must be called by a worksheet shape."
    End If

    buttonKey = private_Button_NormalizeShapeRef(ActiveSheet.Name & "!" & callerShapeName)
    If Not gButtonsDict.Exists(buttonKey) Then
        Err.Raise VBA.vbObjectError + 4311, , "Button binding not found: " & buttonKey
    End If

    Set private_Button_ResolveCaller = gButtonsDict(buttonKey)
End Function


Private Function private_Button_NormalizeShapeRef(ByVal shapeRef As String) As String
    Dim separatorPos As Long
    Dim sheetName As String
    Dim shapeName As String

    shapeRef = VBA.Trim$(shapeRef)
    separatorPos = VBA.InStr(1, shapeRef, "!")
    If separatorPos <= 1 Or separatorPos = VBA.Len(shapeRef) Then
        Err.Raise VBA.vbObjectError + 4313, , "Button ShapeRef must use SheetName!ShapeName format: " & shapeRef
    End If

    sheetName = VBA.Trim$(VBA.Left$(shapeRef, separatorPos - 1))
    shapeName = VBA.Trim$(VBA.Mid$(shapeRef, separatorPos + 1))
    If VBA.Len(sheetName) = 0 Or VBA.Len(shapeName) = 0 Then
        Err.Raise VBA.vbObjectError + 4313, , "Button ShapeRef must use SheetName!ShapeName format: " & shapeRef
    End If

    private_Button_NormalizeShapeRef = sheetName & "!" & shapeName
End Function
' --------------------------------------
'  } // namespace Button
' --------------------------------------

' --------------------------------------
'  namespace Source {
' --------------------------------------
Private Function private_Source_TryParseFilePathKey(ByVal key As String, ByRef outSourceAlias As String) As Boolean
    Dim remainder As String
    Dim separatorPos As Long
    Dim propName As String

    outSourceAlias = ""
    remainder = VBA.Mid$(key, VBA.Len("Source.") + 1)
    separatorPos = VBA.InStr(1, remainder, ".")
    If separatorPos <= 1 Then Exit Function

    outSourceAlias = VBA.Left$(remainder, separatorPos - 1)
    propName = VBA.Mid$(remainder, separatorPos + 1)
    private_Source_TryParseFilePathKey = (propName = "FilePath")
End Function


Private Function private_Source_TryParseTableRefKey(ByVal key As String, ByRef outSourceAlias As String, ByRef outTableAlias As String) As Boolean
    Dim remainder As String
    Dim separatorPos As Long
    Dim propName As String
    Dim tableStart As Long
    Dim tableEnd As Long

    outSourceAlias = ""
    outTableAlias = ""
    remainder = VBA.Mid$(key, VBA.Len("Source.") + 1)
    separatorPos = VBA.InStr(1, remainder, ".")
    If separatorPos <= 1 Then Exit Function

    outSourceAlias = VBA.Left$(remainder, separatorPos - 1)
    propName = VBA.Mid$(remainder, separatorPos + 1)
    If VBA.Left$(propName, VBA.Len("Tables[")) <> "Tables[" Then Exit Function

    tableStart = VBA.Len("Tables[") + 1
    tableEnd = VBA.InStr(tableStart, propName, "]")
    If tableEnd <= tableStart Then Exit Function
    If VBA.Mid$(propName, tableEnd + 1) <> ".Ref" Then Exit Function

    outTableAlias = VBA.Mid$(propName, tableStart, tableEnd - tableStart)
    private_Source_TryParseTableRefKey = True
End Function


Private Sub private_Source_SetFilePath(ByVal sourceAlias As String, ByVal filePath As String)
    Dim sourceDict As Object

    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then
        Err.Raise VBA.vbObjectError + 4110, , "Source FilePath is empty for source: " & sourceAlias
    End If

    Set sourceDict = private_Source_EnsureByAlias(sourceAlias)
    sourceDict("FilePath") = filePath
    If gSourcesDict.Exists(filePath) Then
        If Not gSourcesDict(filePath) Is sourceDict Then
            Err.Raise VBA.vbObjectError + 4111, , "Duplicate source FilePath: " & filePath
        End If
    Else
        gSourcesDict.Add filePath, sourceDict
    End If
End Sub


Private Sub private_Source_AddTableRef(ByVal sourceAlias As String, ByVal tableAlias As String, ByVal tableRef As String)
    Dim sourceDict As Object
    Dim tablesDict As Object

    tableRef = VBA.Trim$(tableRef)
    If VBA.Len(tableRef) = 0 Then
        Err.Raise VBA.vbObjectError + 4112, , "Source table Ref is empty: Source." & sourceAlias & ".Tables[" & tableAlias & "].Ref"
    End If

    Set sourceDict = private_Source_EnsureByAlias(sourceAlias)
    Set tablesDict = sourceDict("TablesDict")
    tablesDict(tableAlias) = tableRef
End Sub


Private Function private_Source_EnsureByAlias(ByVal sourceAlias As String) As Object
    Dim sourceDict As Object

    sourceAlias = VBA.Trim$(sourceAlias)
    If VBA.Len(sourceAlias) = 0 Then
        Err.Raise VBA.vbObjectError + 4113, , "Source alias is empty."
    End If

    If Not gSourceAliasesDict.Exists(sourceAlias) Then
        Set sourceDict = private_Dict_CreateTextMap()
        sourceDict("Alias") = sourceAlias
        Set sourceDict("TablesDict") = private_Dict_CreateTextMap()
        gSourceAliasesDict.Add sourceAlias, sourceDict
    End If

    Set private_Source_EnsureByAlias = gSourceAliasesDict(sourceAlias)
End Function


Private Sub private_Source_ResolveTableRef( _
    ByVal sourceTableRefKey As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outSourceDict As Object, _
    ByRef outTableRef As String _
)
    Dim tablesDict As Object

    If Not private_Source_TryParseTableRefKey(sourceTableRefKey, outSourceAlias, outTableAlias) Then
        Err.Raise VBA.vbObjectError + 4120, , "Query From must reference Source.<Alias>.Tables[<TableAlias>].Ref: " & sourceTableRefKey
    End If

    If Not gSourceAliasesDict.Exists(outSourceAlias) Then
        Err.Raise VBA.vbObjectError + 4121, , "Query From references unknown source: " & outSourceAlias
    End If

    Set outSourceDict = gSourceAliasesDict(outSourceAlias)
    Set tablesDict = outSourceDict("TablesDict")
    If Not tablesDict.Exists(outTableAlias) Then
        Err.Raise VBA.vbObjectError + 4122, , "Query From references unknown table alias '" & outTableAlias & "' for source '" & outSourceAlias & "'."
    End If

    outTableRef = tablesDict(outTableAlias)
End Sub
' --------------------------------------
'  } // namespace Source
' --------------------------------------

' --------------------------------------
'  namespace Sql {
' --------------------------------------
Private Function private_Sql_Build(ByVal queryDict As Object, ByVal tableRef As String) As String
    Dim selectPart As String
    Dim whereTemplate As String
    Dim wherePart As String
    Dim wherePartAlternative As String
    Dim queryAlias As String
    Dim sqlAlternative As String

    selectPart = private_Dict_GetValue(queryDict, "Select", "*")
    queryAlias = private_Dict_GetValue(queryDict, "Alias", "<unknown>")
    whereTemplate = private_Dict_GetValue(queryDict, "Where", "")
    wherePart = private_Sql_ResolveCellRefs(whereTemplate, queryAlias, wherePartAlternative)

#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogInfo "sql:where-template alias='" & VBA.Replace$(queryAlias, "'", "''") & "' where='" & VBA.Replace$(whereTemplate, "'", "''") & "'"
    fn_Diagnostic_LogInfo "sql:where-resolved alias='" & VBA.Replace$(queryAlias, "'", "''") & "' where='" & VBA.Replace$(wherePart, "'", "''") & "'"
#End If

    private_Sql_Build = "SELECT " & selectPart & " FROM [" & tableRef & "]"
    If VBA.Len(wherePart) > 0 Then
        private_Sql_Build = private_Sql_Build & " WHERE " & wherePart
    End If

    Debug.Print private_Sql_Build
#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogInfo "sql:built alias='" & VBA.Replace$(queryAlias, "'", "''") & "' sql='" & VBA.Replace$(private_Sql_Build, "'", "''") & "'"
#End If

    sqlAlternative = ""
    If VBA.Len(wherePartAlternative) > 0 And wherePartAlternative <> wherePart Then
        sqlAlternative = "SELECT " & selectPart & " FROM [" & tableRef & "]"
        sqlAlternative = sqlAlternative & " WHERE " & wherePartAlternative
#If LOGGING_DEBUG_ENABLED Then
        fn_Diagnostic_LogInfo "sql:where-resolved-alt alias='" & VBA.Replace$(queryAlias, "'", "''") & "' where='" & VBA.Replace$(wherePartAlternative, "'", "''") & "'"
        fn_Diagnostic_LogInfo "sql:built-alt alias='" & VBA.Replace$(queryAlias, "'", "''") & "' sql='" & VBA.Replace$(sqlAlternative, "'", "''") & "'"
#End If
    End If

    queryDict("SqlAlternative") = sqlAlternative
End Function


Private Function private_Sql_ResolveCellRefs(ByVal text As String, Optional ByVal queryAlias As String = "", Optional ByRef outAlternativeText As String = "") As String
    Dim p1 As Long
    Dim p2 As Long
    Dim p1Alternative As Long
    Dim p2Alternative As Long
    Dim token As String
    Dim tokenValue As String
    Dim escapedTokenValue As String
    Dim tokenLiteralText As String
    Dim tokenLiteralNumber As String
    Dim replacementValue As String
    Dim replacementAlternativeValue As String
    Dim hasAlternative As Boolean
    Dim hasLeftQuote As Boolean
    Dim hasRightQuote As Boolean
    Dim resolved As String
    Dim resolvedAlternative As String
    Dim normalizedAlias As String

    normalizedAlias = VBA.Trim$(VBA.CStr(queryAlias))
    If VBA.Len(normalizedAlias) = 0 Then normalizedAlias = "<unknown>"

    resolved = text
    resolvedAlternative = text
#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogInfo "sql:placeholders:start alias='" & VBA.Replace$(normalizedAlias, "'", "''") & "' text='" & VBA.Replace$(resolved, "'", "''") & "'"
#End If

    Do
        p1 = VBA.InStr(1, resolved, "{")
        If p1 = 0 Then Exit Do

        p2 = VBA.InStr(p1, resolved, "}")
        If p2 = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            fn_Diagnostic_LogWarning "sql:placeholders:unclosed-token alias='" & VBA.Replace$(normalizedAlias, "'", "''") & "' text='" & VBA.Replace$(resolved, "'", "''") & "'"
#End If
            Exit Do
        End If

        token = VBA.Mid$(resolved, p1 + 1, p2 - p1 - 1)
        tokenValue = private_CellRef_Read(token)
        escapedTokenValue = private_Sql_EscapeValue(tokenValue)
        tokenLiteralText = "'" & escapedTokenValue & "'"
        tokenLiteralNumber = escapedTokenValue
        replacementValue = escapedTokenValue
        replacementAlternativeValue = replacementValue

        hasLeftQuote = (p1 > 1 And VBA.Mid$(resolved, p1 - 1, 1) = "'")
        hasRightQuote = (p2 < VBA.Len(resolved) And VBA.Mid$(resolved, p2 + 1, 1) = "'")
        If VBA.IsNumeric(tokenValue) And Not (hasLeftQuote And hasRightQuote) Then
            replacementAlternativeValue = tokenLiteralText
            hasAlternative = True
    #If LOGGING_DEBUG_ENABLED Then
            fn_Diagnostic_LogInfo "sql:placeholder:alt-literal alias='" & VBA.Replace$(normalizedAlias, "'", "''") & "' token='" & VBA.Replace$(token, "'", "''") & "' alt='" & VBA.Replace$(replacementAlternativeValue, "'", "''") & "'"
    #End If
        End If

#If LOGGING_DEBUG_ENABLED Then
        fn_Diagnostic_LogInfo "sql:placeholder:resolved alias='" & VBA.Replace$(normalizedAlias, "'", "''") & "' token='" & VBA.Replace$(token, "'", "''") & "' value='" & VBA.Replace$(tokenValue, "'", "''") & "' escaped='" & VBA.Replace$(escapedTokenValue, "'", "''") & "' literal-text='" & VBA.Replace$(tokenLiteralText, "'", "''") & "' literal-number='" & VBA.Replace$(tokenLiteralNumber, "'", "''") & "'"
#End If

        resolved = VBA.Left$(resolved, p1 - 1) & _
               replacementValue & _
                   VBA.Mid$(resolved, p2 + 1)

        p1Alternative = VBA.InStr(1, resolvedAlternative, "{")
        p2Alternative = VBA.InStr(p1Alternative, resolvedAlternative, "}")
        resolvedAlternative = VBA.Left$(resolvedAlternative, p1Alternative - 1) & _
                      replacementAlternativeValue & _
                      VBA.Mid$(resolvedAlternative, p2Alternative + 1)
    Loop

#If LOGGING_DEBUG_ENABLED Then
    If VBA.InStr(1, resolved, "{") > 0 Or VBA.InStr(1, resolved, "}") > 0 Then
        fn_Diagnostic_LogWarning "sql:placeholders:residual-markers alias='" & VBA.Replace$(normalizedAlias, "'", "''") & "' text='" & VBA.Replace$(resolved, "'", "''") & "'"
    End If
    fn_Diagnostic_LogInfo "sql:placeholders:done alias='" & VBA.Replace$(normalizedAlias, "'", "''") & "' text='" & VBA.Replace$(resolved, "'", "''") & "'"
    If hasAlternative Then
        fn_Diagnostic_LogInfo "sql:placeholders:done-alt alias='" & VBA.Replace$(normalizedAlias, "'", "''") & "' text='" & VBA.Replace$(resolvedAlternative, "'", "''") & "'"
    End If
#End If

    If hasAlternative Then
        outAlternativeText = resolvedAlternative
    Else
        outAlternativeText = ""
    End If

    private_Sql_ResolveCellRefs = resolved
End Function


Private Function private_Sql_EscapeValue(ByVal value As String) As String
    private_Sql_EscapeValue = VBA.Replace(value, "'", "''")
End Function
' --------------------------------------
'  } // namespace Sql
' --------------------------------------

' --------------------------------------
'  namespace CellRef {
' --------------------------------------
Private Function private_CellRef_Read(ByVal refText As String) As String
    Dim p As Long
    Dim sheetName As String
    Dim cellAddress As String
    Dim rawValue As Variant
    Dim rawTypeName As String
    Dim isNumericValue As Boolean

    refText = VBA.Trim$(VBA.CStr(refText))

#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogInfo "cellref:read:start ref='" & VBA.Replace$(refText, "'", "''") & "'"
#End If

    p = VBA.InStr(1, refText, "!")
    If p = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        fn_Diagnostic_LogError "cellref:invalid-format ref='" & VBA.Replace$(refText, "'", "''") & "'"
#End If
        Err.Raise VBA.vbObjectError + 2000, , "Invalid cell reference: " & refText
    End If

    sheetName = VBA.Trim$(VBA.Left$(refText, p - 1))
    cellAddress = VBA.Trim$(VBA.Mid$(refText, p + 1))

#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogInfo "cellref:read:resolved ref='" & VBA.Replace$(refText, "'", "''") & "' sheet='" & VBA.Replace$(sheetName, "'", "''") & "' cell='" & VBA.Replace$(cellAddress, "'", "''") & "'"
#End If

    On Error GoTo ErrHandler
    rawValue = ThisWorkbook.Worksheets(sheetName).Range(cellAddress).Value
    rawTypeName = VBA.TypeName(rawValue)
    isNumericValue = VBA.IsNumeric(rawValue)
    private_CellRef_Read = VBA.Trim$(private_Text_NullToEmptyString(rawValue))

#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogInfo "cellref:read:value ref='" & VBA.Replace$(refText, "'", "''") & "' value='" & VBA.Replace$(private_CellRef_Read, "'", "''") & "' raw-type='" & VBA.Replace$(rawTypeName, "'", "''") & "' is-numeric='" & VBA.LCase$(VBA.CStr(isNumericValue)) & "'"
#End If
    Exit Function

ErrHandler:
#If LOGGING_DEBUG_ENABLED Then
    fn_Diagnostic_LogError "cellref:read-failed ref='" & VBA.Replace$(refText, "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Function private_CellRef_LooksLikeRef(ByVal refText As String) As Boolean
    Dim p As Long
    Dim sheetName As String
    Dim cellAddress As String

    p = VBA.InStr(1, refText, "!")
    If p <= 1 Or p = VBA.Len(refText) Then Exit Function

    sheetName = VBA.Trim$(VBA.Left$(refText, p - 1))
    cellAddress = VBA.Trim$(VBA.Mid$(refText, p + 1))
    private_CellRef_LooksLikeRef = (VBA.Len(sheetName) > 0 And VBA.Len(cellAddress) > 0)
End Function
' --------------------------------------
'  } // namespace CellRef
' --------------------------------------

' --------------------------------------
'  namespace Output {
' --------------------------------------
Private Sub private_Output_WriteResult(ByVal outputRef As String, ByVal value As String)
    Dim p As Long
    Dim sheetName As String
    Dim cellAddress As String

    p = VBA.InStr(1, outputRef, "!")
    If p = 0 Then
        Err.Raise VBA.vbObjectError + 2001, , "Invalid output reference: " & outputRef
    End If

    sheetName = VBA.Left$(outputRef, p - 1)
    cellAddress = VBA.Mid$(outputRef, p + 1)

    With ThisWorkbook.Worksheets(sheetName).Range(cellAddress).MergeArea
        .ClearContents
        .Cells(1, 1).Value = value
        .WrapText = True
    End With
End Sub
' --------------------------------------
'  } // namespace Output
' --------------------------------------

' --------------------------------------
'  namespace Diagnostic {
' --------------------------------------
Public Sub fn_Diagnostic_LogInfo(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogEvent VBA.CStr(messageText)
#End If
End Sub


Public Sub fn_Diagnostic_LogError(ByVal messageText As String)
    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogEvent "error: " & messageText
#End If
End Sub


Public Sub fn_Diagnostic_LogWarning(ByVal messageText As String)
    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogEvent "warning: " & messageText
#End If
End Sub


Public Sub fn_Diagnostic_ClearLog()
    private_Diagnostic_ClearLogFile
End Sub


Private Sub private_Diagnostic_LogEvent(ByVal messageText As String)
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object
    Dim lineText As String

    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
    If VBA.Len(VBA.Trim$(ThisWorkbook.Path)) = 0 Then Exit Sub

    logPath = ThisWorkbook.Path & "\\" & DIAGNOSTIC_LOG_FILE_REL_PATH
    folderPath = VBA.Left$(logPath, VBA.InStrRev(logPath, "\\") - 1)

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If VBA.Len(folderPath) > 0 Then
            If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
        End If
    End If

    lineText = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    Set stream = fso.OpenTextFile(logPath, 8, True)
    If Not stream Is Nothing Then
        stream.WriteLine lineText
        stream.Close
    End If
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub private_Diagnostic_ClearLogFile()
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object

    If VBA.Len(VBA.Trim$(ThisWorkbook.Path)) = 0 Then Exit Sub

    logPath = ThisWorkbook.Path & "\\" & DIAGNOSTIC_LOG_FILE_REL_PATH
    folderPath = VBA.Left$(logPath, VBA.InStrRev(logPath, "\\") - 1)

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If VBA.Len(folderPath) > 0 Then
            If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
        End If
    End If

    Set stream = fso.OpenTextFile(logPath, 2, True)
    If Not stream Is Nothing Then stream.Close
    Err.Clear
    On Error GoTo 0
End Sub
' --------------------------------------
'  } // namespace Diagnostic
' --------------------------------------

' --------------------------------------
'  namespace Handler {
' --------------------------------------
Private Function private_Handler_BuildStateDict(ByVal stateRs As Object) As Object
    Dim stateDict As Object
    Dim rowDict As Object
    Dim taxId As String
    Dim dativeName As String
    Dim genitiveName As String
    Dim genitiveShortName As String

    If stateRs.EOF Then
        Err.Raise VBA.vbObjectError + 5110, , "State query returned no rows."
    End If

    Set stateDict = private_Dict_CreateTextMap()
    stateRs.MoveFirst

    Do Until stateRs.EOF
        taxId = private_Recordset_FieldTextOrIndex(stateRs, "TaxId", 0)
        dativeName = private_Recordset_FieldTextOrIndex(stateRs, "DativeName", 1)
        If private_Recordset_HasField(stateRs, "GenitiveShortName") Then
            genitiveShortName = private_Recordset_FieldText(stateRs, "GenitiveShortName")
        Else
            genitiveName = private_Recordset_FieldTextOrIndex(stateRs, "GenitiveName", 2)
            genitiveShortName = private_Text_ToShortName(genitiveName)
        End If

        If VBA.Len(taxId) > 0 And VBA.Len(dativeName) > 0 And VBA.Len(genitiveShortName) > 0 Then
            Set rowDict = private_Dict_CreateTextMap()
            rowDict("DativeName") = dativeName
            rowDict("GenitiveShortName") = genitiveShortName
            If stateDict.Exists(taxId) Then stateDict.Remove taxId
            stateDict.Add taxId, rowDict
        End If
        stateRs.MoveNext
    Loop

    If stateDict.Count = 0 Then
        Err.Raise VBA.vbObjectError + 5111, , "State query returned no usable TaxId/DativeName/GenitiveName rows."
    End If

    Set private_Handler_BuildStateDict = stateDict
End Function


Private Function private_Handler_BuildRankDict(ByVal rankRs As Object) As Object
    Dim rankDict As Object
    Dim rowDict As Object
    Dim rankText As String
    Dim rankGenitive As String
    Dim rankDative As String

    If rankRs.EOF Then
        Err.Raise VBA.vbObjectError + 5120, , "RankQuery returned no rows."
    End If

    Set rankDict = private_Dict_CreateTextMap()
    rankRs.MoveFirst

    Do Until rankRs.EOF
        rankText = private_Text_NormalizeLookupToken(private_Recordset_FieldTextOrIndex(rankRs, "Rank", 0))
        rankGenitive = private_Recordset_FieldTextOrIndex(rankRs, "RankGenitive", 1)
        rankDative = private_Recordset_FieldTextOrIndex(rankRs, "RankDative", 2)
        If VBA.Len(rankText) > 0 And VBA.Len(rankGenitive) > 0 And VBA.Len(rankDative) > 0 Then
            Set rowDict = private_Dict_CreateTextMap()
            rowDict("RankGenitive") = rankGenitive
            rowDict("RankDative") = rankDative
            If rankDict.Exists(rankText) Then rankDict.Remove rankText
            rankDict.Add rankText, rowDict
        End If
        rankRs.MoveNext
    Loop

    If rankDict.Count = 0 Then
        Err.Raise VBA.vbObjectError + 5121, , "RankQuery returned no usable Rank/RankGenitive/RankDative rows."
    End If

    Set private_Handler_BuildRankDict = rankDict
End Function


Private Function private_Handler_BuildPositionDict(ByVal positionRs As Object) As Object
    Dim positionDict As Object
    Dim positionCode As String
    Dim positionDative As String

    If positionRs.EOF Then
        Err.Raise VBA.vbObjectError + 5130, , "PositionQuery returned no rows."
    End If

    Set positionDict = private_Dict_CreateTextMap()
    positionRs.MoveFirst

    Do Until positionRs.EOF
        positionCode = private_Recordset_FieldTextOrIndex(positionRs, "PositionCode", 0)
        positionDative = private_Recordset_FieldTextOrIndex(positionRs, "PositionDative", 1)
        If VBA.Len(positionCode) > 0 And VBA.Len(positionDative) > 0 Then
            positionDict(positionCode) = positionDative
        End If
        positionRs.MoveNext
    Loop

    If positionDict.Count = 0 Then
        Err.Raise VBA.vbObjectError + 5131, , "PositionQuery returned no usable PositionCode/PositionDative rows."
    End If

    Set private_Handler_BuildPositionDict = positionDict
End Function


Private Function private_Handler_BuildGdoOrderItem(ByVal paymentsRs As Object, ByVal stateRowDict As Object, ByVal rankRowDict As Object, ByVal positionDative As String, ByVal taxId As String, ByVal includeTaxId As Boolean) As String
    Dim reportText As String
    Dim reportDateText As String
    Dim orderText As String

    reportText = private_Handler_RequireText(private_Recordset_FieldTextOrIndex(paymentsRs, "Report", 5), "Payments row has empty Report.")
    reportDateText = private_Handler_RequireText(private_Text_FormatDisplayDate(private_Recordset_FieldValueOrIndex(paymentsRs, "ReportDate", 6)), "Payments row has empty ReportDate.")

    orderText = private_Text_CapitalizeFirst(rankRowDict("RankDative")) & " " & _
                stateRowDict("DativeName") & ", "

    If includeTaxId Then
        orderText = orderText & taxId & ", "
    End If

    orderText = orderText & private_Text_LowerFirst(positionDative) & "."

    orderText = orderText & _
                VBA.vbCrLf & _
                private_Handler_BuildGdoBasisText(rankRowDict("RankGenitive"), stateRowDict("GenitiveShortName"), reportText, reportDateText)

    private_Handler_BuildGdoOrderItem = orderText
End Function


Private Function private_HandlerArgs_GetBoolean(ByVal handlerArgsDict As Object, ByVal argName As String, ByVal defaultValue As Boolean) As Boolean
    Dim valueText As String

    If handlerArgsDict Is Nothing Then
        private_HandlerArgs_GetBoolean = defaultValue
        Exit Function
    End If

    If Not handlerArgsDict.Exists(argName) Then
        private_HandlerArgs_GetBoolean = defaultValue
        Exit Function
    End If

    valueText = VBA.LCase$(VBA.Trim$(VBA.CStr(handlerArgsDict(argName))))
    If VBA.Len(valueText) = 0 Then
        private_HandlerArgs_GetBoolean = defaultValue
        Exit Function
    End If

    Select Case valueText
        Case "true", "1", "yes", "y", "on", "да", "так"
            private_HandlerArgs_GetBoolean = True
        Case "false", "0", "no", "n", "off", "нет", "ні"
            private_HandlerArgs_GetBoolean = False
        Case Else
            Err.Raise VBA.vbObjectError + 5150, , "Handler arg '" & argName & "' must be boolean, got: " & handlerArgsDict(argName)
    End Select
End Function


Private Function private_Handler_BuildGdoBasisText(ByVal rankGenitive As String, ByVal genitiveShortName As String, ByVal reportText As String, ByVal reportDateText As String) As String
    private_Handler_BuildGdoBasisText = private_Text_UaPidstavaRaport() & _
                                        private_Text_LowerFirst(rankGenitive) & " " & _
                                        genitiveShortName & " (" & _
                                        private_Text_UaIncomingNo() & reportText & _
                                        private_Text_UaFrom() & reportDateText & ")."
End Function


Private Function private_Handler_RequireText(ByVal value As String, ByVal errDescription As String) As String
    private_Handler_RequireText = VBA.Trim$(value)
    If VBA.Len(private_Handler_RequireText) = 0 Then
        Err.Raise VBA.vbObjectError + 5140, , errDescription
    End If
End Function


Private Function private_Handler_ResolveName(ByVal handlerName As String) As String
    Dim separatorPos As Long
    Dim modulePrefix As String
    Dim procedureName As String

    handlerName = VBA.Trim$(handlerName)
    separatorPos = VBA.InStrRev(handlerName, ".")
    If separatorPos > 0 Then
        modulePrefix = VBA.Left$(handlerName, separatorPos)
        procedureName = VBA.Mid$(handlerName, separatorPos + 1)
    Else
        procedureName = handlerName
    End If

    If VBA.Left$(procedureName, VBA.Len("fn_Handler_")) = "fn_Handler_" Then
        private_Handler_ResolveName = modulePrefix & procedureName
        Exit Function
    End If

    If VBA.Left$(procedureName, VBA.Len("fn_Handle_")) = "fn_Handle_" Then
        procedureName = "fn_Handler_" & VBA.Mid$(procedureName, VBA.Len("fn_Handle_") + 1)
    ElseIf VBA.Left$(procedureName, VBA.Len("Handler_")) = "Handler_" Then
        procedureName = "fn_" & procedureName
    ElseIf VBA.Left$(procedureName, VBA.Len("Handle_")) = "Handle_" Then
        procedureName = "fn_Handler_" & VBA.Mid$(procedureName, VBA.Len("Handle_") + 1)
    ElseIf VBA.Left$(procedureName, 3) <> "fn_" Then
        procedureName = "fn_Handler_" & procedureName
    End If

    private_Handler_ResolveName = modulePrefix & procedureName
End Function
' --------------------------------------
'  } // namespace Handler
' --------------------------------------

' --------------------------------------
'  namespace Recordset {
' --------------------------------------
Private Sub private_Recordset_RequireFieldCount(ByVal rs As Object, ByVal recordsetAlias As String, ByVal minFieldCount As Long)
    If rs.Fields.Count < minFieldCount Then
        Err.Raise VBA.vbObjectError + 5201, , "Recordset '" & recordsetAlias & "' must return at least " & minFieldCount & " fields, but returned " & rs.Fields.Count & "."
    End If
End Sub


Private Sub private_Recordset_RequireFields(ByVal rs As Object, ByVal recordsetAlias As String, ByVal fieldsList As Variant)
    Dim i As Long

    For i = LBound(fieldsList) To UBound(fieldsList)
        private_Recordset_RequireField rs, recordsetAlias, VBA.CStr(fieldsList(i))
    Next i
End Sub


Private Sub private_Recordset_RequireField(ByVal rs As Object, ByVal recordsetAlias As String, ByVal fieldName As String)
    Dim fieldObj As Object

    On Error Resume Next
    Err.Clear
    Set fieldObj = rs.Fields(fieldName)
    If Err.Number <> 0 Or fieldObj Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Err.Raise VBA.vbObjectError + 5200, , "Required field '" & fieldName & "' not found in recordset: " & recordsetAlias
    End If
    On Error GoTo 0
End Sub


Private Function private_Recordset_FieldText(ByVal rs As Object, ByVal fieldName As String) As String
    private_Recordset_FieldText = VBA.Trim$(private_Text_NullToEmptyString(rs.Fields(fieldName).Value))
End Function


Private Function private_Recordset_FieldTextOrIndex(ByVal rs As Object, ByVal fieldName As String, ByVal fieldIndex As Long) As String
    If private_Recordset_HasField(rs, fieldName) Then
        private_Recordset_FieldTextOrIndex = private_Recordset_FieldText(rs, fieldName)
    Else
        private_Recordset_FieldTextOrIndex = VBA.Trim$(private_Text_NullToEmptyString(rs.Fields(fieldIndex).Value))
    End If
End Function


Private Function private_Recordset_FieldValueOrIndex(ByVal rs As Object, ByVal fieldName As String, ByVal fieldIndex As Long) As Variant
    If private_Recordset_HasField(rs, fieldName) Then
        private_Recordset_FieldValueOrIndex = rs.Fields(fieldName).Value
    Else
        private_Recordset_FieldValueOrIndex = rs.Fields(fieldIndex).Value
    End If
End Function


Private Function private_Recordset_HasField(ByVal rs As Object, ByVal fieldName As String) As Boolean
    Dim fieldObj As Object

    On Error Resume Next
    Err.Clear
    Set fieldObj = rs.Fields(fieldName)
    private_Recordset_HasField = (Err.Number = 0 And Not fieldObj Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function
' --------------------------------------
'  } // namespace Recordset
' --------------------------------------

' --------------------------------------
'  namespace Dict {
' --------------------------------------
Private Function private_Dict_CreateTextMap() As Object
    Set private_Dict_CreateTextMap = VBA.CreateObject("Scripting.Dictionary")
    private_Dict_CreateTextMap.CompareMode = VBA.vbTextCompare
End Function


Private Function private_Dict_Require(ByVal targetDict As Object, ByVal key As String, ByVal errPrefix As String) As Object
    If Not targetDict.Exists(key) Then
        Err.Raise VBA.vbObjectError + 3000, , errPrefix & ": " & key
    End If
    Set private_Dict_Require = targetDict(key)
End Function


Private Function private_Dict_RequireText(ByVal targetDict As Object, ByVal key As String, ByVal errPrefix As String) As String
    If Not targetDict.Exists(key) Then
        Err.Raise VBA.vbObjectError + 3001, , errPrefix
    End If

    private_Dict_RequireText = VBA.Trim$(VBA.CStr(targetDict(key)))
    If VBA.Len(private_Dict_RequireText) = 0 Then
        Err.Raise VBA.vbObjectError + 3002, , errPrefix
    End If
End Function


Private Function private_Dict_GetValue(ByVal targetDict As Object, ByVal key As String, ByVal defaultValue As String) As String
    If targetDict.Exists(key) Then
        private_Dict_GetValue = targetDict(key)
    Else
        private_Dict_GetValue = defaultValue
    End If
End Function


Private Function private_Dict_HasText(ByVal targetDict As Object, ByVal key As String) As Boolean
    If targetDict.Exists(key) Then
        private_Dict_HasText = VBA.Len(VBA.Trim$(VBA.CStr(targetDict(key)))) > 0
    End If
End Function
' --------------------------------------
'  } // namespace Dict
' --------------------------------------

' --------------------------------------
'  namespace Text {
' --------------------------------------
' Преобразует Null/Empty значения из SQL recordset в пустую строку,
' чтобы VBA.CStr не падал на Null при сборке текста результата.
Private Function private_Text_NullToEmptyString(ByVal value As Variant) As String
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        private_Text_NullToEmptyString = ""
    Else
        private_Text_NullToEmptyString = VBA.CStr(value)
    End If
End Function


Private Function private_Text_FormatDisplayDate(ByVal value As Variant) As String
    Dim textValue As String

    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function

    If VBA.IsDate(value) Then
        private_Text_FormatDisplayDate = VBA.Format$(VBA.CDate(value), "dd.mm.yyyy")
        Exit Function
    End If

    textValue = VBA.Trim$(VBA.CStr(value))
    If VBA.Len(textValue) = 0 Then Exit Function

    If VBA.IsNumeric(textValue) And VBA.CDbl(textValue) >= 20000 And VBA.CDbl(textValue) <= 60000 Then
        private_Text_FormatDisplayDate = VBA.Format$(VBA.CDate(VBA.CDbl(textValue)), "dd.mm.yyyy")
    Else
        private_Text_FormatDisplayDate = textValue
    End If
End Function


Private Function private_Text_CapitalizeFirst(ByVal text As String) As String
    text = VBA.Trim$(text)
    If VBA.Len(text) = 0 Then Exit Function

    private_Text_CapitalizeFirst = VBA.UCase$(VBA.Left$(text, 1)) & VBA.Mid$(text, 2)
End Function


Private Function private_Text_LowerFirst(ByVal text As String) As String
    text = VBA.Trim$(text)
    If VBA.Len(text) = 0 Then Exit Function

    private_Text_LowerFirst = VBA.LCase$(VBA.Left$(text, 1)) & VBA.Mid$(text, 2)
End Function


Private Function private_Text_ToShortName(ByVal fullName As String) As String
    Dim parts As Variant
    Dim firstInitial As String
    Dim patronymicInitial As String

    fullName = private_Text_NormalizeSpaces(fullName)
    If VBA.Len(fullName) = 0 Then Exit Function

    parts = VBA.Split(fullName, " ")
    If UBound(parts) < 2 Then
        private_Text_ToShortName = fullName
        Exit Function
    End If

    firstInitial = VBA.UCase$(VBA.Left$(VBA.CStr(parts(1)), 1))
    patronymicInitial = VBA.UCase$(VBA.Left$(VBA.CStr(parts(2)), 1))
    private_Text_ToShortName = private_Text_CapitalizeWord(VBA.CStr(parts(0))) & " " & firstInitial & "." & patronymicInitial & "."
End Function


Private Function private_Text_NormalizeSpaces(ByVal text As String) As String
    text = VBA.Trim$(text)
    Do While VBA.InStr(1, text, "  ") > 0
        text = VBA.Replace(text, "  ", " ")
    Loop
    private_Text_NormalizeSpaces = text
End Function


Private Function private_Text_NormalizeLookupToken(ByVal text As String) As String
    text = private_Text_NullToEmptyString(text)
    text = VBA.Trim$(text)
    text = VBA.Replace(text, VBA.ChrW$(160), " ")
    text = VBA.Replace(text, VBA.ChrW$(8239), " ")
    text = VBA.Replace(text, VBA.ChrW$(173), "")
    text = VBA.Replace(text, VBA.ChrW$(8208), "-")
    text = VBA.Replace(text, VBA.ChrW$(8209), "-")
    text = VBA.Replace(text, VBA.ChrW$(8210), "-")
    text = VBA.Replace(text, VBA.ChrW$(8211), "-")
    text = VBA.Replace(text, VBA.ChrW$(8212), "-")
    text = VBA.Replace(text, VBA.ChrW$(8722), "-")
    text = VBA.Replace(text, VBA.ChrW$(8217), "'")
    text = VBA.Replace(text, VBA.ChrW$(96), "'")
    text = private_Text_NormalizeSpaces(text)
    private_Text_NormalizeLookupToken = text
End Function


Private Function private_Text_CapitalizeWord(ByVal text As String) As String
    text = VBA.Trim$(text)
    If VBA.Len(text) = 0 Then Exit Function

    private_Text_CapitalizeWord = VBA.UCase$(VBA.Left$(text, 1)) & VBA.LCase$(VBA.Mid$(text, 2))
End Function


Private Function private_Text_UaPidstavaRaport() As String
    private_Text_UaPidstavaRaport = VBA.ChrW$(1055) & VBA.ChrW$(1110) & VBA.ChrW$(1076) & VBA.ChrW$(1089) & VBA.ChrW$(1090) & VBA.ChrW$(1072) & VBA.ChrW$(1074) & VBA.ChrW$(1072) & ": " & VBA.ChrW$(1088) & VBA.ChrW$(1072) & VBA.ChrW$(1087) & VBA.ChrW$(1086) & VBA.ChrW$(1088) & VBA.ChrW$(1090) & " "
End Function


Private Function private_Text_UaIncomingNo() As String
    private_Text_UaIncomingNo = VBA.ChrW$(1074) & VBA.ChrW$(1093) & ". " & VBA.ChrW$(8470) & " "
End Function


Private Function private_Text_UaFrom() As String
    private_Text_UaFrom = " " & VBA.ChrW$(1074) & VBA.ChrW$(1110) & VBA.ChrW$(1076) & " "
End Function
' --------------------------------------
'  } // namespace Text
' --------------------------------------
