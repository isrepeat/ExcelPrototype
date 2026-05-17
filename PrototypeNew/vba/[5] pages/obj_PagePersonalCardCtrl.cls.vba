VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PagePersonalCardCtrl"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PagePersonalCard.Controller"
Private Const SQL_TABLES_RUNTIME_KEY As String = "RuntimeItems.PersonalCard.Tables"
Private Const DEMO_TABLES_RUNTIME_KEY As String = "RuntimeItems.Test.Tables"
Private Const TABLE_HEADER_TEXT As String = "Name | Role | Country | Team | Level | Status | Since"

Private m_Page As obj_IPage
Private m_ConfigTable As obj_ConfigTable
' Парсер конфига вынесен в отдельную зависимость:
' контроллер оркестрирует сценарий, парсер извлекает SQL-параметры из конфига.
Private m_CfgPersonalCardParser As obj_CfgPersonalCardParser
Private m_IsDataReady As Boolean
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Property Get IsDataReady() As Boolean
    IsDataReady = m_IsDataReady
End Property

Public Property Get HasConfigTable() As Boolean
    HasConfigTable = Not m_ConfigTable Is Nothing
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal page As obj_IPage, _
    Optional ByVal configTable As obj_ConfigTable = Nothing _
) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePersonalCardCtrl.Initialize"
#End If
    Dim cfgParser As obj_CfgPersonalCardParser
    Dim pageBase As obj_PageBase

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PagePersonalCardCtrl initialization failed because page is not specified."
#End If
        Exit Function
    End If

    m_IsDisposed = False
    Set m_Page = page
    Set m_ConfigTable = configTable
    ' Подключаем персональный парсер к текущей модели конфига в памяти.
    Set cfgParser = New obj_CfgPersonalCardParser
    If Not cfgParser.Initialize(m_ConfigTable) Then Exit Function
    Set m_CfgPersonalCardParser = cfgParser
    m_IsDataReady = False
    Set pageBase = m_Page.GetPageBase()

    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePersonalCardCtrl.Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    m_IsDataReady = False
    ' Явно закрываем жизненный цикл парсера.
    If Not m_CfgPersonalCardParser Is Nothing Then m_CfgPersonalCardParser.Dispose
    Set m_CfgPersonalCardParser = Nothing
    Set m_ConfigTable = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Public Function RunPipeline( _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePersonalCardCtrl.RunPipeline"
#End If
    If m_Page Is Nothing Then Exit Function

    If Not PrepareSqlTablesRuntime(False) Then Exit Function
    If notifyChange Then
        If Not rt_PageManager.fn_RenderPage(m_Page, "personalcard:run-pipeline") Then Exit Function
    End If

    If m_IsDataReady Then
        rt_Messaging.fn_ShowStatusBarSuccess _
            "PersonalCard pipeline has been executed.", _
            4
    Else
        rt_Messaging.fn_ShowStatusBarWarning _
            "PersonalCard pipeline executed, but no rows were found for the current key.", _
            4
    End If
    RunPipeline = True
End Function

Public Function RerenderPage( _
    Optional ByVal notifyStatus As Boolean = True _
) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePersonalCardCtrl.RerenderPage"
#End If
    If m_Page Is Nothing Then Exit Function
    If Not rt_PageManager.fn_RenderPage(m_Page, "personalcard:manual-rerender") Then Exit Function

    If notifyStatus Then
        rt_Messaging.fn_ShowStatusBarNotice "PersonalCard page has been rerendered.", 2
    End If
    RerenderPage = True
End Function

Public Function PrepareDemoTablesRuntime( _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePersonalCardCtrl.PrepareDemoTablesRuntime"
#End If
    If Not private_RegisterDemoTableItems(notifyChange) Then
        m_IsDataReady = False
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to register demo table items for PersonalCard page."
#End If
        MsgBox "PrototypeNew: failed to register demo table items for PersonalCard page.", vbExclamation, "PrototypeNew / PersonalCard runtime"
        Exit Function
    End If

    PrepareDemoTablesRuntime = True
End Function

Public Function PrepareSqlTablesRuntime( _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePersonalCardCtrl.PrepareSqlTablesRuntime"
#End If
    If Not private_RegisterSqlTableItems(notifyChange) Then
        m_IsDataReady = False
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to register SQL table items for PersonalCard page."
#End If
        Exit Function
    End If

    PrepareSqlTablesRuntime = True
End Function

' //
' // Internal
' //
Private Function private_RegisterDemoTableItems( _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim tables As Collection
    Dim normalizedKey As String

    Set tables = private_BuildDemoTableItems()
    If tables Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function
    normalizedKey = VBA.LCase$(VBA.Trim$(DEMO_TABLES_RUNTIME_KEY))

    m_IsDataReady = False

    ' Точечная очистка:
    ' - удаляем только тестовые таблицы;
    ' - чистим только временные runtime-ключи, которые создают layout-рендереры.
    If Not runtimeSources.RemoveItemsSource(normalizedKey) Then Exit Function
    If Not runtimeSources.RemoveTemporaryItemsSources() Then Exit Function
    If Not runtimeSources.RemoveTemporaryObjectsSources() Then Exit Function

    If Not runtimeSources.SetItemsSource(normalizedKey, tables, notifyChange) Then Exit Function
    m_IsDataReady = (tables.Count > 0)

    private_RegisterDemoTableItems = True
End Function

Private Function private_BuildDemoTableItems() As Collection
    Dim result As Collection
    Dim tableIndex As Long
    Dim rowsCount As Long
    Dim teamName As String

    Set result = New Collection

    For tableIndex = 1 To 20
        rowsCount = 3 + ((tableIndex - 1) Mod 3)
        teamName = "Team " & VBA.Format$(tableIndex, "00")

        result.Add private_CreateDemoTable( _
            "People / " & teamName, _
            TABLE_HEADER_TEXT, _
            private_CreateGeneratedRowsForTeam(tableIndex, rowsCount, teamName))
    Next tableIndex

    Set private_BuildDemoTableItems = result
End Function

Private Function private_CreateGeneratedRowsForTeam( _
    ByVal tableIndex As Long, _
    ByVal rowsCount As Long, _
    ByVal teamName As String _
) As Collection
    Dim result As Collection
    Dim rowIndex As Long
    Dim personName As String
    Dim roleName As String
    Dim countryName As String
    Dim levelName As String
    Dim statusName As String
    Dim sinceYear As String

    Set result = New Collection

    For rowIndex = 1 To rowsCount
        personName = "Person " & VBA.Format$(tableIndex, "00") & "-" & VBA.CStr(rowIndex)
        roleName = private_GetRoleByIndex(tableIndex + rowIndex)
        countryName = private_GetCountryByIndex(tableIndex + rowIndex)
        levelName = "L" & VBA.CStr(((tableIndex + rowIndex) Mod 4) + 1)

        If (tableIndex + rowIndex) Mod 5 = 0 Then
            statusName = "On Hold"
        Else
            statusName = "Active"
        End If

        sinceYear = VBA.CStr(2014 + ((tableIndex + rowIndex) Mod 11))

        result.Add private_CreateDemoRowModel( _
            personName, _
            roleName, _
            countryName, _
            teamName, _
            levelName, _
            statusName, _
            sinceYear)
    Next rowIndex

    Set private_CreateGeneratedRowsForTeam = result
End Function

Private Function private_GetRoleByIndex(ByVal idx As Long) As String
    Select Case ((idx - 1) Mod 7) + 1
        Case 1: private_GetRoleByIndex = "Team Lead"
        Case 2: private_GetRoleByIndex = "Analyst"
        Case 3: private_GetRoleByIndex = "Developer"
        Case 4: private_GetRoleByIndex = "QA"
        Case 5: private_GetRoleByIndex = "DevOps"
        Case 6: private_GetRoleByIndex = "Support"
        Case Else: private_GetRoleByIndex = "Manager"
    End Select
End Function

Private Function private_GetCountryByIndex(ByVal idx As Long) As String
    Select Case ((idx - 1) Mod 6) + 1
        Case 1: private_GetCountryByIndex = "Ukraine"
        Case 2: private_GetCountryByIndex = "Poland"
        Case 3: private_GetCountryByIndex = "Romania"
        Case 4: private_GetCountryByIndex = "Germany"
        Case 5: private_GetCountryByIndex = "Czechia"
        Case Else: private_GetCountryByIndex = "Slovakia"
    End Select
End Function

Private Function private_CreateDemoTable( _
    ByVal sectionTitle As String, _
    ByVal headerText As String, _
    ByVal rows As Collection _
) As obj_TableDynamic
    Dim tableObj As obj_TableDynamic
    Dim rowObj As obj_Row
    Dim colObj As obj_Column
    Dim headerTokens As Variant
    Dim colIndex As Long

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = VBA.CStr(sectionTitle)

    headerTokens = VBA.Split(VBA.CStr(headerText), "|")
    For colIndex = LBound(headerTokens) To UBound(headerTokens)
        Set colObj = New obj_Column
        colObj.Position = colIndex + 1
        colObj.Name = VBA.Trim$(VBA.CStr(headerTokens(colIndex)))
        If VBA.Len(colObj.Name) = 0 Then colObj.Name = "Col" & VBA.CStr(colObj.Position)
        If Not tableObj.AddColumn(colObj) Then Exit Function
    Next colIndex

    If rows Is Nothing Then
        Set private_CreateDemoTable = tableObj
        Exit Function
    End If

    For Each rowObj In rows
        If rowObj Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: table row is not specified."
#End If
            Exit Function
        End If
        If rowObj.CellCount < tableObj.ColumnCount Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: table row has fewer cells than table columns."
#End If
            Exit Function
        End If

        If Not tableObj.AddRow(rowObj) Then Exit Function
    Next rowObj

    Set private_CreateDemoTable = tableObj
End Function

Private Function private_CreateDemoRowModel( _
    ByVal c1 As String, _
    ByVal c2 As String, _
    ByVal c3 As String, _
    ByVal c4 As String, _
    ByVal c5 As String, _
    ByVal c6 As String, _
    ByVal c7 As String _
) As obj_Row
    Dim rowObj As obj_Row

    Set rowObj = New obj_Row
    rowObj.AddCell c1
    rowObj.AddCell c2
    rowObj.AddCell c3
    rowObj.AddCell c4
    rowObj.AddCell c5
    rowObj.AddCell c6
    rowObj.AddCell c7

    Set private_CreateDemoRowModel = rowObj
End Function

Private Function private_RegisterSqlTableItems( _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim tables As Collection
    Dim normalizedKey As String

    Set tables = private_BuildTablesFromConfigTable()
    If tables Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function
    normalizedKey = VBA.LCase$(VBA.Trim$(SQL_TABLES_RUNTIME_KEY))

    m_IsDataReady = False

    If Not runtimeSources.RemoveItemsSource(normalizedKey) Then Exit Function
    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(DEMO_TABLES_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.RemoveTemporaryItemsSources() Then Exit Function
    If Not runtimeSources.RemoveTemporaryObjectsSources() Then Exit Function

    If Not runtimeSources.SetItemsSource(normalizedKey, tables, notifyChange) Then Exit Function
    m_IsDataReady = (tables.Count > 0)

    private_RegisterSqlTableItems = True
End Function

Private Function private_BuildTablesFromConfigTable() As Collection
    Dim parser As obj_CfgPersonalCardParser
    Dim sqlParams As obj_SqlParams
    Dim sectionTitle As String
    Dim sqlTable As obj_TableDynamic
    Dim result As Collection
    Dim rowsCount As Long

    ' 1) Парсим конфиг в SQL-параметры запроса.
    Set parser = m_CfgPersonalCardParser
    If parser Is Nothing Then Exit Function
    If Not parser.TryBuildSqlParams("Daily", "DailyEvents", sqlParams) Then Exit Function
    If sqlParams Is Nothing Then Exit Function
    sectionTitle = private_BuildSectionTitleFromSqlParams(sqlParams)

    ' 2) Выполняем запрос через общий SQL-движок для внешних Excel.
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequest(sqlParams, sqlTable, rowsCount) Then Exit Function

    Set result = New Collection
    If Not sqlTable Is Nothing Then
        ' 3) Применяем метаданные отображения (title) и публикуем непустой результат.
        If VBA.Len(sectionTitle) > 0 Then sqlTable.SectionTitle = sectionTitle
        If rowsCount > 0 Then result.Add sqlTable
    End If

    Set private_BuildTablesFromConfigTable = result
End Function

Private Function private_BuildSectionTitleFromSqlParams(ByVal sqlParams As obj_SqlParams) As String
    Dim rawSheetName As String
    Dim dollarPos As Long

    If sqlParams Is Nothing Then Exit Function

    rawSheetName = VBA.Trim$(sqlParams.SheetName)
    If VBA.Len(rawSheetName) = 0 Then Exit Function

    If VBA.Left$(rawSheetName, 1) = "[" And VBA.Right$(rawSheetName, 1) = "]" Then
        rawSheetName = VBA.Trim$(VBA.Mid$(rawSheetName, 2, VBA.Len(rawSheetName) - 2))
    End If

    dollarPos = VBA.InStr(1, rawSheetName, "$", VBA.vbBinaryCompare)
    If dollarPos > 1 Then
        private_BuildSectionTitleFromSqlParams = VBA.Trim$(VBA.Left$(rawSheetName, dollarPos - 1))
    Else
        private_BuildSectionTitleFromSqlParams = rawSheetName
    End If
End Function
