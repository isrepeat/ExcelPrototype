VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PagePersonalCardCtrl"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PagePersonalCard.Controller"
Private Const TABLES_RUNTIME_KEY As String = "RuntimeItems.Test.Tables"
Private Const TABLE_HEADER_TEXT As String = "Name | Role | Country | Team | Level | Status | Since"

Private m_Page As obj_IPage
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
' // API
' //
Public Function Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePersonalCardCtrl.Initialize"
#End If
    Dim pageBase As obj_PageBase

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PagePersonalCardCtrl initialization failed because page is not specified."
#End If
        MsgBox "PrototypeNew: PagePersonalCardCtrl initialization failed because page is not specified.", vbExclamation, "PrototypeNew / PersonalCard runtime"
        Exit Function
    End If

    m_IsDisposed = False
    Set m_Page = page
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
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Function PrepareDemoTablesRuntime( _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePersonalCardCtrl.PrepareDemoTablesRuntime"
#End If
    If Not private_RegisterDemoTableItems(notifyChange) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to register demo table items for PersonalCard page."
#End If
        MsgBox "PrototypeNew: failed to register demo table items for PersonalCard page.", vbExclamation, "PrototypeNew / PersonalCard runtime"
        Exit Function
    End If

    PrepareDemoTablesRuntime = True
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
    Set runtimeSources = pageBase.RuntimeSources
    normalizedKey = VBA.LCase$(VBA.Trim$(TABLES_RUNTIME_KEY))
    
    ' Точечная очистка:
    ' - удаляем только test-таблицы;
    ' - чистим только временные runtime-ключи, которые создают layout-рендереры.
    If Not runtimeSources.RemoveItemsSource(normalizedKey) Then Exit Function
    If Not runtimeSources.RemoveTemporaryItemsSources() Then Exit Function
    If Not runtimeSources.RemoveTemporaryObjectsSources() Then Exit Function

    If Not runtimeSources.SetItemsSource(normalizedKey, tables, notifyChange) Then Exit Function

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
