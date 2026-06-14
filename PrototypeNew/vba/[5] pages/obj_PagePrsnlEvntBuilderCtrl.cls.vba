VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PagePrsnlEvntBuilderCtrl"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PrsnlEvntBuilder.Controller"
Private Const CANDIDATE_TABLES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.EntityLookup.CandidateTables"
Private Const DUMMY_TABLES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.DummyTables"

Private m_Page As obj_IPage
Private m_LookupFeature As obj_EntityLookupFeature
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
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Property Get LookupFeature() As obj_EntityLookupFeature
    Set LookupFeature = m_LookupFeature
End Property

' //
' // API
' //
Public Function Initialize(ByVal page As Object) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePrsnlEvntBuilderCtrl.Initialize"
#End If
    Dim pageBase As obj_PageBase
    Dim pageInterface As obj_IPage

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PagePrsnlEvntBuilderCtrl initialization failed because page is not specified."
#End If
        Exit Function
    End If
    On Error Resume Next
    Set pageInterface = page
    On Error GoTo 0
    If pageInterface Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PagePrsnlEvntBuilderCtrl initialization failed because page does not implement obj_IPage."
#End If
        Exit Function
    End If

    m_IsDisposed = False
    Set m_Page = pageInterface

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function

    Set m_LookupFeature = New obj_EntityLookupFeature
    If Not m_LookupFeature.Initialize( _
        pageInterface, _
        CANDIDATE_TABLES_RUNTIME_KEY, _
        "prsnlevntbuilder:entitylookup") Then Exit Function

    If Not private_RegisterDummyTables(False) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePrsnlEvntBuilderCtrl.Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    If Not m_LookupFeature Is Nothing Then m_LookupFeature.Dispose
    Set m_LookupFeature = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    UpdateData = m_LookupFeature.UpdateData(configControl)
End Function

Public Function PrepareRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    If Not m_LookupFeature.PrepareLookupRuntime(notifyChange) Then Exit Function
    If Not private_RegisterDummyTables(notifyChange) Then Exit Function
    PrepareRuntime = True
End Function

Public Function ClearLookupCandidates(Optional ByVal renderNow As Boolean = True) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    ClearLookupCandidates = m_LookupFeature.ClearLookupCandidates(renderNow)
End Function

Public Function SearchCandidates( _
    ByVal lookupKey As String, _
    ByVal queryText As String, _
    ByRef outCandidateCount As Long, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    SearchCandidates = m_LookupFeature.SearchCandidates(lookupKey, queryText, outCandidateCount, notifyChange)
End Function

Public Function TryGetLookupKeys(ByRef outLookupKeys As Collection) As Boolean
    Set outLookupKeys = Nothing
    If m_LookupFeature Is Nothing Then Exit Function
    TryGetLookupKeys = m_LookupFeature.TryGetLookupKeys(outLookupKeys)
End Function

' //
' // Internal
' //
Private Function private_RegisterDummyTables(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim dummyTables As Collection

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set dummyTables = New Collection
    dummyTables.Add private_BuildDummyTable()

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(DUMMY_TABLES_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(DUMMY_TABLES_RUNTIME_KEY), dummyTables, notifyChange) Then Exit Function

    private_RegisterDummyTables = True
End Function

Private Function private_BuildDummyTable() As obj_TableDynamic
    Dim tableObj As obj_TableDynamic
    Dim rowObj As obj_Row
    Dim colIndex As Long
    Dim rowIndex As Long

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = "PrsnlEvntBuilder dummy table"

    For colIndex = 1 To 6
        If Not private_AddColumn(tableObj, "Column " & VBA.CStr(colIndex)) Then Exit Function
    Next colIndex

    For rowIndex = 1 To 8
        Set rowObj = New obj_Row
        For colIndex = 1 To 6
            rowObj.PushCellRaw "R" & VBA.CStr(rowIndex) & "C" & VBA.CStr(colIndex)
        Next colIndex
        If Not tableObj.PushRow(rowObj) Then Exit Function
    Next rowIndex

    Set private_BuildDummyTable = tableObj
End Function

Private Function private_AddColumn( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal columnName As String _
) As Boolean
    Dim colObj As obj_Column

    If tableObj Is Nothing Then Exit Function
    Set colObj = New obj_Column
    colObj.Name = VBA.Trim$(columnName)
    If VBA.Len(colObj.Name) = 0 Then colObj.Name = "Column " & VBA.CStr(tableObj.ColumnCount + 1)
    colObj.Position = tableObj.ColumnCount + 1
    private_AddColumn = tableObj.PushColumn(colObj)
End Function
