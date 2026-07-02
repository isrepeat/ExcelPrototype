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
Private Const HOTKEYS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.Hotkeys"
Private Const HOTKEY_ACTION_1 As String = "Action 1"
Private Const HOTKEY_ACTION_2 As String = "Action 2"
Private Const HOTKEY_SELECT_FORM_ROW As String = "Select Form Row"
Private Const LOOKUP_CANDIDATES_CONTROL_NAME As String = "LookupCandidatesTable"
Private Const EVENT_DRAFT_FORM_CONTAINER_NAME As String = "EventDraftForm"
Private Const EVENT_DRAFT_VALUES_CONTAINER_NAME As String = "EventDraftValues"

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
    If Not private_EnsureHotkeyRows(False) Then Exit Function
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
    If Not private_EnsureHotkeyRows(notifyChange) Then Exit Function
    PrepareRuntime = True
End Function

Public Function ClearLookupCandidates(Optional ByVal renderNow As Boolean = True) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    ClearLookupCandidates = m_LookupFeature.ClearLookupCandidates(renderNow)
End Function

Public Function RuntimeHandleHotkeyAction(ByVal actionId As Variant) As Boolean
    Dim pageBase As obj_PageBase
    Dim selectionObj As Object
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim actionText As String
    Dim cellValue As String

    ' Это page-specific action target для HotkeysControl.
    ' HotkeysControl передает только настроенный Action text; контроллер решает,
    ' что этот action значит на PrsnlEvntBuilder. Другие страницы могут переиспользовать
    ' HotkeysControl со своим actionMethod/dataContext.
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    Set selectionObj = Application.Selection
    If Not TypeOf selectionObj Is Range Then Exit Function

    Set targetCell = selectionObj.Cells(1, 1)
    If targetCell Is Nothing Then Exit Function
    If Not (targetCell.Worksheet Is ws) Then Exit Function

    actionText = VBA.Trim$(VBA.CStr(actionId))
    cellValue = VBA.CStr(targetCell.Value2)

    ' Page-specific actions branch by stable action ids and read current sheet state.
    Select Case VBA.LCase$(actionText)
        Case VBA.LCase$(HOTKEY_ACTION_1)
            If Not private_TryAcceptCandidateRowFromSelection(targetCell) Then
                RuntimeHandleHotkeyAction = True
                Exit Function
            End If

        Case VBA.LCase$(HOTKEY_ACTION_2)
            targetCell.Interior.Color = VBA.RGB(126, 36, 121)
            targetCell.Font.Color = VBA.RGB(255, 255, 255)

        Case VBA.LCase$(HOTKEY_SELECT_FORM_ROW)
            If Not private_TrySelectScopedRowFromSelection(targetCell) Then Exit Function

        Case Else
            Exit Function
    End Select

    rt_Messaging.fn_ShowStatusBarSuccess actionText & ": " & targetCell.Address(False, False) & " = '" & cellValue & "'", 3
    RuntimeHandleHotkeyAction = True
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

Private Function private_TryAcceptCandidateRowFromSelection(ByVal targetCell As Range) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim candidateRowsArea As Range
    Dim candidateRowRange As Range
    Dim draftValuesRange As Range
    Dim rowOffset As Long
    Dim sourceRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim colIndex As Long
    Dim previousEnableEvents As Boolean

    If targetCell Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not (targetCell.Worksheet Is ws) Then Exit Function

    ' TableList already registers rendered data rows as controlPart=rows.
    ' This keeps the hotkey independent from hard-coded row/column numbers.
    If Not private_TryResolveCandidateRowsArea(ws, targetCell, candidateRowsArea) Then
        rt_Messaging.fn_ShowStatusBarWarning "Select a candidate row cell first.", 3
        Exit Function
    End If

    ' EventDraftValues is the named layout container for the form value row.
    ' The controller only needs this container range, not individual input names.
    If Not pageBase.TryGetLayoutContainerRange(EVENT_DRAFT_VALUES_CONTAINER_NAME, draftValuesRange) Then
        rt_Messaging.fn_ShowStatusBarWarning "Event draft values container is not rendered.", 3
        Exit Function
    End If

    ' Convert the active cell position into the concrete rendered candidate row.
    ' The candidate rows area can move after rerender, so the relative row offset is used.
    rowOffset = targetCell.Row - candidateRowsArea.Row + 1
    If rowOffset <= 0 Or rowOffset > candidateRowsArea.Rows.Count Then Exit Function

    sourceRow = candidateRowsArea.Row + rowOffset - 1
    Set candidateRowRange = ws.Range( _
        ws.Cells(sourceRow, candidateRowsArea.Column), _
        ws.Cells(sourceRow, candidateRowsArea.Column + candidateRowsArea.Columns.Count - 1))

    ' LookupCandidates visually aligns result columns under the form columns.
    ' Therefore the safest generic mapping is the intersection of absolute Excel columns.
    firstCol = private_MaxLong(candidateRowRange.Column, draftValuesRange.Column)
    lastCol = private_MinLong( _
        candidateRowRange.Column + candidateRowRange.Columns.Count - 1, _
        draftValuesRange.Column + draftValuesRange.Columns.Count - 1)
    If lastCol < firstCol Then
        rt_Messaging.fn_ShowStatusBarWarning "Selected candidate row does not overlap the event form.", 3
        Exit Function
    End If

    ' Values are written directly into sheet cells. Disable events so this accept action
    ' does not recursively trigger input onChange/search/rerender for each copied cell.
    previousEnableEvents = Application.EnableEvents
    On Error GoTo RestoreEventsAndFail
    Application.EnableEvents = False
    For colIndex = firstCol To lastCol
        ws.Cells(draftValuesRange.Row, colIndex).Value2 = ws.Cells(sourceRow, colIndex).Value2
    Next colIndex
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0

    candidateRowRange.Select
    private_TryAcceptCandidateRowFromSelection = True
    Exit Function

RestoreEventsAndFail:
    On Error Resume Next
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
End Function

Private Function private_TrySelectScopedRowFromSelection(ByVal targetCell As Range) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim formRange As Range
    Dim scopedRow As Range
    Dim listObj As ListObject

    If targetCell Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not (targetCell.Worksheet Is ws) Then Exit Function

    If pageBase.TryGetLayoutContainerRange(EVENT_DRAFT_FORM_CONTAINER_NAME, formRange) Then
        If Not formRange Is Nothing Then
            Set scopedRow = Nothing
            On Error Resume Next
            If Not Application.Intersect(targetCell, formRange) Is Nothing Then
                Set scopedRow = Application.Intersect(targetCell.EntireRow, formRange)
            End If
            On Error GoTo 0
            If Not scopedRow Is Nothing Then
                scopedRow.Select
                private_TrySelectScopedRowFromSelection = True
                Exit Function
            End If
        End If
    End If

    For Each listObj In ws.ListObjects
        If listObj Is Nothing Then GoTo ContinueListObject
        If listObj.Range Is Nothing Then GoTo ContinueListObject
        Set scopedRow = Nothing
        On Error Resume Next
        If Not Application.Intersect(targetCell, listObj.Range) Is Nothing Then
            Set scopedRow = Application.Intersect(targetCell.EntireRow, listObj.Range)
        End If
        On Error GoTo 0
        If Not scopedRow Is Nothing Then
            scopedRow.Select
            private_TrySelectScopedRowFromSelection = True
            Exit Function
        End If

ContinueListObject:
    Next listObj

    targetCell.EntireRow.Select
    private_TrySelectScopedRowFromSelection = True
End Function

Private Function private_TryResolveCandidateRowsArea( _
    ByVal ws As Worksheet, _
    ByVal targetCell As Range, _
    ByRef outRowsArea As Range _
) As Boolean
    Dim rowsScope As Range
    Dim columnScope As Range
    Dim area As Range

    Set outRowsArea = Nothing
    If ws Is Nothing Then Exit Function
    If targetCell Is Nothing Then Exit Function

    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, _
        "tablelist", _
        LOOKUP_CANDIDATES_CONTROL_NAME, _
        "rows", _
        rowsScope, _
        columnScope) Then Exit Function
    If rowsScope Is Nothing Then Exit Function

    For Each area In rowsScope.Areas
        If Not Application.Intersect(targetCell, area) Is Nothing Then
            Set outRowsArea = area
            private_TryResolveCandidateRowsArea = True
            Exit Function
        End If
    Next area
End Function

Private Function private_MaxLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        private_MaxLong = leftValue
    Else
        private_MaxLong = rightValue
    End If
End Function

Private Function private_MinLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue <= rightValue Then
        private_MinLong = leftValue
    Else
        private_MinLong = rightValue
    End If
End Function

Private Function private_EnsureHotkeyRows(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim hotkeyRows As Collection
    Dim existingRows As Collection
    Dim hasChanges As Boolean

    ' Сеем default-строки хоткеев только когда page runtime source отсутствует/пустой.
    ' После Apply HotkeysControl пишет отредактированные строки обратно в тот же
    ' RuntimeItems key, поэтому rerender/PrepareRuntime не должны их перетирать.
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    If runtimeSources.TryGetItemsSourceByKey(VBA.LCase$(HOTKEYS_RUNTIME_KEY), existingRows, True) Then
        If Not existingRows Is Nothing Then
            If existingRows.Count > 0 Then
                Set hotkeyRows = existingRows
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_ACTION_1, "CTRL+ENTER", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_ACTION_2, "CTRL+SHIFT+R", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE", hasChanges) Then Exit Function
                If hasChanges Then
                    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY)) Then Exit Function
                    If Not runtimeSources.SetItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY), hotkeyRows, notifyChange) Then Exit Function
                End If
                private_EnsureHotkeyRows = True
                Exit Function
            End If
        End If
    Else
        Exit Function
    End If

    Set hotkeyRows = New Collection
    ' Defaults — это только стартовые данные страницы. Активными они становятся
    ' после render HotkeysControl и RuntimeRegisterBoundRows, где регистрируются
    ' routes для этой страницы.
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_ACTION_1, "CTRL+ENTER") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_ACTION_2, "CTRL+SHIFT+R") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE") Then Exit Function

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY), hotkeyRows, notifyChange) Then Exit Function

    private_EnsureHotkeyRows = True
End Function

Private Function private_EnsureHotkeyRow( _
    ByVal hotkeyRows As Collection, _
    ByVal actionId As String, _
    ByVal defaultHotkey As String, _
    ByRef ioHasChanges As Boolean _
) As Boolean
    If hotkeyRows Is Nothing Then Exit Function
    If Not private_HotkeyRowsContainAction(hotkeyRows, actionId) Then
        If Not private_AddHotkeyRow(hotkeyRows, actionId, defaultHotkey) Then Exit Function
        ioHasChanges = True
    End If

    private_EnsureHotkeyRow = True
End Function

Private Function private_HotkeyRowsContainAction( _
    ByVal hotkeyRows As Collection, _
    ByVal actionId As String _
) As Boolean
    Dim rowItem As Variant
    Dim configEntry As obj_ConfigEntry

    If hotkeyRows Is Nothing Then Exit Function
    actionId = VBA.Trim$(actionId)
    If VBA.Len(actionId) = 0 Then Exit Function

    For Each rowItem In hotkeyRows
        If Not VBA.IsObject(rowItem) Then GoTo ContinueRow
        Set configEntry = Nothing
        On Error Resume Next
        Set configEntry = rowItem
        On Error GoTo 0
        If configEntry Is Nothing Then GoTo ContinueRow
        If VBA.StrComp(VBA.Trim$(configEntry.Key), actionId, VBA.vbTextCompare) = 0 Then
            private_HotkeyRowsContainAction = True
            Exit Function
        End If

ContinueRow:
    Next rowItem
End Function

Private Function private_AddHotkeyRow( _
    ByVal hotkeyRows As Collection, _
    ByVal actionId As String, _
    ByVal defaultHotkey As String _
) As Boolean
    Dim configEntry As obj_ConfigEntry

    If hotkeyRows Is Nothing Then Exit Function
    Set configEntry = New obj_ConfigEntry
    configEntry.Attr = VBA.vbNullString
    configEntry.Key = VBA.Trim$(actionId)
    configEntry.Value = VBA.Trim$(defaultHotkey)
    hotkeyRows.Add configEntry
    private_AddHotkeyRow = True
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
