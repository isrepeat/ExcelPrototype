VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageComparingCtrl"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PageComparing.Controller"
Private Const RUNTIME_ERROR_TITLE As String = "PrototypeNew / Comparing runtime"
Private Const DIFF_TABLES_ITEMS_SOURCE_KEY As String = "RuntimeItems.Comparing.DiffTables"
Private Const DIFF_CONTEXT_ROWS As Long = 1
Private Const OLD_ROW_COLUMN_ALIAS As String = "OldRow"
Private Const NEW_ROW_COLUMN_ALIAS As String = "NewRow"

Private m_Page As obj_IPage
Private m_ConfigTable As obj_ConfigTable
Private m_CfgParser As obj_ComparingCfgParser
Private m_IsConfigReady As Boolean
Private m_StatusText As String
Private m_IsCondensedDiffView As Boolean
Private m_LastOutputColumns As Collection
Private m_LastOutputColumnFormats As Object
Private m_LastDiffRows As Collection
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    m_StatusText = "Comparing config is not loaded yet."
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

Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Property Get IsConfigReady() As Boolean
    IsConfigReady = m_IsConfigReady
End Property

Public Property Get StatusText() As String
    StatusText = m_StatusText
End Property

Public Property Get IsCondensedDiffView() As Boolean
    IsCondensedDiffView = m_IsCondensedDiffView
End Property

Public Function Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageComparingCtrl.Initialize"
#End If
    Dim pageBase As obj_PageBase
    Dim emptyItems As Collection

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PageComparingCtrl initialization failed because page is not specified."
#End If
        Exit Function
    End If

    m_IsDisposed = False
    Set m_Page = page
    Set m_ConfigTable = Nothing
    Set m_CfgParser = Nothing
    Set m_LastOutputColumns = Nothing
    Set m_LastOutputColumnFormats = Nothing
    Set m_LastDiffRows = Nothing
    m_IsCondensedDiffView = True
    m_IsConfigReady = False
    m_StatusText = "Comparing config is not loaded yet."

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If pageBase.RuntimeSources Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function
    Set emptyItems = New Collection
    If Not pageBase.RuntimeSources.SetItemsSource(DIFF_TABLES_ITEMS_SOURCE_KEY, emptyItems, False) Then Exit Function

    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageComparingCtrl.Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    m_IsConfigReady = False
    If Not m_CfgParser Is Nothing Then m_CfgParser.Dispose
    Set m_CfgParser = Nothing
    Set m_ConfigTable = Nothing
    Set m_LastOutputColumns = Nothing
    Set m_LastOutputColumnFormats = Nothing
    Set m_LastDiffRows = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable
    Dim cfgParser As obj_ComparingCfgParser

    m_IsConfigReady = False
    m_StatusText = "Comparing config is not loaded yet."

    If configControl Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PageComparingCtrl.UpdateData failed because config control is not specified."
#End If
        m_StatusText = "Parent DevConfig control was not found."
        Exit Function
    End If

    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    If configTable Is Nothing Then Exit Function

    Set cfgParser = New obj_ComparingCfgParser
    If Not cfgParser.Initialize(configTable) Then Exit Function
    If Not cfgParser.TryValidateSkeleton(m_StatusText) Then Exit Function

    On Error Resume Next
    If Not m_CfgParser Is Nothing Then m_CfgParser.Dispose
    On Error GoTo 0

    Set m_ConfigTable = configTable
    Set m_CfgParser = cfgParser
    m_IsConfigReady = True
    UpdateData = True
End Function

Public Function RunPipeline(Optional ByVal notifyChange As Boolean = True) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageComparingCtrl.RunPipeline"
#End If
    Dim leftTableRef As String
    Dim rightTableRef As String
    Dim keyColumns As Collection
    ' CompareColumns из профиля: видимые бизнес-колонки результата и набор
    ' полей, по которым строка считается modified.
    Dim compareColumns As Collection
    ' Метаданные к CompareColumns, например OutDate{fmt:Date}. Сравнение и SQL
    ' их не используют; map нужен только чтобы пометить UI-колонки как datelike.
    Dim compareColumnFormats As Object
    ' Фактически читаемые из источника колонки: compareColumns плюс keyColumns,
    ' которых нет в видимом результате, но которые нужны для сопоставления строк.
    Dim loadColumns As Collection
    ' Видимый порядок результата. Сейчас совпадает с compareColumns, отдельно
    ' держим как границу между будущим UI-выводом и правилами сравнения.
    Dim outputColumns As Collection
    Dim ignoreCase As Boolean
    Dim trimText As Boolean
    Dim leftSqlParams As obj_SqlParams
    Dim rightSqlParams As obj_SqlParams
    Dim leftData As obj_TableData
    Dim rightData As obj_TableData
    Dim diffRows As Collection
    Dim statusText As String
    Dim visibleRowCount As Long
    Dim totalStart As Single
    Dim stageStart As Single

    If m_Page Is Nothing Then Exit Function
    If m_CfgParser Is Nothing Then
        rt_Messaging.fn_ShowStatusBarWarning "Comparing config is not ready.", 4
        Exit Function
    End If
    If Not m_IsConfigReady Then
        rt_Messaging.fn_ShowStatusBarWarning "Comparing config is not ready.", 4
        Exit Function
    End If

    totalStart = VBA.Timer
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "comparing:pipeline start notifyChange=" & VBA.CStr(notifyChange)
#End If

    stageStart = VBA.Timer
    If Not m_CfgParser.TryGetCompareSettings(leftTableRef, rightTableRef, keyColumns, compareColumns, compareColumnFormats, ignoreCase, trimText) Then Exit Function
    private_LogPipelineStep "settings", stageStart, _
        "left='" & leftTableRef & _
        "' right='" & rightTableRef & "'"

    stageStart = VBA.Timer
    If Not private_TryBuildOutputColumns(compareColumns, outputColumns) Then Exit Function
    If Not private_TryBuildLoadColumns(compareColumns, keyColumns, loadColumns) Then Exit Function
    If Not m_CfgParser.TryBuildTableSqlParams(leftTableRef, loadColumns, leftSqlParams) Then Exit Function
    If Not m_CfgParser.TryBuildTableSqlParams(rightTableRef, loadColumns, rightSqlParams) Then Exit Function
    private_LogPipelineStep "build-sql-params", stageStart, _
        "outputColumns=" & VBA.CStr(outputColumns.Count) & _
        " loadColumns=" & VBA.CStr(loadColumns.Count)

    stageStart = VBA.Timer
    ' Для Comparing читаем внешние Excel-таблицы сразу в легкий obj_TableData.
    ' Он хранит только 2D Variant-массив, без тысяч obj_Row/obj_Cell, поэтому
    ' сравнение больших таблиц не платит за создание и последующее разрушение
    ' тяжелой объектной модели.
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequestData(leftSqlParams, leftData) Then Exit Function
    private_LogPipelineStep "read-left", stageStart, "rows=" & VBA.CStr(leftData.RowCount) & " columns=" & VBA.CStr(leftData.ColumnCount)

    stageStart = VBA.Timer
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequestData(rightSqlParams, rightData) Then Exit Function
    private_LogPipelineStep "read-right", stageStart, "rows=" & VBA.CStr(rightData.RowCount) & " columns=" & VBA.CStr(rightData.ColumnCount)

    Set m_LastOutputColumns = outputColumns
    Set m_LastOutputColumnFormats = compareColumnFormats

    stageStart = VBA.Timer
    If Not private_TryBuildDiffRowsFromData( _
        leftData, _
        rightData, _
        loadColumns, _
        outputColumns, _
        keyColumns, _
        compareColumns, _
        ignoreCase, _
        trimText, _
        diffRows, _
        statusText) Then Exit Function
    private_LogPipelineStep "build-diff", stageStart, "rows=" & VBA.CStr(diffRows.Count)
    m_StatusText = statusText
    Set m_LastDiffRows = diffRows

    stageStart = VBA.Timer
    If Not private_TryRefreshDiffTableItemsSource(visibleRowCount) Then Exit Function
    private_LogPipelineStep "set-result-source", stageStart, _
        "visibleRows=" & VBA.CStr(visibleRowCount) & _
        " mode='" & private_GetDiffViewModeName() & "'"

    If notifyChange Then
        stageStart = VBA.Timer
        If Not rt_PageManager.fn_RenderPage(m_Page, "comparing:run-pipeline") Then Exit Function
        private_LogPipelineStep "page-rerender", stageStart
    End If

    rt_Messaging.fn_ShowStatusBarNotice m_StatusText, 4
    private_LogPipelineStep "total", totalStart, m_StatusText
    RunPipeline = True
End Function

Public Function RuntimeToggleDiffView(Optional ByVal arg As Variant) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageComparingCtrl.RuntimeToggleDiffView"
#End If
    RuntimeToggleDiffView = ToggleDiffView(True)
End Function

Public Function ToggleDiffView(Optional ByVal notifyChange As Boolean = True) As Boolean
    If m_Page Is Nothing Then Exit Function
    m_IsCondensedDiffView = Not m_IsCondensedDiffView

    ' Не кешируем полные obj_TableDynamic между кликами: при повторном Run Excel/VBA
    ' может долго освобождать старые тысячи obj_Row/obj_Cell. Переключение режима
    ' поэтому пересчитывает pipeline заново, зато второй Run не платит за старый кеш.
    ToggleDiffView = RunPipeline(notifyChange)
End Function

Private Function private_TryBuildOutputColumns( _
    ByVal compareColumns As Collection, _
    ByRef outColumns As Collection _
) As Boolean
    Dim aliasItem As Variant
    Dim aliasText As String
    Dim seen As Object

    Set outColumns = New Collection
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    For Each aliasItem In compareColumns
        aliasText = VBA.Trim$(VBA.CStr(aliasItem))
        If VBA.Len(aliasText) = 0 Then GoTo ContinueCompare
        If Not seen.Exists(aliasText) Then
            seen.Add aliasText, True
            outColumns.Add aliasText
        End If
ContinueCompare:
    Next aliasItem

    private_TryBuildOutputColumns = (outColumns.Count > 0)
End Function

Private Function private_TryBuildLoadColumns( _
    ByVal compareColumns As Collection, _
    ByVal keyColumns As Collection, _
    ByRef outColumns As Collection _
) As Boolean
    Dim aliasItem As Variant
    Dim aliasText As String
    Dim seen As Object

    Set outColumns = New Collection
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    ' CompareColumns задает видимый порядок результата. KeyColumns добавляем
    ' только в конец и только если их нет среди видимых колонок: они нужны для
    ' сопоставления строк, но сами по себе не управляют порядком UI-таблицы.
    For Each aliasItem In compareColumns
        aliasText = VBA.Trim$(VBA.CStr(aliasItem))
        If VBA.Len(aliasText) = 0 Then GoTo ContinueCompare
        If Not seen.Exists(aliasText) Then
            seen.Add aliasText, True
            outColumns.Add aliasText
        End If
ContinueCompare:
    Next aliasItem

    For Each aliasItem In keyColumns
        aliasText = VBA.Trim$(VBA.CStr(aliasItem))
        If VBA.Len(aliasText) = 0 Then GoTo ContinueKey
        If Not seen.Exists(aliasText) Then
            seen.Add aliasText, True
            outColumns.Add aliasText
        End If
ContinueKey:
    Next aliasItem

    private_TryBuildLoadColumns = (outColumns.Count > 0)
End Function

Private Function private_TryBuildDiffRowsFromData( _
    ByVal leftData As obj_TableData, _
    ByVal rightData As obj_TableData, _
    ByVal loadColumns As Collection, _
    ByVal outputColumns As Collection, _
    ByVal keyColumns As Collection, _
    ByVal compareColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean, _
    ByRef outDiffRows As Collection, _
    ByRef outStatusText As String _
) As Boolean
    Dim eventTypes() As String
    Dim eventLeftRows() As Long
    Dim eventRightRows() As Long
    Dim eventCount As Long
    Dim maxEvents As Long
    Dim leftKeys() As String
    Dim rightKeys() As String
    Dim leftKeyIndex As Object
    Dim rightKeyIndex As Object
    Dim keyColumnIndexes As Collection
    Dim compareColumnIndexes As Object
    Dim lcsKeys As Collection
    Dim matchKey As Variant
    Dim oldPos As Long
    Dim newPos As Long
    Dim oldMatch As Long
    Dim newMatch As Long
    Dim addedCount As Long
    Dim deletedCount As Long
    Dim modifiedCount As Long
    Dim movedCount As Long
    Dim unchangedCount As Long
    Dim stageStart As Single
    Dim eventStageStart As Single
    Dim rowType As String
    Dim oldKey As String
    Dim newKey As String

    Set outDiffRows = New Collection
    outStatusText = VBA.vbNullString
    If leftData Is Nothing Then Exit Function
    If rightData Is Nothing Then Exit Function

    If Not private_TryBuildColumnIndexList(loadColumns, keyColumns, keyColumnIndexes) Then Exit Function
    If Not private_TryBuildCompareColumnIndexSet(outputColumns, compareColumns, compareColumnIndexes) Then Exit Function

    ' Pipeline сравнения:
    ' 1) строим ключи строк по keyColumns для Old/New;
    ' 2) через LCS/LIS оставляем максимальную общую последовательность ключей;
    ' 3) строки с ключами вне LCS, но в обеих таблицах, маркируем как moved-pair;
    ' 4) между совпавшими ключами фиксируем added/deleted/modified/unchanged;
    ' 5) материализуем в UI все события или condensed-контекст вокруг изменений.
    maxEvents = leftData.RowCount + rightData.RowCount
    If maxEvents <= 0 Then
        private_TryBuildDiffRowsFromData = True
        Exit Function
    End If
    ReDim eventTypes(1 To maxEvents)
    ReDim eventLeftRows(1 To maxEvents)
    ReDim eventRightRows(1 To maxEvents)

    stageStart = VBA.Timer
    If Not private_TryBuildKeyIndexFromData(leftData, keyColumnIndexes, ignoreCase, trimText, leftKeys, leftKeyIndex, "left") Then Exit Function
    private_LogPipelineStep "build-diff:left-key-index", stageStart, "rows=" & VBA.CStr(leftData.RowCount)

    stageStart = VBA.Timer
    If Not private_TryBuildKeyIndexFromData(rightData, keyColumnIndexes, ignoreCase, trimText, rightKeys, rightKeyIndex, "right") Then Exit Function
    private_LogPipelineStep "build-diff:right-key-index", stageStart, "rows=" & VBA.CStr(rightData.RowCount)

    stageStart = VBA.Timer
    If Not private_TryBuildLcsKeys(leftKeys, rightKeys, lcsKeys) Then Exit Function
    private_LogPipelineStep "build-diff:lcs-keys", stageStart, "matches=" & VBA.CStr(lcsKeys.Count)

    stageStart = VBA.Timer
    eventStageStart = stageStart
    oldPos = 1
    newPos = 1
    For Each matchKey In lcsKeys
        oldMatch = CLng(leftKeyIndex(VBA.CStr(matchKey)))
        newMatch = CLng(rightKeyIndex(VBA.CStr(matchKey)))

        Do While oldPos < oldMatch
            oldKey = leftKeys(oldPos)
            If rightKeyIndex.Exists(oldKey) Then
                private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "deleted-moved", oldPos, 0
                deletedCount = deletedCount + 1
                movedCount = movedCount + 1
            Else
                private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "deleted", oldPos, 0
                deletedCount = deletedCount + 1
            End If
            oldPos = oldPos + 1
        Loop

        Do While newPos < newMatch
            newKey = rightKeys(newPos)
            If leftKeyIndex.Exists(newKey) Then
                private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "added-moved", 0, newPos
                addedCount = addedCount + 1
            Else
                private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "added", 0, newPos
                addedCount = addedCount + 1
            End If
            newPos = newPos + 1
        Loop

        If private_RowHasCompareChangesInData(leftData, oldMatch, rightData, newMatch, outputColumns, compareColumnIndexes, ignoreCase, trimText) Then
            rowType = "modified"
            modifiedCount = modifiedCount + 1
        Else
            rowType = "unchanged"
            unchangedCount = unchangedCount + 1
        End If
        private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, rowType, oldMatch, newMatch

        oldPos = oldMatch + 1
        newPos = newMatch + 1
    Next matchKey

    Do While oldPos <= leftData.RowCount
        oldKey = leftKeys(oldPos)
        If rightKeyIndex.Exists(oldKey) Then
            private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "deleted-moved", oldPos, 0
            deletedCount = deletedCount + 1
            movedCount = movedCount + 1
        Else
            private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "deleted", oldPos, 0
            deletedCount = deletedCount + 1
        End If
        oldPos = oldPos + 1
    Loop

    Do While newPos <= rightData.RowCount
        newKey = rightKeys(newPos)
        If leftKeyIndex.Exists(newKey) Then
            private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "added-moved", 0, newPos
            addedCount = addedCount + 1
        Else
            private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "added", 0, newPos
            addedCount = addedCount + 1
        End If
        newPos = newPos + 1
    Loop
    private_LogPipelineStep "build-diff:events", eventStageStart, _
        "events=" & VBA.CStr(eventCount) & _
        " added=" & VBA.CStr(addedCount) & _
        " deleted=" & VBA.CStr(deletedCount) & _
        " modified=" & VBA.CStr(modifiedCount) & _
        " moved=" & VBA.CStr(movedCount) & _
        " unchanged=" & VBA.CStr(unchangedCount)

    stageStart = VBA.Timer
    If Not private_TryMaterializeDiffEventsFromData( _
        leftData, _
        rightData, _
        outputColumns, _
        compareColumns, _
        ignoreCase, _
        trimText, _
        eventTypes, _
        eventLeftRows, _
        eventRightRows, _
        eventCount, _
        m_IsCondensedDiffView, _
        outDiffRows) Then Exit Function

    outStatusText = "Compare done: added=" & VBA.CStr(addedCount) & _
        ", deleted=" & VBA.CStr(deletedCount) & _
        ", modified=" & VBA.CStr(modifiedCount) & _
        ", moved=" & VBA.CStr(movedCount) & _
        ", unchanged=" & VBA.CStr(unchangedCount) & "."
    private_LogPipelineStep "build-diff:materialize-rows", stageStart, _
        "events=" & VBA.CStr(eventCount) & _
        " rows=" & VBA.CStr(outDiffRows.Count) & _
        " mode='" & private_GetDiffViewModeName() & "'"
    private_TryBuildDiffRowsFromData = True
End Function

Private Function private_TryBuildDiffRows( _
    ByVal leftTable As obj_TableDynamic, _
    ByVal rightTable As obj_TableDynamic, _
    ByVal outputColumns As Collection, _
    ByVal keyColumns As Collection, _
    ByVal compareColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean, _
    ByRef outDiffRows As Collection, _
    ByRef outStatusText As String _
) As Boolean
    Dim eventTypes() As String
    Dim eventLeftRows() As Long
    Dim eventRightRows() As Long
    Dim eventCount As Long
    Dim maxEvents As Long
    Dim leftKeys() As String
    Dim rightKeys() As String
    Dim leftKeyIndex As Object
    Dim rightKeyIndex As Object
    Dim compareColumnIndexes As Object
    Dim lcsKeys As Collection
    Dim matchKey As Variant
    Dim oldPos As Long
    Dim newPos As Long
    Dim oldMatch As Long
    Dim newMatch As Long
    Dim addedCount As Long
    Dim deletedCount As Long
    Dim modifiedCount As Long
    Dim unchangedCount As Long
    Dim stageStart As Single
    Dim eventStageStart As Single
    Dim rowType As String

    Set outDiffRows = New Collection
    outStatusText = VBA.vbNullString
    If leftTable Is Nothing Then Exit Function
    If rightTable Is Nothing Then Exit Function

    maxEvents = leftTable.RowCount + rightTable.RowCount
    If maxEvents <= 0 Then
        private_TryBuildDiffRows = True
        Exit Function
    End If
    ReDim eventTypes(1 To maxEvents)
    ReDim eventLeftRows(1 To maxEvents)
    ReDim eventRightRows(1 To maxEvents)

    stageStart = VBA.Timer
    If Not private_TryBuildKeyIndex(leftTable, keyColumns, ignoreCase, trimText, leftKeys, leftKeyIndex, "left") Then Exit Function
    private_LogPipelineStep "build-diff:left-key-index", stageStart, "rows=" & VBA.CStr(leftTable.RowCount)

    stageStart = VBA.Timer
    If Not private_TryBuildKeyIndex(rightTable, keyColumns, ignoreCase, trimText, rightKeys, rightKeyIndex, "right") Then Exit Function
    private_LogPipelineStep "build-diff:right-key-index", stageStart, "rows=" & VBA.CStr(rightTable.RowCount)

    stageStart = VBA.Timer
    If Not private_TryBuildLcsKeys(leftKeys, rightKeys, lcsKeys) Then Exit Function
    private_LogPipelineStep "build-diff:lcs-keys", stageStart, "matches=" & VBA.CStr(lcsKeys.Count)
    If Not private_TryBuildCompareColumnIndexSet(outputColumns, compareColumns, compareColumnIndexes) Then Exit Function

    stageStart = VBA.Timer
    eventStageStart = stageStart
    oldPos = 1
    newPos = 1
    For Each matchKey In lcsKeys
        oldMatch = CLng(leftKeyIndex(VBA.CStr(matchKey)))
        newMatch = CLng(rightKeyIndex(VBA.CStr(matchKey)))

        Do While oldPos < oldMatch
            private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "deleted", oldPos, 0
            deletedCount = deletedCount + 1
            oldPos = oldPos + 1
        Loop

        Do While newPos < newMatch
            private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "added", 0, newPos
            addedCount = addedCount + 1
            newPos = newPos + 1
        Loop

        If private_RowHasCompareChangesIndexed(leftTable.Rows.Item(oldMatch), rightTable.Rows.Item(newMatch), outputColumns, compareColumnIndexes, ignoreCase, trimText) Then
            rowType = "modified"
            modifiedCount = modifiedCount + 1
        Else
            rowType = "unchanged"
            unchangedCount = unchangedCount + 1
        End If
        private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, rowType, oldMatch, newMatch

        oldPos = oldMatch + 1
        newPos = newMatch + 1
    Next matchKey

    Do While oldPos <= leftTable.RowCount
        private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "deleted", oldPos, 0
        deletedCount = deletedCount + 1
        oldPos = oldPos + 1
    Loop

    Do While newPos <= rightTable.RowCount
        private_AddDiffEvent eventTypes, eventLeftRows, eventRightRows, eventCount, "added", 0, newPos
        addedCount = addedCount + 1
        newPos = newPos + 1
    Loop
    private_LogPipelineStep "build-diff:events", eventStageStart, _
        "events=" & VBA.CStr(eventCount) & _
        " added=" & VBA.CStr(addedCount) & _
        " deleted=" & VBA.CStr(deletedCount) & _
        " modified=" & VBA.CStr(modifiedCount) & _
        " unchanged=" & VBA.CStr(unchangedCount)

    stageStart = VBA.Timer
    If Not private_TryMaterializeDiffEvents( _
        leftTable, _
        rightTable, _
        outputColumns, _
        compareColumns, _
        ignoreCase, _
        trimText, _
        eventTypes, _
        eventLeftRows, _
        eventRightRows, _
        eventCount, _
        m_IsCondensedDiffView, _
        outDiffRows) Then Exit Function

    outStatusText = "Compare done: added=" & VBA.CStr(addedCount) & _
        ", deleted=" & VBA.CStr(deletedCount) & _
        ", modified=" & VBA.CStr(modifiedCount) & _
        ", unchanged=" & VBA.CStr(unchangedCount) & "."
    private_LogPipelineStep "build-diff:materialize-rows", stageStart, _
        "events=" & VBA.CStr(eventCount) & _
        " rows=" & VBA.CStr(outDiffRows.Count) & _
        " mode='" & private_GetDiffViewModeName() & "'"
    private_TryBuildDiffRows = True
End Function

Private Sub private_AddDiffEvent( _
    ByRef eventTypes() As String, _
    ByRef eventLeftRows() As Long, _
    ByRef eventRightRows() As Long, _
    ByRef eventCount As Long, _
    ByVal rowType As String, _
    ByVal leftRowIndex As Long, _
    ByVal rightRowIndex As Long _
)
    eventCount = eventCount + 1
    eventTypes(eventCount) = rowType
    eventLeftRows(eventCount) = leftRowIndex
    eventRightRows(eventCount) = rightRowIndex
End Sub

Private Function private_TryMaterializeDiffEvents( _
    ByVal leftTable As obj_TableDynamic, _
    ByVal rightTable As obj_TableDynamic, _
    ByVal outputColumns As Collection, _
    ByVal compareColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean, _
    ByRef eventTypes() As String, _
    ByRef eventLeftRows() As Long, _
    ByRef eventRightRows() As Long, _
    ByVal eventCount As Long, _
    ByVal condensedView As Boolean, _
    ByRef outDiffRows As Collection _
) As Boolean
    Dim keepEvents() As Boolean
    Dim eventIndex As Long
    Dim hasVisibleEvents As Boolean
    Dim inCollapsedRun As Boolean

    Set outDiffRows = New Collection
    If eventCount <= 0 Then
        private_TryMaterializeDiffEvents = True
        Exit Function
    End If

    If condensedView Then
        ReDim keepEvents(1 To eventCount)
        private_MarkCondensedVisibleEvents eventTypes, eventCount, keepEvents, hasVisibleEvents
    End If

    For eventIndex = 1 To eventCount
        If Not condensedView Then
            If Not private_AddDiffRowFromEvent( _
                leftTable, rightTable, outDiffRows, eventTypes(eventIndex), _
                eventLeftRows(eventIndex), eventRightRows(eventIndex), _
                outputColumns, compareColumns, ignoreCase, trimText) Then Exit Function
        ElseIf Not hasVisibleEvents Then
            If Not private_AddDiffRowFromEvent( _
                leftTable, rightTable, outDiffRows, eventTypes(eventIndex), _
                eventLeftRows(eventIndex), eventRightRows(eventIndex), _
                outputColumns, compareColumns, ignoreCase, trimText) Then Exit Function
        ElseIf keepEvents(eventIndex) Then
            If Not private_AddDiffRowFromEvent( _
                leftTable, rightTable, outDiffRows, eventTypes(eventIndex), _
                eventLeftRows(eventIndex), eventRightRows(eventIndex), _
                outputColumns, compareColumns, ignoreCase, trimText) Then Exit Function
            inCollapsedRun = False
        ElseIf Not inCollapsedRun Then
            outDiffRows.Add private_CreateEllipsisDiffRow(outputColumns.Count)
            inCollapsedRun = True
        End If
    Next eventIndex

    private_TryMaterializeDiffEvents = True
End Function

Private Sub private_MarkCondensedVisibleEvents( _
    ByRef eventTypes() As String, _
    ByVal eventCount As Long, _
    ByRef keepEvents() As Boolean, _
    ByRef outHasVisibleEvents As Boolean _
)
    Dim eventIndex As Long
    Dim contextIndex As Long
    Dim contextStart As Long
    Dim contextEnd As Long

    outHasVisibleEvents = False
    For eventIndex = 1 To eventCount
        If private_IsTouchedDiffEventType(eventTypes(eventIndex)) Then
            contextStart = eventIndex - DIFF_CONTEXT_ROWS
            If contextStart < 1 Then contextStart = 1
            contextEnd = eventIndex + DIFF_CONTEXT_ROWS
            If contextEnd > eventCount Then contextEnd = eventCount

            For contextIndex = contextStart To contextEnd
                keepEvents(contextIndex) = True
            Next contextIndex
            outHasVisibleEvents = True
        End If
    Next eventIndex
End Sub

Private Function private_IsTouchedDiffEventType(ByVal rowType As String) As Boolean
    rowType = VBA.LCase$(VBA.Trim$(rowType))
    private_IsTouchedDiffEventType = (VBA.Len(rowType) > 0 And rowType <> "unchanged" And rowType <> "ellipsis")
End Function

Private Function private_AddDiffRowFromEvent( _
    ByVal leftTable As obj_TableDynamic, _
    ByVal rightTable As obj_TableDynamic, _
    ByVal diffRows As Collection, _
    ByVal rowType As String, _
    ByVal leftRowIndex As Long, _
    ByVal rightRowIndex As Long, _
    ByVal outputColumns As Collection, _
    ByVal compareColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean _
) As Boolean
    Dim leftRow As obj_Row
    Dim rightRow As obj_Row

    If leftRowIndex > 0 Then Set leftRow = leftTable.Rows.Item(leftRowIndex)
    If rightRowIndex > 0 Then Set rightRow = rightTable.Rows.Item(rightRowIndex)

    private_AddDiffRowFromEvent = private_AddDiffRow( _
        diffRows, _
        rowType, _
        leftRow, _
        rightRow, _
        outputColumns, _
        compareColumns, _
        ignoreCase, _
        trimText, _
        (VBA.StrComp(rowType, "modified", VBA.vbTextCompare) = 0))
End Function

Private Function private_TryMaterializeDiffEventsFromData( _
    ByVal leftData As obj_TableData, _
    ByVal rightData As obj_TableData, _
    ByVal outputColumns As Collection, _
    ByVal compareColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean, _
    ByRef eventTypes() As String, _
    ByRef eventLeftRows() As Long, _
    ByRef eventRightRows() As Long, _
    ByVal eventCount As Long, _
    ByVal condensedView As Boolean, _
    ByRef outDiffRows As Collection _
) As Boolean
    Dim keepEvents() As Boolean
    Dim eventIndex As Long
    Dim hasVisibleEvents As Boolean
    Dim inCollapsedRun As Boolean

    Set outDiffRows = New Collection
    If eventCount <= 0 Then
        private_TryMaterializeDiffEventsFromData = True
        Exit Function
    End If

    ' Condensed view не меняет сам diff. Он только скрывает длинные участки
    ' unchanged-строк, оставляя рядом с изменениями небольшой контекст и одну
    ' строку-заглушку "...", чтобы пользователь видел разрыв последовательности.
    If condensedView Then
        ReDim keepEvents(1 To eventCount)
        private_MarkCondensedVisibleEvents eventTypes, eventCount, keepEvents, hasVisibleEvents
    End If

    For eventIndex = 1 To eventCount
        If Not condensedView Then
            If Not private_AddDiffRowFromDataEvent( _
                leftData, rightData, outDiffRows, eventTypes(eventIndex), _
                eventLeftRows(eventIndex), eventRightRows(eventIndex), _
                outputColumns, compareColumns, ignoreCase, trimText) Then Exit Function
        ElseIf Not hasVisibleEvents Then
            If Not private_AddDiffRowFromDataEvent( _
                leftData, rightData, outDiffRows, eventTypes(eventIndex), _
                eventLeftRows(eventIndex), eventRightRows(eventIndex), _
                outputColumns, compareColumns, ignoreCase, trimText) Then Exit Function
        ElseIf keepEvents(eventIndex) Then
            If Not private_AddDiffRowFromDataEvent( _
                leftData, rightData, outDiffRows, eventTypes(eventIndex), _
                eventLeftRows(eventIndex), eventRightRows(eventIndex), _
                outputColumns, compareColumns, ignoreCase, trimText) Then Exit Function
            inCollapsedRun = False
        ElseIf Not inCollapsedRun Then
            outDiffRows.Add private_CreateEllipsisDiffRow(outputColumns.Count)
            inCollapsedRun = True
        End If
    Next eventIndex

    private_TryMaterializeDiffEventsFromData = True
End Function

Private Function private_AddDiffRowFromDataEvent( _
    ByVal leftData As obj_TableData, _
    ByVal rightData As obj_TableData, _
    ByVal diffRows As Collection, _
    ByVal rowType As String, _
    ByVal leftRowIndex As Long, _
    ByVal rightRowIndex As Long, _
    ByVal outputColumns As Collection, _
    ByVal compareColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean _
) As Boolean
    Dim rowInfo As Object
    Dim values As Collection
    Dim changedCols As Object
    Dim colIndex As Long
    Dim aliasText As String
    Dim sourceRowIndex As Long
    Dim useRightSource As Boolean
    Dim collectChangedCells As Boolean

    Set rowInfo = VBA.CreateObject("Scripting.Dictionary")
    Set values = New Collection
    Set changedCols = VBA.CreateObject("Scripting.Dictionary")
    changedCols.CompareMode = 1

    If rightRowIndex > 0 Then
        sourceRowIndex = rightRowIndex
        useRightSource = True
    Else
        sourceRowIndex = leftRowIndex
        useRightSource = False
    End If
    If sourceRowIndex <= 0 Then Exit Function

    ' OldRow/NewRow показывают реальные позиции строки в источниках. Перемещения
    ' остаются git-like парой deleted-moved/added-moved, поэтому каждая половина
    ' пары хранит только свой источник, но row type дает отдельный маркер стиля.
    collectChangedCells = _
        (leftRowIndex > 0 And _
         rightRowIndex > 0 And _
         VBA.StrComp(rowType, "modified", VBA.vbTextCompare) = 0)
    For colIndex = 1 To outputColumns.Count
        If useRightSource Then
            values.Add rightData.ValueAt(sourceRowIndex, colIndex)
        Else
            values.Add leftData.ValueAt(sourceRowIndex, colIndex)
        End If
        If collectChangedCells Then
            aliasText = VBA.Trim$(VBA.CStr(outputColumns.Item(colIndex)))
            If private_CollectionContains(compareColumns, aliasText) Then
                If private_NormalizeColumnCompareValue(leftData.ValueAt(leftRowIndex, colIndex), ignoreCase, trimText) <> _
                   private_NormalizeColumnCompareValue(rightData.ValueAt(rightRowIndex, colIndex), ignoreCase, trimText) Then
                    changedCols.Add VBA.CStr(colIndex), True
                End If
            End If
        End If
    Next colIndex

    rowInfo("RowType") = rowType
    rowInfo("OldRow") = VBA.vbNullString
    rowInfo("NewRow") = VBA.vbNullString
    If leftRowIndex > 0 Then rowInfo("OldRow") = VBA.CStr(leftRowIndex)
    If rightRowIndex > 0 Then rowInfo("NewRow") = VBA.CStr(rightRowIndex)
    Set rowInfo("Values") = values
    Set rowInfo("ChangedCols") = changedCols
    diffRows.Add rowInfo
    private_AddDiffRowFromDataEvent = True
End Function

Private Function private_TryBuildCompareColumnIndexSet( _
    ByVal outputColumns As Collection, _
    ByVal compareColumns As Collection, _
    ByRef outCompareColumnIndexes As Object _
) As Boolean
    Dim colIndex As Long
    Dim aliasText As String

    Set outCompareColumnIndexes = VBA.CreateObject("Scripting.Dictionary")
    outCompareColumnIndexes.CompareMode = 1
    If outputColumns Is Nothing Then Exit Function

    For colIndex = 1 To outputColumns.Count
        aliasText = VBA.Trim$(VBA.CStr(outputColumns.Item(colIndex)))
        If private_CollectionContains(compareColumns, aliasText) Then
            outCompareColumnIndexes(VBA.CStr(colIndex)) = True
        End If
    Next colIndex

    private_TryBuildCompareColumnIndexSet = True
End Function

Private Function private_TryBuildColumnIndexList( _
    ByVal outputColumns As Collection, _
    ByVal requestedColumns As Collection, _
    ByRef outColumnIndexes As Collection _
) As Boolean
    Dim requestedItem As Variant
    Dim requestedAlias As String
    Dim colIndex As Long
    Dim aliasText As String

    Set outColumnIndexes = New Collection
    If outputColumns Is Nothing Then Exit Function
    If requestedColumns Is Nothing Then Exit Function

    For Each requestedItem In requestedColumns
        requestedAlias = VBA.Trim$(VBA.CStr(requestedItem))
        If VBA.Len(requestedAlias) = 0 Then GoTo ContinueRequested
        If Not private_TryFindOutputColumnIndex(outputColumns, requestedAlias, colIndex) Then
            private_ShowCompareError "Column alias was not found in compare output columns: " & requestedAlias
            Exit Function
        End If
        outColumnIndexes.Add colIndex
ContinueRequested:
    Next requestedItem

    private_TryBuildColumnIndexList = True
End Function

Private Function private_TryFindOutputColumnIndex( _
    ByVal outputColumns As Collection, _
    ByVal aliasText As String, _
    ByRef outColumnIndex As Long _
) As Boolean
    Dim colIndex As Long

    outColumnIndex = 0
    If outputColumns Is Nothing Then Exit Function
    aliasText = VBA.Trim$(VBA.CStr(aliasText))
    For colIndex = 1 To outputColumns.Count
        If VBA.StrComp(VBA.Trim$(VBA.CStr(outputColumns.Item(colIndex))), aliasText, VBA.vbTextCompare) = 0 Then
            outColumnIndex = colIndex
            private_TryFindOutputColumnIndex = True
            Exit Function
        End If
    Next colIndex
End Function

Private Function private_TryBuildKeyIndexFromData( _
    ByVal tableData As obj_TableData, _
    ByVal keyColumnIndexes As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean, _
    ByRef outKeys() As String, _
    ByRef outIndex As Object, _
    ByVal tableLabel As String _
) As Boolean
    Dim rowIndex As Long
    Dim keyText As String

    Set outIndex = VBA.CreateObject("Scripting.Dictionary")
    outIndex.CompareMode = 0
    If tableData Is Nothing Then Exit Function
    If tableData.RowCount <= 0 Then
        ReDim outKeys(0 To 0)
        private_TryBuildKeyIndexFromData = True
        Exit Function
    End If
    ReDim outKeys(1 To tableData.RowCount)

    For rowIndex = 1 To tableData.RowCount
        keyText = private_BuildRowKeyFromData(tableData, rowIndex, keyColumnIndexes, ignoreCase, trimText)
        If VBA.Len(keyText) = 0 Then
            private_ShowCompareError "Empty key in " & tableLabel & " table at row " & VBA.CStr(rowIndex) & "."
            Exit Function
        End If
        If outIndex.Exists(keyText) Then
            private_ShowCompareError "Duplicate key in " & tableLabel & " table: " & keyText
            Exit Function
        End If
        outKeys(rowIndex) = keyText
        outIndex.Add keyText, rowIndex
    Next rowIndex

    private_TryBuildKeyIndexFromData = True
End Function

Private Function private_BuildRowKeyFromData( _
    ByVal tableData As obj_TableData, _
    ByVal rowIndex As Long, _
    ByVal keyColumnIndexes As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean _
) As String
    Dim keyItem As Variant
    Dim colIndex As Long
    Dim partText As String

    If tableData Is Nothing Then Exit Function
    If keyColumnIndexes Is Nothing Then Exit Function
    For Each keyItem In keyColumnIndexes
        colIndex = CLng(keyItem)
        partText = private_NormalizeColumnCompareValue(tableData.ValueAt(rowIndex, colIndex), ignoreCase, trimText)
        If VBA.Len(private_BuildRowKeyFromData) > 0 Then private_BuildRowKeyFromData = private_BuildRowKeyFromData & VBA.ChrW$(30)
        private_BuildRowKeyFromData = private_BuildRowKeyFromData & partText
    Next keyItem
End Function

Private Function private_TryBuildKeyIndex( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal keyColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean, _
    ByRef outKeys() As String, _
    ByRef outIndex As Object, _
    ByVal tableLabel As String _
) As Boolean
    Dim rowIndex As Long
    Dim keyText As String

    Set outIndex = VBA.CreateObject("Scripting.Dictionary")
    outIndex.CompareMode = 0
    If tableObj Is Nothing Then Exit Function
    If tableObj.RowCount <= 0 Then
        ReDim outKeys(0 To 0)
        private_TryBuildKeyIndex = True
        Exit Function
    End If
    ReDim outKeys(1 To tableObj.RowCount)

    For rowIndex = 1 To tableObj.RowCount
        keyText = private_BuildRowKey(tableObj.Rows.Item(rowIndex), tableObj, keyColumns, ignoreCase, trimText)
        If VBA.Len(keyText) = 0 Then
            private_ShowCompareError "Empty key in " & tableLabel & " table at row " & VBA.CStr(rowIndex) & "."
            Exit Function
        End If
        If outIndex.Exists(keyText) Then
            private_ShowCompareError "Duplicate key in " & tableLabel & " table: " & keyText
            Exit Function
        End If
        outKeys(rowIndex) = keyText
        outIndex.Add keyText, rowIndex
    Next rowIndex

    private_TryBuildKeyIndex = True
End Function

Private Function private_TryBuildLcsKeys( _
    ByRef leftKeys() As String, _
    ByRef rightKeys() As String, _
    ByRef outLcsKeys As Collection _
) As Boolean
    Dim n As Long
    Dim m As Long
    Dim i As Long
    Dim matchCount As Long
    Dim keyText As String
    Dim rightPositions As Object
    Dim seqPositions() As Long
    Dim seqKeys() As String
    Dim tailsPositions() As Long
    Dim tailsIndexes() As Long
    Dim previousIndexes() As Long
    Dim lisLength As Long
    Dim insertAt As Long
    Dim currentIndex As Long
    Dim reversed As Collection

    Set outLcsKeys = New Collection
    n = private_ArrayLength(leftKeys)
    m = private_ArrayLength(rightKeys)
    If n <= 0 Or m <= 0 Then
        private_TryBuildLcsKeys = True
        Exit Function
    End If

    ' Ключи уже проверены на уникальность в каждой таблице, поэтому LCS можно считать
    ' как LIS по позициям правой таблицы. Это сохраняет порядок строк, но не строит
    ' огромную матрицу n*m, которая на 10k строк становится слишком дорогой.
    Set rightPositions = VBA.CreateObject("Scripting.Dictionary")
    rightPositions.CompareMode = 0
    For i = 1 To m
        rightPositions.Add rightKeys(i), i
    Next i

    ReDim seqPositions(1 To n)
    ReDim seqKeys(1 To n)
    For i = 1 To n
        keyText = leftKeys(i)
        If rightPositions.Exists(keyText) Then
            matchCount = matchCount + 1
            seqPositions(matchCount) = CLng(rightPositions(keyText))
            seqKeys(matchCount) = keyText
        End If
    Next i

    If matchCount <= 0 Then
        private_TryBuildLcsKeys = True
        Exit Function
    End If

    ReDim tailsPositions(1 To matchCount)
    ReDim tailsIndexes(1 To matchCount)
    ReDim previousIndexes(1 To matchCount)

    For i = 1 To matchCount
        insertAt = private_FindLisInsertPosition(tailsPositions, lisLength, seqPositions(i))
        tailsPositions(insertAt) = seqPositions(i)
        tailsIndexes(insertAt) = i
        If insertAt > 1 Then previousIndexes(i) = tailsIndexes(insertAt - 1)
        If insertAt > lisLength Then lisLength = insertAt
    Next i

    Set reversed = New Collection
    currentIndex = tailsIndexes(lisLength)
    Do While currentIndex > 0
        reversed.Add seqKeys(currentIndex)
        currentIndex = previousIndexes(currentIndex)
    Loop

    For i = reversed.Count To 1 Step -1
        outLcsKeys.Add VBA.CStr(reversed.Item(i))
    Next i

    private_TryBuildLcsKeys = True
End Function

Private Function private_FindLisInsertPosition( _
    ByRef tailsPositions() As Long, _
    ByVal lisLength As Long, _
    ByVal value As Long _
) As Long
    Dim lo As Long
    Dim hi As Long
    Dim mid As Long

    If lisLength <= 0 Then
        private_FindLisInsertPosition = 1
        Exit Function
    End If

    lo = 1
    hi = lisLength
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If tailsPositions(mid) < value Then
            lo = mid + 1
        Else
            hi = mid - 1
        End If
    Loop

    private_FindLisInsertPosition = lo
End Function

Private Function private_AddDiffRow( _
    ByVal diffRows As Collection, _
    ByVal rowType As String, _
    ByVal leftRow As obj_Row, _
    ByVal rightRow As obj_Row, _
    ByVal outputColumns As Collection, _
    ByVal compareColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean, _
    ByVal collectChangedCells As Boolean _
) As Boolean
    Dim rowInfo As Object
    Dim values As Collection
    Dim changedCols As Object
    Dim sourceRow As obj_Row
    Dim colIndex As Long
    Dim aliasText As String

    Set rowInfo = VBA.CreateObject("Scripting.Dictionary")
    Set values = New Collection
    Set changedCols = VBA.CreateObject("Scripting.Dictionary")
    changedCols.CompareMode = 1

    If rightRow Is Nothing Then
        Set sourceRow = leftRow
    Else
        Set sourceRow = rightRow
    End If
    If sourceRow Is Nothing Then Exit Function

    For colIndex = 1 To outputColumns.Count
        values.Add sourceRow.GetCellValue(colIndex)
        aliasText = VBA.Trim$(VBA.CStr(outputColumns.Item(colIndex)))
        If collectChangedCells Then
            If private_CollectionContains(compareColumns, aliasText) Then
                If private_NormalizeCompareValue(leftRow.GetCellValue(colIndex), ignoreCase, trimText) <> _
                   private_NormalizeCompareValue(rightRow.GetCellValue(colIndex), ignoreCase, trimText) Then
                    changedCols.Add VBA.CStr(colIndex), True
                End If
            End If
        End If
    Next colIndex

    rowInfo("RowType") = rowType
    rowInfo("OldRow") = VBA.vbNullString
    rowInfo("NewRow") = VBA.vbNullString
    Set rowInfo("Values") = values
    Set rowInfo("ChangedCols") = changedCols
    diffRows.Add rowInfo
    private_AddDiffRow = True
End Function

Private Function private_TryBuildDiffTableItems( _
    ByVal outputColumns As Collection, _
    ByVal outputColumnFormats As Object, _
    ByVal diffRows As Collection, _
    ByRef outTableItems As Collection _
) As Boolean
    Dim tableObj As obj_TableDynamic
    Dim colObj As obj_Column
    Dim rowObj As obj_Row
    Dim cellObj As obj_Cell
    Dim rowInfo As Object
    Dim values As Collection
    Dim changedCols As Object
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim columnAlias As String
    Dim rawValue As Variant

    Set outTableItems = Nothing
    If outputColumns Is Nothing Then Exit Function

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = "Comparison result"

    ' UI-таблица получает две служебные колонки OldRow/NewRow. Они не участвуют
    ' в сравнении и не сдвигают changedCols: подсветка измененных ячеек продолжает
    ' использовать индексы исходных outputColumns.
    Set colObj = New obj_Column
    colObj.Position = 1
    colObj.Name = OLD_ROW_COLUMN_ALIAS
    Call colObj.AddAlias(OLD_ROW_COLUMN_ALIAS)
    If Not tableObj.PushColumn(colObj) Then Exit Function

    Set colObj = New obj_Column
    colObj.Position = 2
    colObj.Name = NEW_ROW_COLUMN_ALIAS
    Call colObj.AddAlias(NEW_ROW_COLUMN_ALIAS)
    If Not tableObj.PushColumn(colObj) Then Exit Function

    For colIndex = 1 To outputColumns.Count
        columnAlias = VBA.Trim$(VBA.CStr(outputColumns.Item(colIndex)))
        Set colObj = New obj_Column
        colObj.Position = colIndex + 2
        colObj.Name = columnAlias
        colObj.FormatKind = private_GetColumnFormatKind(outputColumnFormats, columnAlias)
        If VBA.Len(columnAlias) > 0 Then Call colObj.AddAlias(columnAlias)
        If Not tableObj.PushColumn(colObj) Then Exit Function
    Next colIndex

    If Not diffRows Is Nothing Then
        For rowIndex = 1 To diffRows.Count
            Set rowInfo = diffRows.Item(rowIndex)
            Set values = rowInfo("Values")
            Set changedCols = rowInfo("ChangedCols")

            Set rowObj = New obj_Row
            rowObj.Desc = "diff:" & VBA.LCase$(VBA.Trim$(VBA.CStr(rowInfo("RowType"))))

            Set cellObj = New obj_Cell
            cellObj.Value = VBA.vbNullString
            If rowInfo.Exists("OldRow") Then cellObj.Value = VBA.CStr(rowInfo("OldRow"))
            If Not rowObj.PushCell(cellObj) Then Exit Function

            Set cellObj = New obj_Cell
            cellObj.Value = VBA.vbNullString
            If rowInfo.Exists("NewRow") Then cellObj.Value = VBA.CStr(rowInfo("NewRow"))
            If Not rowObj.PushCell(cellObj) Then Exit Function

            For colIndex = 1 To outputColumns.Count
                rawValue = values.Item(colIndex)
                Set cellObj = New obj_Cell
                cellObj.Value = VBA.CStr(rawValue)
                If Not changedCols Is Nothing Then
                    If changedCols.Exists(VBA.CStr(colIndex)) Then cellObj.Desc = "diff:changed"
                End If
                If Not rowObj.PushCell(cellObj) Then Exit Function
            Next colIndex

            If Not tableObj.PushRow(rowObj) Then Exit Function
        Next rowIndex
    End If

    Set outTableItems = New Collection
    outTableItems.Add tableObj
    private_TryBuildDiffTableItems = True
End Function

Private Function private_TryRefreshDiffTableItemsSource(ByRef outVisibleRowCount As Long) As Boolean
    Dim visibleDiffRows As Collection
    Dim diffTableItems As Collection
    Dim stageStart As Single

    outVisibleRowCount = 0
    If m_LastOutputColumns Is Nothing Then Exit Function
    If m_LastDiffRows Is Nothing Then Exit Function

    stageStart = VBA.Timer
    Set visibleDiffRows = m_LastDiffRows
    If Not visibleDiffRows Is Nothing Then outVisibleRowCount = visibleDiffRows.Count
    private_LogPipelineStep "refresh-diff-source:visible-rows", stageStart, _
        "sourceRows=" & VBA.CStr(m_LastDiffRows.Count) & _
        " visibleRows=" & VBA.CStr(outVisibleRowCount) & _
        " mode='" & private_GetDiffViewModeName() & "'"

    stageStart = VBA.Timer
    If Not private_TryBuildDiffTableItems(m_LastOutputColumns, m_LastOutputColumnFormats, visibleDiffRows, diffTableItems) Then Exit Function
    private_LogPipelineStep "refresh-diff-source:build-table-items", stageStart, _
        "visibleRows=" & VBA.CStr(outVisibleRowCount) & _
        " columns=" & VBA.CStr(m_LastOutputColumns.Count + 2)

    stageStart = VBA.Timer
    If Not private_TrySetDiffTableItemsSource(diffTableItems) Then Exit Function
    private_LogPipelineStep "refresh-diff-source:set-items-source", stageStart, _
        "tables=" & VBA.CStr(diffTableItems.Count)

    private_TryRefreshDiffTableItemsSource = True
End Function

Private Function private_CreateEllipsisDiffRow(ByVal outputColumnCount As Long) As Object
    Dim rowInfo As Object
    Dim values As Collection
    Dim changedCols As Object
    Dim colIndex As Long
    Dim ellipsisColumn As Long

    Set rowInfo = VBA.CreateObject("Scripting.Dictionary")
    Set values = New Collection
    Set changedCols = VBA.CreateObject("Scripting.Dictionary")
    changedCols.CompareMode = 1

    ellipsisColumn = (outputColumnCount + 1) \ 2
    If ellipsisColumn <= 0 Then ellipsisColumn = 1

    For colIndex = 1 To outputColumnCount
        If colIndex = ellipsisColumn Then
            values.Add "..."
        Else
            values.Add VBA.vbNullString
        End If
    Next colIndex

    rowInfo("RowType") = "ellipsis"
    rowInfo("OldRow") = VBA.vbNullString
    rowInfo("NewRow") = VBA.vbNullString
    Set rowInfo("Values") = values
    Set rowInfo("ChangedCols") = changedCols
    Set private_CreateEllipsisDiffRow = rowInfo
End Function

Private Function private_GetDiffViewModeName() As String
    If m_IsCondensedDiffView Then
        private_GetDiffViewModeName = "condensed"
    Else
        private_GetDiffViewModeName = "full"
    End If
End Function

Private Function private_TrySetDiffTableItemsSource(ByVal tableItems As Collection) As Boolean
    Dim pageBase As obj_PageBase

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If pageBase.RuntimeSources Is Nothing Then Exit Function
    If tableItems Is Nothing Then Set tableItems = New Collection

    private_TrySetDiffTableItemsSource = pageBase.RuntimeSources.SetItemsSource(DIFF_TABLES_ITEMS_SOURCE_KEY, tableItems, False)
End Function

Private Function private_RowHasCompareChanges( _
    ByVal leftRow As obj_Row, _
    ByVal rightRow As obj_Row, _
    ByVal outputColumns As Collection, _
    ByVal compareColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean _
) As Boolean
    Dim colIndex As Long
    Dim aliasText As String

    If leftRow Is Nothing Then Exit Function
    If rightRow Is Nothing Then Exit Function

    For colIndex = 1 To outputColumns.Count
        aliasText = VBA.Trim$(VBA.CStr(outputColumns.Item(colIndex)))
        If private_CollectionContains(compareColumns, aliasText) Then
            If private_NormalizeCompareValue(leftRow.GetCellValue(colIndex), ignoreCase, trimText) <> _
               private_NormalizeCompareValue(rightRow.GetCellValue(colIndex), ignoreCase, trimText) Then
                private_RowHasCompareChanges = True
                Exit Function
            End If
        End If
    Next colIndex
End Function

Private Function private_RowHasCompareChangesIndexed( _
    ByVal leftRow As obj_Row, _
    ByVal rightRow As obj_Row, _
    ByVal outputColumns As Collection, _
    ByVal compareColumnIndexes As Object, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean _
) As Boolean
    Dim colIndex As Long

    If leftRow Is Nothing Then Exit Function
    If rightRow Is Nothing Then Exit Function
    If outputColumns Is Nothing Then Exit Function
    If compareColumnIndexes Is Nothing Then Exit Function

    For colIndex = 1 To outputColumns.Count
        If compareColumnIndexes.Exists(VBA.CStr(colIndex)) Then
            If private_NormalizeCompareValue(leftRow.GetCellValue(colIndex), ignoreCase, trimText) <> _
               private_NormalizeCompareValue(rightRow.GetCellValue(colIndex), ignoreCase, trimText) Then
                private_RowHasCompareChangesIndexed = True
                Exit Function
            End If
        End If
    Next colIndex
End Function

Private Function private_RowHasCompareChangesInData( _
    ByVal leftData As obj_TableData, _
    ByVal leftRowIndex As Long, _
    ByVal rightData As obj_TableData, _
    ByVal rightRowIndex As Long, _
    ByVal outputColumns As Collection, _
    ByVal compareColumnIndexes As Object, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean _
) As Boolean
    Dim colIndex As Long

    If leftData Is Nothing Then Exit Function
    If rightData Is Nothing Then Exit Function
    If outputColumns Is Nothing Then Exit Function
    If compareColumnIndexes Is Nothing Then Exit Function

    For colIndex = 1 To outputColumns.Count
        If compareColumnIndexes.Exists(VBA.CStr(colIndex)) Then
            If private_NormalizeColumnCompareValue(leftData.ValueAt(leftRowIndex, colIndex), ignoreCase, trimText) <> _
               private_NormalizeColumnCompareValue(rightData.ValueAt(rightRowIndex, colIndex), ignoreCase, trimText) Then
                private_RowHasCompareChangesInData = True
                Exit Function
            End If
        End If
    Next colIndex
End Function

Private Function private_BuildRowKey( _
    ByVal rowObj As obj_Row, _
    ByVal tableObj As obj_TableDynamic, _
    ByVal keyColumns As Collection, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean _
) As String
    Dim keyItem As Variant
    Dim keyAlias As String
    Dim colIndex As Long
    Dim partText As String

    If rowObj Is Nothing Then Exit Function
    If tableObj Is Nothing Then Exit Function

    For Each keyItem In keyColumns
        keyAlias = VBA.Trim$(VBA.CStr(keyItem))
        If VBA.Len(keyAlias) = 0 Then GoTo ContinueKey
        If Not tableObj.TryGetColumnIndexByAlias(keyAlias, colIndex) Then
            private_ShowCompareError "Key column alias was not found in loaded table: " & keyAlias
            Exit Function
        End If
        partText = private_NormalizeCompareValue(rowObj.GetCellValue(colIndex), ignoreCase, trimText)
        If VBA.Len(private_BuildRowKey) > 0 Then private_BuildRowKey = private_BuildRowKey & VBA.ChrW$(30)
        private_BuildRowKey = private_BuildRowKey & partText
ContinueKey:
    Next keyItem
End Function

Private Function private_NormalizeCompareValue( _
    ByVal valueText As String, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean _
) As String
    private_NormalizeCompareValue = VBA.CStr(valueText)
    If trimText Then
        private_NormalizeCompareValue = VBA.Replace$(private_NormalizeCompareValue, VBA.vbCr, " ")
        private_NormalizeCompareValue = VBA.Replace$(private_NormalizeCompareValue, VBA.vbLf, " ")
        private_NormalizeCompareValue = VBA.Replace$(private_NormalizeCompareValue, VBA.vbTab, " ")
        private_NormalizeCompareValue = VBA.Replace$(private_NormalizeCompareValue, VBA.ChrW$(160), " ")
        private_NormalizeCompareValue = VBA.Replace$(private_NormalizeCompareValue, "  ", " ")
        private_NormalizeCompareValue = VBA.Replace$(private_NormalizeCompareValue, "  ", " ")
        private_NormalizeCompareValue = VBA.Trim$(private_NormalizeCompareValue)
    End If
    If ignoreCase Then private_NormalizeCompareValue = VBA.LCase$(private_NormalizeCompareValue)
End Function

Private Function private_NormalizeColumnCompareValue( _
    ByVal valueIn As Variant, _
    ByVal ignoreCase As Boolean, _
    ByVal trimText As Boolean _
) As String
    If VBA.IsError(valueIn) Then
        private_NormalizeColumnCompareValue = "#ERR"
        Exit Function
    End If
    If VBA.IsNull(valueIn) Or VBA.IsEmpty(valueIn) Then
        private_NormalizeColumnCompareValue = VBA.vbNullString
        Exit Function
    End If

    ' Сравнение идет по фактическому значению, которое пришло из SQL/Excel,
    ' а текстовые правила ignoreCase/trimText применяются уже поверх общего
    ' CStr-представления.
    private_NormalizeColumnCompareValue = private_NormalizeCompareValue(VBA.CStr(valueIn), ignoreCase, trimText)
End Function

Private Function private_GetColumnFormatKind(ByVal columnFormats As Object, ByVal columnAlias As String) As String
    columnAlias = VBA.Trim$(VBA.CStr(columnAlias))
    If VBA.Len(columnAlias) = 0 Then Exit Function
    If columnFormats Is Nothing Then Exit Function
    If columnFormats.Exists(columnAlias) Then private_GetColumnFormatKind = VBA.LCase$(VBA.Trim$(VBA.CStr(columnFormats(columnAlias))))
End Function

Private Function private_CollectionContains(ByVal items As Collection, ByVal valueText As String) As Boolean
    Dim item As Variant

    If items Is Nothing Then Exit Function
    valueText = VBA.Trim$(VBA.CStr(valueText))
    For Each item In items
        If VBA.StrComp(VBA.Trim$(VBA.CStr(item)), valueText, VBA.vbTextCompare) = 0 Then
            private_CollectionContains = True
            Exit Function
        End If
    Next item
End Function

Private Function private_ArrayLength(ByRef values() As String) As Long
    Dim lowerBound As Long
    Dim upperBound As Long

    On Error GoTo EH
    lowerBound = LBound(values)
    upperBound = UBound(values)
    If lowerBound = 0 Then
        If upperBound = 0 Then
            If VBA.Len(values(0)) = 0 Then
                private_ArrayLength = 0
                Exit Function
            End If
        End If
    End If
    private_ArrayLength = upperBound - lowerBound + 1
    Exit Function
EH:
    private_ArrayLength = 0
End Function

Private Sub private_LogPipelineStep( _
    ByVal stepName As String, _
    ByVal startedAt As Single, _
    Optional ByVal details As String = "" _
)
#If LOGGING_DEBUG_ENABLED Then
    Dim messageText As String

    messageText = "comparing:pipeline step='" & private_LogSafeText(stepName) & _
        "' ms=" & VBA.Format$(private_ElapsedMs(startedAt, VBA.Timer), "0")
    details = VBA.Trim$(VBA.CStr(details))
    If VBA.Len(details) > 0 Then messageText = messageText & " " & details
    ex_Core.fn_Diagnostic_LogInfo messageText
#End If
End Sub

Private Function private_ElapsedMs(ByVal startedAt As Single, ByVal finishedAt As Single) As Double
    If finishedAt < startedAt Then finishedAt = finishedAt + 86400!
    private_ElapsedMs = (CDbl(finishedAt) - CDbl(startedAt)) * 1000#
End Function

Private Function private_LogSafeText(ByVal valueText As String) As String
    private_LogSafeText = VBA.Replace$(VBA.CStr(valueText), "'", "''")
End Function

Private Sub private_ShowCompareError(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PageComparingCtrl: " & VBA.CStr(messageText)
#End If
    VBA.MsgBox "PrototypeNew: " & VBA.CStr(messageText), vbExclamation, RUNTIME_ERROR_TITLE
End Sub

Public Function RerenderPage(Optional ByVal notifyStatus As Boolean = True) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageComparingCtrl.RerenderPage"
#End If
    If m_Page Is Nothing Then Exit Function
    If Not rt_PageManager.fn_RenderPage(m_Page, "comparing:manual-rerender") Then Exit Function

    If notifyStatus Then
        rt_Messaging.fn_ShowStatusBarNotice "Comparing page has been rerendered.", 2
    End If
    RerenderPage = True
End Function
