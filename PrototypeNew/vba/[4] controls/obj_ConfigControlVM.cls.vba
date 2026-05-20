VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IControl

Private Const CONFIG_COL_COUNT As Long = 3

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_Page As obj_IPage
Private m_RuntimeControlKey As String
Private m_RuntimeTableName As String
Private m_ItemsSourceRaw As String
Private m_TableNameRaw As String
Private m_ControlLayout As obj_ControlLayout
Private m_ConfigTableViewItem As obj_ConfigTableViewItem
Private m_IsConfigured As Boolean

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
    obj_IControl_Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IControl_Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    m_IsConfigured = False
    Set m_Page = page
    obj_IControl_Initialize = True
End Function

Private Sub obj_IControl_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_ControlLayout = Nothing
    Set m_ConfigTableViewItem = Nothing
    Set m_Page = Nothing
    m_ControlName = VBA.vbNullString
    m_ItemsSourceRaw = VBA.vbNullString
    m_TableNameRaw = VBA.vbNullString
    m_RuntimeControlKey = VBA.vbNullString
    m_RuntimeTableName = VBA.vbNullString
    m_IsConfigured = False
    On Error GoTo 0
End Sub

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim pageBase As obj_PageBase
    Dim resolvedItems As Collection
    Dim configTable As obj_ConfigTable

    m_IsConfigured = False
    Set m_ControlLayout = Nothing
    Set m_ConfigTableViewItem = Nothing
    Set m_ControlBase = Nothing
    m_RuntimeControlKey = VBA.vbNullString
    m_RuntimeTableName = VBA.vbNullString

    Set pageBase = m_Page.GetPageBase()
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "Config", "config", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: itemsSource is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    m_TableNameRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "tableName")))

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Config", m_ControlName, "style") Then Exit Sub

    If (m_ControlLayout.ColEnd - m_ControlLayout.ColStart + 1) < CONFIG_COL_COUNT Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: control '" & m_ControlName & "' requires at least 3 columns (Attr, Key, Value)."
#End If
        Exit Sub
    End If

    Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then Exit Sub
    m_RuntimeControlKey = private_BuildRuntimeControlKey()
    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, m_ItemsSourceRaw, resolvedItems) Then Exit Sub
    If Not private_TryBuildConfigTable(resolvedItems, configTable) Then Exit Sub

    Set m_ConfigTableViewItem = New obj_ConfigTableViewItem
    If Not m_ConfigTableViewItem.Initialize(configTable) Then Exit Sub

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim boundsRange As Range
    Dim writeRange As Range
    Dim valueBlock As Variant
    Dim rowsToWrite As Long
    Dim maxRows As Long
    Dim dataRows As Long
    Dim idx As Long
    Dim rowOut As Long
    Dim configEntryViewItem As obj_ConfigEntryViewItem
    Dim entryItems As list__obj_ConfigEntryViewItem
    Dim configEntry As obj_ConfigEntry
    Dim attrToken As String
    Dim absRow As Long
    Dim tableObj As ListObject
    Dim targetTableName As String
    Dim hashRows As Collection
    Dim rxRows As Collection
    Dim page As obj_PageBase

    ' Базовые проверки: контрол должен быть настроен и привязан к странице/листу.
    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If

    Set page = Nothing
    If Not m_ControlBase Is Nothing Then Set page = m_ControlBase.PageBase
    If page Is Nothing Then
        Set page = m_Page.GetPageBase()
    End If
    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: page is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(page, m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: sheet '" & m_ControlLayout.LayoutSheetName & "' was not found for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    If m_ConfigTableViewItem Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: view item is not configured for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If
    If Not m_ConfigTableViewItem.TryResyncEntryItemsFromModel() Then Exit Sub
    Set entryItems = m_ConfigTableViewItem.EntryItems
    If entryItems Is Nothing Then Exit Sub

    ' Собираем номера строк для специальных style-part (attrhash/attrrx).
    Set hashRows = New Collection
    Set rxRows = New Collection

    ' Ограничиваем объем вывода доступной высотой контрола.
    maxRows = m_ControlLayout.RowEnd - m_ControlLayout.RowStart + 1
    dataRows = entryItems.Count
    rowsToWrite = 1 + dataRows
    If rowsToWrite < 2 Then rowsToWrite = 2
    If rowsToWrite > maxRows Then rowsToWrite = maxRows

    ' Готовим буфер данных целиком и одной операцией выгружаем его в диапазон.
    ReDim valueBlock(1 To rowsToWrite, 1 To CONFIG_COL_COUNT)
    valueBlock(1, 1) = "Attr"
    valueBlock(1, 2) = "Key"
    valueBlock(1, 3) = "Value"

    For idx = 1 To entryItems.Count
        Set configEntryViewItem = entryItems.Item(idx)
        rowOut = idx + 1
        If rowOut > rowsToWrite Then Exit For

        If configEntryViewItem Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Config: entry view item is Nothing in control '" & m_ControlName & "'."
#End If
            GoTo ContinueItem
        End If

        Set configEntry = configEntryViewItem.Model
        If configEntry Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Config: entry view item has no model in control '" & m_ControlName & "'."
#End If
            GoTo ContinueItem
        End If

        valueBlock(rowOut, 1) = configEntry.Attr
        valueBlock(rowOut, 2) = configEntry.Key
        valueBlock(rowOut, 3) = configEntry.Value

        absRow = m_ControlLayout.RowStart + rowOut - 1
        attrToken = VBA.LCase$(VBA.Trim$(configEntry.Attr))
        Select Case attrToken
            Case "#"
                private_HandleAttrHash configEntry
                hashRows.Add absRow

            Case "rx"
                private_HandleAttrRx configEntry
                rxRows.Add absRow
        End Select

ContinueItem:
    Next idx

    ' Чистим целевую область и удаляем пересекающиеся таблицы, чтобы избежать конфликтов ListObject.
    Set boundsRange = ws.Range( _
        ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart), _
        ws.Cells(m_ControlLayout.RowEnd, m_ControlLayout.ColStart + CONFIG_COL_COUNT - 1))

    If Not private_TryDeleteIntersectingTables(ws, boundsRange) Then Exit Sub

    boundsRange.UnMerge
    boundsRange.ClearContents

    Set writeRange = ws.Range( _
        ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart), _
        ws.Cells(m_ControlLayout.RowStart + rowsToWrite - 1, m_ControlLayout.ColStart + CONFIG_COL_COUNT - 1))
    writeRange.Value2 = valueBlock

    ' Преобразуем диапазон в ListObject, чтобы получить табличный рендер и фильтры Excel.
    On Error GoTo EH_TABLE
    Set tableObj = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=writeRange, XlListObjectHasHeaders:=xlYes)
    On Error GoTo 0

    targetTableName = private_BuildTableName(ws)
    If VBA.Len(targetTableName) > 0 Then
        On Error Resume Next
        tableObj.Name = targetTableName
        On Error GoTo 0
    End If

    On Error Resume Next
    tableObj.TableStyle = "TableStyleMedium2"
    tableObj.ShowAutoFilter = True
    On Error GoTo 0
    m_RuntimeTableName = VBA.Trim$(tableObj.Name)

    ' Регистрируем именованные части колонок, чтобы pipeline мог адресно задавать их стиль/ширину.
    If Not private_RegisterColumnPart(ws, "attr", writeRange.Columns(1)) Then Exit Sub
    If Not private_RegisterColumnPart(ws, "key", writeRange.Columns(2)) Then Exit Sub
    If Not private_RegisterColumnPart(ws, "value", writeRange.Columns(3)) Then Exit Sub

    ' Регистрируем строковые части для специальных атрибутов (# и rx).
    If Not private_RegisterAttrRows(ws, "attrhash", hashRows) Then Exit Sub
    If Not private_RegisterAttrRows(ws, "attrrx", rxRows) Then Exit Sub
    If Not private_TryRegisterRuntimeControl() Then Exit Sub
    Exit Sub

EH_TABLE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Config: failed to create table for control '" & m_ControlName & "': " & Err.Description
#End If
End Sub

Private Function private_RegisterColumnPart( _
    ByVal ws As Worksheet, _
    ByVal partName As String, _
    ByVal columnRange As Range _
) As Boolean
    If ws Is Nothing Then Exit Function
    If columnRange Is Nothing Then Exit Function

    If Not ex_ControlPartsRuntime.fn_RegisterControlPart( _
        ws, _
        "config", _
        m_ControlName, _
        VBA.LCase$(VBA.Trim$(partName)), _
        columnRange) Then Exit Function

    private_RegisterColumnPart = True
End Function

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource", "tablename"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

' //
' // API
' //
Public Function TrySetRenderedValueByKey( _
    ByVal configKey As String, _
    ByVal newValue As String _
) As Boolean
    Dim tableObj As ListObject
    Dim headerRange As Range
    Dim dataRange As Range
    Dim colCount As Long
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim keyColumnIndex As Long
    Dim valueColumnIndex As Long
    Dim rawHeader As String
    Dim normalizedHeader As String
    Dim keyCellText As String

    configKey = VBA.Trim$(configKey)
    If VBA.Len(configKey) = 0 Then Exit Function

    If Not private_TryResolveRenderedTableObject(tableObj) Then Exit Function
    Set headerRange = tableObj.HeaderRowRange
    If headerRange Is Nothing Then Exit Function

    colCount = headerRange.Columns.Count
    If colCount <= 0 Then Exit Function

    For colIndex = 1 To colCount
        rawHeader = VBA.Trim$(VBA.CStr(headerRange.Cells(1, colIndex).Value2))
        If Not private_TryNormalizeProfileAttrName(rawHeader, normalizedHeader) Then Exit Function

        Select Case VBA.LCase$(normalizedHeader)
            Case "key"
                keyColumnIndex = colIndex
            Case "value"
                valueColumnIndex = colIndex
        End Select
    Next colIndex

    If keyColumnIndex <= 0 Then Exit Function
    If valueColumnIndex <= 0 Then Exit Function

    Set dataRange = tableObj.DataBodyRange
    If dataRange Is Nothing Then Exit Function

    For rowIndex = 1 To dataRange.Rows.Count
        keyCellText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, keyColumnIndex).Value2))
        If VBA.StrComp(keyCellText, configKey, VBA.vbTextCompare) <> 0 Then GoTo ContinueRow
        dataRange.Cells(rowIndex, valueColumnIndex).Value2 = newValue
        TrySetRenderedValueByKey = True
        Exit Function
ContinueRow:
    Next rowIndex
End Function

Public Function TryGetRenderedConfigEntries(ByRef outEntries As Collection) As Boolean
    Dim tableObj As ListObject
    Dim headerRange As Range
    Dim dataRange As Range
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim attrColumnIndex As Long
    Dim keyColumnIndex As Long
    Dim valueColumnIndex As Long
    Dim rawHeader As String
    Dim normalizedAttrName As String
    Dim cellText As String
    Dim rowHasAnyValue As Boolean
    Dim attrText As String
    Dim keyText As String
    Dim valueText As String
    Dim configEntry As obj_ConfigEntry

    Set outEntries = New Collection

    If Not private_TryResolveRenderedTableObject(tableObj) Then Exit Function

    Set headerRange = tableObj.HeaderRowRange
    If headerRange Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' has no header row for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    colCount = headerRange.Columns.Count
    If colCount <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' has no columns for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    For colIndex = 1 To colCount
        rawHeader = VBA.Trim$(VBA.CStr(headerRange.Cells(1, colIndex).Value2))
        If Not private_TryNormalizeProfileAttrName(rawHeader, normalizedAttrName) Then Exit Function

        Select Case VBA.LCase$(normalizedAttrName)
            Case "attr"
                If attrColumnIndex > 0 Then
#If LOGGING_DEBUG_ENABLED Then
                    ex_Core.fn_Diagnostic_LogError "Config: duplicate 'Attr' header in table '" & tableObj.Name & "' for control '" & m_ControlName & "'."
#End If
                    Exit Function
                End If
                attrColumnIndex = colIndex

            Case "key"
                If keyColumnIndex > 0 Then
#If LOGGING_DEBUG_ENABLED Then
                    ex_Core.fn_Diagnostic_LogError "Config: duplicate 'Key' header in table '" & tableObj.Name & "' for control '" & m_ControlName & "'."
#End If
                    Exit Function
                End If
                keyColumnIndex = colIndex

            Case "value"
                If valueColumnIndex > 0 Then
#If LOGGING_DEBUG_ENABLED Then
                    ex_Core.fn_Diagnostic_LogError "Config: duplicate 'Value' header in table '" & tableObj.Name & "' for control '" & m_ControlName & "'."
#End If
                    Exit Function
                End If
                valueColumnIndex = colIndex
        End Select
    Next colIndex

    If keyColumnIndex <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' must contain 'Key' column for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    Set dataRange = tableObj.DataBodyRange
    If dataRange Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' has no data rows for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    rowCount = dataRange.Rows.Count
    For rowIndex = 1 To rowCount
        rowHasAnyValue = False
        For colIndex = 1 To colCount
            cellText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, colIndex).Value2))
            If VBA.Len(cellText) > 0 Then
                rowHasAnyValue = True
                Exit For
            End If
        Next colIndex
        If Not rowHasAnyValue Then GoTo ContinueRow

        If attrColumnIndex > 0 Then
            attrText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, attrColumnIndex).Value2))
        Else
            attrText = VBA.vbNullString
        End If
        keyText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, keyColumnIndex).Value2))
        If valueColumnIndex > 0 Then
            valueText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, valueColumnIndex).Value2))
        Else
            valueText = VBA.vbNullString
        End If

        If VBA.Len(keyText) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Config: key is empty in row " & VBA.CStr(rowIndex + 1) & " of table '" & tableObj.Name & "'."
#End If
            Exit Function
        End If

        Set configEntry = New obj_ConfigEntry
        configEntry.Attr = attrText
        configEntry.Key = keyText
        configEntry.Value = valueText
        outEntries.Add configEntry
ContinueRow:
    Next rowIndex

    If outEntries.Count = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' has no non-empty key rows."
#End If
        Exit Function
    End If

    TryGetRenderedConfigEntries = True
End Function

Public Function TryBuildConfigTableFromRendered(ByRef outConfigTable As obj_ConfigTable) As Boolean
    Dim configEntries As Collection
    Dim configEntryItem As Variant
    Dim configEntry As obj_ConfigEntry
    Dim result As obj_ConfigTable

    Set outConfigTable = Nothing
    Set configEntries = Nothing
    If Not TryGetRenderedConfigEntries(configEntries) Then Exit Function

    Set result = New obj_ConfigTable
    If Not result.Initialize() Then Exit Function

    If Not configEntries Is Nothing Then
        For Each configEntryItem In configEntries
            Set configEntry = Nothing
            Set configEntry = configEntryItem
            If configEntry Is Nothing Then GoTo ContinueConfigEntry
            If Not result.AddRow(configEntry.Attr, configEntry.Key, configEntry.Value) Then Exit Function
ContinueConfigEntry:
        Next configEntryItem
    End If

    Set outConfigTable = result
    TryBuildConfigTableFromRendered = True
End Function

Public Function TryGetConfigTable(ByRef outConfigTable As obj_ConfigTable) As Boolean
    Set outConfigTable = Nothing

    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: control '" & m_ControlName & "' is not configured."
#End If
        Exit Function
    End If

    If m_ConfigTableViewItem Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table view item is not initialized for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    Set outConfigTable = m_ConfigTableViewItem.Model
    If outConfigTable Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: model table is not initialized for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    TryGetConfigTable = True
End Function

Public Function TryBuildRenderedConfigNode(ByVal dom As Object, ByRef outConfigNode As Object) As Boolean
    Dim tableObj As ListObject
    Dim headerRange As Range
    Dim dataRange As Range
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim colCount As Long
    Dim keyColumnIndex As Long
    Dim rowWrittenCount As Long
    Dim rawHeader As String
    Dim attrName As String
    Dim cellText As String
    Dim itemNode As Object
    Dim hasAnyValue As Boolean
    Dim normalizedHeaderNames() As String

    ' Строим "source" XML-узел из текущего отрендеренного состояния таблицы.
    ' Контракт метода: вернуть контейнер <config> с дочерними <item .../>.
    ' Важно: метод не изменяет внешний profile-node и не сохраняет файл.
    Set outConfigNode = Nothing
    If dom Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: DOM is not specified for config node build in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    If Not private_TryResolveRenderedTableObject(tableObj) Then Exit Function

    Set headerRange = tableObj.HeaderRowRange
    If headerRange Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' has no header row in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    colCount = headerRange.Columns.Count
    If colCount <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' has no columns in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    ' Нормализуем заголовки таблицы в имена XML-атрибутов:
    ' Attr/marker -> attr, Key/name -> key, Value -> value, остальные -> sanitize.
    ReDim normalizedHeaderNames(1 To colCount)
    For colIndex = 1 To colCount
        rawHeader = VBA.Trim$(VBA.CStr(headerRange.Cells(1, colIndex).Value2))
        If Not private_TryNormalizeProfileAttrName(rawHeader, attrName) Then Exit Function
        normalizedHeaderNames(colIndex) = attrName
        If VBA.StrComp(attrName, "key", VBA.vbTextCompare) = 0 Then keyColumnIndex = colIndex
    Next colIndex

    If keyColumnIndex <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' must contain 'Key' column for profile save in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    Set dataRange = tableObj.DataBodyRange
    If dataRange Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' has no data rows for profile save in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    On Error GoTo EH_XML
    ' Source-контейнер создаем в универсальном виде.
    ' Его дальше интерпретирует caller
    Set outConfigNode = dom.createElement("config")
    rowWrittenCount = 0

    For rowIndex = 1 To dataRange.Rows.Count
        hasAnyValue = False
        For colIndex = 1 To colCount
            cellText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, colIndex).Value2))
            If VBA.Len(cellText) > 0 Then
                hasAnyValue = True
                Exit For
            End If
        Next colIndex
        If Not hasAnyValue Then GoTo ContinueRow

        cellText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, keyColumnIndex).Value2))
        If VBA.Len(cellText) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Config: key is empty in row " & VBA.CStr(rowIndex + 1) & " of table '" & tableObj.Name & "'."
#End If
            Exit Function
        End If

        ' Каждая непустая строка таблицы превращается в отдельный <item/>.
        Set itemNode = dom.createElement("item")

        For colIndex = 1 To colCount
            attrName = normalizedHeaderNames(colIndex)
            If VBA.Len(attrName) = 0 Then GoTo ContinueCol

            cellText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, colIndex).Value2))
            If VBA.StrComp(attrName, "attr", VBA.vbTextCompare) = 0 And VBA.Len(cellText) = 0 Then
                GoTo ContinueCol
            End If

            ' Значения колонок пишем как XML-атрибуты item.
            itemNode.setAttribute attrName, cellText
ContinueCol:
        Next colIndex

        outConfigNode.appendChild itemNode
        rowWrittenCount = rowWrittenCount + 1
ContinueRow:
    Next rowIndex

    If rowWrittenCount = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & tableObj.Name & "' does not contain non-empty rows for profile save in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    TryBuildRenderedConfigNode = True
    Exit Function

EH_XML:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Config: failed to build XML node from rendered table in control '" & m_ControlName & "': " & Err.Description
#End If
End Function

' //
' // Internal
' //
Private Function private_TryBuildConfigTable(ByVal sourceItems As Collection, ByRef outTable As obj_ConfigTable) As Boolean
    Dim sourceConfigEntry As obj_ConfigEntry

    Set outTable = Nothing
    If sourceItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: itemsSource is not resolved for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    Set outTable = New obj_ConfigTable

    For Each sourceConfigEntry In sourceItems
        If sourceConfigEntry Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Config: itemsSource contains empty obj_ConfigEntry in control '" & m_ControlName & "'."
#End If
            GoTo ContinueSourceItem
        End If
        If Not outTable.AddItem(sourceConfigEntry) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Config: failed to append resolved obj_ConfigEntry to table for control '" & m_ControlName & "'."
#End If
            Exit Function
        End If

ContinueSourceItem:
    Next sourceConfigEntry

    private_TryBuildConfigTable = True
End Function

Private Function private_TryResolveRenderedTableObject(ByRef outTableObj As ListObject) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet

    Set outTableObj = Nothing
    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: control '" & m_ControlName & "' is not configured for save."
#End If
        Exit Function
    End If

    Set pageBase = Nothing
    If Not m_ControlBase Is Nothing Then Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then
        Set pageBase = m_Page.GetPageBase()
    End If
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: page is not specified for save in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    If m_ControlLayout Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: control layout is not configured for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    Set ws = private_GetWorksheetByName(pageBase, m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: sheet '" & m_ControlLayout.LayoutSheetName & "' was not found for save in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    If VBA.Len(VBA.Trim$(m_RuntimeTableName)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: runtime table name is not initialized for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    On Error Resume Next
    Set outTableObj = ws.ListObjects(m_RuntimeTableName)
    On Error GoTo 0
    If outTableObj Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: table '" & m_RuntimeTableName & "' was not found for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    private_TryResolveRenderedTableObject = True
End Function

Private Function private_TryNormalizeProfileAttrName(ByVal rawHeader As String, ByRef outAttrName As String) As Boolean
    Dim normalizedHeader As String

    outAttrName = VBA.vbNullString
    normalizedHeader = VBA.LCase$(VBA.Trim$(rawHeader))
    If VBA.Len(normalizedHeader) = 0 Then
        private_TryNormalizeProfileAttrName = True
        Exit Function
    End If

    Select Case normalizedHeader
        Case "attr", "marker"
            outAttrName = "attr"

        Case "key", "name"
            outAttrName = "key"

        Case "value"
            outAttrName = "value"

        Case Else
            outAttrName = private_SanitizeXmlAttrNameToken(normalizedHeader)
            If VBA.Len(outAttrName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "Config: header '" & rawHeader & "' cannot be converted into valid XML attribute name for control '" & m_ControlName & "'."
#End If
                Exit Function
            End If
    End Select

    private_TryNormalizeProfileAttrName = True
End Function

Private Function private_SanitizeXmlAttrNameToken(ByVal rawName As String) As String
    Dim i As Long
    Dim ch As String
    Dim sanitized As String

    rawName = VBA.Trim$(rawName)
    If VBA.Len(rawName) = 0 Then Exit Function

    For i = 1 To VBA.Len(rawName)
        ch = VBA.Mid$(rawName, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or _
           (ch >= "0" And ch <= "9") Or ch = "_" Or ch = "-" Then
            sanitized = sanitized & ch
        Else
            sanitized = sanitized & "_"
        End If
    Next i

    If VBA.Len(sanitized) = 0 Then Exit Function
    If (VBA.Left$(sanitized, 1) >= "0" And VBA.Left$(sanitized, 1) <= "9") Or VBA.Left$(sanitized, 1) = "-" Then
        sanitized = "c_" & sanitized
    End If

    private_SanitizeXmlAttrNameToken = sanitized
End Function

Private Function private_BuildRuntimeControlKey() As String
    If m_ControlLayout Is Nothing Then Exit Function
    private_BuildRuntimeControlKey = "config|" & VBA.LCase$(VBA.Trim$(m_ControlLayout.LayoutSheetName & "|" & m_ControlName))
End Function

Private Function private_TryRegisterRuntimeControl() As Boolean
    Dim pageBase As obj_PageBase

    Set pageBase = m_Page.GetPageBase()

    If VBA.Len(VBA.Trim$(m_RuntimeControlKey)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: runtime control key is empty in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    If Not pageBase.RegisterControl(m_RuntimeControlKey, Me) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Config: failed to register runtime control key '" & m_RuntimeControlKey & "' for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    private_TryRegisterRuntimeControl = True
End Function

Private Sub private_HandleAttrHash(ByVal configEntry As obj_ConfigEntry)
    If configEntry Is Nothing Then Exit Sub
    ' Placeholder for special '#' semantics.
    ' Current behavior: the row is registered into attrhash part for style rules.
End Sub

Private Sub private_HandleAttrRx(ByVal configEntry As obj_ConfigEntry)
    If configEntry Is Nothing Then Exit Sub
    ' Placeholder for special 'rx' semantics.
    ' Current behavior: the row is registered into attrrx part for style rules.
End Sub

Private Function private_RegisterAttrRows(ByVal ws As Worksheet, ByVal partName As String, ByVal rowNumbers As Collection) As Boolean
    Dim rowNumber As Variant
    Dim rowRange As Range
    Dim unionRange As Range

    If ws Is Nothing Then Exit Function
    If rowNumbers Is Nothing Then
        private_RegisterAttrRows = True
        Exit Function
    End If
    If rowNumbers.Count = 0 Then
        private_RegisterAttrRows = True
        Exit Function
    End If

    For Each rowNumber In rowNumbers
        Set rowRange = ws.Range( _
            ws.Cells(VBA.CLng(rowNumber), m_ControlLayout.ColStart), _
            ws.Cells(VBA.CLng(rowNumber), m_ControlLayout.ColStart + CONFIG_COL_COUNT - 1))

        If unionRange Is Nothing Then
            Set unionRange = rowRange
        Else
            Set unionRange = Application.Union(unionRange, rowRange)
        End If
    Next rowNumber

    If unionRange Is Nothing Then
        private_RegisterAttrRows = True
        Exit Function
    End If

    If Not ex_ControlPartsRuntime.fn_RegisterControlPart( _
        ws, _
        "config", _
        m_ControlName, _
        VBA.LCase$(VBA.Trim$(partName)), _
        unionRange) Then Exit Function

    private_RegisterAttrRows = True
End Function

Private Function private_TryDeleteIntersectingTables(ByVal ws As Worksheet, ByVal targetRange As Range) As Boolean
    Dim i As Long
    Dim tableObj As ListObject

    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function

    For i = ws.ListObjects.Count To 1 Step -1
        Set tableObj = ws.ListObjects(i)

        If Not tableObj Is Nothing Then
            If Not Application.Intersect(tableObj.Range, targetRange) Is Nothing Then
                tableObj.Delete
            End If
        End If
    Next i

    private_TryDeleteIntersectingTables = True
End Function

Private Function private_BuildTableName(ByVal ws As Worksheet) As String
    Dim baseName As String
    Dim candidate As String
    Dim suffix As Long

    If ws Is Nothing Then Exit Function

    baseName = m_TableNameRaw
    If VBA.Len(VBA.Trim$(baseName)) = 0 Then baseName = "cfg_" & m_ControlName
    baseName = private_SanitizeTableName(baseName)
    If VBA.Len(baseName) = 0 Then baseName = "cfgTable"

    candidate = baseName
    suffix = 1

    Do While private_TableNameExists(ws, candidate)
        suffix = suffix + 1
        candidate = VBA.Left$(baseName, 240) & "_" & VBA.CStr(suffix)
    Loop

    private_BuildTableName = candidate
End Function

Private Function private_SanitizeTableName(ByVal rawName As String) As String
    Dim i As Long
    Dim ch As String
    Dim outName As String

    rawName = VBA.Trim$(rawName)
    If VBA.Len(rawName) = 0 Then Exit Function

    For i = 1 To VBA.Len(rawName)
        ch = VBA.Mid$(rawName, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or ch = "_" Then
            outName = outName & ch
        Else
            outName = outName & "_"
        End If
    Next i

    If VBA.Len(outName) = 0 Then Exit Function
    If Not ((VBA.Left$(outName, 1) >= "A" And VBA.Left$(outName, 1) <= "Z") Or _
            (VBA.Left$(outName, 1) >= "a" And VBA.Left$(outName, 1) <= "z") Or _
            VBA.Left$(outName, 1) = "_") Then
        outName = "cfg_" & outName
    End If

    If VBA.Len(outName) > 255 Then outName = VBA.Left$(outName, 255)
    private_SanitizeTableName = outName
End Function

Private Function private_TableNameExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    Dim tableObj As ListObject

    If ws Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(tableName)) = 0 Then Exit Function

    For Each tableObj In ws.ListObjects
        If VBA.StrComp(tableObj.Name, tableName, VBA.vbTextCompare) = 0 Then
            private_TableNameExists = True
            Exit Function
        End If
    Next tableObj
End Function

Private Function private_GetWorksheetByName(ByVal page As obj_PageBase, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If page Is Nothing Then Exit Function
    Set ws = page.Worksheet
    If ws Is Nothing Then Exit Function

    sheetName = VBA.LCase$(VBA.Trim$(sheetName))
    If VBA.Len(sheetName) > 0 Then
        If VBA.StrComp(VBA.LCase$(VBA.Trim$(ws.Name)), sheetName, VBA.vbTextCompare) <> 0 Then Exit Function
    End If

    Set private_GetWorksheetByName = ws
End Function
