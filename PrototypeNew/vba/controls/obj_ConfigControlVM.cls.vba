VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigControlVM"
Option Explicit
Implements obj_IControl

Private Const CONFIG_COL_COUNT As Long = 3

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_TableNameRaw As String
Private m_ControlLayout As obj_ControlLayout
Private m_ConfigTableViewItem As obj_ConfigTableViewItem
Private m_IsConfigured As Boolean

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal page As obj_PageBase, ByVal controlNode As Object)
    Dim pageBase As obj_PageBase
    Dim resolvedItems As Collection
    Dim configTable As obj_ConfigTable

    m_IsConfigured = False
    Set m_ControlLayout = Nothing
    Set m_ConfigTableViewItem = Nothing
    Set m_ControlBase = Nothing

    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Configure(page, controlNode, "Config", "config", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
        ex_Core.m_Diagnostic_LogError "Config: itemsSource is not specified for control '" & m_ControlName & "'."
        Exit Sub
    End If

    m_TableNameRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "tableName")))

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Config", m_ControlName, "style") Then Exit Sub

    If (m_ControlLayout.ColEnd - m_ControlLayout.ColStart + 1) < CONFIG_COL_COUNT Then
        ex_Core.m_Diagnostic_LogError "Config: control '" & m_ControlName & "' requires at least 3 columns (Attr, Key, Value)."
        Exit Sub
    End If

    Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then Exit Sub
    If Not ex_RuntimeSourceResolver.m_TryResolveItemsSource(pageBase.RuntimeSources, m_ItemsSourceRaw, resolvedItems) Then Exit Sub
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

    If Not m_IsConfigured Then
        ex_Core.m_Diagnostic_LogError "Config: control '" & m_ControlName & "' is not configured."
        Exit Sub
    End If

    Set page = Nothing
    If Not m_ControlBase Is Nothing Then Set page = m_ControlBase.PageBase
    If page Is Nothing Then
        ex_Core.m_Diagnostic_LogError "Config: page is not specified for control '" & m_ControlName & "'."
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(page, m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then
        ex_Core.m_Diagnostic_LogError "Config: sheet '" & m_ControlLayout.LayoutSheetName & "' was not found for control '" & m_ControlName & "'."
        Exit Sub
    End If

    If m_ConfigTableViewItem Is Nothing Then
        ex_Core.m_Diagnostic_LogError "Config: view item is not configured for control '" & m_ControlName & "'."
        Exit Sub
    End If
    If Not m_ConfigTableViewItem.TryResyncEntryItemsFromModel() Then Exit Sub
    Set entryItems = m_ConfigTableViewItem.EntryItems
    If entryItems Is Nothing Then Exit Sub

    Set hashRows = New Collection
    Set rxRows = New Collection

    maxRows = m_ControlLayout.RowEnd - m_ControlLayout.RowStart + 1
    dataRows = entryItems.Count
    rowsToWrite = 1 + dataRows
    If rowsToWrite < 2 Then rowsToWrite = 2
    If rowsToWrite > maxRows Then rowsToWrite = maxRows

    ReDim valueBlock(1 To rowsToWrite, 1 To CONFIG_COL_COUNT)
    valueBlock(1, 1) = "Attr"
    valueBlock(1, 2) = "Key"
    valueBlock(1, 3) = "Value"

    For idx = 1 To entryItems.Count
        Set configEntryViewItem = entryItems.Item(idx)
        rowOut = idx + 1
        If rowOut > rowsToWrite Then Exit For

        If configEntryViewItem Is Nothing Then
            ex_Core.m_Diagnostic_LogError "Config: entry view item is Nothing in control '" & m_ControlName & "'."
            GoTo ContinueItem
        End If

        Set configEntry = configEntryViewItem.Model
        If configEntry Is Nothing Then
            ex_Core.m_Diagnostic_LogError "Config: entry view item has no model in control '" & m_ControlName & "'."
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

    If Not private_RegisterAttrRows(ws, "attrhash", hashRows) Then Exit Sub
    If Not private_RegisterAttrRows(ws, "attrrx", rxRows) Then Exit Sub
    Exit Sub

EH_TABLE:
    ex_Core.m_Diagnostic_LogError "Config: failed to create table for control '" & m_ControlName & "': " & Err.Description
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource", "tablename"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

' //
' // API
' //

'
' //
' // Internal
' //
Private Function private_TryBuildConfigTable(ByVal sourceItems As Collection, ByRef outTable As obj_ConfigTable) As Boolean
    Dim sourceConfigEntry As obj_ConfigEntry

    Set outTable = Nothing
    If sourceItems Is Nothing Then
        ex_Core.m_Diagnostic_LogError "Config: itemsSource is not resolved for control '" & m_ControlName & "'."
        Exit Function
    End If

    Set outTable = New obj_ConfigTable

    For Each sourceConfigEntry In sourceItems
        If sourceConfigEntry Is Nothing Then
            ex_Core.m_Diagnostic_LogError "Config: itemsSource contains empty obj_ConfigEntry in control '" & m_ControlName & "'."
            GoTo ContinueSourceItem
        End If
        If Not outTable.AddItem(sourceConfigEntry) Then
            ex_Core.m_Diagnostic_LogError "Config: failed to append resolved obj_ConfigEntry to table for control '" & m_ControlName & "'."
            Exit Function
        End If

ContinueSourceItem:
    Next sourceConfigEntry

    private_TryBuildConfigTable = True
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

    If Not ex_ControlPartsRuntime.m_RegisterControlPart( _
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
