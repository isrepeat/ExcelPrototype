VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigControlVM"
Option Explicit
Implements obj_IControl

Private Const CONFIG_COL_COUNT As Long = 3

Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_TableNameRaw As String
Private m_Layout As obj_ControlLayout
Private m_ConfigItems As Collection
Private m_IsConfigured As Boolean

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    m_IsConfigured = False
    Set m_Layout = Nothing
    Set m_ConfigItems = Nothing

    If controlNode Is Nothing Then
        MsgBox "Config: control node is not specified.", vbExclamation
        Exit Sub
    End If

    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then m_ControlName = "config"

    m_ItemsSourceRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource")))
    If Len(m_ItemsSourceRaw) = 0 Then
        MsgBox "Config: itemsSource is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    m_TableNameRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "tableName")))

    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.m_TryReadFromNode(controlNode, "Config", m_ControlName, "style") Then Exit Sub

    If (m_Layout.ColEnd - m_Layout.ColStart + 1) < CONFIG_COL_COUNT Then
        MsgBox "Config: control '" & m_ControlName & "' requires at least 3 columns (Attr, Key, Value).", vbExclamation
        Exit Sub
    End If

    If Not ex_ListItemsSourceRuntime.m_TryResolveItemsSource(m_ItemsSourceRaw, m_ConfigItems) Then Exit Sub

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim boundsRange As Range
    Dim writeRange As Range
    Dim valueBlock As Variant
    Dim rowsToWrite As Long
    Dim maxRows As Long
    Dim dataRows As Long
    Dim idx As Long
    Dim rowOut As Long
    Dim configItemRaw As Variant
    Dim configView As obj_ConfigViewItem
    Dim attrToken As String
    Dim absRow As Long
    Dim tableObj As ListObject
    Dim targetTableName As String
    Dim hashRows As Collection
    Dim rxRows As Collection

    If Not m_IsConfigured Then
        MsgBox "Config: control '" & m_ControlName & "' is not configured.", vbExclamation
        Exit Sub
    End If

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "Config: workbook is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = mp_GetWorksheetByName(wb, m_Layout.LayoutSheet)
    If ws Is Nothing Then
        MsgBox "Config: sheet '" & m_Layout.LayoutSheet & "' was not found for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If m_ConfigItems Is Nothing Then
        MsgBox "Config: itemsSource is not resolved for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set hashRows = New Collection
    Set rxRows = New Collection

    maxRows = m_Layout.RowEnd - m_Layout.RowStart + 1
    dataRows = m_ConfigItems.Count
    rowsToWrite = 1 + dataRows
    If rowsToWrite < 2 Then rowsToWrite = 2
    If rowsToWrite > maxRows Then rowsToWrite = maxRows

    ReDim valueBlock(1 To rowsToWrite, 1 To CONFIG_COL_COUNT)
    valueBlock(1, 1) = "Attr"
    valueBlock(1, 2) = "Key"
    valueBlock(1, 3) = "Value"

    idx = 0
    For Each configItemRaw In m_ConfigItems
        idx = idx + 1
        rowOut = idx + 1
        If rowOut > rowsToWrite Then Exit For

        Set configView = Nothing
        If Not mp_TryResolveConfigViewItem(configItemRaw, configView) Then Exit Sub
        If configView Is Nothing Then GoTo ContinueItem

        valueBlock(rowOut, 1) = configView.Attr
        valueBlock(rowOut, 2) = configView.Key
        valueBlock(rowOut, 3) = configView.Value

        absRow = m_Layout.RowStart + rowOut - 1
        attrToken = LCase$(Trim$(configView.Attr))
        Select Case attrToken
            Case "#"
                mp_HandleAttrHash configView
                hashRows.Add absRow

            Case "rx"
                mp_HandleAttrRx configView
                rxRows.Add absRow
        End Select

ContinueItem:
    Next configItemRaw

    Set boundsRange = ws.Range( _
        ws.Cells(m_Layout.RowStart, m_Layout.ColStart), _
        ws.Cells(m_Layout.RowEnd, m_Layout.ColStart + CONFIG_COL_COUNT - 1))

    If Not mp_TryDeleteIntersectingTables(ws, boundsRange) Then Exit Sub

    boundsRange.UnMerge
    boundsRange.ClearContents

    Set writeRange = ws.Range( _
        ws.Cells(m_Layout.RowStart, m_Layout.ColStart), _
        ws.Cells(m_Layout.RowStart + rowsToWrite - 1, m_Layout.ColStart + CONFIG_COL_COUNT - 1))
    writeRange.Value2 = valueBlock

    On Error GoTo EH_TABLE
    Set tableObj = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=writeRange, XlListObjectHasHeaders:=xlYes)
    On Error GoTo 0

    targetTableName = mp_BuildTableName(ws)
    If Len(targetTableName) > 0 Then
        On Error Resume Next
        tableObj.Name = targetTableName
        On Error GoTo 0
    End If

    On Error Resume Next
    tableObj.TableStyle = "TableStyleMedium2"
    tableObj.ShowAutoFilter = True
    On Error GoTo 0

    If Not mp_RegisterAttrRows(ws, "attrhash", hashRows) Then Exit Sub
    If Not mp_RegisterAttrRows(ws, "attrrx", rxRows) Then Exit Sub
    Exit Sub

EH_TABLE:
    MsgBox "Config: failed to create table for control '" & m_ControlName & "': " & Err.Description, vbExclamation
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case LCase$(Trim$(attrName))
        Case "itemssource", "tablename"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

' //
' // Internal
' //
Private Function mp_TryResolveConfigViewItem(ByVal rawItem As Variant, ByRef outView As obj_ConfigViewItem) As Boolean
    Dim cfgModel As obj_Config

    If Not IsObject(rawItem) Then
        MsgBox "Config: itemsSource entry must be an object.", vbExclamation
        Exit Function
    End If

    Select Case LCase$(TypeName(rawItem))
        Case "obj_configviewitem"
            Set outView = rawItem
            mp_TryResolveConfigViewItem = True

        Case "obj_config"
            Set cfgModel = rawItem
            Set outView = New obj_ConfigViewItem
            Set outView.Model = cfgModel
            mp_TryResolveConfigViewItem = True

        Case Else
            MsgBox "Config: unsupported itemsSource type '" & TypeName(rawItem) & "'. Expected obj_ConfigViewItem or obj_Config.", vbExclamation
    End Select
End Function

Private Sub mp_HandleAttrHash(ByVal configView As obj_ConfigViewItem)
    If configView Is Nothing Then Exit Sub
    ' Placeholder for special '#' semantics.
    ' Current behavior: the row is registered into attrhash part for style rules.
End Sub

Private Sub mp_HandleAttrRx(ByVal configView As obj_ConfigViewItem)
    If configView Is Nothing Then Exit Sub
    ' Placeholder for special 'rx' semantics.
    ' Current behavior: the row is registered into attrrx part for style rules.
End Sub

Private Function mp_RegisterAttrRows(ByVal ws As Worksheet, ByVal partName As String, ByVal rowNumbers As Collection) As Boolean
    Dim rowNumber As Variant
    Dim rowRange As Range
    Dim unionRange As Range

    If ws Is Nothing Then Exit Function
    If rowNumbers Is Nothing Then
        mp_RegisterAttrRows = True
        Exit Function
    End If
    If rowNumbers.Count = 0 Then
        mp_RegisterAttrRows = True
        Exit Function
    End If

    For Each rowNumber In rowNumbers
        Set rowRange = ws.Range( _
            ws.Cells(CLng(rowNumber), m_Layout.ColStart), _
            ws.Cells(CLng(rowNumber), m_Layout.ColStart + CONFIG_COL_COUNT - 1))

        If unionRange Is Nothing Then
            Set unionRange = rowRange
        Else
            Set unionRange = Application.Union(unionRange, rowRange)
        End If
    Next rowNumber

    If unionRange Is Nothing Then
        mp_RegisterAttrRows = True
        Exit Function
    End If

    If Not ex_ControlPartsRuntime.m_RegisterControlPart( _
        ws, _
        "config", _
        m_ControlName, _
        LCase$(Trim$(partName)), _
        unionRange) Then Exit Function

    mp_RegisterAttrRows = True
End Function

Private Function mp_TryDeleteIntersectingTables(ByVal ws As Worksheet, ByVal targetRange As Range) As Boolean
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

    mp_TryDeleteIntersectingTables = True
End Function

Private Function mp_BuildTableName(ByVal ws As Worksheet) As String
    Dim baseName As String
    Dim candidate As String
    Dim suffix As Long

    If ws Is Nothing Then Exit Function

    baseName = m_TableNameRaw
    If Len(Trim$(baseName)) = 0 Then baseName = "cfg_" & m_ControlName
    baseName = mp_SanitizeTableName(baseName)
    If Len(baseName) = 0 Then baseName = "cfgTable"

    candidate = baseName
    suffix = 1

    Do While mp_TableNameExists(ws, candidate)
        suffix = suffix + 1
        candidate = Left$(baseName, 240) & "_" & CStr(suffix)
    Loop

    mp_BuildTableName = candidate
End Function

Private Function mp_SanitizeTableName(ByVal rawName As String) As String
    Dim i As Long
    Dim ch As String
    Dim outName As String

    rawName = Trim$(rawName)
    If Len(rawName) = 0 Then Exit Function

    For i = 1 To Len(rawName)
        ch = Mid$(rawName, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or ch = "_" Then
            outName = outName & ch
        Else
            outName = outName & "_"
        End If
    Next i

    If Len(outName) = 0 Then Exit Function
    If Not ((Left$(outName, 1) >= "A" And Left$(outName, 1) <= "Z") Or _
            (Left$(outName, 1) >= "a" And Left$(outName, 1) <= "z") Or _
            Left$(outName, 1) = "_") Then
        outName = "cfg_" & outName
    End If

    If Len(outName) > 255 Then outName = Left$(outName, 255)
    mp_SanitizeTableName = outName
End Function

Private Function mp_TableNameExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    Dim tableObj As ListObject

    If ws Is Nothing Then Exit Function
    If Len(Trim$(tableName)) = 0 Then Exit Function

    For Each tableObj In ws.ListObjects
        If StrComp(tableObj.Name, tableName, vbTextCompare) = 0 Then
            mp_TableNameExists = True
            Exit Function
        End If
    Next tableObj
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
