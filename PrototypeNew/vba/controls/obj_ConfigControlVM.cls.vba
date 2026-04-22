VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigControlVM"
Option Explicit
Implements obj_IControl

Private Const CONFIG_COL_COUNT As Long = 3

Private m_Base As obj_ControlBase
Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_TableNameRaw As String
Private m_Layout As obj_ControlLayout
Private m_ConfigItems As Collection
Private m_IsConfigured As Boolean

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal page As obj_PageBase, ByVal controlNode As Object)
    Dim currentPage As obj_PageBase

    m_IsConfigured = False
    Set m_Layout = Nothing
    Set m_ConfigItems = Nothing
    Set m_Base = Nothing

    Set m_Base = New obj_ControlBase
    If Not m_Base.Configure(page, controlNode, "Config", "config", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
        VBA.MsgBox "Config: itemsSource is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    m_TableNameRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "tableName")))

    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.TryReadFromNode(controlNode, "Config", m_ControlName, "style") Then Exit Sub

    If (m_Layout.ColEnd - m_Layout.ColStart + 1) < CONFIG_COL_COUNT Then
        VBA.MsgBox "Config: control '" & m_ControlName & "' requires at least 3 columns (Attr, Key, Value).", VBA.vbExclamation
        Exit Sub
    End If

    Set currentPage = m_Base.PageBase
    If currentPage Is Nothing Then Exit Sub
    If Not currentPage.RuntimeSources.TryResolveItemsSource(m_ItemsSourceRaw, m_ConfigItems) Then Exit Sub

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
    Dim configItemRaw As Variant
    Dim configView As obj_ConfigViewItem
    Dim attrToken As String
    Dim absRow As Long
    Dim tableObj As ListObject
    Dim targetTableName As String
    Dim hashRows As Collection
    Dim rxRows As Collection
    Dim page As obj_PageBase

    If Not m_IsConfigured Then
        VBA.MsgBox "Config: control '" & m_ControlName & "' is not configured.", VBA.vbExclamation
        Exit Sub
    End If

    Set page = Nothing
    If Not m_Base Is Nothing Then Set page = m_Base.PageBase
    If page Is Nothing Then
        VBA.MsgBox "Config: page is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(page, m_Layout.LayoutSheetName)
    If ws Is Nothing Then
        VBA.MsgBox "Config: sheet '" & m_Layout.LayoutSheetName & "' was not found for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If m_ConfigItems Is Nothing Then
        VBA.MsgBox "Config: itemsSource is not resolved for control '" & m_ControlName & "'.", VBA.vbExclamation
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
        If Not private_TryResolveConfigViewItem(configItemRaw, configView) Then Exit Sub
        If configView Is Nothing Then GoTo ContinueItem

        valueBlock(rowOut, 1) = configView.Attr
        valueBlock(rowOut, 2) = configView.Key
        valueBlock(rowOut, 3) = configView.Value

        absRow = m_Layout.RowStart + rowOut - 1
        attrToken = VBA.LCase$(VBA.Trim$(configView.Attr))
        Select Case attrToken
            Case "#"
                private_HandleAttrHash configView
                hashRows.Add absRow

            Case "rx"
                private_HandleAttrRx configView
                rxRows.Add absRow
        End Select

ContinueItem:
    Next configItemRaw

    Set boundsRange = ws.Range( _
        ws.Cells(m_Layout.RowStart, m_Layout.ColStart), _
        ws.Cells(m_Layout.RowEnd, m_Layout.ColStart + CONFIG_COL_COUNT - 1))

    If Not private_TryDeleteIntersectingTables(ws, boundsRange) Then Exit Sub

    boundsRange.UnMerge
    boundsRange.ClearContents

    Set writeRange = ws.Range( _
        ws.Cells(m_Layout.RowStart, m_Layout.ColStart), _
        ws.Cells(m_Layout.RowStart + rowsToWrite - 1, m_Layout.ColStart + CONFIG_COL_COUNT - 1))
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
    VBA.MsgBox "Config: failed to create table for control '" & m_ControlName & "': " & Err.Description, VBA.vbExclamation
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
' (No public API yet.)
'
' //
' // Internal
' //
Private Function private_TryResolveConfigViewItem(ByVal rawItem As Variant, ByRef outView As obj_ConfigViewItem) As Boolean
    Dim cfgModel As obj_Config

    If Not VBA.IsObject(rawItem) Then
        VBA.MsgBox "Config: itemsSource entry must be an object.", VBA.vbExclamation
        Exit Function
    End If

    Select Case VBA.LCase$(VBA.TypeName(rawItem))
        Case "obj_configviewitem"
            Set outView = rawItem
            private_TryResolveConfigViewItem = True

        Case "obj_config"
            Set cfgModel = rawItem
            Set outView = New obj_ConfigViewItem
            Set outView.Model = cfgModel
            private_TryResolveConfigViewItem = True

        Case Else
            VBA.MsgBox "Config: unsupported itemsSource type '" & VBA.TypeName(rawItem) & "'. Expected obj_ConfigViewItem or obj_Config.", VBA.vbExclamation
    End Select
End Function

Private Sub private_HandleAttrHash(ByVal configView As obj_ConfigViewItem)
    If configView Is Nothing Then Exit Sub
    ' Placeholder for special '#' semantics.
    ' Current behavior: the row is registered into attrhash part for style rules.
End Sub

Private Sub private_HandleAttrRx(ByVal configView As obj_ConfigViewItem)
    If configView Is Nothing Then Exit Sub
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
            ws.Cells(VBA.CLng(rowNumber), m_Layout.ColStart), _
            ws.Cells(VBA.CLng(rowNumber), m_Layout.ColStart + CONFIG_COL_COUNT - 1))

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
