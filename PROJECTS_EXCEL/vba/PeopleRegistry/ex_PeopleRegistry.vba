Attribute VB_Name = "ex_PeopleRegistry"
Option Explicit

Private Const HISTORY_PASSWORD As String = "PeopleRegistry.History"

Private m_sourceIDs As Object
Private m_sourceNames As Object
Private m_originAliases As Object
Private m_savedID As String
Private m_savedName As String
Private m_savedOrigin As String
Private m_config As Object

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub PreviewIDs()
    private_Run False
End Sub

Public Sub SaveIDs()
    private_Run True
End Sub

Public Sub ReplaceID()
    private_ReplaceID
End Sub

Public Sub ToggleHistoryProtection()
    Dim context As Object
    Dim sheet As Worksheet
    Dim entered As Variant
    Dim errorText As String

    If Not ex_RuntimeLifecycle.fn_TryEnter(context) Then Exit Sub
    Set m_config = Nothing
    On Error GoTo Failed
    Set sheet = private_Table("History").Parent
    If sheet.ProtectContents Then
        entered = Application.InputBox(Prompt:=private_Cfg("Message.PasswordPrompt"), _
            Title:=private_Cfg("Button.UnlockHistory"), Type:=2)
        If VarType(entered) = vbBoolean Then GoTo Cleanup
        If Len(CStr(entered)) = 0 Then GoTo Cleanup
        If StrComp(CStr(entered), HISTORY_PASSWORD, vbBinaryCompare) <> 0 Then
            ex_WindowsUi.fn_ShowMessage private_Cfg("Message.PasswordInvalid"), vbExclamation
            GoTo Cleanup
        End If
        sheet.Unprotect Password:=CStr(entered)
        private_SetupHistoryButton sheet
    Else
        sheet.Cells.Locked = True
        private_ProtectHistory sheet
    End If
Cleanup:
    ex_RuntimeLifecycle.fn_Leave context
    Exit Sub
Failed:
    errorText = Err.Description
    ex_RuntimeLifecycle.fn_Leave context
    ex_WindowsUi.fn_ShowMessage errorText, vbExclamation
End Sub

Public Sub ClearResults()
    Dim context As Object
    Dim index As Long
    Dim errorText As String

    If Not ex_RuntimeLifecycle.fn_TryEnter(context) Then Exit Sub
    Set m_config = Nothing
    On Error GoTo Failed
    For index = 1 To 3
        private_FormCell(index).MergeArea.ClearContents
    Next index
    private_ReplacementCell(1).ClearContents
    private_ReplacementCell(2).ClearContents
    m_savedID = vbNullString
    m_savedName = vbNullString
    m_savedOrigin = vbNullString
    ex_RuntimeLifecycle.fn_Leave context
    Exit Sub
Failed:
    errorText = Err.Description
    ex_RuntimeLifecycle.fn_Leave context
    ex_WindowsUi.fn_ShowMessage errorText, vbExclamation
End Sub

Public Function fn_Initialize() As Boolean
    Dim historyTable As ListObject
    Dim history As Object
    Dim errorText As String
    Dim oldUpdating As Boolean

    oldUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Set m_config = Nothing
    Application.ScreenUpdating = False
    Set historyTable = private_Table("History")
    historyTable.Parent.Unprotect Password:=HISTORY_PASSWORD
    historyTable.Parent.Cells.Locked = True
    historyTable.ShowAutoFilter = True
    private_SetupHistoryButton historyTable.Parent
    If historyTable.ListColumns.Count <> 5 Then
        Err.Raise vbObjectError + 2052, , private_Cfg("Message.SchemaInvalid")
    End If
    Set history = private_LoadHistory(historyTable)
    Set m_originAliases = private_LoadOriginAliases()
    private_LocalizeTables
    private_ApplyTheme
    private_ApplyLayout
    private_SetupForm
    private_SetupButtons ThisWorkbook.Worksheets(private_Cfg("Sheet.Input"))
    private_ProtectHistory historyTable.Parent
    Application.ScreenUpdating = oldUpdating
    fn_Initialize = True
    Exit Function
Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not historyTable Is Nothing Then
        Err.Clear
        private_ProtectHistory historyTable.Parent
        If Err.Number <> 0 Then
            errorText = errorText & vbLf & Replace(private_Cfg("Message.ProtectionFailed"), "{error}", Err.Description)
        End If
    End If
    On Error GoTo 0
    Application.ScreenUpdating = oldUpdating
    ex_WindowsUi.fn_ShowMessage errorText, vbCritical
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_FormCell(ByVal index As Long) As Range
    Dim addresses As Variant

    addresses = Split(private_Cfg("Form.Cells"), "|")
    Set private_FormCell = ThisWorkbook.Worksheets(private_Cfg("Sheet.Input")).Range(CStr(addresses(index - 1)))
End Function

Private Function private_ReplacementCell(ByVal index As Long) As Range
    Dim addresses As Variant

    addresses = Split(private_Cfg("Form.ReplaceCells"), "|")
    Set private_ReplacementCell = ThisWorkbook.Worksheets(private_Cfg("Sheet.Input")).Range(CStr(addresses(index - 1)))
End Function

Private Sub private_SetupForm()
    Dim labels As Variant
    Dim headers As Variant
    Dim infoCells As Variant
    Dim infoKeys As Variant
    Dim sheet As Worksheet
    Dim index As Long

    Set sheet = ThisWorkbook.Worksheets(private_Cfg("Sheet.Input"))
    labels = Split(private_Cfg("Form.Labels"), "|")
    headers = Split(private_Cfg("Form.Headers"), "|")
    sheet.Range(private_Cfg("Layout.ResultRange")).Merge
    For index = 1 To 3
        sheet.Range(CStr(labels(index - 1))).Value2 = headers(index - 1)
        If index <= 2 Then private_FormCell(index).EntireRow.RowHeight = CDbl(private_Cfg("Layout.FieldHeight"))
        With private_FormCell(index).MergeArea
            .NumberFormat = "@"
            .WrapText = True
            .VerticalAlignment = xlCenter
            .Borders.Color = private_Color("Theme.Border")
            .Interior.Color = private_Color("Theme.Header")
            If index <= 2 Then .Interior.Color = private_Color("Theme.Alternate")
        End With
    Next index
    private_FormCell(3).MergeArea.VerticalAlignment = xlTop
    labels = Split(private_Cfg("Form.ReplaceLabels"), "|")
    headers = Split(private_Cfg("Form.ReplaceHeaders"), "|")
    For index = 1 To 2
        sheet.Range(CStr(labels(index - 1))).Value2 = headers(index - 1)
        sheet.Range(CStr(labels(index - 1))).Font.Bold = True
        With private_ReplacementCell(index)
            .EntireRow.RowHeight = CDbl(private_Cfg("Layout.FieldHeight"))
            .NumberFormat = "@"
            .WrapText = True
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = private_Color("Theme.Border")
            .Interior.Color = private_Color("Theme.Alternate")
        End With
    Next index
    infoCells = Split(private_Cfg("Form.InfoCells"), "|")
    infoKeys = Array("Title.Input", "Help.Input", "Title.Result", "Help.Result", "Title.Replace", "Help.Replace")
    For index = 0 To 5
        sheet.Range(CStr(infoCells(index))).Value2 = private_Cfg(CStr(infoKeys(index)))
        sheet.Range(CStr(infoCells(index))).Font.Color = private_Color("Theme.Muted")
        If index = 0 Or index = 2 Or index = 4 Then
            sheet.Range(CStr(infoCells(index))).Font.Color = private_Color("Theme.Accent")
            sheet.Range(CStr(infoCells(index))).Font.Bold = True
            sheet.Range(CStr(infoCells(index))).Font.Size = Val(private_Cfg("Theme.TitleFontSize"))
        End If
    Next index
End Sub

Private Sub private_ApplyLayout()
    Dim roles As Variant
    Dim role As Variant
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim entry As Variant
    Dim parts As Variant
    Dim sheetName As String
    Dim addresses As Variant
    Dim address As Variant

    roles = Array("Input", "History", "Aliases", "Config")
    For Each role In roles
        sheetName = "wsConfig"
        If role <> "Config" Then sheetName = private_Cfg("Sheet." & CStr(role))
        Set sheet = ThisWorkbook.Worksheets(sheetName)
        sheet.Columns.ColumnWidth = Val(private_Cfg("Layout.DefaultColumnWidth"))
        sheet.Rows.RowHeight = Val(private_Cfg("Layout." & CStr(role) & ".DefaultRowHeight"))
        sheet.Cells.Font.Size = Val(private_Cfg("Theme.FontSize"))
        For Each entry In Split(private_Cfg("Layout." & CStr(role) & ".ColumnWidths"), "|")
            parts = Split(CStr(entry), "=")
            sheet.Columns(CStr(parts(0))).ColumnWidth = Val(CStr(parts(1)))
        Next entry
        If role = "Aliases" Or role = "Config" Then
            If StrComp(private_Cfg("Layout." & CStr(role) & ".AutoFitRows"), "True", vbTextCompare) = 0 Then
                For Each table In sheet.ListObjects
                    table.Range.EntireRow.AutoFit
                Next table
                If role = "Aliases" Then sheet.Rows(2).AutoFit
            End If
        End If
    Next role
    Set sheet = ThisWorkbook.Worksheets(private_Cfg("Sheet.Input"))
    For Each entry In Split(private_Cfg("Layout.Input.RowHeights"), "|")
        parts = Split(CStr(entry), "=")
        sheet.Rows(CLng(parts(0))).RowHeight = Val(CStr(parts(1)))
    Next entry
    With sheet.Range(private_Cfg("Layout.Input.ViewRange"))
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = CLng(private_Cfg("Layout.Input.ContentAlignment"))
    End With
    addresses = Split(private_Cfg("Form.Labels") & "|" & private_Cfg("Form.ReplaceLabels"), "|")
    For Each address In addresses
        sheet.Range(CStr(address)).HorizontalAlignment = CLng(private_Cfg("Layout.Input.LabelAlignment"))
    Next address
    With sheet.PageSetup
        .PrintArea = private_Cfg("Layout.Input.PrintArea")
        .PaperSize = CLng(private_Cfg("Layout.Input.PaperSize"))
        .Orientation = CLng(private_Cfg("Layout.Input.Orientation"))
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    sheet.Activate
    With Application.ActiveWindow
        .Zoom = CLng(private_Cfg("Layout.Zoom"))
        .DisplayGridlines = False
        .FreezePanes = False
        .ScrollRow = sheet.Range(private_Cfg("Layout.Input.ScrollCell")).Row
        .ScrollColumn = sheet.Range(private_Cfg("Layout.Input.ScrollCell")).Column
    End With
    sheet.Range(private_Cfg("Layout.Input.ActiveCell")).Select
End Sub

Private Sub private_LocalizeTables()
    Dim kind As Variant
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim target As ListObject
    Dim headers As Variant
    Dim index As Long
    Dim oldUpdating As Boolean
    Dim errorNumber As Long
    Dim errorText As String

    oldUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    ' Resolve tables by stable technical names before changing localized worksheet names.
    For Each kind In Array("History", "Aliases")
        Set target = Nothing
        For Each sheet In ThisWorkbook.Worksheets
            For Each table In sheet.ListObjects
                If table.Name = private_Cfg("Table." & kind) Then
                    Set target = table
                    Exit For
                End If
            Next table
            If Not target Is Nothing Then
                Exit For
            End If
        Next sheet
        If target Is Nothing Then
            Err.Raise vbObjectError + 2052, , private_Cfg("Message.SchemaInvalid")
        End If
        headers = Split(private_Cfg("Header." & kind), "|")
        If target.ListColumns.Count <> UBound(headers) + 1 Then
            Err.Raise vbObjectError + 2052, , private_Cfg("Message.SchemaInvalid")
        End If
        target.Parent.Name = private_Cfg("Sheet." & kind)
        target.HeaderRowRange.Cells(1, 1).Offset(-2).Value2 = private_Cfg("Title." & kind)
        For index = 0 To UBound(headers)
            target.ListColumns(index + 1).Name = headers(index)
        Next index
    Next kind
    Application.ScreenUpdating = oldUpdating
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Application.ScreenUpdating = oldUpdating
    Err.Raise errorNumber, "private_LocalizeTables", errorText
End Sub

Private Sub private_ApplyTheme()
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim style As TableStyle
    Dim candidate As TableStyle
    Dim button As Shape
    Dim oldUpdating As Boolean
    Dim errorNumber As Long
    Dim errorText As String

    oldUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    For Each candidate In ThisWorkbook.TableStyles
        If candidate.Name = "PeopleRegistryDark" Then
            Set style = candidate
            Exit For
        End If
    Next candidate
    If style Is Nothing Then
        Set style = ThisWorkbook.TableStyles.Add("PeopleRegistryDark")
    End If
    With style.TableStyleElements(xlWholeTable)
        .Interior.Color = private_Color("Theme.Background")
        .Font.Color = private_Color("Theme.Foreground")
    End With
    With style.TableStyleElements(xlHeaderRow)
        .Interior.Color = private_Color("Theme.Header")
        .Font.Color = private_Color("Theme.Foreground")
        .Font.Bold = True
    End With
    With style.TableStyleElements(xlRowStripe1)
        .StripeSize = 1
        .Interior.Color = private_Color("Theme.Background")
        .Font.Color = private_Color("Theme.Foreground")
    End With
    With style.TableStyleElements(xlRowStripe2)
        .StripeSize = 1
        .Interior.Color = private_Color("Theme.Alternate")
        .Font.Color = private_Color("Theme.Foreground")
    End With
    For Each sheet In ThisWorkbook.Worksheets
        sheet.Cells.Interior.Color = private_Color("Theme.Background")
        sheet.Cells.Font.Color = private_Color("Theme.Foreground")
        sheet.Cells.Font.Name = private_Cfg("Theme.FontName")
        sheet.Tab.Color = private_Color("Theme.Accent")
        For Each table In sheet.ListObjects
            ' Direct fills override table banding; remove them before applying the style.
            table.Range.Interior.Pattern = xlPatternNone
            table.TableStyle = style.Name
            table.ShowTableStyleRowStripes = True
            table.ShowTableStyleColumnStripes = False
            table.Range.Font.Color = private_Color("Theme.Foreground")
            table.HeaderRowRange.Interior.Color = private_Color("Theme.Header")
            table.HeaderRowRange.Font.Bold = True
            table.Range.Borders.Color = private_Color("Theme.Border")
            table.Range.Borders.LineStyle = xlContinuous
        Next table
        For Each button In sheet.Shapes
            If button.Name = "btnPreview" Or button.Name = "btnSave" Or button.Name = "btnClear" Then
                button.AutoShapeType = msoShapeRectangle
                button.Fill.ForeColor.RGB = private_Color("Theme.Alternate")
                button.Line.ForeColor.RGB = private_Color("Theme.Border")
                button.TextFrame2.TextRange.Font.Name = private_Cfg("Theme.FontName")
                button.TextFrame2.TextRange.Font.Bold = msoTrue
                button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = private_Color("Theme.Accent")
            End If
        Next button
    Next sheet
    Application.ScreenUpdating = oldUpdating
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Application.ScreenUpdating = oldUpdating
    Err.Raise errorNumber, "private_ApplyTheme", errorText
End Sub

Private Function private_Color(ByVal key As String) As Long
    Dim value As String

    value = UCase$(Trim$(private_Cfg(key)))
    If Len(value) <> 7 Or Left$(value, 1) <> "#" Or _
            Mid$(value, 2) Like "*[!0-9A-F]*" Then
        Err.Raise vbObjectError + 2053, , Replace(private_Cfg("Message.InvalidColor"), "{key}", key)
    End If
    private_Color = RGB( _
        CLng("&H" & Mid$(value, 2, 2)), _
        CLng("&H" & Mid$(value, 4, 2)), _
        CLng("&H" & Mid$(value, 6, 2)))
End Function

Private Sub private_SetupButtons(ByVal sheet As Worksheet)
    Dim names As Variant
    Dim commands As Variant
    Dim captions As Variant
    Dim anchors As Variant
    Dim button As Shape
    Dim candidate As Shape
    Dim index As Long
    Dim oldUpdating As Boolean
    Dim errorNumber As Long
    Dim errorText As String

    oldUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    For index = sheet.Shapes.Count To 1 Step -1
        If sheet.Shapes(index).Name = "btnPreview" Or sheet.Shapes(index).Name = "btnClear" Then
            sheet.Shapes(index).Delete
        End If
    Next index
    names = Array("btnSave", "btnReplace")
    commands = Array("SaveIDs", "ReplaceID")
    captions = Array("Button.Save", "Button.Replace")
    anchors = Split(private_Cfg("Layout.ButtonCells"), "|")
    For index = 0 To UBound(names)
        Set button = Nothing
        For Each candidate In sheet.Shapes
            If candidate.Name = names(index) Then
                Set button = candidate
                Exit For
            End If
        Next candidate
        If button Is Nothing Then
            Set button = sheet.Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 1, 1)
            button.Name = names(index)
        End If
        button.Left = sheet.Range(CStr(anchors(index))).Left
        button.Top = sheet.Range(CStr(anchors(index))).Top
        button.Width = CDbl(private_Cfg("Layout.ButtonWidth"))
        button.Height = CDbl(private_Cfg("Layout.ButtonHeight"))
        button.AutoShapeType = msoShapeRectangle
        button.Fill.ForeColor.RGB = private_Color("Theme.Alternate")
        button.Line.ForeColor.RGB = private_Color("Theme.Border")
        button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = private_Color("Theme.Accent")
        button.TextFrame2.TextRange.Text = private_Cfg(CStr(captions(index)))
        button.TextFrame2.TextRange.Font.Size = Val(private_Cfg("Theme.ButtonFontSize"))
        button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
            "'!ex_PeopleRegistry." & commands(index)
        button.Placement = xlFreeFloating
    Next index
    Application.ScreenUpdating = oldUpdating
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Application.ScreenUpdating = oldUpdating
    Err.Raise errorNumber, "private_SetupButtons", errorText
End Sub

Private Sub private_Run(ByVal commit As Boolean)
    Dim historyTable As ListObject
    Dim history As Object
    Dim context As Object
    Dim values As Variant
    Dim nameParts As Variant
    Dim oldUpdating As Boolean
    Dim oldEvents As Boolean
    Dim errorText As String
    Dim saving As Boolean

    If Not ex_RuntimeLifecycle.fn_TryEnter(context) Then Exit Sub
    Set m_config = Nothing
    oldUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ReDim values(1 To 1, 1 To 7)
    values(1, 1) = private_FormCell(1).Value2
    values(1, 2) = private_FormCell(2).Value2
    private_FormCell(3).MergeArea.ClearContents
    If Len(private_Normalize(private_InputText(values, 1))) = 0 Then
        Err.Raise vbObjectError + 2001, , private_Cfg("Message.NameRequired")
    End If
    If Len(private_Normalize(private_InputText(values, 2))) = 0 Then
        Err.Raise vbObjectError + 2010, , private_Cfg("Message.OriginRequired")
    End If
    nameParts = Split(private_Normalize(private_InputText(values, 1)), " ")
    If Len(private_Token(CStr(nameParts(0)))) = 0 Or _
        Len(private_Token(private_Normalize(private_InputText(values, 2)))) = 0 Then
        Err.Raise vbObjectError + 2030, , private_Cfg("Message.EmptyToken")
    End If
    If StrComp(private_Normalize(private_InputText(values, 1)), m_savedName, vbTextCompare) = 0 And _
        StrComp(private_Normalize(private_InputText(values, 2)), m_savedOrigin, vbTextCompare) = 0 Then
        values(1, 4) = m_savedID
    Else
        values(1, 4) = vbNullString
    End If
    If commit And (ThisWorkbook.ReadOnly Or Len(ThisWorkbook.Path) = 0) Then
        Err.Raise vbObjectError + 2000, , private_Cfg("Message.ReadOnly")
    End If
    Set historyTable = private_Table("History")
    Set history = private_LoadHistory(historyTable)
    Set m_originAliases = private_LoadOriginAliases()
    If Not ex_PR_Source.fn_Load(m_sourceIDs, m_sourceNames, errorText) Then
        Err.Raise vbObjectError + 2102, , errorText
    End If
    If private_ProcessRow(values, history, commit) Then
        If commit Then
            saving = True
            private_WriteResult values
            private_SaveHistory historyTable, history, CStr(values(1, 5))
            m_savedID = CStr(values(1, 5))
            m_savedName = private_Normalize(private_InputText(values, 1))
            m_savedOrigin = private_Normalize(private_InputText(values, 2))
            saving = False
        End If
    End If
    private_WriteResult values
Cleanup:
    Application.ScreenUpdating = oldUpdating
    Application.EnableEvents = oldEvents
    ex_RuntimeLifecycle.fn_Leave context
    Exit Sub
Failed:
    errorText = Err.Description
    If saving Then errorText = Replace(private_Cfg("Message.SaveFailed"), "{error}", errorText)
    On Error Resume Next
    values(1, 5) = vbNullString
    values(1, 6) = private_Cfg("Status.Error")
    values(1, 7) = errorText
    private_WriteResult values
    On Error GoTo 0
    GoTo Cleanup
End Sub

Private Sub private_ReplaceID()
    Dim context As Object
    Dim historyTable As ListObject
    Dim history As Object
    Dim values As Variant
    Dim record As Variant
    Dim oldValue As Variant
    Dim newValue As Variant
    Dim oldID As String
    Dim newID As String
    Dim currentID As String
    Dim fullName As String
    Dim origin As String
    Dim message As String
    Dim errorText As String
    Dim oldEvents As Boolean
    Dim oldUpdating As Boolean
    Dim saving As Boolean

    If Not ex_RuntimeLifecycle.fn_TryEnter(context) Then Exit Sub
    Set m_config = Nothing
    oldEvents = Application.EnableEvents
    oldUpdating = Application.ScreenUpdating
    ReDim values(1 To 1, 1 To 7)
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    private_FormCell(3).MergeArea.ClearContents
    oldValue = private_ReplacementCell(1).Value2
    newValue = private_ReplacementCell(2).Value2
    If IsError(oldValue) Or IsError(newValue) Then
        Err.Raise vbObjectError + 2070, , private_Cfg("Message.CellError")
    End If
    oldID = private_Normalize(CStr(oldValue))
    newID = private_Normalize(CStr(newValue))
    If Len(oldID) = 0 Or Len(newID) = 0 Then
        Err.Raise vbObjectError + 2071, , private_Cfg("Message.ReplaceRequired")
    End If
    private_ValidateNewID newID
    If ThisWorkbook.ReadOnly Or Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 2000, , private_Cfg("Message.ReadOnly")
    End If
    Set historyTable = private_Table("History")
    Set history = private_LoadHistory(historyTable)
    If Not ex_PR_Source.fn_Load(m_sourceIDs, m_sourceNames, errorText) Then
        Err.Raise vbObjectError + 2102, , errorText
    End If
    If history.Exists(oldID) Then
        record = history(oldID)
        currentID = CStr(record(0))
        record = history(currentID)
        fullName = CStr(record(1))
        origin = CStr(record(2))
    ElseIf m_sourceIDs.Exists(oldID) Then
        record = m_sourceIDs(oldID)
        currentID = oldID
        fullName = CStr(record(0))
        origin = CStr(record(1))
    Else
        Err.Raise vbObjectError + 2005, , private_Cfg("Message.UnknownID")
    End If
    Call private_SourceMatches(oldID, fullName)
    Call private_SourceMatches(currentID, fullName)
    If StrComp(newID, currentID, vbTextCompare) = 0 Then
        values(1, 5) = currentID
        values(1, 6) = private_Cfg("Status.Saved")
        If Not history.Exists(currentID) Then values(1, 6) = private_Cfg("Status.Source")
        values(1, 7) = private_Cfg("Message.SameID")
        GoTo Report
    End If
    ' Never merge saved groups or infer identity from matching names in the source.
    If history.Exists(newID) Or m_sourceIDs.Exists(newID) Then
        Err.Raise vbObjectError + 2072, , private_Cfg("Message.IDInUse")
    End If
    message = Replace(private_Cfg("Message.ConfirmReplacement"), "{name}", fullName)
    message = Replace(message, "{old}", currentID)
    message = Replace(message, "{new}", newID)
    message = Replace(message, "{newline}", vbLf)
    If ex_WindowsUi.fn_ShowMessage(message, vbYesNo Or vbExclamation Or vbDefaultButton2, _
        private_Cfg("Title.Replace")) <> vbYes Then
        values(1, 6) = private_Cfg("Status.Cancelled")
        values(1, 7) = private_Cfg("Message.ReplaceCancelled")
        GoTo Report
    End If
    values(1, 1) = fullName
    values(1, 2) = origin
    values(1, 3) = newID
    values(1, 4) = oldID
    If private_ProcessRow(values, history, True) Then
        saving = True
        private_ReplacementCell(1).Value2 = CStr(values(1, 5))
        private_ReplacementCell(2).ClearContents
        private_WriteResult values
        private_SaveHistory historyTable, history
        saving = False
        If StrComp(m_savedID, currentID, vbTextCompare) = 0 Then m_savedID = CStr(values(1, 5))
    End If
Report:
    private_WriteResult values
Cleanup:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldUpdating
    ex_RuntimeLifecycle.fn_Leave context
    Exit Sub
Failed:
    errorText = Err.Description
    If saving Then
        errorText = Replace(private_Cfg("Message.SaveFailed"), "{error}", errorText)
        On Error Resume Next
        private_ReplacementCell(1).Value2 = oldValue
        private_ReplacementCell(2).Value2 = newValue
        On Error GoTo 0
    End If
    values(1, 5) = vbNullString
    values(1, 6) = private_Cfg("Status.Error")
    values(1, 7) = errorText
    GoTo Report
End Sub

Private Sub private_WriteResult(ByRef values As Variant)
    Dim text As String

    text = Replace(private_Cfg("Form.ResultTemplate"), "{id}", CStr(values(1, 5)))
    text = Replace(text, "{status}", CStr(values(1, 6)))
    text = Replace(text, "{message}", CStr(values(1, 7)))
    private_FormCell(3).Value2 = Replace(text, "{newline}", vbLf)
End Sub

Private Function private_ProcessRow( _
    ByRef inputRow As Variant, _
    ByVal history As Object, _
    ByVal commit As Boolean _
) As Boolean
    Dim fullName As String
    Dim origin As String
    Dim newID As String
    Dim existingID As String
    Dim candidates As String
    Dim selectedID As String
    Dim resultID As String
    Dim message As String
    Dim record As Variant
    Dim key As Variant
    Dim isNew As Boolean
    Dim isMigration As Boolean
    Dim sourceKnown As Boolean
    Dim mutationStarted As Boolean

    On Error GoTo Failed
    fullName = private_Normalize(private_InputText(inputRow, 1))
    origin = private_Normalize(private_InputText(inputRow, 2))
    newID = private_Normalize(private_InputText(inputRow, 3))
    existingID = private_Normalize(private_InputText(inputRow, 4))
    If Len(fullName) = 0 Then
        Err.Raise vbObjectError + 2001, , private_Cfg("Message.NameRequired")
    End If
    If Len(newID) > 0 Then
        private_ValidateNewID newID
    End If

    ' Resolve explicit identifiers; matching names alone never merge people.
    If Len(existingID) > 0 Then
        sourceKnown = private_SourceMatches(existingID, fullName)
        If history.Exists(existingID) Then
            record = history(existingID)
            private_CheckName CStr(record(1)), fullName, existingID
            selectedID = CStr(record(0))
        ElseIf sourceKnown Then
            selectedID = existingID
        Else
            Err.Raise vbObjectError + 2005, , private_Cfg("Message.UnknownID")
        End If
    End If
    If Len(newID) > 0 Then
        sourceKnown = private_SourceMatches(newID, fullName)
        If history.Exists(newID) Then
            record = history(newID)
            private_CheckName CStr(record(1)), fullName, newID
            If Len(selectedID) > 0 And StrComp(selectedID, CStr(record(0)), vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 2007, , private_Cfg("Message.IDConflict")
            End If
            selectedID = CStr(record(0))
        ElseIf sourceKnown Then
            If Len(selectedID) = 0 Then
                selectedID = newID
            End If
        End If
    End If
    isNew = Len(selectedID) = 0

    ' A blank new ID generates for a new person and resolves an existing person's current ID.
    If Len(selectedID) = 0 Then
        If Len(newID) > 0 Then
            resultID = newID
        Else
            If Len(origin) = 0 Then
                Err.Raise vbObjectError + 2010, , private_Cfg("Message.OriginRequired")
            End If
            resultID = private_UniqueID(private_BaseID(fullName, origin), history)
        End If
    Else
        sourceKnown = private_SourceMatches(selectedID, fullName)
        resultID = selectedID
        If Len(newID) > 0 And StrComp(newID, selectedID, vbTextCompare) <> 0 Then
            If Not history.Exists(newID) Then
                resultID = newID
                isMigration = True
            End If
        End If
    End If

    If Len(origin) = 0 And Len(selectedID) > 0 Then
        If history.Exists(selectedID) Then
            record = history(selectedID)
            origin = CStr(record(2))
        ElseIf m_sourceIDs.Exists(selectedID) Then
            record = m_sourceIDs(selectedID)
            origin = CStr(record(1))
        End If
    End If

    ' Confirm same-name creation before mutating history, on preview and save alike.
    If isNew Then
        candidates = private_Candidates(fullName, history)
        If Len(candidates) > 0 Then
            message = Replace(private_Cfg("Message.ConfirmNewPerson"), "{candidates}", candidates)
            If ex_WindowsUi.fn_ShowMessage(message, vbYesNo Or vbExclamation Or vbDefaultButton2, _
                private_Cfg("Title.Input")) <> vbYes Then
                inputRow(1, 5) = vbNullString
                inputRow(1, 6) = private_Cfg("Status.Cancelled")
                inputRow(1, 7) = private_Cfg("Message.Cancelled")
                Exit Function
            End If
        End If
    End If

    ' Flatten replacements directly to the current ID; no internal person keys or chains.
    If isMigration Then
        mutationStarted = True
        For Each key In history.Keys
            record = history(key)
            If StrComp(CStr(record(0)), selectedID, vbTextCompare) = 0 Then
                record(0) = resultID
                history(key) = record
            End If
        Next key
        If Not history.Exists(selectedID) Then
            history.Add selectedID, Array(resultID, fullName, origin, private_Timestamp())
        End If
        history.Add resultID, Array(resultID, fullName, origin, private_Timestamp())
    ElseIf isNew Then
        mutationStarted = True
        history.Add resultID, Array(resultID, fullName, origin, private_Timestamp())
    End If

    inputRow(1, 5) = resultID
    If isMigration Then
        If commit Then
            message = private_Cfg("Message.Migrated")
        Else
            message = private_Cfg("Message.PreviewMigrated")
        End If
    ElseIf isNew Then
        If commit Then
            message = private_Cfg("Message.Created")
        Else
            message = private_Cfg("Message.PreviewCreated")
        End If
    ElseIf m_sourceIDs.Exists(resultID) Then
        message = private_Cfg("Message.SourceFound")
    Else
        message = private_Cfg("Message.Found")
    End If
    If m_sourceIDs.Exists(resultID) And Not isMigration Then
        inputRow(1, 6) = private_Cfg("Status.Source")
    ElseIf commit Then
        inputRow(1, 6) = private_Cfg("Status.Saved")
    Else
        inputRow(1, 6) = private_Cfg("Status.Preview")
    End If
    inputRow(1, 7) = message
    private_ProcessRow = True
    Exit Function
Failed:
    message = Err.Description
    If mutationStarted Then
        ' A failed mutation aborts before any history is persisted.
        Err.Raise vbObjectError + 2012, "private_ProcessRow", message
    End If
    inputRow(1, 5) = vbNullString
    inputRow(1, 6) = private_Cfg("Status.Error")
    inputRow(1, 7) = message
End Function

Private Function private_InputText( _
    ByRef values As Variant, _
    ByVal column As Long _
) As String
    If IsError(values(1, column)) Then
        Err.Raise vbObjectError + 2040, , private_Cfg("Message.CellError")
    End If
    private_InputText = CStr(values(1, column))
End Function

Private Sub private_ValidateNewID( _
    ByVal identifier As String _
)
    Dim index As Long
    Dim character As String
    Dim invalidMessage As String

    invalidMessage = Replace(private_Cfg("Message.InvalidID"), "{max}", private_Cfg("Limit.IDLength"))
    If Len(identifier) > CLng(private_Cfg("Limit.IDLength")) Then
        Err.Raise vbObjectError + 2061, , invalidMessage
    End If
    For index = 1 To Len(identifier)
        character = Mid$(identifier, index, 1)
        If AscW(character) >= 0 And AscW(character) <= 32 Then
            Err.Raise vbObjectError + 2061, , invalidMessage
        End If
        If InStr(1, private_Cfg("Value.InvalidIDCharacters"), character, vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 2061, , invalidMessage
        End If
    Next index
End Sub

Private Function private_LoadHistory(ByVal table As ListObject) As Object
    Dim history As Object
    Dim values As Variant
    Dim row As Long
    Dim column As Long
    Dim populated As Boolean
    Dim identifier As String
    Dim currentID As String
    Dim fullName As String
    Dim key As Variant
    Dim record As Variant
    Dim current As Variant

    Set history = private_Dictionary()
    If table.ListColumns.Count <> 5 Then
        Err.Raise vbObjectError + 2052, , private_Cfg("Message.SchemaInvalid")
    End If
    If Not table.DataBodyRange Is Nothing Then
        values = table.DataBodyRange.Value2
        For row = 1 To UBound(values, 1)
            populated = False
            For column = 1 To 5
                If IsError(values(row, column)) Then
                    Err.Raise vbObjectError + 2040, , private_Cfg("Message.CellError")
                End If
                If Not IsEmpty(values(row, column)) Then populated = True
            Next column
            If populated Then
                identifier = private_Normalize(CStr(values(row, 1)))
                currentID = private_Normalize(CStr(values(row, 2)))
                fullName = private_Normalize(CStr(values(row, 3)))
                If Len(identifier) = 0 Or Len(currentID) = 0 Or Len(fullName) = 0 Or history.Exists(identifier) Then
                    Err.Raise vbObjectError + 2020, , private_Cfg("Message.RegistryInvalid")
                End If
                history.Add identifier, Array(currentID, fullName, CStr(values(row, 4)), CStr(values(row, 5)))
            End If
        Next row
    End If
    For Each key In history.Keys
        record = history(key)
        If Not history.Exists(CStr(record(0))) Then
            Err.Raise vbObjectError + 2021, , private_Cfg("Message.RegistryInvalid")
        End If
        current = history(CStr(record(0)))
        If StrComp(CStr(current(0)), CStr(record(0)), vbTextCompare) <> 0 Or _
                StrComp(CStr(current(1)), CStr(record(1)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 2022, , private_Cfg("Message.RegistryInvalid")
        End If
    Next key
    Set private_LoadHistory = history
End Function

Private Sub private_SetupHistoryButton(ByVal sheet As Worksheet)
    Dim button As Shape
    Dim candidate As Shape
    Dim captionKey As String

    For Each candidate In sheet.Shapes
        If candidate.Name = "btnHistoryLock" Then
            Set button = candidate
            Exit For
        End If
    Next candidate
    If button Is Nothing Then
        Set button = sheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 1, 1)
        button.Name = "btnHistoryLock"
    End If
    captionKey = "Button.LockHistory"
    If sheet.ProtectContents Then captionKey = "Button.UnlockHistory"
    button.Left = sheet.Range(private_Cfg("Layout.HistoryButtonCell")).Left + Val(private_Cfg("Layout.HistoryButtonOffsetX"))
    button.Top = sheet.Range(private_Cfg("Layout.HistoryButtonCell")).Top + Val(private_Cfg("Layout.HistoryButtonOffsetY"))
    button.Width = CDbl(private_Cfg("Layout.HistoryButtonWidth"))
    button.Height = Val(private_Cfg("Layout.HistoryButtonHeight"))
    button.Fill.ForeColor.RGB = private_Color("Theme.Alternate")
    button.Line.ForeColor.RGB = private_Color("Theme.Border")
    button.TextFrame2.TextRange.Text = private_Cfg(captionKey)
    button.TextFrame2.TextRange.Font.Size = Val(private_Cfg("Theme.ButtonFontSize"))
    button.TextFrame2.TextRange.Font.Name = private_Cfg("Theme.FontName")
    button.TextFrame2.TextRange.Font.Bold = msoTrue
    button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = private_Color("Theme.Accent")
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!ex_PeopleRegistry.ToggleHistoryProtection"
    button.Placement = xlFreeFloating
    button.Locked = True
End Sub

Private Sub private_ProtectHistory(ByVal sheet As Worksheet)
    sheet.Protect Password:=HISTORY_PASSWORD, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, _
        AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, _
        AllowInsertingColumns:=False, AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
        AllowDeletingColumns:=False, AllowDeletingRows:=False, AllowSorting:=False, _
        AllowFiltering:=True, AllowUsingPivotTables:=False
    sheet.EnableSelection = xlNoRestrictions
    private_SetupHistoryButton sheet
End Sub

Private Sub private_SaveHistory( _
    ByVal table As ListObject, _
    ByVal history As Object, _
    Optional ByVal newRegistrationID As String = vbNullString _
)
    Dim original As Variant
    Dim oldCount As Long
    Dim values As Variant
    Dim key As Variant
    Dim record As Variant
    Dim row As Long
    Dim column As Long
    Dim errorText As String
    Dim rollbackError As String
    Dim mutationStarted As Boolean

    oldCount = table.ListRows.Count
    If oldCount > 0 Then
        original = table.DataBodyRange.Value2
    End If
    On Error GoTo Failed
    table.Parent.Unprotect Password:=HISTORY_PASSWORD
    mutationStarted = True
    ' Registration only appends a canonical row; replacement still rewrites all affected links.
    If Len(newRegistrationID) > 0 And history.Count = oldCount + 1 Then
        record = history(newRegistrationID)
        If StrComp(CStr(record(0)), newRegistrationID, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 2020, , private_Cfg("Message.RegistryInvalid")
        End If
        table.Resize table.HeaderRowRange.Resize(history.Count + 1, 5)
        ReDim values(1 To 1, 1 To 5)
        values(1, 1) = newRegistrationID
        For column = 0 To 3
            values(1, column + 2) = CStr(record(column))
        Next column
        With table.HeaderRowRange.Offset(oldCount + 1).Resize(1, 5)
            .NumberFormat = "@"
            .Value2 = values
        End With
    ElseIf Len(newRegistrationID) = 0 Or history.Count <> oldCount Then
        If Not table.DataBodyRange Is Nothing Then
            table.DataBodyRange.ClearContents
        End If
        table.Resize table.HeaderRowRange.Resize(history.Count + 1, 5)
        If history.Count > 0 Then
            ReDim values(1 To history.Count, 1 To 5)
            For Each key In history.Keys
                row = row + 1
                record = history(key)
                values(row, 1) = CStr(key)
                For column = 0 To 3
                    values(row, column + 2) = CStr(record(column))
                Next column
            Next key
            ' An empty resized table can retain DataBodyRange = Nothing until data is written.
            With table.HeaderRowRange.Offset(1).Resize(history.Count, 5)
                .NumberFormat = "@"
                .Value2 = values
            End With
        End If
    End If
    private_ProtectHistory table.Parent
    ThisWorkbook.Save
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    Err.Clear
    If mutationStarted Then
        table.Parent.Unprotect Password:=HISTORY_PASSWORD
        If Not table.DataBodyRange Is Nothing Then
            table.DataBodyRange.ClearContents
        End If
        table.Resize table.HeaderRowRange.Resize(oldCount + 1, 5)
        If oldCount > 0 Then
            table.HeaderRowRange.Offset(1).Resize(oldCount, 5).Value2 = original
        End If
    End If
    If Err.Number <> 0 Then
        rollbackError = Err.Description
    End If
    Err.Clear
    private_ProtectHistory table.Parent
    If Err.Number <> 0 Then
        rollbackError = rollbackError & vbLf & Replace(private_Cfg("Message.ProtectionFailed"), "{error}", Err.Description)
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 2023, "private_SaveHistory", errorText & vbCrLf & rollbackError
End Sub

Private Sub private_CheckName( _
    ByVal expected As String, _
    ByVal actual As String, _
    ByVal identifier As String _
)
    If StrComp(private_Normalize(expected), actual, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2104, , Replace(private_Cfg("Message.SourceMismatch"), "{id}", identifier)
    End If
End Sub

Private Function private_Candidates( _
    ByVal fullName As String, _
    ByVal history As Object _
) As String
    Dim candidates As Object
    Dim source As Object
    Dim key As Variant
    Dim record As Variant
    Dim count As Long
    Dim result As String

    Set candidates = private_Dictionary()
    If m_sourceNames.Exists(fullName) Then
        Set source = m_sourceNames(fullName)
        For Each key In source.Keys
            candidates(CStr(key)) = CStr(source(key))
        Next key
    End If
    For Each key In history.Keys
        record = history(key)
        If StrComp(CStr(record(1)), fullName, vbTextCompare) = 0 Then
            candidates(CStr(record(0))) = CStr(record(2))
        End If
    Next key
    For Each key In candidates.Keys
        count = count + 1
        If count > CLng(private_Cfg("Limit.Candidates")) Then
            Exit For
        End If
        result = result & CStr(key) & " [" & CStr(candidates(key)) & "]; "
    Next key
    private_Candidates = result
End Function

Private Function private_Timestamp() As String
    private_Timestamp = Format$(Now, private_Cfg("Format.Timestamp"))
End Function

Private Function private_SourceMatches( _
    ByVal identifier As String, _
    ByVal fullName As String _
) As Boolean
    Dim record As Variant

    If Not m_sourceIDs.Exists(identifier) Then
        Exit Function
    End If
    record = m_sourceIDs(identifier)
    If CBool(record(2)) Then
        Err.Raise vbObjectError + 2103, , Replace(private_Cfg("Message.SourceAmbiguous"), "{id}", identifier)
    End If
    If StrComp(CStr(record(0)), fullName, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2104, , Replace(private_Cfg("Message.SourceMismatch"), "{id}", identifier)
    End If
    private_SourceMatches = True
End Function

Private Function private_BaseID( _
    ByVal fullName As String, _
    ByVal origin As String) As String

    Dim words As Variant
    Dim result As String
    Dim index As Long
    Dim originToken As String

    words = Split(fullName, " ")
    result = private_Token(CStr(words(0)))
    result = UCase$(Left$(result, 1)) & LCase$(Mid$(result, 2))
    For index = 1 To UBound(words)
        result = result & UCase$(Left$(private_Token(CStr(words(index))), 1))
    Next index
    originToken = private_Token(origin)
    If m_originAliases.Exists(private_OriginKey(origin)) Then
        originToken = CStr(m_originAliases(private_OriginKey(origin)))
    End If
    result = result & private_Cfg("Value.Separator") & originToken
    If Len(private_Token(CStr(words(0)))) = 0 Or Len(originToken) = 0 Then
        Err.Raise vbObjectError + 2030, , private_Cfg("Message.EmptyToken")
    End If
    private_BaseID = Left$(result, CLng(private_Cfg("Limit.IDLength")))
End Function

Private Function private_OriginKey(ByVal value As String) As String
    value = Replace(private_Normalize(value), ChrW(&H2019), "'")
    value = Replace(value, ChrW(&H2BC), "'")
    private_OriginKey = private_Token(value)
End Function

Private Function private_LoadOriginAliases() As Object
    Dim aliases As Object
    Dim table As ListObject
    Dim values As Variant
    Dim row As Long
    Dim messageTemplate As String
    Dim maxLength As Long
    Dim headerRow As Long
    Dim fullName As String
    Dim shortCode As String
    Dim key As String
    Dim message As String

    Set aliases = private_Dictionary()
    Set table = private_Table("Aliases")
    If table.ListColumns.Count <> 2 Then
        Err.Raise vbObjectError + 2052, , private_Cfg("Message.SchemaInvalid")
    End If
    messageTemplate = private_Cfg("Message.AliasInvalid")
    maxLength = CLng(private_Cfg("Limit.AliasLength"))
    headerRow = table.HeaderRowRange.Row
    If Not table.DataBodyRange Is Nothing Then
        values = table.DataBodyRange.Value2
        For row = 1 To UBound(values, 1)
            If IsError(values(row, 1)) Or IsError(values(row, 2)) Then
                Err.Raise vbObjectError + 2040, , private_Cfg("Message.CellError")
            End If
            If Not IsEmpty(values(row, 1)) Or Not IsEmpty(values(row, 2)) Then
                fullName = private_Normalize(CStr(values(row, 1)))
                shortCode = private_Normalize(CStr(values(row, 2)))
                key = private_OriginKey(fullName)
                message = Replace(messageTemplate, "{row}", CStr(headerRow + row))
                If Len(key) = 0 Or Len(shortCode) = 0 Or _
                        Len(shortCode) > maxLength Or _
                        StrComp(private_Token(shortCode), shortCode, vbBinaryCompare) <> 0 Then
                    Err.Raise vbObjectError + 2063, , message
                End If
                If aliases.Exists(key) Then
                    If StrComp(CStr(aliases(key)), shortCode, vbTextCompare) <> 0 Then
                        Err.Raise vbObjectError + 2063, , message
                    End If
                Else
                    aliases.Add key, shortCode
                End If
            End If
        Next row
    End If
    Set private_LoadOriginAliases = aliases
End Function

Private Function private_UniqueID( _
    ByVal baseID As String, _
    ByVal history As Object) As String

    Dim candidate As String
    Dim suffix As String
    Dim counter As Long
    Dim maxLength As Long

    candidate = baseID
    counter = 1
    maxLength = CLng(private_Cfg("Limit.IDLength"))
    Do While history.Exists(candidate) Or m_sourceIDs.Exists(candidate)
        counter = counter + 1
        suffix = private_Cfg("Value.Separator") & CStr(counter)
        candidate = Left$(baseID, maxLength - Len(suffix)) & suffix
    Loop
    private_UniqueID = candidate
End Function

Private Function private_Normalize(ByVal value As String) As String
    value = Replace(value, ChrW(160), " ")
    value = Replace(value, ChrW(&H202F), " ")
    value = Replace(value, vbTab, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Trim$(value)
    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop
    private_Normalize = value
End Function

Private Function private_Token(ByVal value As String) As String
    Dim index As Long
    Dim character As String
    Dim result As String
    Dim removedCharacters As String

    removedCharacters = private_Cfg("Value.RemovedCharacters")
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If InStr(1, removedCharacters, character, vbBinaryCompare) = 0 Then
            result = result & character
        End If
    Next index
    private_Token = result
End Function

Private Function private_Dictionary() As Object
    Dim dictionary As Object

    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbTextCompare
    Set private_Dictionary = dictionary
End Function

Private Function private_Table(ByVal key As String) As ListObject
    Set private_Table = ThisWorkbook.Worksheets(private_Cfg("Sheet." & key)).ListObjects(private_Cfg("Table." & key))
End Function

Private Function private_Cfg(ByVal key As String) As String
    Dim values As Variant
    Dim loaded As Object
    Dim row As Long
    Dim fullKey As String

    If m_config Is Nothing Then
        Set loaded = private_Dictionary()
        values = ThisWorkbook.Worksheets("wsConfig").ListObjects("tbConfig").DataBodyRange.Value2
        For row = 1 To UBound(values, 1)
            If IsError(values(row, 2)) Or IsError(values(row, 3)) Then
                Err.Raise vbObjectError + 2050, "ex_PeopleRegistry", "Invalid configuration cell."
            End If
            fullKey = CStr(values(row, 2))
            If loaded.Exists(fullKey) Then
                Err.Raise vbObjectError + 2050, "ex_PeopleRegistry", "Duplicate configuration key: " & fullKey
            End If
            loaded.Add fullKey, CStr(values(row, 3))
        Next row
        Set m_config = loaded
    End If
    fullKey = "PeopleRegistry::" & key
    If m_config.Exists(fullKey) Then
        private_Cfg = CStr(m_config(fullKey))
        If Len(private_Cfg) > 0 Then Exit Function
    End If
    Err.Raise vbObjectError + 2050, "ex_PeopleRegistry", "Missing configuration key: PeopleRegistry::" & key
End Function