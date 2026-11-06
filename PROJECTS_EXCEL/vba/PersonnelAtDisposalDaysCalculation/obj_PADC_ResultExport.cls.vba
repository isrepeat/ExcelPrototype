VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_ResultExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_configuration As obj_PADC_Configuration

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize( _
    ByVal configuration As obj_PADC_Configuration _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If configuration Is Nothing Then
        Exit Function
    End If
    Set m_configuration = configuration
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_configuration = Nothing
End Sub

Public Sub CreateWorkbook( _
    ByVal values As Variant, _
    ByVal endDay As Long _
)
    Dim workbook As Workbook
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim headers(1 To 1, 1 To 16) As Variant
    Dim headerKeys As Variant
    Dim widths As Variant
    Dim palette As Variant
    Dim count As Long
    Dim i As Long
    Dim previousScreenUpdating As Boolean
    Dim errorNumber As Long
    Dim errorText As String
    Dim outputPath As String
    Dim backgroundRows As Long
    Dim backgroundColumns As Long
    Dim computedColor As Long

    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_ResultExport", "Service is not initialized."
    End If
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    outputPath = private_NextOutputPath()

    ' Prepare a standalone workbook from the last successful calculation.
    count = UBound(values, 1)
    headerKeys = VBA.Array("export.header.Rank", "header.FullName", "header.TaxId", _
        "export.header.Status", "export.header.OpenStatus", _
        "export.header.StartDate", "header.Periods", "header.CountedDays", "header.ThresholdDate", _
        "export.header.From", "export.header.Through", "export.header.Period", _
        "export.header.UnitPeriod", "export.header.Difference", "export.header.Include", "header.Notes")
    For i = 1 To 16
        headers(1, i) = m_configuration.GetText(VBA.CStr(headerKeys(i - 1)))
    Next i
    widths = VBA.Split(m_configuration.GetText("export.widths"), "|")
    palette = VBA.Split(m_configuration.GetText("export.palette"), "|")
    If UBound(widths) <> 15 Or UBound(palette) <> 5 Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed") & "export.widths / export.palette"
    End If
    Set workbook = Application.Workbooks.Add(xlWBATWorksheet)
    Set sheet = workbook.Worksheets(1)
    sheet.Name = m_configuration.GetText("export.sheet")
    sheet.Range("B7:Q7").Value2 = headers
    sheet.Range("D8").Resize(count, 1).NumberFormat = "@"
    sheet.Range("B8").Resize(count, 16).Value = values
    Set table = sheet.ListObjects.Add(xlSrcRange, sheet.Range("B7").Resize(count + 1, 16), , xlYes)
    table.Name = m_configuration.GetText("export.table")
    table.TableStyle = ""

    ' Leave rows without a valid calculation outside the monthly selection.
    sheet.Range("B2:Q2").Merge
    sheet.Range("B2").Value2 = m_configuration.GetText("export.title")
    sheet.Range("B3").Value2 = m_configuration.GetText("export.labels.Date")
    sheet.Range("B4").Value2 = m_configuration.GetText("export.labels.Start")
    sheet.Range("B5").Value2 = m_configuration.GetText("export.labels.End")
    sheet.Range("C3").Value = VBA.CDate(endDay)
    sheet.Range("C4").Formula = m_configuration.GetText("export.formula.Start")
    sheet.Range("C5").Formula = m_configuration.GetText("export.formula.End")
    table.ListColumns(10).DataBodyRange.FormulaR1C1 = m_configuration.GetText("export.formula.From")
    table.ListColumns(11).DataBodyRange.FormulaR1C1 = m_configuration.GetText("export.formula.Through")
    table.ListColumns(12).DataBodyRange.FormulaR1C1 = m_configuration.GetText("export.formula.Period")
    table.ListColumns(14).DataBodyRange.FormulaR1C1 = m_configuration.GetText("export.formula.Difference")
    table.ListColumns(15).DataBodyRange.FormulaR1C1 = m_configuration.GetText("export.formula.Include")

    ' Apply the warm dark palette, alternating rows and orange accents.
    backgroundRows = VBA.CLng(m_configuration.GetText("export.padding.rows"))
    backgroundColumns = VBA.CLng(m_configuration.GetText("export.padding.columns"))
    computedColor = private_Color(m_configuration.GetText("export.computedColor"))
    If backgroundRows < 0 Or backgroundColumns < 0 Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed") & "export.padding"
    End If
    With sheet.Range(sheet.Cells(1, 1), sheet.Cells(count + 7 + backgroundRows, 17 + backgroundColumns))
        .Interior.Color = private_Color(VBA.CStr(palette(0)))
        .Font.Name = m_configuration.GetText("export.font")
        .Font.Size = VBA.CDbl(m_configuration.GetText("export.fontSize"))
        .Font.Color = private_Color(VBA.CStr(palette(4)))
        .VerticalAlignment = xlCenter
    End With
    For i = 1 To 16
        sheet.Columns(i + 1).ColumnWidth = VBA.CDbl(widths(i - 1))
    Next i
    sheet.Columns(1).ColumnWidth = 3
    sheet.Range("B2").Font.Color = private_Color(VBA.CStr(palette(5)))
    sheet.Range("B2").Font.Bold = True
    sheet.Range("B2").Font.Size = 16
    sheet.Rows(2).RowHeight = 30
    sheet.Range("B3:C5").RowHeight = 21
    sheet.Range("C3:C5").NumberFormat = m_configuration.GetText("format.date")
    With table.HeaderRowRange
        .Interior.Color = private_Color(VBA.CStr(palette(3)))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 60
    End With
    table.DataBodyRange.WrapText = True
    table.DataBodyRange.HorizontalAlignment = xlCenter
    For i = 1 To count
        table.DataBodyRange.Rows(i).Interior.Color = private_Color(VBA.CStr(palette(1 + ((i - 1) Mod 2))))
    Next i
    table.ListColumns(2).DataBodyRange.Font.Color = private_Color(VBA.CStr(palette(5)))
    For Each headerKeys In VBA.Array(6, 9, 10, 11)
        table.ListColumns(VBA.CLng(headerKeys)).DataBodyRange.NumberFormat = m_configuration.GetText("format.date")
    Next headerKeys
    table.ListColumns(8).DataBodyRange.NumberFormat = "0"
    table.ListColumns(13).DataBodyRange.NumberFormat = "@"
    ' Muted labels and values distinguish calculated fields from source inputs.
    sheet.Range("B4:C5").Font.Color = computedColor
    sheet.Range("K7:M" & count + 7).Font.Color = computedColor
    sheet.Range("O7:P" & count + 7).Font.Color = computedColor
    ex_PADC.fn_MarkDateSources table, 2, 6, 16, m_configuration
    table.DataBodyRange.Rows.AutoFit
    sheet.Calculate
    workbook.Activate
    sheet.Activate
    workbook.Windows(1).DisplayGridlines = False
    sheet.Range("B2").Select
    workbook.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    On Error Resume Next
    If Not workbook Is Nothing Then
        workbook.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Application.ScreenUpdating = previousScreenUpdating
    VBA.Err.Raise errorNumber, "obj_PADC_ResultExport", errorText
End Sub

' //
' // Private
' //
Private Function private_Color(ByVal hexValue As String) As Long
    private_Color = VBA.RGB(VBA.CLng("&H" & VBA.Left$(hexValue, 2)), _
        VBA.CLng("&H" & VBA.Mid$(hexValue, 3, 2)), VBA.CLng("&H" & VBA.Right$(hexValue, 2)))
End Function

Private Function private_NextOutputPath() As String
    Dim fileSystem As Object
    Dim baseName As String
    Dim suffixTemplate As String
    Dim suffix As String
    Dim candidate As String
    Dim version As Long

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If VBA.Len(ThisWorkbook.Path) = 0 Then
        m_configuration.Fail m_configuration.GetText("text.ExportFolderMissing")
    End If
    If Not fileSystem.FolderExists(ThisWorkbook.Path) Then
        m_configuration.Fail m_configuration.GetText("text.ExportFolderMissing")
    End If
    baseName = m_configuration.GetText("export.fileName")
    suffixTemplate = m_configuration.GetText("export.versionSuffix")
    If VBA.Len(baseName) = 0 Or VBA.InStr(1, suffixTemplate, "{version}", vbBinaryCompare) = 0 Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed") & "export.fileName / export.versionSuffix"
    End If
    version = 1
    Do
        suffix = vbNullString
        If version > 1 Then
            suffix = VBA.Replace(suffixTemplate, "{version}", VBA.CStr(version))
        End If
        candidate = fileSystem.BuildPath(ThisWorkbook.Path, baseName & suffix & ".xlsx")
        If Not fileSystem.FileExists(candidate) And Not fileSystem.FolderExists(candidate) Then
            private_NextOutputPath = candidate
            Exit Function
        End If
        version = version + 1
    Loop
End Function