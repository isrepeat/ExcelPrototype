VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_QueryWorkbookSession"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_app As Object
Private m_book As Workbook
Private m_owned As Boolean
Private m_headers As Variant
Private m_body As Range
Private m_sqlRef As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get Headers() As Variant
    Headers = m_headers
End Property

Public Property Get Body() As Range
    Set Body = m_body
End Property

Public Property Get SqlReference() As String
    SqlReference = m_sqlRef
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_body = Nothing
    On Error Resume Next
    If m_owned And Not m_book Is Nothing Then
        m_book.Close False
    End If
    Set m_book = Nothing
    If Not m_app Is Nothing Then
        m_app.Quit
    End If
    Set m_app = Nothing
    m_owned = False
    m_headers = Empty
    m_sqlRef = VBA.vbNullString
    On Error GoTo 0
End Sub

Public Function TryOpen( _
    ByVal source As obj_TableSource, _
    ByRef diagnostic As String, _
    Optional ByVal savedOnly As Boolean = False _
) As Boolean
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim candidate As ListObject
    Dim table As ListObject
    Dim extent As Range
    Dim header As Range
    Dim i As Long
    Dim count As Long
    Dim canonical As String
    Dim lastRow As Range
    Dim lastColumn As Range
    Dim startedAt As Double

    On Error GoTo EH
    If m_isDisposed Or Not m_isInitialized Then
        diagnostic = "Workbook session is not initialized or is disposed."
        Exit Function
    End If
    If Not m_book Is Nothing Then
        diagnostic = "Workbook session is already open."
        Exit Function
    End If
    canonical = VBA.CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(source.WorkbookPath)
    If Not savedOnly Then
        For Each book In Application.Workbooks
            If VBA.StrComp(book.FullName, canonical, VBA.vbTextCompare) = 0 Then
                Set m_book = book
                Exit For
            End If
        Next book
    End If
    If m_book Is Nothing Then
        If VBA.Len(VBA.Dir$(canonical)) = 0 Then
            Err.Raise VBA.vbObjectError + 2130, , "Workbook not found: " & canonical
        End If
        ex_Core.fn_Diagnostic_WriteLog "QUERY_STAGE_STARTED | Name=CreateExcel"
        startedAt = VBA.Timer
        Set m_app = VBA.CreateObject("Excel.Application")
        ex_Core.fn_Diagnostic_WritePerf "Query.CreateExcel", startedAt
        m_app.Visible = False
        m_app.DisplayAlerts = False
        m_app.EnableEvents = False
        m_app.AutomationSecurity = 3
        ex_Core.fn_Diagnostic_WriteLog "QUERY_STAGE_STARTED | Name=OpenWorkbook"
        startedAt = VBA.Timer
        Set m_book = m_app.Workbooks.Open(Filename:=canonical, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        ex_Core.fn_Diagnostic_WritePerf "Query.OpenWorkbook", startedAt
        m_owned = True
    End If
    ex_Core.fn_Diagnostic_WriteLog "QUERY_STAGE_STARTED | Name=ResolveRange"
    startedAt = VBA.Timer
    If VBA.Len(source.TableName) > 0 Then
        For Each sheet In m_book.Worksheets
            If VBA.Len(source.SheetName) = 0 Or VBA.StrComp(sheet.Name, source.SheetName, VBA.vbTextCompare) = 0 Then
                For Each candidate In sheet.ListObjects
                    If VBA.StrComp(candidate.Name, source.TableName, VBA.vbTextCompare) = 0 Then
                        count = count + 1
                        Set table = candidate
                    End If
                Next candidate
            End If
        Next sheet
        If count <> 1 Then
            Err.Raise VBA.vbObjectError + 2131, , "Table not found or ambiguous: " & source.TableName
        End If
        ReDim m_headers(0 To table.ListColumns.Count - 1)
        For i = 1 To table.ListColumns.Count
            m_headers(i - 1) = table.ListColumns(i).Name
        Next i
        Set m_body = table.DataBodyRange
        Set header = table.Range.Rows(1)
        Set sheet = table.Parent
        If Not table.ShowHeaders Then
            Err.Raise VBA.vbObjectError + 2132, , "Table header row must be shown."
        End If
        Set extent = header
        If Not m_body Is Nothing Then
            Set extent = sheet.Range(header.Cells(1, 1), m_body.Cells(m_body.Rows.Count, m_body.Columns.Count))
        End If
    Else
        Set sheet = m_book.Worksheets(source.SheetName)
        If VBA.StrComp(source.RangeAddress, "auto", VBA.vbTextCompare) = 0 Then
            Set lastRow = sheet.Cells.Find(What:="*", After:=sheet.Cells(1, 1), _
                LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
            Set lastColumn = sheet.Rows(1).Find(What:="*", After:=sheet.Cells(1, 1), _
                LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
            If lastRow Is Nothing Or lastColumn Is Nothing Then
                Err.Raise VBA.vbObjectError + 2133, , "Auto range requires headers in row 1: " & sheet.Name
            End If
            Set extent = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow.Row, lastColumn.Column))
        Else
            Set extent = sheet.Range(source.RangeAddress)
        End If
        If extent.Areas.Count <> 1 Then
            Err.Raise VBA.vbObjectError + 2133, , "A contiguous range is required."
        End If
        Set header = extent.Rows(1)
        ReDim m_headers(0 To header.Columns.Count - 1)
        For i = 1 To header.Columns.Count
            If VBA.IsError(header.Cells(1, i).Value2) Then
                Err.Raise VBA.vbObjectError + 2134, , "Invalid header value."
            End If
            m_headers(i - 1) = VBA.CStr(header.Cells(1, i).Value2)
        Next i
        If extent.Rows.Count > 1 Then
            Set m_body = extent.Offset(1, 0).Resize(extent.Rows.Count - 1, extent.Columns.Count)
        End If
    End If
    m_sqlRef = "[" & VBA.Replace$(sheet.Name, "]", "]]") & "$" & extent.Address(False, False) & "]"
    ex_Core.fn_Diagnostic_WritePerf "Query.ResolveRange", startedAt
    TryOpen = True
    Exit Function
EH:
    diagnostic = "Workbook source: " & Err.Description
    Me.Dispose
End Function