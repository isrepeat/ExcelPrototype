VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_TableQueryService"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Public Backend As en_QueryBackend
Private m_excel As obj_ExcelQueryExecutor
Private m_ado As obj_AdoQueryExecutor

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
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Backend = QueryAuto
    Set m_excel = New obj_ExcelQueryExecutor
    Set m_ado = New obj_AdoQueryExecutor
    If Not m_excel.Initialize() Then
        Me.Dispose
        Exit Function
    End If
    If Not m_ado.Initialize() Then
        Me.Dispose
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
    If Not m_excel Is Nothing Then
        m_excel.Dispose
    End If
    If Not m_ado Is Nothing Then
        m_ado.Dispose
    End If
    Set m_excel = Nothing
    Set m_ado = Nothing
End Sub

Public Function TryExecute( _
    ByVal source As obj_TableSource, _
    ByVal query As obj_TableQuery, _
    ByRef output As obj_DataTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim executor As obj_ITableQueryExecutor
    Dim book As Workbook
    Dim selected As en_QueryBackend
    Dim path As String

    On Error GoTo EH
    Set output = Nothing
    diagnostic = VBA.vbNullString
    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2160, , "Query service is not initialized or is disposed."
    End If
    If source Is Nothing Or query Is Nothing Then
        Err.Raise VBA.vbObjectError + 2161, , "Source and query are required."
    End If
    If Not source.TryValidate(diagnostic) Then
        Exit Function
    End If
    selected = Backend
    If selected = QueryAuto Then
        selected = QueryAdo
        path = VBA.CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(source.WorkbookPath)
        For Each book In Application.Workbooks
            If VBA.StrComp(book.FullName, path, VBA.vbTextCompare) = 0 Then
                selected = QueryExcel
                Exit For
            End If
        Next book
    End If
    Select Case selected
        Case QueryExcel
            Set executor = m_excel
        Case QueryAdo
            Set executor = m_ado
        Case Else
            Err.Raise VBA.vbObjectError + 2162, , "Unsupported query backend."
    End Select
    TryExecute = executor.TryExecute(source, query, output, diagnostic)
    Exit Function
EH:
    diagnostic = "Query service: " & Err.Description
    Set output = Nothing
End Function