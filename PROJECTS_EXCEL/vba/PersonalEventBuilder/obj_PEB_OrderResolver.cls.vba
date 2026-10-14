VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_OrderResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_source As obj_TableSource
Private m_dateHeader As String
Private m_numberHeader As String
Private m_orders As Collection
Private m_daysBefore As Long
Private m_daysAfter As Long

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
Public Function Initialize(ByRef diagnostic As String) As Boolean
    Dim fileSystem As Object
    Dim path As String

    If m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set m_source = New obj_TableSource
    If Not m_source.Initialize() Then Err.Raise 5, , private_Message("OrderSourceError")
    path = private_Config("source.path")
    If VBA.Mid$(path, 2, 1) <> ":" And VBA.Left$(path, 2) <> "\\" Then path = fileSystem.BuildPath(ThisWorkbook.Path, path)
    m_source.WorkbookPath = fileSystem.GetAbsolutePathName(path)
    m_source.SheetName = private_Config("source.sheet")
    m_source.TableName = private_Config("source.table")
    m_dateHeader = private_Config("dateHeader")
    m_numberHeader = private_Config("numberHeader")
    m_daysBefore = VBA.CLng(private_Config("daysBefore"))
    m_daysAfter = VBA.CLng(private_Config("daysAfter"))
    If m_daysBefore < 0 Or m_daysAfter < 0 Then Err.Raise 5, , private_Message("OrderSourceError")
    If Not m_source.TryValidate(diagnostic) Then GoTo EH
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    If VBA.Len(diagnostic) = 0 Then diagnostic = VBA.Err.Description
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    If Not m_source Is Nothing Then m_source.Dispose
    Set m_source = Nothing
    Set m_orders = Nothing
End Sub

Public Sub InvalidateCache()
    Set m_orders = Nothing
End Sub

Public Function TrySearch( _
    ByVal queryText As String, _
    ByVal yearText As String, _
    ByRef matches As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim searchYear As Long
    Dim searchDate As Date
    Dim searchNumber As String
    Dim record As obj_LookupRecord
    Dim number As Variant
    Dim orderDateValue As Variant
    Dim isDateQuery As Boolean
    Dim pairParts As Variant

    On Error GoTo EH
    Set matches = New Collection
    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , private_Message("OrderSourceError")
    queryText = VBA.Trim$(queryText)
    yearText = VBA.Trim$(yearText)
    searchYear = VBA.Year(VBA.Date)
    If VBA.Len(yearText) > 0 Then
        If Not yearText Like "####" Then Err.Raise 5, , private_Message("OrderInvalidYear")
        searchYear = VBA.CLng(yearText)
        If searchYear < 1000 Then Err.Raise 5, , private_Message("OrderInvalidYear")
    End If
    ' Digits identify a number; only DD.MM and DD.MM.YYYY identify dates.
    If VBA.Len(queryText) = 0 Then
        ' An empty query offers a date window around today without selecting a pair.
        searchNumber = VBA.vbNullString
    ElseIf VBA.InStr(1, queryText, "/") > 0 Then
        pairParts = VBA.Split(queryText, "/")
        If UBound(pairParts) <> 1 Then Err.Raise 5, , private_Message("OrderInvalidInput")
        If Not private_IsNumber(VBA.Trim$(pairParts(0))) Then Err.Raise 5, , private_Message("OrderInvalidInput")
        searchNumber = VBA.CStr(VBA.CDec(VBA.Trim$(pairParts(0))))
        If Not private_TryDate(VBA.Trim$(pairParts(1)), searchDate) Then Err.Raise 5, , private_Message("OrderInvalidDate")
        If VBA.Len(yearText) > 0 And VBA.Year(searchDate) <> searchYear Then Err.Raise 5, , private_Message("OrderYearConflict")
        searchYear = VBA.Year(searchDate)
        isDateQuery = True
    ElseIf private_IsNumber(queryText) Then
        searchNumber = VBA.CStr(VBA.CDec(queryText))
    ElseIf queryText Like "##.##" Then
        If Not private_TryDate(queryText & "." & VBA.CStr(searchYear), searchDate) Then Err.Raise 5, , private_Message("OrderInvalidDate")
        isDateQuery = True
    ElseIf queryText Like "##.##.####" Then
        If Not private_TryDate(queryText, searchDate) Then Err.Raise 5, , private_Message("OrderInvalidDate")
        If VBA.Len(yearText) > 0 And VBA.Year(searchDate) <> searchYear Then Err.Raise 5, , private_Message("OrderYearConflict")
        searchYear = VBA.Year(searchDate)
        isDateQuery = True
    Else
        Err.Raise 5, , private_Message("OrderInvalidInput")
    End If

    If m_orders Is Nothing Then
        If Not private_Load(diagnostic) Then Exit Function
    End If
    For Each record In m_orders
        If Not record.TryGetValue("OrderNumber", number) Then Err.Raise 5, , private_Message("OrderSourceError")
        If Not record.TryGetValue("OrderDate", orderDateValue) Then Err.Raise 5, , private_Message("OrderSourceError")
        If VBA.Len(queryText) = 0 Then
            If VBA.CDate(orderDateValue) >= VBA.Date - m_daysBefore And VBA.CDate(orderDateValue) <= VBA.Date + m_daysAfter Then
                If VBA.Len(yearText) = 0 Or VBA.Year(VBA.CDate(orderDateValue)) = searchYear Then matches.Add record
            End If
        ElseIf VBA.Year(VBA.CDate(orderDateValue)) = searchYear Then
            If isDateQuery Then
                If VBA.CDate(orderDateValue) = searchDate Then
                    If VBA.Len(searchNumber) = 0 Or VBA.CStr(number) = searchNumber Then matches.Add record
                End If
            Else
                If VBA.CStr(number) = searchNumber Then matches.Add record
            End If
        End If
    Next record
    TrySearch = True
    Exit Function
EH:
    Set matches = Nothing
    diagnostic = VBA.Err.Description
End Function

' //
' // Private
' //
Private Function private_Load(ByRef diagnostic As String) As Boolean
    Dim tableQueryService As obj_TableQueryService
    Dim tableQuery As obj_TableQuery
    Dim dataTable As obj_DataTable
    Dim unique As Object
    Dim rows As Collection
    Dim fields As Object
    Dim record As obj_LookupRecord
    Dim dateValue As Variant
    Dim number As String
    Dim orderDate As Date
    Dim key As String
    Dim rowIndex As Long
    Dim dateColumn As Long
    Dim numberColumn As Long

    On Error GoTo EH
    Set tableQueryService = New obj_TableQueryService
    Set tableQuery = New obj_TableQuery
    If Not tableQueryService.Initialize() Or Not tableQuery.Initialize() Then Err.Raise 5, , private_Message("OrderSourceError")
    tableQuery.AddColumn m_dateHeader
    tableQuery.AddColumn m_numberHeader
    If Not tableQueryService.TryExecute(m_source, tableQuery, dataTable, diagnostic) Then
        diagnostic = VBA.Replace$(private_Message("OrderReadError"), "{details}", diagnostic)
        GoTo Cleanup
    End If
    dateColumn = ex_TableQuery.fn_HeaderIndex(dataTable.Headers, m_dateHeader)
    numberColumn = ex_TableQuery.fn_HeaderIndex(dataTable.Headers, m_numberHeader)
    Set unique = VBA.CreateObject("Scripting.Dictionary")
    Set rows = New Collection

    ' Deduplicate identical pairs and reject conflicting numbers within a year.
    For rowIndex = 1 To dataTable.RowCount
        number = private_Text(dataTable.ValueAt(rowIndex, numberColumn))
        dateValue = dataTable.ValueAt(rowIndex, dateColumn)
        If VBA.Len(number) > 0 Or VBA.Len(private_Text(dateValue)) > 0 Then
            If Not private_IsNumber(number) Then Err.Raise 5, , VBA.Replace$(private_Message("OrderDataInvalid"), "{row}", VBA.CStr(rowIndex))
            If VBA.VarType(dateValue) = VBA.vbDate Or VBA.IsNumeric(dateValue) Then
                orderDate = VBA.CDate(VBA.Int(VBA.CDbl(VBA.CDate(dateValue))))
            ElseIf Not private_TryDate(private_Text(dateValue), orderDate) Then
                Err.Raise 5, , VBA.Replace$(private_Message("OrderDataInvalid"), "{row}", VBA.CStr(rowIndex))
            End If
            number = VBA.CStr(VBA.CDec(number))
            key = VBA.CStr(VBA.Year(orderDate)) & "/" & number
            If unique.Exists(key) Then
                If unique(key) <> VBA.CDbl(orderDate) Then Err.Raise 5, , VBA.Replace$(private_Message("OrderDuplicate"), "{key}", key)
            Else
                unique.Add key, VBA.CDbl(orderDate)
                Set fields = VBA.CreateObject("Scripting.Dictionary")
                fields("OrderNumber") = number
                fields("OrderDate") = orderDate
                Set record = New obj_LookupRecord
                If Not record.Initialize(key, fields) Then Err.Raise 5, , private_Message("OrderSourceError")
                rows.Add record
            End If
        End If
    Next rowIndex
    Set m_orders = rows
    private_Load = True
Cleanup:
    If Not dataTable Is Nothing Then dataTable.Dispose
    If Not tableQueryService Is Nothing Then tableQueryService.Dispose
    If Not tableQuery Is Nothing Then tableQuery.Dispose
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_TryDate( _
    ByVal text As String, _
    ByRef parsedDate As Date _
) As Boolean
    Dim day As Long
    Dim month As Long
    Dim year As Long

    On Error GoTo EH
    If Not text Like "##.##.####" Then Exit Function
    day = VBA.CLng(VBA.Left$(text, 2))
    month = VBA.CLng(VBA.Mid$(text, 4, 2))
    year = VBA.CLng(VBA.Right$(text, 4))
    If year < 1000 Then Exit Function
    parsedDate = VBA.DateSerial(year, month, day)
    private_TryDate = VBA.Day(parsedDate) = day And VBA.Month(parsedDate) = month And VBA.Year(parsedDate) = year
EH:
End Function

Private Function private_IsNumber(ByVal text As String) As Boolean
    On Error GoTo EH
    If VBA.Len(text) = 0 Then Exit Function
    If text Like "*[!0-9]*" Then Exit Function
    private_IsNumber = VBA.CDec(text) > 0
EH:
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then Err.Raise 5, , private_Message("OrderSourceError")
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    private_Text = VBA.Trim$(VBA.CStr(value))
End Function

Private Function private_Config(ByVal suffix As String) As String
    If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::orders." & suffix, private_Config) Then
        Err.Raise 5, , "Required configuration key not found: PersonalEventBuilder::orders." & suffix
    End If
End Function

Private Function private_Message(ByVal key As String) As String
    If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::text." & key, private_Message) Then
        Err.Raise 5, , "Required configuration text not found: PersonalEventBuilder::text." & key
    End If
End Function