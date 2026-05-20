VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SqlParams"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_SourcePath As String
Private m_SheetName As String
Private m_RangeStartMarker As String
Private m_RangeEndMarker As String
Private m_WhereConditions As String
Private m_SourceColumnHeaders As Collection
Private m_MappedColumnHeaders As Collection
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_SourceColumnHeaders = New Collection
    Set m_MappedColumnHeaders = New Collection
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get SourcePath() As String
    SourcePath = m_SourcePath
End Property

Public Property Let SourcePath(ByVal value As String)
    m_SourcePath = VBA.Trim$(VBA.CStr(value))
End Property

Public Property Get SheetName() As String
    SheetName = m_SheetName
End Property

Public Property Let SheetName(ByVal value As String)
    m_SheetName = VBA.Trim$(VBA.CStr(value))
End Property

Public Property Get RangeStartMarker() As String
    RangeStartMarker = m_RangeStartMarker
End Property

Public Property Let RangeStartMarker(ByVal value As String)
    m_RangeStartMarker = VBA.Trim$(VBA.CStr(value))
End Property

Public Property Get RangeEndMarker() As String
    RangeEndMarker = m_RangeEndMarker
End Property

Public Property Let RangeEndMarker(ByVal value As String)
    m_RangeEndMarker = VBA.Trim$(VBA.CStr(value))
End Property

Public Property Get WhereConditions() As String
    WhereConditions = m_WhereConditions
End Property

Public Property Let WhereConditions(ByVal value As String)
    m_WhereConditions = VBA.Trim$(VBA.CStr(value))
End Property

Public Property Get SourceColumnHeaders() As Collection
    If m_SourceColumnHeaders Is Nothing Then Set m_SourceColumnHeaders = New Collection
    Set SourceColumnHeaders = m_SourceColumnHeaders
End Property

Public Property Get MappedColumnHeaders() As Collection
    If m_MappedColumnHeaders Is Nothing Then Set m_MappedColumnHeaders = New Collection
    Set MappedColumnHeaders = m_MappedColumnHeaders
End Property

Public Property Get ColumnCount() As Long
    If m_SourceColumnHeaders Is Nothing Then Exit Property
    ColumnCount = m_SourceColumnHeaders.Count
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_RangeStartMarker = VBA.vbNullString
    m_RangeEndMarker = VBA.vbNullString
    m_WhereConditions = VBA.vbNullString
    Set m_SourceColumnHeaders = New Collection
    Set m_MappedColumnHeaders = New Collection
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    Set m_SourceColumnHeaders = Nothing
    Set m_MappedColumnHeaders = Nothing
    On Error GoTo 0
End Sub

Public Function AddColumnMapping( _
    ByVal sourceColumnHeader As String, _
    Optional ByVal mappedColumnHeader As String = VBA.vbNullString _
) As Boolean
    sourceColumnHeader = VBA.Trim$(sourceColumnHeader)
    mappedColumnHeader = VBA.Trim$(mappedColumnHeader)

    If VBA.Len(sourceColumnHeader) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SqlParams: sourceColumnHeader is required for mapping."
#End If
        Exit Function
    End If

    If VBA.Len(mappedColumnHeader) = 0 Then mappedColumnHeader = sourceColumnHeader

    If m_SourceColumnHeaders Is Nothing Then Set m_SourceColumnHeaders = New Collection
    If m_MappedColumnHeaders Is Nothing Then Set m_MappedColumnHeaders = New Collection

    m_SourceColumnHeaders.Add sourceColumnHeader
    m_MappedColumnHeaders.Add mappedColumnHeader
    AddColumnMapping = True
End Function

Public Function ClearColumnMappings() As Boolean
    Set m_SourceColumnHeaders = New Collection
    Set m_MappedColumnHeaders = New Collection
    ClearColumnMappings = True
End Function

Public Function TryValidate(ByRef outErrorText As String) As Boolean
    outErrorText = VBA.vbNullString

    If VBA.Len(m_SourcePath) = 0 Then
        outErrorText = "SourcePath is required."
        Exit Function
    End If

    If VBA.Len(m_SheetName) = 0 Then
        outErrorText = "SheetName is required."
        Exit Function
    End If

    If (VBA.Len(m_RangeStartMarker) > 0 Xor VBA.Len(m_RangeEndMarker) > 0) Then
        outErrorText = "RangeStartMarker and RangeEndMarker must be provided together."
        Exit Function
    End If

    If m_SourceColumnHeaders Is Nothing Or m_MappedColumnHeaders Is Nothing Then
        outErrorText = "Column mappings are not initialized."
        Exit Function
    End If

    If m_SourceColumnHeaders.Count <= 0 Then
        outErrorText = "At least one column mapping is required."
        Exit Function
    End If

    If m_SourceColumnHeaders.Count <> m_MappedColumnHeaders.Count Then
        outErrorText = "SourceColumnHeaders and MappedColumnHeaders counts must match."
        Exit Function
    End If

    TryValidate = True
End Function

Public Function fn_ToString() As String
    fn_ToString = _
        "SqlParams{" & _
        "SourcePath=" & private_QuoteValue(m_SourcePath) & "; " & _
        "SheetName=" & private_QuoteValue(m_SheetName) & "; " & _
        "RangeStartMarker=" & private_QuoteValue(m_RangeStartMarker) & "; " & _
        "RangeEndMarker=" & private_QuoteValue(m_RangeEndMarker) & "; " & _
        "WhereConditions=" & private_QuoteValue(m_WhereConditions) & "; " & _
        "SourceColumnHeaders=[" & private_CollectionToDelimitedText(m_SourceColumnHeaders, ", ") & "]; " & _
        "MappedColumnHeaders=[" & private_CollectionToDelimitedText(m_MappedColumnHeaders, ", ") & "]" & _
        "}"
End Function

' //
' // Internal
' //
Private Function private_CollectionToDelimitedText( _
    ByVal sourceItems As Collection, _
    ByVal delimiterText As String _
) As String
    Dim i As Long
    Dim itemText As String

    If sourceItems Is Nothing Then Exit Function
    If VBA.Len(delimiterText) = 0 Then delimiterText = ", "

    For i = 1 To sourceItems.Count
        itemText = VBA.Trim$(VBA.CStr(sourceItems.Item(i)))
        If i > 1 Then private_CollectionToDelimitedText = private_CollectionToDelimitedText & delimiterText
        private_CollectionToDelimitedText = private_CollectionToDelimitedText & private_QuoteValue(itemText)
    Next i
End Function

Private Function private_QuoteValue(ByVal valueText As String) As String
    private_QuoteValue = """" & VBA.Replace$(VBA.CStr(valueText), """", "'") & """"
End Function
