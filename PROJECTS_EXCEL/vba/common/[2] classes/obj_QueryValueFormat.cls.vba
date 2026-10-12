VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_QueryValueFormat"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_decimalSeparator As String
Private m_groupSeparator As String
Private m_dateFormat As String
Private m_dateSystem As String

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
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_isInitialized And Not m_isDisposed
End Property

Public Property Get DecimalSeparator() As String
    DecimalSeparator = m_decimalSeparator
End Property

Public Property Get GroupSeparator() As String
    GroupSeparator = m_groupSeparator
End Property

Public Property Get DateFormat() As String
    DateFormat = m_dateFormat
End Property

Public Property Get DateSystem() As String
    DateSystem = m_dateSystem
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal decimalSeparator As String, _
    ByVal groupSeparator As String, _
    ByVal dateFormat As String, _
    ByVal dateSystem As String _
) As Boolean
    If m_isDisposed Or m_isInitialized Then Exit Function
    If decimalSeparator <> "." And decimalSeparator <> "," And decimalSeparator <> "" Then Exit Function
    If groupSeparator <> "" And groupSeparator <> "." And groupSeparator <> "," And _
            groupSeparator <> " " And groupSeparator <> VBA.ChrW$(160) Then Exit Function
    If groupSeparator <> "" And groupSeparator = decimalSeparator Then Exit Function
    If dateFormat <> "" And dateFormat <> "dd.MM.yyyy" And dateFormat <> "yyyy-MM-dd" Then Exit Function
    If dateSystem <> "" And dateSystem <> "1900" And dateSystem <> "1904" Then Exit Function
    m_decimalSeparator = decimalSeparator
    m_groupSeparator = groupSeparator
    m_dateFormat = dateFormat
    m_dateSystem = dateSystem
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isInitialized = False
    m_isDisposed = True
    m_decimalSeparator = VBA.vbNullString
    m_groupSeparator = VBA.vbNullString
    m_dateFormat = VBA.vbNullString
    m_dateSystem = VBA.vbNullString
End Sub

Public Function InitializeFromConfig(ByVal prefix As String) As Boolean
    Dim decimalSeparator As String
    Dim groupSeparator As String
    Dim dateFormat As String
    Dim dateSystem As String

    If m_isDisposed Or m_isInitialized Then Exit Function
    If VBA.Len(VBA.Trim$(prefix)) = 0 Then Exit Function
    decimalSeparator = private_ReadConfig(prefix & ".decimalSeparator")
    groupSeparator = private_ReadConfig(prefix & ".groupSeparator")
    dateFormat = private_ReadConfig(prefix & ".dateFormat")
    dateSystem = private_ReadConfig(prefix & ".dateSystem")
    If decimalSeparator = "none" Then decimalSeparator = VBA.vbNullString
    Select Case groupSeparator
        Case "none"
            groupSeparator = VBA.vbNullString
        Case "space"
            groupSeparator = " "
        Case "nbsp"
            groupSeparator = VBA.ChrW$(160)
    End Select
    If dateFormat = "none" Then dateFormat = VBA.vbNullString
    If dateSystem = "none" Then dateSystem = VBA.vbNullString
    InitializeFromConfig = Me.Initialize(decimalSeparator, groupSeparator, dateFormat, dateSystem)
    If Not InitializeFromConfig Then
        Err.Raise VBA.vbObjectError + 2108, , "Invalid query value format configuration: " & prefix
    End If
End Function

' //
' // Private
' //
Private Function private_ReadConfig(ByVal key As String) As String
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, private_ReadConfig) Then
        Err.Raise VBA.vbObjectError + 2108, , "Missing query value format configuration key: " & key
    End If
End Function