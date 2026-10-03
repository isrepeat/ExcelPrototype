VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_LookupRecord"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_key As String
Private m_fields As Object
Private m_details As Object

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
Public Property Get Key() As String
    Key = m_key
End Property

Public Property Get DetailNames() As Variant
    If Not m_details Is Nothing Then DetailNames = m_details.Keys
End Property

Public Property Get Fields() As Object
    Dim copy As Object
    Dim name As Variant

    Set copy = VBA.CreateObject("Scripting.Dictionary")
    copy.CompareMode = VBA.vbTextCompare
    If Not m_fields Is Nothing Then
        For Each name In m_fields.Keys
            copy(name) = m_fields(name)
        Next name
    End If
    Set Fields = copy
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal key As String, _
    ByVal fields As Object _
) As Boolean
    Dim name As Variant

    If m_isInitialized Or m_isDisposed Then Exit Function
    If fields Is Nothing Then Exit Function
    Set m_fields = VBA.CreateObject("Scripting.Dictionary")
    m_fields.CompareMode = VBA.vbTextCompare
    For Each name In fields.Keys
        If VBA.IsObject(fields(name)) Or VBA.IsError(fields(name)) Then Exit Function
        m_fields(name) = fields(name)
    Next name
    Set m_details = VBA.CreateObject("Scripting.Dictionary")
    m_details.CompareMode = VBA.vbTextCompare
    m_key = key
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    Set m_fields = Nothing
    Set m_details = Nothing
    m_key = VBA.vbNullString
End Sub

Public Function TryGetValue( _
    ByVal name As String, _
    ByRef value As Variant _
) As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not m_fields.Exists(name) Then Exit Function
    value = m_fields(name)
    TryGetValue = True
End Function

Public Function SetDetail( _
    ByVal name As String, _
    ByVal records As Collection _
) As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If VBA.Len(name) = 0 Or records Is Nothing Then Exit Function
    Set m_details(name) = records
    SetDetail = True
End Function

Public Function GetDetail(ByVal name As String) As Collection
    Dim copy As New Collection
    Dim record As obj_LookupRecord

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If m_details.Exists(name) Then
        For Each record In m_details(name)
            copy.Add record
        Next record
    End If
    Set GetDetail = copy
End Function

Public Function TryMerge( _
    ByVal other As obj_LookupRecord, _
    ByVal overwrite As Boolean _
) As Boolean
    Dim fields As Object
    Dim name As Variant

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If other Is Nothing Then Exit Function
    Set fields = other.Fields
    For Each name In fields.Keys
        If overwrite Or Not m_fields.Exists(name) Then m_fields(name) = fields(name)
    Next name
    TryMerge = True
End Function