VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Column"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_Name As String
Private m_Position As Long
Private m_FormatKind As String
Private m_AliasesByKey As Object

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_AliasesByKey = VBA.CreateObject("Scripting.Dictionary")
    m_AliasesByKey.CompareMode = 1
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Set m_AliasesByKey = Nothing
    On Error GoTo 0
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal value As String)
    m_Name = VBA.Trim$(value)
End Property

Public Property Get Position() As Long
    Position = m_Position
End Property

Public Property Let Position(ByVal value As Long)
    If value > 0 Then
        m_Position = value
    Else
        m_Position = 0
    End If
End Property

Public Property Get FormatKind() As String
    FormatKind = m_FormatKind
End Property

Public Property Let FormatKind(ByVal value As String)
    m_FormatKind = VBA.LCase$(VBA.Trim$(VBA.CStr(value)))
End Property

Public Property Get Aliases() As Collection
    Dim result As Collection
    Dim aliasKey As Variant

    Set result = New Collection
    If Not m_AliasesByKey Is Nothing Then
        For Each aliasKey In m_AliasesByKey.Keys
            result.Add VBA.CStr(m_AliasesByKey(aliasKey))
        Next aliasKey
    End If

    Set Aliases = result
End Property

Public Function AddAlias(ByVal aliasName As String) As Boolean
    Dim normalizedKey As String

    aliasName = VBA.Trim$(VBA.CStr(aliasName))
    If VBA.Len(aliasName) = 0 Then Exit Function

    If m_AliasesByKey Is Nothing Then
        Set m_AliasesByKey = VBA.CreateObject("Scripting.Dictionary")
        m_AliasesByKey.CompareMode = 1
    End If

    normalizedKey = private_NormalizeAliasKey(aliasName)
    If VBA.Len(normalizedKey) = 0 Then Exit Function

    If m_AliasesByKey.Exists(normalizedKey) Then
        AddAlias = True
        Exit Function
    End If

    m_AliasesByKey.Add normalizedKey, aliasName
    AddAlias = True
End Function

Public Function HasAlias(ByVal aliasName As String) As Boolean
    Dim normalizedKey As String

    If m_AliasesByKey Is Nothing Then Exit Function
    normalizedKey = private_NormalizeAliasKey(aliasName)
    If VBA.Len(normalizedKey) = 0 Then Exit Function

    HasAlias = m_AliasesByKey.Exists(normalizedKey)
End Function

Public Sub ClearAliases()
    If m_AliasesByKey Is Nothing Then Exit Sub
    m_AliasesByKey.RemoveAll
End Sub

Private Function private_NormalizeAliasKey(ByVal aliasName As String) As String
    aliasName = VBA.Trim$(VBA.CStr(aliasName))
    If VBA.Len(aliasName) = 0 Then Exit Function
    private_NormalizeAliasKey = VBA.LCase$(aliasName)
End Function
