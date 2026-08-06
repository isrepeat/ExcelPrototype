VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SqlMetaRowItem"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IClonable

Private m_SourceRow As obj_Row
Private m_ItemIdentity As obj_ItemIdentity
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IClonable_Clone(Optional ByVal targetColumnCount As Long = 0) As Object
    Set obj_IClonable_Clone = Me.Clone(targetColumnCount)
End Function

' //
' // Properties
' //
Public Property Get SourceRow() As obj_Row
    Set SourceRow = m_SourceRow
End Property

Public Property Set SourceRow(ByVal value As obj_Row)
    Set m_SourceRow = Nothing
    If value Is Nothing Then Exit Property
    Set m_SourceRow = value.Clone
End Property

Public Property Get ItemIdentity() As obj_ItemIdentity
    Set ItemIdentity = m_ItemIdentity
End Property

Public Property Set ItemIdentity(ByVal value As obj_ItemIdentity)
    Set m_ItemIdentity = Nothing
    If value Is Nothing Then
        Set m_ItemIdentity = New obj_ItemIdentity
        Exit Property
    End If

    Set m_ItemIdentity = value.Clone
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal sourceRow As obj_Row, _
    ByVal itemIdentity As obj_ItemIdentity _
) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    Set Me.SourceRow = sourceRow
    Set Me.ItemIdentity = itemIdentity
    If Not sourceRow Is Nothing Then
        If m_SourceRow Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "SqlMetaRowItem.Initialize: failed to clone SourceRow."
#End If
            Exit Function
        End If
    End If
    If m_ItemIdentity Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SqlMetaRowItem.Initialize: ItemIdentity is Nothing."
#End If
        Exit Function
    End If
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    Set m_SourceRow = Nothing
    Set m_ItemIdentity = Nothing
    On Error GoTo 0
End Sub

Public Function Clone(Optional ByVal targetColumnCount As Long = 0) As obj_SqlMetaRowItem
    Dim result As obj_SqlMetaRowItem

    Set result = New obj_SqlMetaRowItem
    If result Is Nothing Then Exit Function
    If Not result.Initialize(m_SourceRow, m_ItemIdentity) Then Exit Function

    Set Clone = result
End Function
