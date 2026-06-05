VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Cell"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Private Const VIRTUAL_MARKER As String = "__virtual"

Private m_Value As String
Private m_Desc As String
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    ' Не вызываем Dispose из деструктора: при освобождении obj_Row
    ' рантайм сам отпускает obj_Cell, а ручная цепочка Dispose внутри
    ' Class_Terminate усложняет безопасное освобождение строк.
    m_IsDisposed = True
End Sub

' //
' // Properties
' //
Public Property Get Value() As String
    Value = m_Value
End Property

Public Property Let Value(ByVal valueText As String)
    m_Value = VBA.CStr(valueText)
End Property

Public Property Get Desc() As String
    Desc = m_Desc
End Property

Public Property Let Desc(ByVal valueText As String)
    m_Desc = VBA.CStr(valueText)
End Property

Public Property Get IsVirtual() As Boolean
    IsVirtual = (VBA.InStr(1, m_Desc, VIRTUAL_MARKER, VBA.vbTextCompare) > 0)
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
End Sub

Public Function MarkAsVirtual(Optional ByVal extraDesc As String = VBA.vbNullString) As Boolean
    If VBA.Len(VBA.Trim$(extraDesc)) > 0 Then
        m_Desc = VIRTUAL_MARKER & ":" & VBA.Trim$(extraDesc)
    Else
        m_Desc = VIRTUAL_MARKER
    End If
    MarkAsVirtual = True
End Function

Public Function Clone() As obj_Cell
    Dim result As obj_Cell

    Set result = New obj_Cell
    result.Value = m_Value
    result.Desc = m_Desc

    Set Clone = result
End Function
