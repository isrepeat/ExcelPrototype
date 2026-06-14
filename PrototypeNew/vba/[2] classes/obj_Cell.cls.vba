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
Private m_IsButtonView As Boolean
Private m_ButtonActionArg As Variant
Private m_ButtonActionArgIsObject As Boolean
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

Public Property Get IsButtonView() As Boolean
    IsButtonView = m_IsButtonView
End Property

Public Property Let IsButtonView(ByVal value As Boolean)
    m_IsButtonView = VBA.CBool(value)
End Property

Public Property Get ButtonActionArg() As Variant
    If m_ButtonActionArgIsObject Then
        Set ButtonActionArg = m_ButtonActionArg
    Else
        ButtonActionArg = m_ButtonActionArg
    End If
End Property

Public Property Let ButtonActionArg(ByVal value As Variant)
    m_ButtonActionArgIsObject = False
    m_ButtonActionArg = value
End Property

Public Property Set ButtonActionArg(ByVal value As Object)
    m_ButtonActionArgIsObject = True
    Set m_ButtonActionArg = value
End Property

Public Property Get ButtonActionArgIsObject() As Boolean
    ButtonActionArgIsObject = m_ButtonActionArgIsObject
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

Public Function MarkAsButtonView(Optional ByVal actionArg As Variant) As Boolean
    m_IsButtonView = True
    If Not VBA.IsMissing(actionArg) Then
        If VBA.IsObject(actionArg) Then
            Set Me.ButtonActionArg = actionArg
        Else
            Me.ButtonActionArg = actionArg
        End If
    End If
    MarkAsButtonView = True
End Function

Public Function Clone() As obj_Cell
    Dim result As obj_Cell

    Set result = New obj_Cell
    result.Value = m_Value
    result.Desc = m_Desc
    result.IsButtonView = m_IsButtonView
    If m_ButtonActionArgIsObject Then
        Set result.ButtonActionArg = m_ButtonActionArg
    Else
        result.ButtonActionArg = m_ButtonActionArg
    End If

    Set Clone = result
End Function
