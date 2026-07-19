VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Cell"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False
#Const CELL_BUTTON_VIEW_ENABLED = False

Private Const VIRTUAL_MARKER As String = "__virtual"

Private m_Value As String
Private m_Desc As String
Private m_Tags As Collection
#If CELL_BUTTON_VIEW_ENABLED Then
' Private m_IsButtonView As Boolean
' Private m_ButtonActionArg As Variant
' Private m_ButtonActionArgIsObject As Boolean
#End If
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
    Set m_Tags = New Collection
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

#If CELL_BUTTON_VIEW_ENABLED Then
' Public Property Get IsButtonView() As Boolean
'     IsButtonView = m_IsButtonView
' End Property

' Public Property Let IsButtonView(ByVal value As Boolean)
'     m_IsButtonView = VBA.CBool(value)
' End Property

' Public Property Get ButtonActionArg() As Variant
'     If m_ButtonActionArgIsObject Then
'         Set ButtonActionArg = m_ButtonActionArg
'     Else
'         ButtonActionArg = m_ButtonActionArg
'     End If
' End Property

' Public Property Let ButtonActionArg(ByVal value As Variant)
'     m_ButtonActionArgIsObject = False
'     m_ButtonActionArg = value
' End Property

' Public Property Set ButtonActionArg(ByVal value As Object)
'     m_ButtonActionArgIsObject = True
'     Set m_ButtonActionArg = value
' End Property

' Public Property Get ButtonActionArgIsObject() As Boolean
'     ButtonActionArgIsObject = m_ButtonActionArgIsObject
' End Property
#End If

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
    Set m_Tags = Nothing
End Sub

' Семантические теги не задают оформление сами. Renderer публикует каждый
' тег как controlPart `tag-<имя>`, после чего внешний style selector решает,
' как должен выглядеть соответствующий диапазон.
Public Function AddTag(ByVal tagName As String) As Boolean
    Dim normalizedTag As String

    normalizedTag = private_NormalizeTag(tagName)
    If VBA.Len(normalizedTag) = 0 Then Exit Function
    If m_Tags Is Nothing Then Set m_Tags = New Collection
    If Me.HasTag(normalizedTag) Then
        AddTag = True
        Exit Function
    End If

    m_Tags.Add normalizedTag
    AddTag = True
End Function

Public Function HasTag(ByVal tagName As String) As Boolean
    Dim normalizedTag As String
    Dim tagItem As Variant

    normalizedTag = private_NormalizeTag(tagName)
    If VBA.Len(normalizedTag) = 0 Then Exit Function
    If m_Tags Is Nothing Then Exit Function

    For Each tagItem In m_Tags
        If VBA.StrComp(VBA.CStr(tagItem), normalizedTag, VBA.vbBinaryCompare) = 0 Then
            HasTag = True
            Exit Function
        End If
    Next tagItem
End Function

Public Property Get Tags() As Collection
    Dim result As Collection
    Dim tagItem As Variant

    Set result = New Collection
    If Not m_Tags Is Nothing Then
        For Each tagItem In m_Tags
            result.Add VBA.CStr(tagItem)
        Next tagItem
    End If
    Set Tags = result
End Property

Public Function MarkAsVirtual(Optional ByVal extraDesc As String = VBA.vbNullString) As Boolean
    If VBA.Len(VBA.Trim$(extraDesc)) > 0 Then
        m_Desc = VIRTUAL_MARKER & ":" & VBA.Trim$(extraDesc)
    Else
        m_Desc = VIRTUAL_MARKER
    End If
    MarkAsVirtual = True
End Function

#If CELL_BUTTON_VIEW_ENABLED Then
' Public Function MarkAsButtonView(Optional ByVal actionArg As Variant) As Boolean
'     m_IsButtonView = True
'     If Not VBA.IsMissing(actionArg) Then
'         If VBA.IsObject(actionArg) Then
'             Set Me.ButtonActionArg = actionArg
'         Else
'             Me.ButtonActionArg = actionArg
'         End If
'     End If
'     MarkAsButtonView = True
' End Function
#End If

Public Function Clone() As obj_Cell
    Dim result As obj_Cell
    Dim tagItem As Variant

    Set result = New obj_Cell
    result.Value = m_Value
    result.Desc = m_Desc
    If Not m_Tags Is Nothing Then
        For Each tagItem In m_Tags
            If Not result.AddTag(VBA.CStr(tagItem)) Then Exit Function
        Next tagItem
    End If
#If CELL_BUTTON_VIEW_ENABLED Then
'     result.IsButtonView = m_IsButtonView
'     If m_ButtonActionArgIsObject Then
'         Set result.ButtonActionArg = m_ButtonActionArg
'     Else
'         result.ButtonActionArg = m_ButtonActionArg
'     End If
#End If

    Set Clone = result
End Function

Private Function private_NormalizeTag(ByVal tagName As String) As String
    Dim normalizedTag As String
    Dim charIndex As Long
    Dim charText As String

    normalizedTag = VBA.LCase$(VBA.Trim$(tagName))
    If VBA.Len(normalizedTag) = 0 Then Exit Function
    For charIndex = 1 To VBA.Len(normalizedTag)
        charText = VBA.Mid$(normalizedTag, charIndex, 1)
        If Not (charText Like "[a-z0-9_-]") Then Exit Function
    Next charIndex
    private_NormalizeTag = normalizedTag
End Function
