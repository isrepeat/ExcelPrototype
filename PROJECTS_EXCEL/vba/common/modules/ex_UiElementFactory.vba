Attribute VB_Name = "ex_UiElementFactory"
Option Explicit
Private m_factories As Object
Private m_initialized As Boolean

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_factories = Nothing
    m_initialized = False
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Register(ByVal tag As String, ByVal factory As obj_IUiElementFactory)
    If m_factories Is Nothing Then
        Set m_factories = VBA.CreateObject("Scripting.Dictionary")
        m_factories.CompareMode = VBA.vbTextCompare
    End If
    Set m_factories(tag) = factory
End Sub

Public Function fn_Create( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As obj_IUiElement
    Dim tag As Variant
    Dim factory As obj_IUiElementFactory
    Dim builtInFactory As obj_UiElementFactory

    If Not m_initialized Then
        For Each tag In VBA.Array("page", "grid", "stackPanel", "form", "control")
            Set builtInFactory = New obj_UiElementFactory
            builtInFactory.Initialize VBA.LCase$(VBA.CStr(tag))
            fn_Register VBA.CStr(tag), builtInFactory
        Next tag
        m_initialized = True
    End If
    tag = VBA.CStr(definition.baseName)
    If Not m_factories.Exists(tag) Then
        diagnostic = "Unsupported visual tag: " & tag
        Exit Function
    End If
    Set factory = m_factories(tag)
    Set fn_Create = factory.Create(definition, context, source, diagnostic)
End Function

Public Function fn_Attribute(ByVal node As Object, ByVal name As String) As String
    Dim value As Variant

    value = node.getAttribute(name)
    If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then fn_Attribute = VBA.CStr(value)
End Function

Public Function fn_Long( _
    ByVal node As Object, _
    ByVal name As String, _
    ByVal defaultValue As Long _
) As Long
    Dim value As String

    value = fn_Attribute(node, name)
    If VBA.Len(value) = 0 Then
        fn_Long = defaultValue
    ElseIf VBA.IsNumeric(value) Then
        If VBA.CDbl(value) <> VBA.Fix(VBA.CDbl(value)) Then
            VBA.Err.Raise VBA.vbObjectError + 2202, , "Integer layout value required: " & name
        End If
        fn_Long = VBA.CLng(value)
        If fn_Long < 1 Then VBA.Err.Raise VBA.vbObjectError + 2201, , "Positive layout value required: " & name
    Else
        VBA.Err.Raise VBA.vbObjectError + 2202, , "Integer layout value required: " & name
    End If
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------