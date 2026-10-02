Attribute VB_Name = "ex_UiControlFactory"
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
Public Sub fn_Register(ByVal controlType As String, ByVal factory As obj_IUiControlFactory)
    If m_factories Is Nothing Then
        Set m_factories = VBA.CreateObject("Scripting.Dictionary")
        m_factories.CompareMode = VBA.vbTextCompare
    End If
    Set m_factories(controlType) = factory
End Sub

Public Function fn_Create(ByVal controlNode As Object) As obj_IUiControl
    Dim controlType As String
    Dim builtInType As Variant
    Dim builtInFactory As obj_UiControlFactory
    Dim factory As obj_IUiControlFactory

    If Not m_initialized Then
        For Each builtInType In VBA.Array("Label", "Button", "Table", "Input", "Select", "Form")
            If m_factories Is Nothing Then
                Set m_factories = VBA.CreateObject("Scripting.Dictionary")
                m_factories.CompareMode = VBA.vbTextCompare
            End If
            If Not m_factories.Exists(builtInType) Then
                Set builtInFactory = New obj_UiControlFactory
                builtInFactory.Initialize VBA.CStr(builtInType)
                fn_Register VBA.CStr(builtInType), builtInFactory
            End If
        Next builtInType
        m_initialized = True
    End If
    controlType = ex_UiElementFactory.fn_Attribute(controlNode, "type")
    If Not m_factories.Exists(controlType) Then Exit Function
    Set factory = m_factories(controlType)
    Set fn_Create = factory.Create()
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------