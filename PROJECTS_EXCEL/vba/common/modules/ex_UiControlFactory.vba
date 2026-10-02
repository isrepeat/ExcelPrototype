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
Public Sub fn_Register(ByVal controlType As String, ByVal factory As obj_IUiControlFactory, Optional ByVal namespaceUri As String = "urn:excelprototype:controls")
    If m_factories Is Nothing Then
        Set m_factories = VBA.CreateObject("Scripting.Dictionary")
        m_factories.CompareMode = VBA.vbBinaryCompare
    End If
    Set m_factories(namespaceUri & "|" & VBA.LCase$(controlType)) = factory
End Sub

Public Function fn_TryGetSchema(ByVal name As String, ByRef schema As obj_UiMarkupSchema, ByRef diagnostic As String, Optional ByVal namespaceUri As String = "urn:excelprototype:controls") As Boolean
    Dim provider As obj_IUiMarkupSchemaProvider
    Dim factory As obj_IUiControlFactory

    Set schema = Nothing
    private_Initialize
    name = namespaceUri & "|" & VBA.LCase$(name)
    If Not m_factories.Exists(name) Then
        diagnostic = "Unknown Control: " & name
        Exit Function
    End If
    Set factory = m_factories(name)
    If Not TypeOf factory Is obj_IUiMarkupSchemaProvider Then
        diagnostic = "Markup schema provider is required: " & name
        Exit Function
    End If
    Set provider = factory
    Set schema = provider.GetSchema()
    fn_TryGetSchema = Not schema Is Nothing
    If schema Is Nothing Then diagnostic = "Markup schema is missing: " & name
End Function

Public Function fn_Create(ByVal controlNode As Object) As obj_IUiControl
    Dim controlType As String
    Dim builtInType As Variant
    Dim builtInFactory As obj_UiControlFactory
    Dim factory As obj_IUiControlFactory

    private_Initialize
    controlType = "urn:excelprototype:controls|" & VBA.LCase$(ex_UiElementFactory.fn_Attribute(controlNode, "type"))
    If Not m_factories.Exists(controlType) Then Exit Function
    Set factory = m_factories(controlType)
    Set fn_Create = factory.Create()
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_Initialize()
    Dim builtInType As Variant
    Dim builtInFactory As obj_UiControlFactory

    If Not m_initialized Then
        For Each builtInType In VBA.Array("Label", "Button", "Table", "TableList", "Input", "Select", "Form")
            If m_factories Is Nothing Then
                Set m_factories = VBA.CreateObject("Scripting.Dictionary")
                m_factories.CompareMode = VBA.vbBinaryCompare
            End If
            If Not m_factories.Exists("urn:excelprototype:controls|" & VBA.LCase$(VBA.CStr(builtInType))) Then
                Set builtInFactory = New obj_UiControlFactory
                builtInFactory.Initialize VBA.CStr(builtInType)
                fn_Register VBA.CStr(builtInType), builtInFactory
            End If
        Next builtInType
        m_initialized = True
    End If
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------