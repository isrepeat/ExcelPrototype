VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiControlFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControlFactory
Implements obj_IUiMarkupSchemaProvider

Private m_type As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IUiControlFactory_Create() As obj_IUiControl
    Dim control As obj_IUiControl

    Select Case m_type
        Case "label"
            Set control = New obj_UiLabelControl
        Case "button"
            Set control = New obj_UiButtonControl
        Case "table"
            Set control = New obj_UiTableControl
        Case "tablelist"
            Set control = New obj_UiTableListControl
        Case "form"
            Set control = New obj_UiFormControl
        Case "input", "select"
            Set control = New obj_UiFieldControl
    End Select
    If control Is Nothing Then Exit Function
    If Not control.Initialize() Then Exit Function
    Set obj_IUiControlFactory_Create = control
End Function

Private Function obj_IUiMarkupSchemaProvider_GetSchema() As obj_UiMarkupSchema
    Set obj_IUiMarkupSchemaProvider_GetSchema = private_CreateSchema()
End Function

' //
' // API
' //
Public Sub Dispose()
    m_type = VBA.vbNullString
End Sub

Public Sub Initialize(ByVal controlType As String)
    m_type = VBA.LCase$(controlType)
End Sub

' //
' // Private
' //
Private Function private_CreateSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema
    Dim field As obj_UiMarkupSchema

    schema.AddAttribute "name", "string", False, "", False, 0
    schema.AddAttribute "dataContext", "context", False, "", True, 0
    schema.AddAttribute "row", "positive", False, "", False, 0
    schema.AddAttribute "column", "positive", False, "", False, 0
    schema.AddAttribute "rowSpan", "positive", False, "", False, 0
    schema.AddAttribute "columnSpan", "positive", False, "", False, 0
    schema.AddAttribute "name", "string", True, "", False, 0
    schema.AddAttribute "style", "string", False, "", True, 0

    Select Case m_type
        Case "label", "button"
            schema.RequireAnyAttribute "text|caption"
            schema.AddAttribute "text", "string", False, "", True, 0
            schema.AddAttribute "caption", "string", False, "", True, 0
            If m_type = "button" Then
                schema.AddAttribute "command", "string", False, "", True
            End If
        Case "table", "tablelist"
            schema.RequireAnyAttribute "itemsSource"
            schema.AddAttribute "itemsSource", "string", False, "", True, 0
            If m_type = "tablelist" Then schema.AddAttribute "gapRows", "nonnegative", False, "", False, 0
            schema.AddAttribute "showHeaders", "boolean", False, "", False, 0

        Case "input", "select"
            If m_type = "select" Then
                schema.RequireAnyAttribute "items|itemsSource"
            End If
            schema.AddAttribute "value", "string", True, "", True, 0
            schema.AddAttribute "inputType", "enum", False, "text|select|checkbox", False, 0
            schema.AddAttribute "readOnly", "boolean", False, "", False, 0
            schema.AddAttribute "items", "string", False, "", True, 0
            schema.AddAttribute "itemsSource", "string", False, "", True, 0
            schema.AddAttribute "onChange", "string", False, "", True, 0
            schema.AddAttribute "panelStyle", "string", False, "", True, 0
            schema.AddAttribute "itemStyle", "string", False, "", True, 0
            schema.AddAttribute "itemHeight", "number", False, "", False, 0
            schema.AddAttribute "itemMargin", "number", False, "", False, 0

        Case "form"
            schema.AddAttribute "orientation", "enum", True, "horizontal|vertical", False, 0
            schema.AddAttribute "labelColumnSpan", "positive", False, "", False, 0
            schema.AddAttribute "columnSpan", "positive", False, "", False, 0
            schema.AddAttribute "rowSpan", "positive", False, "", False, 0
            schema.AddAttribute "labelPosition", "enum", False, "left|top", False, 0
            schema.AddAttribute "labelStyle", "string", False, "", True, 0
            schema.AddAttribute "fieldStyle", "string", False, "", True, 0
            schema.AddAttribute "onChange", "string", False, "", True, 0
            schema.AddAttribute "readOnly", "boolean", False, "", False, 0
            Set field = private_FieldSchema()
            schema.AddChild "field", field
            schema.AllowControlChildren
        Case Else
            Exit Function
    End Select
    Set private_CreateSchema = schema
End Function

Private Function private_FieldSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema

    schema.AddAttribute "name", "string", True, "", False, 0
    schema.AddAttribute "label", "string", True, "", False, 0
    schema.AddAttribute "type", "enum", True, "text|select|checkbox", False, 0
    schema.AddAttribute "labelColumnSpan", "positive", False, "", False, 0
    schema.AddAttribute "columnSpan", "positive", False, "", False, 0
    schema.AddAttribute "rowSpan", "positive", False, "", False, 0
    schema.AddAttribute "labelPosition", "enum", False, "left|top", False, 0
    schema.AddAttribute "labelStyle", "string", False, "", True, 0
    schema.AddAttribute "fieldStyle", "string", False, "", True, 0
    schema.AddAttribute "onChange", "string", False, "", True, 0
    schema.AddAttribute "readOnly", "boolean", False, "", False, 0
    schema.AddAttribute "required", "boolean", False, "", False, 0
    schema.AddAttribute "value", "string", False, "", True, 0
    schema.AddAttribute "itemsSource", "string", False, "", True, 0
    schema.AddAttribute "items", "string", False, "", True, 0
    schema.AddAttribute "panelStyle", "string", False, "", True, 0
    schema.AddAttribute "itemStyle", "string", False, "", True, 0
    schema.AddAttribute "itemHeight", "number", False, "", False, 0
    schema.AddAttribute "itemMargin", "number", False, "", False, 0
    Set private_FieldSchema = schema
End Function