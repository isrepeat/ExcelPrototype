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

#Const ENABLE_LOGGING_VERBOSE = False

Implements obj_IUiControlFactory
Implements obj_IUiMarkupSchemaProvider

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

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

#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CLASS_NEW_STARTED | Type=" & m_type
#End If
    Select Case m_type
        Case "label"
            Set control = New obj_UiLabelControl
        Case "button"
            Set control = New obj_UiButtonControl
        Case "checkbox"
            Set control = New obj_UiCheckboxControl
        Case "table"
            Set control = New obj_UiTableControl
        Case "candidates"
            Set control = New obj_UiCandidatesControl
        Case "tablelist"
            Set control = New obj_UiTableListControl
        Case "items"
            Set control = New obj_UiItemsControl
        Case "form"
            Set control = New obj_UiFormControl
        Case "input", "select"
            Set control = New obj_UiFieldControl
    End Select
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CLASS_NEW_COMPLETED | Type=" & m_type
#End If
    If control Is Nothing Then Exit Function
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CLASS_INITIALIZE_STARTED | Type=" & m_type
#End If
    If Not control.Initialize() Then Exit Function
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CLASS_INITIALIZE_COMPLETED | Type=" & m_type
#End If
    Set obj_IUiControlFactory_Create = control
End Function

Private Function obj_IUiMarkupSchemaProvider_GetSchema() As obj_UiMarkupSchema
    Set obj_IUiMarkupSchemaProvider_GetSchema = private_CreateSchema()
End Function

' //
' // API
' //
Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    m_type = VBA.vbNullString
End Sub

Public Sub Initialize(ByVal controlType As String)
    If m_isDisposed Or m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2166, , "Object is already initialized or disposed."
    End If
    m_type = VBA.LCase$(controlType)
    m_isInitialized = True
End Sub

' //
' // Private
' //
Private Function private_CreateSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema
    Dim field As obj_UiMarkupSchema
    Dim wrapper As obj_UiMarkupSchema
    Dim policy As obj_UiMarkupSchema
    Dim rule As obj_UiMarkupSchema

    If Not schema.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    schema.AddAttribute "name", "string", False, "", False, 0
    schema.AddAttribute "dataContext", "context", False, "", True, 0
    schema.AddAttribute "row", "positive", False, "", False, 0
    schema.AddAttribute "column", "positive", False, "", False, 0
    schema.AddAttribute "rowSpan", "positive", False, "", False, 0
    schema.AddAttribute "columnSpan", "positive", False, "", False, 0
    schema.AddAttribute "name", "string", True, "", False, 0
    schema.AddAttribute "style", "string", False, "", True, 0

    Select Case m_type
        Case "label", "button", "checkbox"
            schema.RequireAnyAttribute "text|caption"
            schema.AddAttribute "text", "string", False, "", True, 0
            schema.AddAttribute "caption", "string", False, "", True, 0
            If m_type = "checkbox" Then
                schema.AddAttribute "onChange", "string", False, "", True, 0
                schema.AddAttribute "commandParameter", "string", False, "", True, 0
                schema.AddAttribute "indeterminate", "boolean", False, "", True, 0
                schema.AddAttribute "indeterminateGlyph", "string", False, "", True, 0
                schema.AddAttribute "value", "string", True, "", True, 0
                schema.AddAttribute "checkedGlyph", "string", True, "", True, 0
                schema.AddAttribute "uncheckedGlyph", "string", True, "", True, 0
            End If
            If m_type = "button" Then
                private_AddShapeLayoutSchema schema
                schema.AddAttribute "command", "string", False, "", True
                schema.AddAttribute "enabled", "boolean", False, "", True
            End If
        Case "items"
            schema.RequireAnyAttribute "itemsSource"
            schema.AddAttribute "itemsSource", "string", True, "", True, 0
            schema.AddAttribute "gapRows", "nonnegative", False, "", False, 0
            schema.AllowControlChildren
        Case "table", "tablelist", "candidates"
            schema.RequireAnyAttribute "itemsSource"
            schema.AddAttribute "itemsSource", "string", False, "", True, 0
            If m_type = "tablelist" Then schema.AddAttribute "gapRows", "nonnegative", False, "", False, 0
            schema.AddAttribute "showHeaders", "boolean", False, "", False, 0
            schema.AddAttribute "showAutoFilter", "boolean", False, "", False, 0
            schema.AddAttribute "representation", "enum", False, "Raw|Smart", False, 0
            If m_type = "table" Then schema.AddAttribute "excelTableName", "string", False, "", False, 255
            If m_type = "table" Then
                Set wrapper = New obj_UiMarkupSchema
                Set rule = New obj_UiMarkupSchema
                If Not wrapper.Initialize() Or Not rule.Initialize() Then Err.Raise 5, , "Cannot initialize cell template schema."
                rule.AddAttribute "columnAlias", "string", True, "", False, 0
                rule.AddAttribute "mode", "enum", False, "replace|overlay", False, 0
                rule.AllowControlChildren
                wrapper.AddChild "cellTemplate", rule, 1, -1, "urn:excelprototype:controls"
                schema.AddChild "table.cellTemplates", wrapper, 0, 1, "urn:excelprototype:controls"
            End If
            If m_type = "tablelist" Then
                Set wrapper = New obj_UiMarkupSchema
                Set policy = New obj_UiMarkupSchema
                Set rule = New obj_UiMarkupSchema
                If Not wrapper.Initialize() Or Not policy.Initialize() Or Not rule.Initialize() Then
                    Err.Raise VBA.vbObjectError + 2167, , "Table policy schema initialization failed."
                End If
                policy.AddAttribute "defaultRepresentation", "enum", False, "Raw|Smart"
                rule.AddAttribute "index", "tableindices", True
                rule.AddAttribute "representation", "enum", True, "Raw|Smart"
                rule.AddAttribute "excelTableName", "string", False, "", False, 255
                policy.AddChild "rule", rule, 0, -1, "urn:excelprototype:controls"
                wrapper.AddChild "tablePolicy", policy, 1, 1, "urn:excelprototype:controls"
                schema.AddChild "tableList.tablePolicy", wrapper, 0, 1, "urn:excelprototype:controls"
            End If
            If m_type = "candidates" Then
                schema.AddAttribute "alignToForm", "string", True, "", False, 0
                schema.AddAttribute "columnsBefore", "string", False, "", False, 0
                schema.AddAttribute "columnsAfter", "string", False, "", False, 0
                schema.AddAttribute "workspaceFirstColumn", "positive", False, "", False, 0
                schema.AddAttribute "workspaceLastColumn", "positive", False, "", False, 0
            End If
            If m_type = "table" Or m_type = "candidates" Then
                schema.AddAttribute "selectionMode", "enum", m_type = "candidates", "Focus|Confirm", True, 0
                schema.AddAttribute "onSelect", "string", False, "", True, 0
                schema.AddAttribute "selectedItem", "string", False, "", True, 0
                schema.AddAttribute "selectedStyle", "string", False, "", True, 0
            End If

        Case "input", "select"
            private_AddShapeLayoutSchema schema
            If m_type = "select" Then
                schema.RequireAnyAttribute "items|itemsSource"
            End If
            schema.AddAttribute "value", "string", True, "", True, 0
            schema.AddAttribute "inputType", "enum", False, "text|select|checkbox", False, 0
            schema.AddAttribute "readOnly", "boolean", False, "", True, 0
            schema.AddAttribute "editable", "boolean", False, "", False, 0
            schema.AddAttribute "onFocus", "string", False, "", True, 0
            schema.AddAttribute "highlightValue", "string", False, "", True, 0
            schema.AddAttribute "highlightStyle", "string", False, "", True, 0
            schema.AddAttribute "items", "string", False, "", True, 0
            schema.AddAttribute "itemsSource", "string", False, "", True, 0
            schema.AddAttribute "onChange", "string", False, "", True, 0
            schema.AddAttribute "panelStyle", "string", False, "", True, 0
            schema.AddAttribute "itemStyle", "string", False, "", True, 0
            schema.AddAttribute "itemHeight", "number", False, "", False, 0
            schema.AddAttribute "itemMargin", "number", False, "", False, 0

        Case "form"
            schema.AddAttribute "layoutSource", "string", False, "", False, 0
            schema.AddAttribute "layout", "string", False, "", True, 0
            schema.AddAttribute "orientation", "enum", True, "horizontal|vertical", False, 0
            schema.AddAttribute "labelPosition", "enum", False, "left|top", False, 0
            schema.AddAttribute "labelStyle", "string", False, "", True, 0
            schema.AddAttribute "fieldStyle", "string", False, "", True, 0
            schema.AddAttribute "onChange", "string", False, "", True, 0
            schema.AddAttribute "readOnly", "boolean", False, "", True, 0
            Set field = private_FieldSchema()
            schema.AddChild "field", field
            schema.AllowControlChildren
        Case Else
            Exit Function
    End Select
    Set private_CreateSchema = schema
End Function

Private Sub private_AddShapeLayoutSchema(ByVal schema As obj_UiMarkupSchema)
    schema.AddAttribute "shapePlacement", "enum", False, "freeFloating|move|moveAndSize", False, 0
    schema.AddAttribute "shapeWidth", "string", False, "", False, 0
    schema.AddAttribute "shapeHeight", "string", False, "", False, 0
    schema.AddAttribute "shapeHorizontalAlignment", "enum", False, "left|center|right", False, 0
    schema.AddAttribute "shapeVerticalAlignment", "enum", False, "top|center|bottom", False, 0
End Sub

Private Function private_FieldSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema

    If Not schema.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    schema.AddAttribute "name", "string", True, "", False, 0
    schema.AddAttribute "label", "string", True, "", True, 0
    schema.AddAttribute "type", "enum", True, "text|select|checkbox", False, 0
    schema.AddAttribute "labelPosition", "enum", False, "left|top", False, 0
    schema.AddAttribute "columnSpan", "positive", False, "", False, 0
    schema.AddAttribute "labelStyle", "string", False, "", True, 0
    schema.AddAttribute "fieldStyle", "string", False, "", True, 0
    schema.AddAttribute "onFocus", "string", False, "", True, 0
    schema.AddAttribute "onChange", "string", False, "", True, 0
    schema.AddAttribute "readOnly", "boolean", False, "", True, 0
    schema.AddAttribute "required", "boolean", False, "", False, 0
    schema.AddAttribute "value", "string", True, "", True, 0
    schema.AddAttribute "itemsSource", "string", False, "", True, 0
    schema.AddAttribute "items", "string", False, "", True, 0
    schema.AddAttribute "panelStyle", "string", False, "", True, 0
    schema.AddAttribute "itemStyle", "string", False, "", True, 0
    schema.AddAttribute "itemHeight", "number", False, "", False, 0
    schema.AddAttribute "itemMargin", "number", False, "", False, 0
    Set private_FieldSchema = schema
End Function