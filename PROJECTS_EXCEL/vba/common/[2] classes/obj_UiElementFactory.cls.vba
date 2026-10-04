VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiElementFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiElementFactory
Implements obj_IUiMarkupSchemaProvider

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_kind As String

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
Private Function obj_IUiElementFactory_Create( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As obj_IUiElement
    Set obj_IUiElementFactory_Create = Me.Create(definition, context, source, diagnostic)
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
    m_kind = VBA.vbNullString
End Sub

Public Sub Initialize(ByVal kind As String)
    If m_isDisposed Or m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2166, , "Object is already initialized or disposed."
    End If
    m_kind = kind
    m_isInitialized = True
End Sub

Public Function Create( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As obj_IUiElement
    Dim element As obj_IUiElement

    On Error GoTo EH_CONFIGURE
    Select Case m_kind
        Case "page"
            Set element = New obj_UiPageElement
        Case "grid"
            Set element = New obj_UiGridElement
        Case "stackpanel"
            Set element = New obj_UiStackPanelElement
    End Select
    If element Is Nothing Then Exit Function
    If Not element.Configure(definition, context, source, diagnostic) Then
        element.Dispose
        Exit Function
    End If
    context.RegisterElement ex_UiElementFactory.fn_Attribute(definition, "name"), element
    Set Create = element
    Exit Function
EH_CONFIGURE:
    diagnostic = VBA.Err.Description
    If Not element Is Nothing Then element.Dispose
End Function

' //
' // Private
' //
Private Function private_CreateSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema

    If Not schema.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    Select Case m_kind
        Case "page"
            schema.AddAttribute "name", "string", False, "", False, 0
            schema.AddAttribute "dataContext", "context", False, "", True, 0
            schema.AddAttribute "row", "positive", False, "", False, 0
            schema.AddAttribute "column", "positive", False, "", False, 0
            schema.AddAttribute "version", "positive", False, "", False, 0
            schema.AddChild "grid", Nothing, 1, 1
            schema.AddChild "styles", private_StylesSchema(), 0, 1
        Case "grid"
            schema.AddAttribute "name", "string", False, "", False, 0
            schema.AddAttribute "dataContext", "context", False, "", True, 0
            schema.AddAttribute "row", "positive", False, "", False, 0
            schema.AddAttribute "column", "positive", False, "", False, 0
            schema.AddAttribute "anchorCell", "string", False, "", False, 0
            schema.AddChild "grid.rowDefinitions", private_TrackSchema("rowDefinition"), 0, 1
            schema.AddChild "grid.columnDefinitions", private_TrackSchema("columnDefinition"), 0, 1
            schema.AddAttribute "rowSpan", "positive", False, "", False, 0
            schema.AddAttribute "columnSpan", "positive", False, "", False, 0
            schema.AllowVisualChildren
        Case "stackpanel"
            schema.AddAttribute "name", "string", False, "", False, 0
            schema.AddAttribute "dataContext", "context", False, "", True, 0
            schema.AddAttribute "row", "positive", False, "", False, 0
            schema.AddAttribute "column", "positive", False, "", False, 0
            schema.AddAttribute "orientation", "enum", True, "horizontal|vertical", False, 0
            schema.AddAttribute "rowSpan", "positive", False, "", False, 0
            schema.AddAttribute "columnSpan", "positive", False, "", False, 0
            schema.AllowVisualChildren
        Case Else
            Exit Function
    End Select
    Set private_CreateSchema = schema
End Function

Private Function private_TrackSchema(ByVal itemName As String) As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema
    Dim track As New obj_UiMarkupSchema

    If Not schema.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    If Not track.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    track.AddAttribute "size", "gridsize", True, "", False, 0
    schema.AddChild itemName, track, 1
    Set private_TrackSchema = schema
End Function

Private Function private_StylesSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema

    If Not schema.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    schema.AddChild "controlStyle", private_ControlStyleSchema()
    schema.AddChild "stylePipelineStage", private_StylePipelineStageSchema()
    Set private_StylesSchema = schema
End Function

Private Function private_ControlStyleSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema

    If Not schema.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    schema.AddAttribute "name", "string", True, "", False, 0
    schema.AddAttribute "backColor", "string", False, "", False, 0
    schema.AddAttribute "textColor", "string", False, "", False, 0
    schema.AddAttribute "fontColor", "string", False, "", False, 0
    schema.AddAttribute "fontName", "string", False, "", False, 0
    schema.AddAttribute "fontSize", "string", False, "", False, 0
    schema.AddAttribute "fontBold", "string", False, "", False, 0
    schema.AddAttribute "fontItalic", "string", False, "", False, 0
    schema.AddAttribute "borderColor", "string", False, "", False, 0
    schema.AddAttribute "borderWeight", "string", False, "", False, 0
    schema.AddAttribute "borderLineStyle", "string", False, "", False, 0
    schema.AddAttribute "horizontal", "string", False, "", False, 0
    schema.AddAttribute "vertical", "string", False, "", False, 0
    schema.AddAttribute "overflow", "string", False, "", False, 0
    schema.AddAttribute "width", "string", False, "", False, 0
    schema.AddAttribute "rowHeight", "string", False, "", False, 0
    schema.AddAttribute "columnWidth", "string", False, "", False, 0
    schema.AddAttribute "gridLines", "string", False, "", False, 0
    schema.AddAttribute "zoom", "string", False, "", False, 0
    Set private_ControlStyleSchema = schema
End Function

Private Function private_StylePipelineStageSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema

    If Not schema.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    schema.AddAttribute "name", "string", True, "", False, 0
    schema.AddAttribute "enabled", "boolean", False, "", False, 0
    schema.AddChild "layer", private_StyleLayerSchema()
    Set private_StylePipelineStageSchema = schema
End Function

Private Function private_StyleLayerSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema

    If Not schema.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    schema.AddAttribute "name", "string", True, "", False, 0
    schema.AddChild "rule", private_StyleRuleSchema()
    schema.AddAttribute "enabled", "boolean", False, "", False, 0
    Set private_StyleLayerSchema = schema
End Function

Private Function private_StyleRuleSchema() As obj_UiMarkupSchema
    Dim schema As New obj_UiMarkupSchema

    If Not schema.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
    schema.AddAttribute "target", "enum", True, "sheet|usedRange|row|column|cell|range|control|controlPart|layoutContainer|layoutBound|inlinePart", False, 0
    schema.AddAttribute "selector", "string", False, "", False, 0
    schema.AddAttribute "style", "string", False, "", False, 0
    schema.AddAttribute "styles", "styleblock", False, "", False, 0
    schema.AddAttribute "enabled", "boolean", False, "", False, 0
    schema.RequireAnyAttribute "style|styles"
    Set private_StyleRuleSchema = schema
End Function