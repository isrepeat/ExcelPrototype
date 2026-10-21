VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiFormControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_panel As obj_IUiElement
Private m_name As String
Private m_context As obj_UiRenderContext
Private m_row As Long
Private m_column As Long
Private m_definition As Object
Private m_reservedColumns As Long
Private m_fields As Object
Private m_source As String
Private m_retainedFieldsLayout As Boolean

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
Private Function obj_IUiControl_Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    obj_IUiControl_Initialize = True
    m_isInitialized = obj_IUiControl_Initialize
End Function

Private Sub obj_IUiControl_Dispose()
    Me.Dispose
End Sub

Private Function obj_IUiControl_Configure( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim node As Object
    Dim child As obj_IUiElement
    Dim container As obj_IUiContainer
    Dim panelNode As Object

    If Not ex_UiFormLayouts.fn_TryResolve(definition, context.BindingContext, definition, diagnostic) Then Exit Function

    If Not definition.getAttributeNode("rowSpan") Is Nothing Or _
            Not definition.getAttributeNode("labelColumnSpan") Is Nothing Then
        diagnostic = "Form does not support rowSpan or labelColumnSpan."
        Exit Function
    End If
    m_name = ex_UiElementFactory.fn_Attribute(definition, "name")
    If VBA.Len(m_name) = 0 Or VBA.Len(source) = 0 Then
        diagnostic = "Form requires name and dataContext (local or inherited)."
        Exit Function
    End If
    Set m_context = context
    m_source = source
    Set m_fields = CreateObject("Scripting.Dictionary")
    Set m_definition = definition.cloneNode(True)
    m_row = ex_UiElementFactory.fn_Long(definition, "row", 1)
    m_column = ex_UiElementFactory.fn_Long(definition, "column", 1)
    m_reservedColumns = ex_UiElementFactory.fn_Long(definition, "columnSpan", 0)
    If ex_UiElementFactory.fn_Attribute(definition, "placement") = "mapped" Then
        Set m_panel = New obj_UiGridElement
    Else
        Set m_panel = New obj_UiStackPanelElement
    End If
    Set container = m_panel
    Set panelNode = definition.cloneNode(False)
    panelNode.setAttribute "row", "1"
    panelNode.setAttribute "column", "1"
    If Not m_panel.Configure(panelNode, context, source, diagnostic) Then Exit Function
    For Each node In definition.ChildNodes
        If node.NodeType = 1 Then
            Select Case VBA.CStr(node.namespaceURI) & "|" & VBA.LCase$(VBA.CStr(node.baseName))
                Case "urn:excelprototype:profiles|field"
                    Set child = New obj_UiFormField
                    If Not child.Configure(node, context, source, diagnostic) Then
                        child.Dispose
                        Exit Function
                    End If
                    m_fields.Add ex_UiElementFactory.fn_Attribute(node, "name"), child
                Case Else
                    If VBA.CStr(node.namespaceURI) <> "urn:excelprototype:controls" Then
                        diagnostic = "Form only accepts field and controls namespace children: " & m_name
                        Exit Function
                    End If
                    Set child = ex_UiElementFactory.fn_Create(node, context, source, diagnostic)
                    If child Is Nothing Then Exit Function
            End Select
            container.AddChild child
        End If
    Next node
    context.RegisterForm m_name, Me
    obj_IUiControl_Configure = True
End Function

Private Function obj_IUiControl_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiControl_Measure = m_panel.Measure(rows, columns, diagnostic)
    If Not obj_IUiControl_Measure Then Exit Function
    If m_reservedColumns > 0 Then
        If columns > m_reservedColumns Then
            diagnostic = "Form exceeds its reserved columnSpan: " & m_name
            obj_IUiControl_Measure = False
            Exit Function
        End If
        columns = m_reservedColumns
    End If
    rows = rows + m_row - 1
    columns = columns + m_column - 1
End Function

Private Function obj_IUiControl_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim panel As obj_UiStackPanelElement
    Dim grid As obj_UiGridElement
    Dim field As obj_UiFormField
    Dim key As Variant

    obj_IUiControl_Arrange = m_panel.Arrange(row + m_row - 1, column + m_column - 1, diagnostic)
    If obj_IUiControl_Arrange And m_retainedFieldsLayout Then
        m_context.Styles.ForgetControl m_name
        If TypeOf m_panel Is obj_UiGridElement Then
            Set grid = m_panel
            grid.RegisterLayoutPart
        Else
            Set panel = m_panel
            panel.RegisterLayoutPart
        End If
        For Each key In m_fields.Keys
            Set field = m_fields(key)
            field.RefreshLayoutPart
        Next key
        m_retainedFieldsLayout = False
    End If
End Function

Private Function obj_IUiControl_Render(ByRef diagnostic As String) As Boolean
    obj_IUiControl_Render = m_panel.Render(diagnostic)
End Function

Private Function obj_IUiControl_Validate(ByVal errors As Collection) As Boolean
    obj_IUiControl_Validate = m_panel.Validate(errors)
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
    If Not m_context Is Nothing Then m_context.UnregisterForm m_name
    If Not m_panel Is Nothing Then m_panel.Dispose
    Set m_panel = Nothing
    Set m_context = Nothing
    Set m_definition = Nothing
    Set m_fields = Nothing
End Sub

Public Function TryRebuildFields( _
    ByVal definition As Object, _
    ByRef rebuilt As Boolean, _
    ByRef diagnostic As String _
) As Boolean
    Dim resolved As Object
    Dim oldNodes As Object
    Dim nextNodes As Object
    Dim nextFields As Object
    Dim children As Collection
    Dim node As Object
    Dim oldNode As Object
    Dim child As obj_IUiElement
    Dim panel As obj_UiStackPanelElement
    Dim key As Variant
    Dim name As String
    Dim container As obj_IUiContainer
    Dim oldContainer As Object
    Dim nextContainer As Object
    Dim retainedCount As Long
    Dim createdCount As Long
    Dim grid As obj_UiGridElement

    rebuilt = False
    If Not ex_UiFormLayouts.fn_TryResolve(definition, m_context.BindingContext, resolved, diagnostic) Then Exit Function
    TryRebuildFields = True
    ' Only field-only forms with identical container attributes can retain their children.
    Set oldContainer = m_definition.cloneNode(False)
    Set nextContainer = resolved.cloneNode(False)
    If nextContainer.XML <> oldContainer.XML Then Exit Function
    Set oldNodes = CreateObject("Scripting.Dictionary")
    Set nextNodes = CreateObject("Scripting.Dictionary")
    For Each node In m_definition.ChildNodes
        If node.NodeType = 1 Then
            If node.namespaceURI <> "urn:excelprototype:profiles" Or node.baseName <> "field" Then Exit Function
            oldNodes.Add ex_UiElementFactory.fn_Attribute(node, "name"), node
        End If
    Next node
    For Each node In resolved.ChildNodes
        If node.NodeType = 1 Then
            If node.namespaceURI <> "urn:excelprototype:profiles" Or node.baseName <> "field" Then Exit Function
            nextNodes.Add ex_UiElementFactory.fn_Attribute(node, "name"), node
        End If
    Next node
    For Each key In m_fields.Keys
        Set oldNode = oldNodes(key)
        If nextNodes.Exists(key) Then
            Set node = nextNodes(key)
            If oldNode.XML = node.XML Then GoTo RetainField
        End If
        Set child = m_fields(key)
        child.Dispose
RetainField:
    Next key
    Set nextFields = CreateObject("Scripting.Dictionary")
    Set children = New Collection
    Set container = m_panel
    For Each key In nextNodes.Keys
        Set node = nextNodes(key)
        name = CStr(key)
        Set child = Nothing
        If oldNodes.Exists(key) Then
            Set oldNode = oldNodes(key)
            If oldNode.XML = node.XML Then
                Set child = m_fields(key)
                retainedCount = retainedCount + 1
            End If
        End If
        If child Is Nothing Then
            createdCount = createdCount + 1
            Set child = New obj_UiFormField
            ' Keep partially configured replacements owned until the collection is committed.
            container.AddChild child
            If Not child.Configure(node, m_context, m_source, diagnostic) Then
                child.Dispose
                TryRebuildFields = False
                Exit Function
            End If
        End If
        children.Add child
        nextFields.Add name, child
    Next key
    If TypeOf m_panel Is obj_UiGridElement Then
        Set grid = m_panel
        grid.ReplaceChildren children
    Else
        Set panel = m_panel
        panel.ReplaceChildren children
    End If
    Set m_fields = nextFields
    Set m_definition = resolved.cloneNode(True)
    m_retainedFieldsLayout = True
    ex_Core.fn_Diagnostic_WriteLog "TEMP.Form.Reconcile | Name=" & m_name & " | Retained=" & CStr(retainedCount) & " | Created=" & CStr(createdCount)
    rebuilt = True
End Function

Public Function GetDefinition() As Object
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If m_definition Is Nothing Then Exit Function
    Set GetDefinition = m_definition.cloneNode(True)
End Function