VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiControlElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const ENABLE_LOGGING_VERBOSE = False

Implements obj_IUiElement
Implements obj_IUiLayoutSlot
Implements obj_IUiBindingListener

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_definition As Object
Private m_intrinsicRows As String
Private m_intrinsicColumns As String
Private m_context As obj_UiRenderContext
Private m_bindingContext As obj_UiBindingContext
Private m_control As obj_IUiControl
Private m_source As String
Private m_rows As Long
Private m_columns As Long
Private m_area As Range
Private m_localRow As Long
Private m_localColumn As Long
Private m_readOnlyDirty As Boolean
Private m_visualDirty As Boolean
Private m_textDirty As Boolean
Private m_candidateValuesDirty As Boolean
Private m_candidateLayoutDirty As Boolean
Private m_bindingDependencies As Collection
Private m_collapsed As Boolean
Private m_visibilityMeasured As Boolean
Private m_closeControl As obj_IUiControl
Private m_closeName As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get RenderArea() As Range
    Set RenderArea = m_area
End Property

' //
' // Interface
' //
Private Function obj_IUiElement_Configure( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiElement_Configure = Me.Configure(definition, context, source, diagnostic)
End Function

Private Function obj_IUiElement_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim collapsed As Boolean

    If Not private_TryVisibility(collapsed, diagnostic) Then Exit Function
    If collapsed And m_visibilityMeasured And Not m_collapsed Then
        ' Dispose active routes and shape actions before the occupied area is reused.
        m_control.Dispose
        private_DisposeClose
        private_RestoreLocalPosition
        Set m_control = ex_UiControlFactory.fn_Create(m_definition)
        If m_control Is Nothing Then
            diagnostic = "Cannot recreate collapsed control."
            Exit Function
        End If
        If Not m_control.Configure(m_definition, m_context, m_source, diagnostic) Then Exit Function
    End If
    m_collapsed = collapsed
    m_visibilityMeasured = True
    If m_collapsed Then
        rows = 0
        columns = 0
        m_rows = 0
        m_columns = 0
        Set m_area = Nothing
        obj_IUiElement_Measure = True
        Exit Function
    End If
    If VBA.Len(m_intrinsicRows) = 0 Then
        m_definition.removeAttribute "rowSpan"
    Else
        m_definition.setAttribute "rowSpan", m_intrinsicRows
    End If
    If VBA.Len(m_intrinsicColumns) = 0 Then
        m_definition.removeAttribute "columnSpan"
    Else
        m_definition.setAttribute "columnSpan", m_intrinsicColumns
    End If
    obj_IUiElement_Measure = m_control.Measure(rows, columns, diagnostic)
    m_rows = rows
    m_columns = columns
End Function

Private Function obj_IUiElement_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim localRow As Long
    Dim localColumn As Long
    Dim candidates As obj_UiCandidatesControl

    If m_collapsed Then
        Set m_area = Nothing
        obj_IUiElement_Arrange = True
        Exit Function
    End If
    obj_IUiElement_Arrange = m_control.Arrange(row, column, diagnostic)
    Set m_area = Nothing
    If obj_IUiElement_Arrange And m_rows > 0 And m_columns > 0 Then
        localRow = m_localRow
        localColumn = m_localColumn
        If TypeOf m_control Is obj_UiCandidatesControl Then
            Set candidates = m_control
            localColumn = candidates.TableStartColumn
        End If
        Set m_area = m_context.TargetWorksheet.Cells(row + localRow - 1, column + localColumn - 1).Resize( _
            m_rows - localRow + 1, m_columns - localColumn + 1)
        If Not private_ArrangeClose(diagnostic) Then
            obj_IUiElement_Arrange = False
            Exit Function
        End If
    End If
End Function

Private Function obj_IUiElement_Render(ByRef diagnostic As String) As Boolean
    If m_collapsed Then
        obj_IUiElement_Render = True
        Exit Function
    End If
    obj_IUiElement_Render = m_control.Render(diagnostic)
    If obj_IUiElement_Render And Not m_closeControl Is Nothing Then
        m_context.Styles.ForgetControl m_closeName
        obj_IUiElement_Render = m_closeControl.Render(diagnostic)
    End If
    If obj_IUiElement_Render Then
        m_candidateValuesDirty = False
        m_candidateLayoutDirty = False
    End If
    If Not obj_IUiElement_Render Then
        diagnostic = "Cannot render control: " & ex_UiElementFactory.fn_Attribute(m_definition, "name")
    End If
End Function

Private Function obj_IUiElement_Validate(ByVal errors As Collection) As Boolean
    obj_IUiElement_Validate = m_control.Validate(errors)
End Function

Private Sub obj_IUiElement_Dispose()
    Me.Dispose
End Sub

Private Sub obj_IUiLayoutSlot_SetSize( _
    ByVal rows As Long, _
    ByVal columns As Long _
)
    Select Case VBA.LCase$(ex_UiElementFactory.fn_Attribute(m_definition, "type"))
        Case "label", "button", "checkbox", "input", "select"
            m_definition.setAttribute "rowSpan", VBA.CStr(rows)
            m_definition.setAttribute "columnSpan", VBA.CStr(columns)
            m_rows = rows + m_localRow - 1
            m_columns = columns + m_localColumn - 1
    End Select
End Sub

' //
' // API
' //
Public Function Configure( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim attributeNode As Object
    Dim bindingSource As String
    Dim bindingPath As String
    Dim controlName As String

    If m_isDisposed Or m_isInitialized Then
        diagnostic = "Element is already initialized or disposed."
        Exit Function
    End If
    controlName = ex_UiElementFactory.fn_Attribute(definition, "name")
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CLONE_STARTED | Name=" & controlName
#End If
    Set m_definition = definition.cloneNode(True)
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CLONE_COMPLETED | Name=" & controlName
#End If
    m_intrinsicRows = ex_UiElementFactory.fn_Attribute(definition, "rowSpan")
    m_intrinsicColumns = ex_UiElementFactory.fn_Attribute(definition, "columnSpan")
    If VBA.CStr(definition.namespaceURI) = "urn:excelprototype:controls" Then _
        m_definition.setAttribute "type", VBA.CStr(definition.baseName)
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CONTEXT_ATTACH_STARTED | Name=" & controlName
#End If
    Set m_context = context
    m_source = source
    m_localRow = ex_UiElementFactory.fn_Long(definition, "row", 1)
    m_localColumn = ex_UiElementFactory.fn_Long(definition, "column", 1)
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CONTEXT_ATTACH_COMPLETED | Name=" & controlName
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_EVENTS_ATTACH_STARTED | Name=" & controlName
#End If
    Set m_bindingContext = context.BindingContext
    m_bindingContext.Subscribe Me
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_EVENTS_ATTACH_COMPLETED | Name=" & controlName
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_BINDINGS_PARSE_STARTED | Name=" & controlName
#End If
    Set m_bindingDependencies = New Collection
    For Each attributeNode In m_definition.Attributes
        If attributeNode.nodeName <> "dataContext" And VBA.StrComp(VBA.Left$(VBA.Trim$(VBA.CStr(attributeNode.Text)), 9), _
                "{Binding ", VBA.vbTextCompare) = 0 Then
            If Not ex_UiBindingRuntime.fn_TryParseBinding(VBA.CStr(attributeNode.Text), _
                    source, bindingSource, bindingPath, context.BindingContext) Then
                diagnostic = "Invalid binding on control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
                Exit Function
            End If
            attributeNode.Text = "{Binding Source=" & bindingSource & "; Path=" & bindingPath & "}"
            If attributeNode.nodeName <> "selectedItem" Then
                m_bindingDependencies.Add VBA.Array(VBA.CStr(attributeNode.nodeName), bindingSource, bindingPath)
            End If
        End If
    Next attributeNode
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_BINDINGS_PARSE_COMPLETED | Name=" & controlName
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_FACTORY_CREATE_STARTED | Name=" & controlName
#End If
    Set m_control = ex_UiControlFactory.fn_Create(m_definition)
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_FACTORY_CREATE_COMPLETED | Name=" & controlName
#End If
    If m_control Is Nothing Then
        diagnostic = "Cannot create control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
        Exit Function
    End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CONFIGURE_STARTED | Name=" & controlName
#End If
    If Not m_control.Configure(m_definition, context, source, diagnostic) Then
        If VBA.Len(diagnostic) = 0 Then diagnostic = "Cannot configure control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
        Exit Function
    End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_CONFIGURE_COMPLETED | Name=" & controlName
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_REGISTER_STARTED | Name=" & controlName
#End If
    context.RegisterElement ex_UiElementFactory.fn_Attribute(definition, "name"), Me
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_REGISTER_COMPLETED | Name=" & controlName
#End If
    Configure = True
    m_isInitialized = Configure
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    private_DisposeClose
    If Not m_control Is Nothing Then m_control.Dispose
    If Not m_context Is Nothing Then m_context.UnregisterElement ex_UiElementFactory.fn_Attribute(m_definition, "name"), Me
    Set m_control = Nothing
    If Not m_bindingContext Is Nothing Then m_bindingContext.Unsubscribe Me
    Set m_bindingContext = Nothing
    Set m_context = Nothing
    Set m_definition = Nothing
    Set m_bindingDependencies = Nothing
    Set m_area = Nothing
End Sub

Public Function RefreshVisual(ByRef diagnostic As String) As Boolean
    Dim target As obj_IUiBindingTarget
    Dim fieldControl As obj_UiFieldControl
    Dim labelControl As obj_UiLabelControl

    If m_collapsed Then
        RefreshVisual = True
        Exit Function
    End If
    If m_readOnlyDirty And Not m_visualDirty And TypeOf m_control Is obj_UiFieldControl Then
        Set fieldControl = m_control
        RefreshVisual = fieldControl.RefreshReadOnly()
    ElseIf m_textDirty And Not m_visualDirty And TypeOf m_control Is obj_UiLabelControl Then
        Set labelControl = m_control
        RefreshVisual = labelControl.RefreshText()
    ElseIf TypeOf m_control Is obj_IUiBindingTarget Then
        Set target = m_control
        RefreshVisual = target.RefreshBindings(m_context)
    Else
        RefreshVisual = m_control.Render(diagnostic)
    End If
    If Not RefreshVisual And VBA.Len(diagnostic) = 0 Then diagnostic = "Cannot refresh control: " & ex_UiElementFactory.fn_Attribute(m_definition, "name")
    If RefreshVisual Then
        m_readOnlyDirty = False
        m_visualDirty = False
        m_textDirty = False
    End If
End Function

Public Function DependsOnLayout(ByVal name As String) As Boolean
    DependsOnLayout = (VBA.StrComp(ex_UiElementFactory.fn_Attribute(m_definition, "alignToForm"), name, VBA.vbTextCompare) = 0)
End Function

Public Function TryRefreshValues( _
    ByVal pending As Object, _
    ByRef refreshed As Boolean, _
    ByRef diagnostic As String _
) As Boolean
    Dim candidates As obj_UiCandidatesControl
    Dim table As obj_UiTableControl
    Dim dependency As String

    refreshed = False
    TryRefreshValues = True
    If m_collapsed Then Exit Function
    If TypeOf m_control Is obj_UiTableControl Then
        dependency = ex_UiElementFactory.fn_Attribute(m_definition, "name")
        If Not pending.Exists(dependency) Then
            Exit Function
        End If
        If pending(dependency) Then
            Exit Function
        End If
        Set table = m_control
        TryRefreshValues = table.TryRefreshTemplateBindings(refreshed, diagnostic)
        Exit Function
    End If
    If Not TypeOf m_control Is obj_UiCandidatesControl Then Exit Function
    If Not m_candidateValuesDirty Or m_candidateLayoutDirty Then Exit Function
    dependency = ex_UiElementFactory.fn_Attribute(m_definition, "alignToForm")
    If pending.Exists(dependency) Then
        If pending(dependency) Then Exit Function
    End If
    Set candidates = m_control
    TryRefreshValues = candidates.TryRefreshValues(refreshed, diagnostic)
    If refreshed Then m_candidateValuesDirty = False
End Function

Public Function RebuildStructure(ByRef diagnostic As String) As Boolean
    If VBA.LCase$(ex_UiElementFactory.fn_Attribute(m_definition, "type")) <> "form" Then
        RebuildStructure = True
        Exit Function
    End If
    If Not m_control Is Nothing Then m_control.Dispose
    private_RestoreLocalPosition
    Set m_control = ex_UiControlFactory.fn_Create(m_definition)
    If m_control Is Nothing Then
        diagnostic = "Cannot recreate control: " & ex_UiElementFactory.fn_Attribute(m_definition, "name")
        Exit Function
    End If
    RebuildStructure = m_control.Configure(m_definition, m_context, m_source, diagnostic)
End Function

Public Function TryGetItemArea( _
    ByVal index As Long, _
    ByRef area As Range _
) As Boolean
    Dim items As obj_UiItemsControl

    Set area = Nothing
    If Not TypeOf m_control Is obj_UiItemsControl Then Exit Function
    Set items = m_control
    TryGetItemArea = items.TryGetItemArea(index, area)
End Function

Public Function TryRefreshTableRow( _
    ByVal rowIndex As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim table As obj_UiTableControl
    Dim refreshed As Boolean

    If m_isDisposed Or Not TypeOf m_control Is obj_UiTableControl Then
        diagnostic = "The control is not an active table."
        Exit Function
    End If
    Set table = m_control
    If Not table.TryRefreshTemplateBindings(refreshed, diagnostic, rowIndex) Then Exit Function
    If Not refreshed Then
        diagnostic = "The table requires a full refresh before updating a row."
        Exit Function
    End If
    TryRefreshTableRow = True
End Function

' //
' // Private
' //
Private Sub private_RestoreLocalPosition()
    ' Arrange writes sheet coordinates into the node; ConfigurePosition expects local offsets.
    m_definition.setAttribute "row", VBA.CStr(m_localRow)
    m_definition.setAttribute "column", VBA.CStr(m_localColumn)
End Sub

Private Function private_ArrangeClose(ByRef diagnostic As String) As Boolean
    Dim definition As Object
    Dim rows As Long
    Dim columns As Long
    Dim command As String
    Dim caption As String

    private_ArrangeClose = True
    command = ex_UiElementFactory.fn_Attribute(m_definition, "closeCommand")
    If VBA.Len(command) = 0 Then Exit Function
    private_ArrangeClose = False
    If m_area.Column <= 1 Then
        diagnostic = "A close button requires a free column to the left of the table."
        Exit Function
    End If
    If m_closeControl Is Nothing Then
        caption = ex_UiElementFactory.fn_Attribute(m_definition, "closeCaption")
        If VBA.Len(caption) = 0 Then
            diagnostic = "A close button requires closeCaption."
            Exit Function
        End If
        m_closeName = "Close_" & ex_UiElementFactory.fn_Attribute(m_definition, "name")
        If VBA.Len("btn_" & m_closeName) > 31 Then
            diagnostic = "Close button name exceeds the Excel name limit: " & m_closeName
            Exit Function
        End If
        Set definition = m_definition.ownerDocument.createNode(1, "controls:button", "urn:excelprototype:controls")
        definition.setAttribute "name", m_closeName
        definition.setAttribute "type", "button"
        definition.setAttribute "row", "1"
        definition.setAttribute "column", "1"
        definition.setAttribute "rowSpan", "1"
        definition.setAttribute "columnSpan", "1"
        definition.setAttribute "caption", caption
        definition.setAttribute "command", command
        definition.setAttribute "style", ex_UiElementFactory.fn_Attribute(m_definition, "style")
        Set m_closeControl = ex_UiControlFactory.fn_Create(definition)
        If m_closeControl Is Nothing Then
            diagnostic = "Cannot create close button: " & m_closeName
            Exit Function
        End If
        If Not m_closeControl.Configure(definition, m_context, m_source, diagnostic) Then Exit Function
    End If
    If Not m_closeControl.Measure(rows, columns, diagnostic) Then Exit Function
    If Not m_closeControl.Arrange(m_area.Row, m_area.Column - 1, diagnostic) Then Exit Function
    Set m_area = m_area.Offset(0, -1).Resize(m_area.Rows.Count, m_area.Columns.Count + 1)
    private_ArrangeClose = True
End Function

Private Sub private_DisposeClose()
    Dim shape As Shape

    If Not m_closeControl Is Nothing Then m_closeControl.Dispose
    Set m_closeControl = Nothing
    If m_context Is Nothing Or VBA.Len(m_closeName) = 0 Then Exit Sub
    ' Remove style regions before their shape becomes an invalid Excel object.
    m_context.Styles.ForgetControl m_closeName
    m_context.Router.UnregisterShape "btn_" & m_closeName
    m_context.UnregisterShapeLayout "btn_" & m_closeName
    For Each shape In m_context.TargetWorksheet.Shapes
        If shape.Name = "btn_" & m_closeName Then
            shape.Delete
            Exit For
        End If
    Next shape
    m_closeName = VBA.vbNullString
End Sub

Private Function private_TryVisibility( _
    ByRef collapsed As Boolean, _
    ByRef diagnostic As String _
) As Boolean
    Dim expression As String
    Dim value As String

    expression = ex_UiElementFactory.fn_Attribute(m_definition, "visibility")
    collapsed = False
    If VBA.Len(expression) = 0 Then
        private_TryVisibility = True
        Exit Function
    End If
    If Not ex_UiBindingRuntime.fn_TryResolveText(expression, m_bindingContext, value) Then
        diagnostic = "Cannot resolve control visibility."
        Exit Function
    End If
    Select Case value
        Case "Visible"
            collapsed = False
        Case "Collapsed"
            collapsed = True
        Case Else
            diagnostic = "Invalid control visibility: " & value
            Exit Function
    End Select
    private_TryVisibility = True
End Function

Private Sub obj_IUiBindingListener_OnValueChanged( _
    ByVal sourceName As String, _
    ByVal bindingPath As String _
)
    private_bindingContext_ValueChanged sourceName, bindingPath
End Sub

Private Sub private_bindingContext_ValueChanged(ByVal sourceName As String, ByVal bindingPath As String)
    Dim dependency As Variant
    Dim attributeName As String
    Dim source As String
    Dim path As String

    If m_isDisposed Or m_control Is Nothing Then Exit Sub
    If m_bindingDependencies Is Nothing Then Exit Sub
    For Each dependency In m_bindingDependencies
        attributeName = dependency(0)
        source = dependency(1)
        path = dependency(2)
        If VBA.StrComp(source, sourceName, VBA.vbTextCompare) = 0 Then
            If VBA.StrComp(path, bindingPath, VBA.vbTextCompare) = 0 Or _
                    VBA.StrComp(VBA.Left$(path, VBA.Len(bindingPath) + 1), bindingPath & ".", VBA.vbTextCompare) = 0 Then
                If attributeName = "visibility" Then
                    m_candidateLayoutDirty = True
                    m_context.MarkDirty ex_UiElementFactory.fn_Attribute(m_definition, "name"), True
                    GoTo NextAttribute
                End If
                If TypeOf m_control Is obj_UiCandidatesControl Then
                    If attributeName = "itemsSource" Then
                        m_candidateValuesDirty = True
                    Else
                        m_candidateLayoutDirty = True
                    End If
                End If
                If TypeOf m_control Is obj_UiFieldControl Then
                    ' Cell bindings already propagate values without refreshing control presentation.
                    If attributeName = "value" Then GoTo NextAttribute
                    If attributeName = "readOnly" Then
                        m_readOnlyDirty = True
                    Else
                        m_visualDirty = True
                    End If
                ElseIf TypeOf m_control Is obj_UiLabelControl Then
                    If attributeName = "text" Or attributeName = "caption" Then
                        m_textDirty = True
                    Else
                        m_visualDirty = True
                    End If
                Else
                    m_visualDirty = True
                End If
                If TypeOf m_control Is obj_IUiBindingTarget And _
                        attributeName <> "itemsSource" And _
                        attributeName <> "rowSpan" And _
                        attributeName <> "columnSpan" Then
                    m_context.MarkDirty ex_UiElementFactory.fn_Attribute(m_definition, "name"), False
                ElseIf attributeName <> "value" Then
                    m_context.MarkDirty ex_UiElementFactory.fn_Attribute(m_definition, "name"), True
                End If
            End If
        End If
NextAttribute:
    Next dependency
    Exit Sub
End Sub