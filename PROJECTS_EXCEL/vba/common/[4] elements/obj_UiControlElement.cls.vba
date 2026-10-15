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
    For Each attributeNode In m_definition.Attributes
        If attributeNode.nodeName <> "dataContext" And VBA.StrComp(VBA.Left$(VBA.Trim$(VBA.CStr(attributeNode.Text)), 9), _
                "{Binding ", VBA.vbTextCompare) = 0 Then
            If Not ex_UiBindingRuntime.fn_TryParseBinding(VBA.CStr(attributeNode.Text), _
                    source, bindingSource, bindingPath, context.BindingContext) Then
                diagnostic = "Invalid binding on control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
                Exit Function
            End If
            attributeNode.Text = "{Binding Source=" & bindingSource & "; Path=" & bindingPath & "}"
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
    obj_IUiElement_Configure = True
    m_isInitialized = obj_IUiElement_Configure
End Function

Private Function obj_IUiElement_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
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

    obj_IUiElement_Arrange = m_control.Arrange(row, column, diagnostic)
    If obj_IUiElement_Arrange And m_rows > 0 And m_columns > 0 Then
        localRow = m_localRow
        localColumn = m_localColumn
        If TypeOf m_control Is obj_UiCandidatesControl Then
            Set candidates = m_control
            localColumn = candidates.TableStartColumn
        End If
        Set m_area = m_context.TargetWorksheet.Cells(row + localRow - 1, column + localColumn - 1).Resize( _
            m_rows - localRow + 1, m_columns - localColumn + 1)
    End If
End Function

Private Function obj_IUiElement_Render(ByRef diagnostic As String) As Boolean
    obj_IUiElement_Render = m_control.Render(diagnostic)
    If Not obj_IUiElement_Render Then diagnostic = "Cannot render control: " & ex_UiElementFactory.fn_Attribute(m_definition, "name")
End Function

Private Function obj_IUiElement_Validate(ByVal errors As Collection) As Boolean
    obj_IUiElement_Validate = m_control.Validate(errors)
End Function

Private Sub obj_IUiElement_Dispose()
    Me.Dispose
End Sub

Private Sub obj_IUiLayoutSlot_SetSize(ByVal rows As Long, ByVal columns As Long)
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
Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_control Is Nothing Then m_control.Dispose
    If Not m_context Is Nothing Then m_context.UnregisterElement ex_UiElementFactory.fn_Attribute(m_definition, "name"), Me
    Set m_control = Nothing
    If Not m_bindingContext Is Nothing Then m_bindingContext.Unsubscribe Me
    Set m_bindingContext = Nothing
    Set m_context = Nothing
    Set m_definition = Nothing
    Set m_area = Nothing
End Sub

Public Function RefreshVisual(ByRef diagnostic As String) As Boolean
    Dim target As obj_IUiBindingTarget
    Dim fieldControl As obj_UiFieldControl
    Dim labelControl As obj_UiLabelControl

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

Public Function RebuildStructure(ByRef diagnostic As String) As Boolean
    If VBA.LCase$(ex_UiElementFactory.fn_Attribute(m_definition, "type")) <> "form" Then
        RebuildStructure = True
        Exit Function
    End If
    If Not m_control Is Nothing Then m_control.Dispose
    Set m_control = ex_UiControlFactory.fn_Create(m_definition)
    If m_control Is Nothing Then
        diagnostic = "Cannot recreate control: " & ex_UiElementFactory.fn_Attribute(m_definition, "name")
        Exit Function
    End If
    RebuildStructure = m_control.Configure(m_definition, m_context, m_source, diagnostic)
End Function

' //
' // Private
' //
Private Sub obj_IUiBindingListener_OnValueChanged( _
    ByVal sourceName As String, _
    ByVal bindingPath As String _
)
    private_bindingContext_ValueChanged sourceName, bindingPath
End Sub

Private Sub private_bindingContext_ValueChanged(ByVal sourceName As String, ByVal bindingPath As String)
    Dim nodeAttribute As Object
    Dim source As String
    Dim path As String

    If m_isDisposed Or m_control Is Nothing Then Exit Sub
    For Each nodeAttribute In m_definition.Attributes
        If nodeAttribute.nodeName <> "selectedItem" Then
            If ex_UiBindingRuntime.fn_TryParseBinding(VBA.CStr(nodeAttribute.Text), _
                    VBA.vbNullString, source, path) Then
                If VBA.StrComp(source, sourceName, VBA.vbTextCompare) = 0 Then
                    If VBA.StrComp(path, bindingPath, VBA.vbTextCompare) = 0 Or _
                            VBA.StrComp(VBA.Left$(path, VBA.Len(bindingPath) + 1), bindingPath & ".", VBA.vbTextCompare) = 0 Then
                        If TypeOf m_control Is obj_UiFieldControl Then
                            ' Cell bindings already propagate values without refreshing control presentation.
                            If nodeAttribute.nodeName = "value" Then GoTo NextAttribute
                            If nodeAttribute.nodeName = "readOnly" Then
                                m_readOnlyDirty = True
                            Else
                                m_visualDirty = True
                            End If
                        ElseIf TypeOf m_control Is obj_UiLabelControl Then
                            If nodeAttribute.nodeName = "text" Or nodeAttribute.nodeName = "caption" Then
                                m_textDirty = True
                            Else
                                m_visualDirty = True
                            End If
                        Else
                            m_visualDirty = True
                        End If
                        If TypeOf m_control Is obj_IUiBindingTarget And _
                                nodeAttribute.nodeName <> "itemsSource" And _
                                nodeAttribute.nodeName <> "rowSpan" And _
                                nodeAttribute.nodeName <> "columnSpan" Then
                            m_context.MarkDirty ex_UiElementFactory.fn_Attribute(m_definition, "name"), False
                        ElseIf nodeAttribute.nodeName <> "value" Then
                            m_context.MarkDirty ex_UiElementFactory.fn_Attribute(m_definition, "name"), True
                        End If
                    End If
                End If
            End If
        End If
NextAttribute:
    Next nodeAttribute
    Exit Sub
End Sub