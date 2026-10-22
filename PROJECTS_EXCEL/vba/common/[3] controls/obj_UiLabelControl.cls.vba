VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiLabelControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const ENABLE_LOGGING_VERBOSE = False

Implements obj_IUiControl
Implements obj_IUiBindingTarget
Implements obj_IUiEventHandler

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_renderContext As obj_UiRenderContext
Private m_uiControlBase As obj_UiControlBase
Private m_targetRange As Range
Private m_registeredRange As Range

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_LABEL_BASE_NEW_STARTED"
#End If
    Set m_uiControlBase = New obj_UiControlBase
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_LABEL_BASE_NEW_COMPLETED"
#End If
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
    If m_uiControlBase Is Nothing Then Set m_uiControlBase = New obj_UiControlBase
    obj_IUiControl_Initialize = m_uiControlBase.Initialize()
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
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_LABEL_CONTEXT_ATTACH_STARTED"
#End If
    Set m_renderContext = context
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_LABEL_CONTEXT_ATTACH_COMPLETED"
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_LABEL_POSITION_CONFIGURE_STARTED"
#End If
    m_uiControlBase.ConfigurePosition definition
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_LABEL_POSITION_CONFIGURE_COMPLETED"
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_LABEL_CONFIGURE_STARTED"
#End If
    obj_IUiControl_Configure = private_Configure(definition)
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_LABEL_CONFIGURE_COMPLETED | Success=" & VBA.CStr(obj_IUiControl_Configure)
#End If
    If Not obj_IUiControl_Configure Then diagnostic = "Cannot configure control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
End Function

Private Function obj_IUiControl_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim target As Range

    m_uiControlBase.SetPosition 1, 1
    If Not private_Configure(m_uiControlBase.ControlNode) Then Exit Function
    Set target = private_Measure(m_renderContext)
    If target Is Nothing Then
        diagnostic = "Cannot measure control: " & m_uiControlBase.ControlName
        Exit Function
    End If
    m_uiControlBase.GetSize target, rows, columns
    obj_IUiControl_Measure = True
End Function

Private Function obj_IUiControl_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim target As Range

    m_uiControlBase.ArrangePosition row, column
    If Not private_Configure(m_uiControlBase.ControlNode) Then Exit Function
    Set target = private_Measure(m_renderContext)
    obj_IUiControl_Arrange = Not target Is Nothing
    If target Is Nothing Then diagnostic = "Cannot arrange control: " & m_uiControlBase.ControlName
End Function

Private Function obj_IUiControl_Render(ByRef diagnostic As String) As Boolean
    obj_IUiControl_Render = private_Render(m_renderContext)
    If Not obj_IUiControl_Render Then diagnostic = "Cannot render control: " & m_uiControlBase.ControlName
End Function

Private Function obj_IUiControl_Validate(ByVal errors As Collection) As Boolean
    obj_IUiControl_Validate = True
End Function

Private Function obj_IUiBindingTarget_RefreshBindings( _
    ByVal context As obj_UiRenderContext _
) As Boolean
    obj_IUiBindingTarget_RefreshBindings = private_Render(context)
End Function

Private Function obj_IUiEventHandler_HandleEvent( _
    ByVal kind As String, _
    ByVal payload As Variant _
) As Boolean
    Dim previousEvents As Boolean
    Dim caption As String

    If Not m_isInitialized Or m_isDisposed Or kind <> "change" Then Exit Function
    If m_registeredRange Is Nothing Then Exit Function
    If Not m_uiControlBase.TryGetCaption(m_renderContext.BindingContext, caption) Then Exit Function
    previousEvents = Application.EnableEvents
    On Error GoTo EH_HEADER
    Application.EnableEvents = False
    m_registeredRange.Cells(1, 1).Value2 = caption
    Application.EnableEvents = previousEvents
    obj_IUiEventHandler_HandleEvent = True
    Exit Function
EH_HEADER:
    Application.EnableEvents = previousEvents
    Err.Raise Err.Number, "Label.HandleEvent", Err.Description
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
    private_UnregisterHeader
    If Not m_uiControlBase Is Nothing Then m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
    Set m_renderContext = Nothing
    Set m_targetRange = Nothing
End Sub

Public Function RefreshText() As Boolean
    Dim caption As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If m_targetRange Is Nothing Then Exit Function
    If Not m_uiControlBase.TryGetCaption(m_renderContext.BindingContext, caption) Then Exit Function
    m_targetRange.Cells(1, 1).Value2 = caption
    RefreshText = True
End Function

' //
' // Private
' //
Private Sub private_UnregisterHeader()
    Dim cell As Range

    If m_registeredRange Is Nothing Then Exit Sub
    For Each cell In m_registeredRange.Cells
        m_renderContext.Router.UnregisterCell cell
    Next cell
    Set m_registeredRange = Nothing
End Sub

Private Function private_Configure(ByVal controlNode As Object) As Boolean
    private_Configure = m_uiControlBase.Configure(controlNode)
End Function

Private Function private_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim startedAt As Double
    Dim targetRange As Range
    Dim cell As Range

    startedAt = VBA.Timer
    Dim captionText As String

    Set targetRange = m_targetRange
    If targetRange Is Nothing Then Set targetRange = private_Measure(uiRenderContext)
    If Not m_uiControlBase.TryGetCaption(uiRenderContext.BindingContext, captionText) Then Exit Function
    If targetRange.Cells.CountLarge > 1 Then targetRange.Merge
    targetRange.Value2 = captionText
    uiRenderContext.Styles.ApplyControlStyle targetRange, Nothing, m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    If ex_UiElementFactory.fn_Attribute(m_uiControlBase.ControlNode, "readOnly") = "true" Then
        private_UnregisterHeader
        Set m_registeredRange = targetRange
        For Each cell In m_registeredRange.Cells
            uiRenderContext.Router.RegisterCell cell, Me
        Next cell
    End If
    private_Render = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Label.Render", startedAt
End Function

Private Function private_Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    If m_uiControlBase Is Nothing Then Exit Function
    Set m_targetRange = m_uiControlBase.Measure(uiRenderContext)
    Set private_Measure = m_targetRange
End Function