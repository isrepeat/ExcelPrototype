VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_TDB_BasisReorder"
Option Explicit

Implements obj_IUiEventHandler

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private WithEvents m_application As Excel.Application
Private m_pageBase As obj_PageBase
Private m_context As obj_UiRenderContext
Private m_area As Range
Private m_up As Shape
Private m_down As Shape
Private m_identifier As String

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
Private Function obj_IUiEventHandler_HandleEvent( _
    ByVal kind As String, _
    ByVal payload As Variant _
) As Boolean
    Dim direction As Long

    If Not m_isInitialized Or m_isDisposed Or kind <> "click" Then Exit Function
    If VBA.CStr(payload) = "btn_TdbBasisUp" Then direction = -1
    If VBA.CStr(payload) = "btn_TdbBasisDown" Then direction = 1
    If direction = 0 Then Exit Function
    obj_IUiEventHandler_HandleEvent = private_Move(direction)
End Function

' //
' // API
' //
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    If m_isInitialized Or m_isDisposed Or pageBase Is Nothing Then Exit Function
    Set m_pageBase = pageBase
    Set m_application = Application
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    private_Hide
    Set m_application = Nothing
    Set m_context = Nothing
    Set m_pageBase = Nothing
    m_isInitialized = False
    m_isDisposed = True
End Sub

Public Sub ClearFocus()
    private_Hide
End Sub

Public Function FocusCell(ByVal target As Object) As Boolean
    Dim tableArea As Range
    Dim selectedArea As Range
    Dim chain As Collection
    Dim index As Long

    On Error GoTo Failed
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    private_Hide
    FocusCell = True
    If target Is Nothing Then Exit Function
    If Not TypeOf target Is Range Then Exit Function
    If target.Areas.Count <> 1 Or target.Rows.Count <> 1 Then Exit Function
    If Not ex_UiRuntime.fn_TryGetContext(target.Worksheet, m_context) Then Exit Function
    If Not m_context.BindingContext Is m_pageBase.BindingContext Then Exit Function
    If Not m_context.TryGetElementArea("BasisList", tableArea) Then Exit Function
    Set selectedArea = target.Cells(1, 1).MergeArea
    If Application.Intersect(selectedArea, tableArea) Is Nothing Then Exit Function
    If selectedArea.Column <> tableArea.Column Then Exit Function
    Set chain = ex_TDB_BasisChain.fn_Read(m_pageBase.BindingContext)
    index = selectedArea.Row - tableArea.Row
    If index < 1 Or index > chain.Count Then Exit Function
    m_identifier = VBA.CStr(chain(index)("RowId"))
    Set m_area = selectedArea
    FocusCell = private_Show()
    Exit Function
Failed:
    FocusCell = False
    private_Report VBA.Err.Description
End Function

' //
' // Private
' //
Private Sub m_application_SheetSelectionChange( _
    ByVal sheet As Object, _
    ByVal target As Range _
)
    If m_area Is Nothing Then Exit Sub
    If Not sheet Is m_area.Parent Then
        private_Hide
    ElseIf Application.Intersect(target, m_area) Is Nothing Then
        private_Hide
    End If
End Sub

Private Sub m_application_SheetDeactivate(ByVal sheet As Object)
    private_Hide
End Sub

Private Sub private_Hide()
    On Error Resume Next
    private_ShowRemoveButton 0
    If Not m_context Is Nothing Then
        m_context.Router.UnregisterShape "btn_TdbBasisUp"
        m_context.Router.UnregisterShape "btn_TdbBasisDown"
        m_context.Styles.ForgetShape "btn_TdbBasisUp"
        m_context.Styles.ForgetShape "btn_TdbBasisDown"
    End If
    If Not m_up Is Nothing Then m_up.Delete
    If Not m_down Is Nothing Then m_down.Delete
    Set m_up = Nothing
    Set m_down = Nothing
    Set m_area = Nothing
    m_identifier = VBA.vbNullString
    On Error GoTo 0
End Sub

Private Function private_Index(ByVal chain As Collection) As Long
    Dim index As Long

    For index = 1 To chain.Count
        If VBA.CStr(chain(index)("RowId")) = m_identifier Then
            private_Index = index
            Exit Function
        End If
    Next index
End Function

Private Function private_Show() As Boolean
    Dim chain As Collection
    Dim index As Long
    Dim gutter As Range
    Dim gap As Double
    Dim height As Double

    Set chain = ex_TDB_BasisChain.fn_Read(m_pageBase.BindingContext)
    index = private_Index(chain)
    If index = 0 Then
        private_Hide
        private_Show = True
        Exit Function
    End If
    Set gutter = m_area.Worksheet.Cells(m_area.Row, m_area.Column - 1)
    gap = VBA.CDbl(ex_TravelDocumentsBuilder.fn_Config("basisChain.reorder.gap"))
    height = (m_area.Height - gap) / 2
    If height <= 0 Or gutter.Width <= 0 Then Exit Function
    Set m_up = private_CreateArrow(gutter, "btn_TdbBasisUp", "upGlyph", m_area.Top, height, index > 1)
    Set m_down = private_CreateArrow(gutter, "btn_TdbBasisDown", "downGlyph", m_area.Top + height + gap, height, index < chain.Count)
    private_ShowRemoveButton index + 1
    private_Show = True
End Function

Private Sub private_ShowRemoveButton(ByVal visibleRow As Long)
    Dim shape As Shape
    Dim selectedName As String
    Dim tableArea As Range

    If m_context Is Nothing Then Exit Sub
    If Not m_context.TryGetElementArea("BasisList", tableArea) Then Exit Sub
    selectedName = "chk_BasisList_r" & VBA.CStr(visibleRow) & "c" & VBA.CStr(tableArea.Columns.Count)
    For Each shape In tableArea.Worksheet.Shapes
        If VBA.Left$(shape.Name, 15) = "chk_BasisList_r" Then
            If shape.Name <> "chk_BasisList_r1c" & VBA.CStr(tableArea.Columns.Count) Then
                shape.Visible = (shape.Name = selectedName)
            End If
        End If
    Next shape
End Sub

Private Function private_CreateArrow( _
    ByVal gutter As Range, _
    ByVal name As String, _
    ByVal glyphKey As String, _
    ByVal top As Double, _
    ByVal height As Double, _
    ByVal enabled As Boolean _
) As Shape
    Dim arrow As Shape
    Dim document As Object
    Dim node As Object
    Dim caption As String

    caption = ex_TravelDocumentsBuilder.fn_Config("basisChain.reorder." & glyphKey)
    Set document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    Set node = document.createElement("button")
    node.setAttribute "name", name
    node.setAttribute "style", ex_TravelDocumentsBuilder.fn_Config("basisChain.reorder.style")
    If Not enabled Then
        node.setAttribute "textColor", ex_TravelDocumentsBuilder.fn_Config("basisChain.reorder.disabledColor")
    End If
    Set arrow = gutter.Worksheet.Shapes.AddShape(msoShapeRectangle, gutter.Left, top, gutter.Width, height)
    arrow.Name = name
    arrow.Placement = xlMoveAndSize
    m_context.Styles.ApplyControlStyle gutter, arrow, node, m_pageBase.BindingContext
    m_context.Styles.ApplyButtonCaption arrow, caption
    arrow.TextFrame2.MarginLeft = 0
    arrow.TextFrame2.MarginRight = 0
    arrow.TextFrame2.MarginTop = 0
    arrow.TextFrame2.MarginBottom = 0
    arrow.Visible = msoTrue
    arrow.ZOrder msoBringToFront
    arrow.OnAction = "ex_UiBridge.fn_OnShapeClick"
    m_context.Router.RegisterShape name, Me
    Set private_CreateArrow = arrow
End Function

Private Function private_Move(ByVal direction As Long) As Boolean
    Dim chain As Collection
    Dim row As Object
    Dim index As Long
    Dim target As Range
    Dim identifier As String
    Dim diagnostic As String
    Dim eventsBefore As Boolean
    Dim guarded As Boolean

    On Error GoTo Failed
    If m_area Is Nothing Then Exit Function
    Set chain = ex_TDB_BasisChain.fn_Read(m_pageBase.BindingContext)
    index = private_Index(chain)
    If index = 0 Then
        private_Hide
        private_Move = True
        Exit Function
    End If
    If index + direction < 1 Or index + direction > chain.Count Then
        private_Move = True
        Exit Function
    End If
    Set target = m_area.Worksheet.Cells(m_area.Row + direction, m_area.Column)
    identifier = m_identifier
    Set row = chain(index)
    chain.Remove index
    If direction < 0 Then
        chain.Add row, Before:=index - 1
    Else
        chain.Add row, After:=index
    End If
    private_Hide
    If Not ex_TDB_BasisChain.fn_Publish(m_pageBase.BindingContext) Then Exit Function
    If Not m_pageBase.RefreshContent("BasisList", diagnostic) Then VBA.Err.Raise 5, , diagnostic
    eventsBefore = Application.EnableEvents
    guarded = True
    Application.EnableEvents = False
    target.MergeArea.Select
    Application.EnableEvents = eventsBefore
    guarded = False
    Set m_area = target.MergeArea
    m_identifier = identifier
    private_Move = private_Show()
    Exit Function
Failed:
    If guarded Then Application.EnableEvents = eventsBefore
    private_Report VBA.Err.Description
End Function

Private Sub private_Report(ByVal diagnostic As String)
    private_Hide
    ex_Core.fn_Diagnostic_WriteLog "TDB_BASIS_REORDER_FAILED | " & diagnostic
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation
End Sub