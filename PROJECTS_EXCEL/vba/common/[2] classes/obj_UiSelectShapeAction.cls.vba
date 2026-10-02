VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiSelectShapeAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiEventHandler

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_controlId As Long
Private m_worksheet As Worksheet
Private m_targetCell As Range
Private m_headerShapeName As String
Private m_panelShapeName As String
Private m_itemShapeNames As Collection
Private m_shapeNames As Collection
Private m_items As Collection
Private m_selectedValue As String
Private WithEvents m_uiCellBinding As obj_UiCellBinding
Private m_isExpanded As Boolean

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
    If kind = "dismiss" Then
        obj_IUiEventHandler_HandleEvent = Me.CollapseDropdown()
    Else
        obj_IUiEventHandler_HandleEvent = Me.HandleShapeClick(VBA.CStr(payload))
    End If
End Function

' //
' // Properties
' //
Public Property Get ControlId() As Long
    ControlId = m_controlId
End Property

Public Property Get WorksheetName() As String
    If Not m_worksheet Is Nothing Then WorksheetName = m_worksheet.Name
End Property

Public Property Get ShapeNames() As Collection
    Set ShapeNames = m_shapeNames
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal controlId As Long, _
    ByVal targetCell As Range, _
    ByVal headerShapeName As String, _
    ByVal panelShapeName As String, _
    ByVal itemShapeNames As Collection, _
    ByVal shapeNames As Collection, _
    ByVal items As Collection, _
    ByVal selectedValue As String, _
    ByVal uiCellBinding As obj_UiCellBinding _
) As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If controlId <= 0 Or targetCell Is Nothing Then Exit Function
    If itemShapeNames Is Nothing Or shapeNames Is Nothing Or items Is Nothing Then Exit Function
    If uiCellBinding Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(headerShapeName)) = 0 Or _
       VBA.Len(VBA.Trim$(panelShapeName)) = 0 Then Exit Function
    If itemShapeNames.Count <> items.Count Then Exit Function

    m_controlId = controlId
    Set m_worksheet = targetCell.Parent
    Set m_targetCell = targetCell.Cells(1, 1)
    m_headerShapeName = headerShapeName
    m_panelShapeName = panelShapeName
    Set m_itemShapeNames = itemShapeNames
    Set m_shapeNames = shapeNames
    Set m_items = items
    m_selectedValue = selectedValue
    Set m_uiCellBinding = uiCellBinding
    m_isExpanded = False
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_uiCellBinding = Nothing
    Set m_items = Nothing
    Set m_shapeNames = Nothing
    Set m_itemShapeNames = Nothing
    Set m_targetCell = Nothing
    Set m_worksheet = Nothing
    m_headerShapeName = VBA.vbNullString
    m_panelShapeName = VBA.vbNullString
    m_selectedValue = VBA.vbNullString
    m_controlId = 0
    m_isExpanded = False
End Sub

Public Function HandleShapeClick(ByVal callerShapeName As String) As Boolean
    Dim itemIndex As Long

    If m_isDisposed Then Exit Function
    If VBA.StrComp(callerShapeName, m_headerShapeName, VBA.vbTextCompare) = 0 Then
        HandleShapeClick = private_ToggleDropdown()
        Exit Function
    End If
    If VBA.StrComp(callerShapeName, m_panelShapeName, VBA.vbTextCompare) = 0 Then
        HandleShapeClick = Me.CollapseDropdown()
        Exit Function
    End If
    For itemIndex = 1 To m_itemShapeNames.Count
        If VBA.StrComp(callerShapeName, VBA.CStr(m_itemShapeNames(itemIndex)), _
                VBA.vbTextCompare) = 0 Then
            HandleShapeClick = private_SelectItem(itemIndex)
            Exit Function
        End If
    Next itemIndex
End Function

Public Function CollapseDropdown() As Boolean
    If m_isDisposed Then
        CollapseDropdown = True
        Exit Function
    End If
    If Not m_isExpanded Then
        CollapseDropdown = True
        Exit Function
    End If
    m_isExpanded = False
    private_UpdateHeader
    CollapseDropdown = private_ApplyVisibility()
End Function

' //
' // Private
' //
Private Function private_ToggleDropdown() As Boolean
    m_isExpanded = Not m_isExpanded
    private_UpdateHeader
    private_ToggleDropdown = private_ApplyVisibility()
End Function

Private Function private_SelectItem(ByVal itemIndex As Long) As Boolean
    Dim previousEnableEvents As Boolean
    Dim selectedValue As String

    If itemIndex <= 0 Or itemIndex > m_items.Count Then Exit Function
    If m_targetCell Is Nothing Or m_uiCellBinding Is Nothing Then Exit Function
    selectedValue = VBA.CStr(m_items(itemIndex))
    m_selectedValue = selectedValue
    m_isExpanded = False
    private_UpdateHeader
    If Not private_ApplyVisibility() Then Exit Function

    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    m_targetCell.Value2 = selectedValue
    Application.EnableEvents = previousEnableEvents
    If Not m_uiCellBinding.HandleCellChange(m_targetCell) Then Exit Function
    private_SelectItem = True
    Exit Function
EH:
    Application.EnableEvents = previousEnableEvents
    ex_Core.fn_Diagnostic_WriteLog "UI_SELECT_ITEM_ERROR | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
End Function

Private Function private_ApplyVisibility() As Boolean
    Dim panelShape As Shape
    Dim itemShape As Shape
    Dim itemIndex As Long

    If m_worksheet Is Nothing Then Exit Function
    On Error GoTo EH
    Set panelShape = m_worksheet.Shapes(m_panelShapeName)
    panelShape.Visible = VBA.IIf(m_isExpanded, msoTrue, msoFalse)
    If m_isExpanded Then panelShape.ZOrder msoBringToFront
    For itemIndex = 1 To m_itemShapeNames.Count
        Set itemShape = m_worksheet.Shapes(VBA.CStr(m_itemShapeNames(itemIndex)))
        itemShape.Visible = VBA.IIf(m_isExpanded, msoTrue, msoFalse)
        If m_isExpanded Then itemShape.ZOrder msoBringToFront
    Next itemIndex
    private_ApplyVisibility = True
    Exit Function
EH:
    ex_Core.fn_Diagnostic_WriteLog "UI_SELECT_VISIBILITY_ERROR | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
End Function

Private Sub private_UpdateHeader()
    Dim headerShape As Shape
    Dim arrowText As String

    If m_worksheet Is Nothing Then Exit Sub
    On Error Resume Next
    Set headerShape = m_worksheet.Shapes(m_headerShapeName)
    If headerShape Is Nothing Then Exit Sub
    If m_isExpanded Then
        arrowText = VBA.ChrW(&H25B2)
    Else
        arrowText = VBA.ChrW(&H25BC)
    End If
    headerShape.TextFrame2.TextRange.Text = m_selectedValue & " " & arrowText
    On Error GoTo 0
End Sub

Private Sub m_uiCellBinding_ValueRefreshed()
    private_uiCellBinding_ValueRefreshed
End Sub

Private Sub private_uiCellBinding_ValueRefreshed()
    If m_isDisposed Or m_targetCell Is Nothing Then Exit Sub
    m_selectedValue = VBA.CStr(m_targetCell.Value2)
    private_UpdateHeader
End Sub