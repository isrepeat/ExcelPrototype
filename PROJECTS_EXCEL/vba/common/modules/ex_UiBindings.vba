Option Explicit

Private shapeCommands As Object
Private cellBindings As Object
Private selectControls As Object
Private selectActionsByShape As Object
Private nextSelectControlId As Long

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    private_DisposeSelectControls
    Set shapeCommands = Nothing
    Set cellBindings = Nothing
    Set selectControls = Nothing
    Set selectActionsByShape = Nothing
    nextSelectControlId = 0
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Reset()
    private_DisposeSelectControls
    Set shapeCommands = VBA.CreateObject("Scripting.Dictionary")
    shapeCommands.CompareMode = VBA.vbTextCompare
    Set cellBindings = VBA.CreateObject("Scripting.Dictionary")
    cellBindings.CompareMode = VBA.vbTextCompare
    Set selectControls = VBA.CreateObject("Scripting.Dictionary")
    selectControls.CompareMode = VBA.vbTextCompare
    Set selectActionsByShape = VBA.CreateObject("Scripting.Dictionary")
    selectActionsByShape.CompareMode = VBA.vbTextCompare
    nextSelectControlId = 0
End Sub

Public Sub fn_Register(ByVal shapeName As String, ByVal uiCommand As obj_UiCommand)
    If uiCommand Is Nothing Then Exit Sub
    If shapeCommands Is Nothing Then fn_Reset
    Set shapeCommands(shapeName) = uiCommand
End Sub

Public Function fn_TryGetCommand(ByVal shapeName As String, ByRef outUiCommand As obj_UiCommand) As Boolean
    Set outUiCommand = Nothing
    If shapeCommands Is Nothing Then Exit Function
    If Not shapeCommands.Exists(shapeName) Then Exit Function
    Set outUiCommand = shapeCommands(shapeName)
    fn_TryGetCommand = True
End Function

Public Sub fn_ClearCellBindings(ByVal worksheetName As String)
    Dim bindingKey As Variant
    Dim uiCellBinding As obj_UiCellBinding
    Dim keysToRemove As Collection
    Dim keyToRemove As Variant

    If cellBindings Is Nothing Then Exit Sub
    Set keysToRemove = New Collection
    For Each bindingKey In cellBindings.Keys
        Set uiCellBinding = cellBindings(bindingKey)
        If VBA.StrComp(uiCellBinding.WorksheetName, worksheetName, VBA.vbTextCompare) = 0 Then
            keysToRemove.Add VBA.CStr(bindingKey)
        End If
    Next bindingKey
    For Each keyToRemove In keysToRemove
        Set uiCellBinding = cellBindings(keyToRemove)
        uiCellBinding.Dispose
        cellBindings.Remove VBA.CStr(keyToRemove)
    Next keyToRemove
End Sub

Public Function fn_RegisterCellBinding(ByVal uiCellBinding As obj_UiCellBinding) As Boolean
    Dim bindingKey As String

    If uiCellBinding Is Nothing Then Exit Function
    If cellBindings Is Nothing Then fn_Reset
    bindingKey = private_CellBindingKey(uiCellBinding.WorksheetName, uiCellBinding.CellAddress)
    If VBA.Len(bindingKey) = 0 Then Exit Function
    If cellBindings.Exists(bindingKey) Then
        Set cellBindings(bindingKey) = uiCellBinding
    Else
        cellBindings.Add bindingKey, uiCellBinding
    End If
    fn_RegisterCellBinding = True
End Function

Public Function fn_NextSelectControlId() As Long
    If shapeCommands Is Nothing Then fn_Reset
    nextSelectControlId = nextSelectControlId + 1
    fn_NextSelectControlId = nextSelectControlId
End Function

Public Function fn_RegisterSelectControl( _
    ByVal uiSelectAction As obj_UiSelectShapeAction _
) As Boolean
    Dim controlKey As String
    Dim shapeName As Variant

    If uiSelectAction Is Nothing Then Exit Function
    If selectControls Is Nothing Then fn_Reset
    controlKey = VBA.CStr(uiSelectAction.ControlId)
    If selectControls.Exists(controlKey) Then Exit Function
    Set selectControls(controlKey) = uiSelectAction
    If selectActionsByShape Is Nothing Then
        Set selectActionsByShape = VBA.CreateObject("Scripting.Dictionary")
        selectActionsByShape.CompareMode = VBA.vbTextCompare
    End If
    For Each shapeName In uiSelectAction.ShapeNames
        Set selectActionsByShape(VBA.CStr(shapeName)) = uiSelectAction
    Next shapeName
    fn_RegisterSelectControl = True
End Function

Public Function fn_HandleSelectShapeClick(ByVal shapeName As String) As Boolean
    Dim uiSelectAction As obj_UiSelectShapeAction

    If selectActionsByShape Is Nothing Then Exit Function
    If Not selectActionsByShape.Exists(shapeName) Then Exit Function
    Set uiSelectAction = selectActionsByShape(shapeName)
    uiSelectAction.HandleShapeClick shapeName
    fn_HandleSelectShapeClick = True
End Function

Public Sub fn_ClearSelectControls(ByVal worksheetName As String)
    Dim controlKey As Variant
    Dim uiSelectAction As obj_UiSelectShapeAction
    Dim controlKeysToRemove As Collection
    Dim shapeName As Variant

    If selectControls Is Nothing Then Exit Sub
    Set controlKeysToRemove = New Collection
    For Each controlKey In selectControls.Keys
        Set uiSelectAction = selectControls(controlKey)
        If VBA.StrComp(uiSelectAction.WorksheetName, worksheetName, VBA.vbTextCompare) = 0 Then
            uiSelectAction.CollapseDropdown
            For Each shapeName In uiSelectAction.ShapeNames
                If Not selectActionsByShape Is Nothing Then
                    If selectActionsByShape.Exists(VBA.CStr(shapeName)) Then _
                        selectActionsByShape.Remove VBA.CStr(shapeName)
                End If
                If Not shapeCommands Is Nothing Then
                    If shapeCommands.Exists(VBA.CStr(shapeName)) Then _
                        shapeCommands.Remove VBA.CStr(shapeName)
                End If
            Next shapeName
            uiSelectAction.Dispose
            controlKeysToRemove.Add VBA.CStr(controlKey)
        End If
    Next controlKey
    For Each controlKey In controlKeysToRemove
        selectControls.Remove VBA.CStr(controlKey)
    Next controlKey
End Sub

Public Sub fn_CollapseSelectControls()
    Dim controlKey As Variant
    Dim uiSelectAction As obj_UiSelectShapeAction

    If selectControls Is Nothing Then Exit Sub
    For Each controlKey In selectControls.Keys
        Set uiSelectAction = selectControls(controlKey)
        uiSelectAction.CollapseDropdown
    Next controlKey
End Sub

Public Function fn_HandleCellChange(ByVal target As Range) As Boolean
    Dim changedCell As Range
    Dim bindingKey As String
    Dim uiCellBinding As obj_UiCellBinding
    Dim previousEnableEvents As Boolean
    Dim wasHandled As Boolean

    If target Is Nothing Or cellBindings Is Nothing Then Exit Function
    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    For Each changedCell In target.Cells
        bindingKey = private_CellBindingKey(changedCell.Parent.Name, changedCell.Address(False, False))
        If cellBindings.Exists(bindingKey) Then
            Set uiCellBinding = cellBindings(bindingKey)
            If Not uiCellBinding.HandleCellChange(changedCell) Then GoTo CleanExit
            wasHandled = True
        End If
    Next changedCell
    fn_HandleCellChange = wasHandled
CleanExit:
    Application.EnableEvents = previousEnableEvents
    Exit Function
EH:
    ex_Core.fn_Diagnostic_WriteLog "UI_CELL_CHANGE_ERROR | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    Resume CleanExit
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_CellBindingKey(ByVal worksheetName As String, ByVal cellAddress As String) As String
    worksheetName = VBA.Trim$(worksheetName)
    cellAddress = VBA.Trim$(cellAddress)
    If VBA.Len(worksheetName) = 0 Or VBA.Len(cellAddress) = 0 Then Exit Function
    private_CellBindingKey = worksheetName & "!" & cellAddress
End Function

Private Sub private_DisposeSelectControls()
    Dim controlKey As Variant
    Dim uiSelectAction As obj_UiSelectShapeAction

    If selectControls Is Nothing Then Exit Sub
    For Each controlKey In selectControls.Keys
        Set uiSelectAction = selectControls(controlKey)
        uiSelectAction.CollapseDropdown
        uiSelectAction.Dispose
    Next controlKey
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------