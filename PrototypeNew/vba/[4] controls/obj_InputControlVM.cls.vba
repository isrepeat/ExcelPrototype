VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_InputControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IControl

Private m_ControlBase As obj_ControlBase
Private m_ControlLayout As obj_ControlLayout
Private m_ControlName As String
Private m_ValueRaw As String
Private m_ValueResolved As String
Private m_OnChangeRaw As String
Private m_OnChangeMacroRef As String
Private m_CallbackContext As Object
Private m_RuntimeControlKey As String
Private m_IsConfigured As Boolean
Private m_Page As obj_IPage

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    obj_IControl_Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IControl_Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    m_IsConfigured = False
    Set m_Page = page
    obj_IControl_Initialize = True
End Function

Private Sub obj_IControl_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_ControlLayout = Nothing
    Set m_CallbackContext = Nothing
    Set m_Page = Nothing
    m_RuntimeControlKey = VBA.vbNullString
    m_IsConfigured = False
    On Error GoTo 0
End Sub

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim dataContext As Object
    Dim pageBase As obj_PageBase

    m_IsConfigured = False
    Set m_ControlLayout = Nothing
    Set m_ControlBase = Nothing
    Set m_CallbackContext = Nothing
    m_ValueResolved = VBA.vbNullString
    m_RuntimeControlKey = VBA.vbNullString

    Set pageBase = m_Page.GetPageBase()
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "Input", "input", m_ControlName) Then Exit Sub

    m_ValueRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "value"))
    If VBA.Len(VBA.Trim$(m_ValueRaw)) = 0 Then
        m_ValueRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "text"))
    End If

    Set dataContext = m_ControlBase.DataContext
    If dataContext Is Nothing Then Set dataContext = m_Page
    Set m_CallbackContext = dataContext

    If VBA.Len(VBA.Trim$(m_ValueRaw)) > 0 Then
        If Not ex_BindingRuntime.fn_TryResolveTextBinding(m_ValueRaw, dataContext, m_ValueResolved) Then Exit Sub
    End If

    m_OnChangeRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "onChange"))
    If VBA.Len(VBA.Trim$(m_OnChangeRaw)) = 0 Then
        ' Совместимость со старой схемой атрибутов:
        ' поддерживаем и onChangeMacro, и onChange.
        m_OnChangeRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "onChangeMacro"))
    End If
    m_OnChangeMacroRef = VBA.vbNullString
    If VBA.Len(VBA.Trim$(m_OnChangeRaw)) > 0 Then
        If Not private_TryResolveCallbackRef(m_OnChangeRaw, m_CallbackContext, m_OnChangeMacroRef) Then Exit Sub
    End If

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Input", m_ControlName, "style") Then Exit Sub

    m_RuntimeControlKey = "input|" & VBA.LCase$(VBA.Trim$(m_ControlLayout.LayoutSheetName & "|" & m_ControlName))
    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim inputCell As Range
    Dim currentValue As String
    Dim pageBase As obj_PageBase

    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Input: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If

    Set pageBase = Nothing
    If Not m_ControlBase Is Nothing Then Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Input: page is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(pageBase, m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Input: sheet '" & m_ControlLayout.LayoutSheetName & "' was not found for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    On Error GoTo EH_RANGE
    Set inputCell = ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart)
    On Error GoTo 0

    If inputCell Is Nothing Then Exit Sub

    If m_ControlLayout.RowEnd <> m_ControlLayout.RowStart Or m_ControlLayout.ColEnd <> m_ControlLayout.ColStart Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogWarning "Input: control '" & m_ControlName & "' is forced to a single cell (" & inputCell.Address(False, False) & ")."
#End If
    End If

    ' Для input всегда фиксируем текстовый формат, чтобы Excel не съедал пользовательский ввод
    ' (даты/коды/лидирующие нули) до того, как onChange обработает значение.
    inputCell.NumberFormat = "@"
    inputCell.HorizontalAlignment = xlHAlignLeft
    inputCell.VerticalAlignment = xlVAlignCenter
    inputCell.WrapText = False

    currentValue = VBA.Trim$(VBA.CStr(inputCell.Value2))

    ' Начальный value из XML применяем только при первом рендере пустой ячейки,
    ' чтобы не затирать уже введенное пользователем значение на следующих rerender.
    If VBA.Len(currentValue) = 0 And VBA.Len(VBA.Trim$(m_ValueResolved)) > 0 Then
        inputCell.Value2 = m_ValueResolved
    End If

    If Not private_TryBindRuntimeRoute(inputCell) Then Exit Sub
    Exit Sub

EH_RANGE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Input: failed to resolve target cell for control '" & m_ControlName & "'."
#End If
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "value", "text", "onchange", "onchangemacro"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

' //
' // API
' //
Public Function RuntimeHandleCellChange(Optional ByVal changedCellAddress As String = VBA.vbNullString) As Boolean
    If VBA.Len(VBA.Trim$(m_OnChangeMacroRef)) = 0 Then
        RuntimeHandleCellChange = True
        Exit Function
    End If

    changedCellAddress = VBA.Trim$(changedCellAddress)
    RuntimeHandleCellChange = rt_Bridge.fn_RunCallback(m_OnChangeMacroRef, m_CallbackContext, changedCellAddress)
End Function

' //
' // Internal
' //
Private Function private_TryBindRuntimeRoute(ByVal inputCell As Range) As Boolean
    Dim pageBase As obj_PageBase

    If inputCell Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(m_RuntimeControlKey)) = 0 Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    ' Регистрируем route по адресу ячейки (а не по shape), т.к. вход идет из Workbook_SheetChange.
    If Not pageBase.RegisterControl(m_RuntimeControlKey, Me) Then Exit Function
    If Not pageBase.RegisterCellRoute(inputCell.Address(False, False), m_RuntimeControlKey, "RuntimeHandleCellChange", True, inputCell.Address(False, False)) Then Exit Function

    private_TryBindRuntimeRoute = True
End Function

Private Function private_TryResolveCallbackRef( _
    ByVal rawText As String, _
    ByVal dataContext As Object, _
    ByRef outCallbackRef As String _
) As Boolean
    Dim resolvedValue As Variant

    outCallbackRef = VBA.vbNullString
    rawText = VBA.Trim$(rawText)
    If VBA.Len(rawText) = 0 Then
        private_TryResolveCallbackRef = True
        Exit Function
    End If

    If Not ex_BindingRuntime.fn_TryResolveValueBinding(rawText, dataContext, resolvedValue) Then Exit Function
    If VBA.IsObject(resolvedValue) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Input: callback binding must resolve to scalar value for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    outCallbackRef = VBA.Trim$(VBA.CStr(resolvedValue))
    If VBA.Len(outCallbackRef) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Input: callback binding resolved to empty value for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    private_TryResolveCallbackRef = True
End Function

Private Function private_GetWorksheetByName(ByVal page As obj_PageBase, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If page Is Nothing Then Exit Function
    Set ws = page.Worksheet
    If ws Is Nothing Then Exit Function

    sheetName = VBA.LCase$(VBA.Trim$(sheetName))
    If VBA.Len(sheetName) > 0 Then
        If VBA.StrComp(VBA.LCase$(VBA.Trim$(ws.Name)), sheetName, VBA.vbTextCompare) <> 0 Then Exit Function
    End If

    Set private_GetWorksheetByName = ws
End Function
