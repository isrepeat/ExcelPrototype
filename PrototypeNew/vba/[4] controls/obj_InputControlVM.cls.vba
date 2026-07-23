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
Private m_OnChangeArgRaw As String
Private m_OnChangeArgResolved As Variant
Private m_HasOnChangeArg As Boolean
Private m_OnChangeMacroRef As String
Private m_CallbackContext As Object
Private m_RuntimeControlKey As String
Private m_IsConfigured As Boolean
Private m_Page As obj_IPage
Private m_RenderedInputRange As Range

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
    Set m_RenderedInputRange = Nothing
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
    Set m_RenderedInputRange = Nothing
    m_ValueResolved = VBA.vbNullString
    m_OnChangeArgRaw = VBA.vbNullString
    m_HasOnChangeArg = False
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
    m_OnChangeArgRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "onChangeArg"))
    If VBA.Len(VBA.Trim$(m_OnChangeArgRaw)) > 0 Then
        If Not ex_BindingRuntime.fn_TryResolveValueBinding(m_OnChangeArgRaw, dataContext, m_OnChangeArgResolved) Then Exit Sub
        m_HasOnChangeArg = True
    End If

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Input", m_ControlName, "style") Then Exit Sub

    m_RuntimeControlKey = "input|" & VBA.LCase$(VBA.Trim$(m_ControlLayout.LayoutSheetName & "|" & m_ControlName))
    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim inputCell As Range
    Dim inputRange As Range
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
    Set inputRange = ws.Range(ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart), ws.Cells(m_ControlLayout.RowEnd, m_ControlLayout.ColEnd))
    If inputRange.Cells.CountLarge > 1 Then inputRange.Merge
    Set inputCell = ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart)
    On Error GoTo 0

    If inputCell Is Nothing Then Exit Sub
    Set m_RenderedInputRange = inputRange

    ' Для input всегда фиксируем текстовый формат, чтобы Excel не съедал пользовательский ввод
    ' (даты/коды/лидирующие нули) до того, как onChange обработает значение.
    inputRange.NumberFormat = "@"
    inputRange.HorizontalAlignment = xlHAlignLeft
    inputRange.VerticalAlignment = xlVAlignCenter
    inputRange.WrapText = False
    If Not private_ApplyPresetStyle(inputRange, m_ControlLayout.StyleName) Then Exit Sub
    If Not private_RegisterControlPart(ws, inputRange) Then Exit Sub

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

Private Function obj_IControl_Measure( _
    ByVal controlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanColls As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    outSpanRows = 1
    outSpanColls = 1
    obj_IControl_Measure = True
End Function

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "value", "text", "onchange", "onchangemacro", "onchangearg"
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
    Dim callbackPayload As Object

    If VBA.Len(VBA.Trim$(m_OnChangeMacroRef)) = 0 Then
        RuntimeHandleCellChange = True
        Exit Function
    End If

    changedCellAddress = VBA.Trim$(changedCellAddress)
    If m_HasOnChangeArg Then
        Set callbackPayload = private_BuildCellChangePayload(changedCellAddress)
        RuntimeHandleCellChange = rt_Bridge.fn_RunCallback(m_OnChangeMacroRef, m_CallbackContext, callbackPayload)
    Else
        RuntimeHandleCellChange = rt_Bridge.fn_RunCallback(m_OnChangeMacroRef, m_CallbackContext, changedCellAddress)
    End If
End Function

Public Function TryGetValue(ByRef outValue As String) As Boolean
    outValue = VBA.vbNullString
    If m_RenderedInputRange Is Nothing Then Exit Function

    ' Контрол владеет ссылкой на свой rendered Range. Consumers не должны
    ' восстанавливать ячейку через глобальный индекс visual parts: тот индекс
    ' предназначен для styles/layout и может меняться при partial render.
    outValue = VBA.Trim$(VBA.CStr(m_RenderedInputRange.Cells(1, 1).Value2))
    TryGetValue = True
End Function

Public Function ClearValue() As Boolean
    If m_RenderedInputRange Is Nothing Then Exit Function
    m_RenderedInputRange.ClearContents
    ClearValue = True
End Function

' //
' // Internal
' //
Private Function private_ApplyPresetStyle(ByVal targetRange As Range, ByVal styleName As String) As Boolean
    If targetRange Is Nothing Then Exit Function

    Select Case VBA.LCase$(VBA.Trim$(styleName))
        Case VBA.vbNullString
            ' no-op

        Case "lookupinput", "inputfield"
            targetRange.Interior.Color = VBA.RGB(255, 247, 214)
            targetRange.Font.Color = VBA.RGB(31, 35, 41)
            targetRange.Font.Bold = True
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = VBA.RGB(245, 158, 11)
            targetRange.Borders.Weight = xlMedium

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Input: unsupported style '" & styleName & "' for control '" & m_ControlName & "'."
#End If
            Exit Function
    End Select

    private_ApplyPresetStyle = True
End Function

Private Function private_RegisterControlPart( _
    ByVal ws As Worksheet, _
    ByVal inputRange As Range _
) As Boolean
    If ws Is Nothing Then Exit Function
    If inputRange Is Nothing Then Exit Function

    private_RegisterControlPart = ex_ControlPartsRuntime.fn_RegisterControlPart( _
        ws, _
        "input", _
        m_ControlName, _
        "cell", _
        inputRange)
End Function

Private Function private_BuildCellChangePayload(ByVal changedCellAddress As String) As Object
    Dim payload As Object

    Set payload = ex_Helpers.fn_CreateDictionaryTextCompare()
    payload("ChangedCellAddress") = VBA.Trim$(changedCellAddress)
    payload("Arg") = m_OnChangeArgResolved

    Set private_BuildCellChangePayload = payload
End Function

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
