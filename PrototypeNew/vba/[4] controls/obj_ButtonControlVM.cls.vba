VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ButtonControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IControl
Implements obj_ISerializable

Private Const DEFAULT_CAPTION As String = "Update Code"
Private Const INLINE_PART_BUTTON As String = "button"

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_CaptionRaw As String
Private m_CaptionInlineSource As String
Private m_OnClickRaw As String
Private m_ControlLayout As obj_ControlLayout
Private m_CaptionText As String
Private m_CaptionInlineTextPart As obj_InlineTextPart
Private m_OnClickMacroRef As String
Private m_OnClickCallbackContext As Object
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
    Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
' Configure поднимает контракт контрола из XML:
' читаем attrs -> резолвим биндинги -> нормализуем layout -> готовим runtime key.
Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim dataContext As Object
    Dim callbackContext As Object
    Dim captionTextResolved As String
    Dim onClickResolved As Variant
    Dim pageBase As obj_PageBase

    m_IsConfigured = False
    Set m_ControlLayout = Nothing
    Set m_ControlBase = Nothing
    Set m_CaptionInlineTextPart = New obj_InlineTextPart
    m_CaptionInlineSource = VBA.vbNullString
    m_CaptionText = VBA.vbNullString
    m_RuntimeControlKey = VBA.vbNullString

    Set pageBase = m_Page.GetPageBase()
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "Button", "button", m_ControlName) Then Exit Sub
    Set m_OnClickCallbackContext = Nothing

    m_CaptionRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "caption"))
    If VBA.Len(m_CaptionRaw) = 0 Then m_CaptionRaw = DEFAULT_CAPTION

    m_OnClickRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "onClick"))
    If VBA.Len(VBA.Trim$(m_OnClickRaw)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: onClick is required for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set dataContext = m_ControlBase.DataContext
    If dataContext Is Nothing Then Set dataContext = Me

    If Not ex_BindingRuntime.fn_TryResolveTextBinding(m_CaptionRaw, dataContext, captionTextResolved) Then Exit Sub
    m_CaptionInlineSource = captionTextResolved
    If Not private_TryResolveCaptionInlineText(pageBase, m_CaptionInlineSource) Then Exit Sub

    Set callbackContext = m_ControlBase.DataContext
    If Not ex_BindingRuntime.fn_TryResolveValueBinding(m_OnClickRaw, callbackContext, onClickResolved) Then Exit Sub
    If VBA.IsObject(onClickResolved) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: onClick binding must resolve to scalar callback value for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If
    m_OnClickMacroRef = VBA.Trim$(VBA.CStr(onClickResolved))
    m_OnClickMacroRef = VBA.Trim$(m_OnClickMacroRef)
    If VBA.Len(m_OnClickMacroRef) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: onClick resolved to empty callback for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If
    Set m_OnClickCallbackContext = callbackContext

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Button", m_ControlName, "style") Then Exit Sub
    m_RuntimeControlKey = "button|" & VBA.LCase$(VBA.Trim$(m_ControlLayout.LayoutSheetName & "|" & m_ControlName))

    m_IsConfigured = True
End Sub

' Render создает/пересоздает shape-кнопку на целевом диапазоне листа
' и привязывает route клика через rt_Bridge -> PageBase маршрутизацию.
Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim buttonName As String
    Dim targetRange As Range
    Dim metaMap As Object
    Dim pageBase As obj_PageBase

    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If

    Set pageBase = Nothing
    If Not m_ControlBase Is Nothing Then Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then
        Set pageBase = m_Page.GetPageBase()
    End If
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: page is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set ws = pageBase.Worksheet
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: page worksheet is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart), ws.Cells(m_ControlLayout.RowEnd, m_ControlLayout.ColEnd))
    On Error GoTo 0

    If targetRange Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: failed to resolve target range for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If
    If targetRange.Width <= 0# Or targetRange.Height <= 0# Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: target range has non-positive width/height for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    buttonName = "btn_" & m_ControlName

    Set shp = private_GetUiShapeByName(ws, buttonName)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
        shp.Name = buttonName
    Else
        shp.Left = targetRange.Left
        shp.Top = targetRange.Top
        shp.Width = targetRange.Width
        shp.Height = targetRange.Height
    End If
    shp.Placement = xlMoveAndSize
    If Not private_TryBindRuntimeRoute(shp) Then Exit Sub

    On Error Resume Next
    shp.TextFrame2.TextRange.Text = m_CaptionText
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame.Characters.Text = m_CaptionText
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    If Not private_RegisterCaptionInlineRuns(pageBase, shp) Then Exit Sub

    Set metaMap = VBA.CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.control") = m_ControlName
    If VBA.Len(VBA.Trim$(m_ControlLayout.StyleName)) > 0 Then
        metaMap("pn.style") = m_ControlLayout.StyleName
    Else
        metaMap("pn.style") = VBA.vbNullString
    End If
    If Not ex_ShapeMetaRuntime.fn_TrySetShapeMetaValues(shp, metaMap) Then Exit Sub
    On Error GoTo EH_BUTTON

    Exit Sub

EH_BUTTON:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Button: failed to render control '" & m_ControlName & "': " & Err.Description
#End If
    Exit Sub

EH_RANGE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Button: failed to resolve target range for control '" & m_ControlName & "': " & Err.Description
#End If
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "caption", "onclick"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

Private Function obj_ISerializable_GetSerializableTypeRoot() As String
    obj_ISerializable_GetSerializableTypeRoot = "button"
End Function

Private Function obj_ISerializable_TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    obj_ISerializable_TrySerializeSnapshot = Me.TrySerializeSnapshot(outSnapshotXml)
End Function

Private Function obj_ISerializable_TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
    obj_ISerializable_TryDeserializeSnapshot = Me.TryDeserializeSnapshot(snapshotXml)
End Function

Private Function obj_ISerializable_TryRestoreState() As Boolean
    obj_ISerializable_TryRestoreState = True
End Function

' //
' // API
' //
Public Function Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    Set m_Page = page
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Err.Clear
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_ControlLayout = Nothing
    Set m_CaptionInlineTextPart = Nothing
    Set m_OnClickCallbackContext = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

' Callstack[1]: Shape.OnAction -> rt_Bridge.fn_OnShapeClick -> rt_PageManager.fn_TryGetPageByWorksheet -> page.DispatchShapeClick -> obj_PageMain.obj_IPage_DispatchShapeClick -> obj_PageBase.DispatchShapeClick -> obj_PageBase.private_TryInvokeControlAction -> obj_ButtonControlVM.RuntimeHandleClick
Public Function RuntimeHandleClick() As Boolean
    If Not rt_Bridge.fn_RunCallback(m_OnClickMacroRef, m_OnClickCallbackContext) Then Exit Function
    RuntimeHandleClick = True
End Function

' Snapshot хранит уже резолвленные runtime-значения,
' чтобы после reload можно было восстановить кнопку без повторного парсинга XML-контракта.
Public Function TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    Dim sheetName As String
    Dim shapeName As String
    Dim styleName As String
    Dim isConfiguredText As String

    outSnapshotXml = VBA.vbNullString

    If VBA.Len(VBA.Trim$(m_ControlName)) = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(m_RuntimeControlKey)) = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(m_OnClickMacroRef)) = 0 Then Exit Function
    If m_ControlLayout Is Nothing Then Exit Function

    sheetName = VBA.Trim$(m_ControlLayout.LayoutSheetName)
    If VBA.Len(sheetName) = 0 Then Exit Function
    shapeName = "btn_" & m_ControlName
    styleName = VBA.Trim$(m_ControlLayout.StyleName)
    isConfiguredText = VBA.IIf(m_IsConfigured, "true", "false")

    outSnapshotXml = _
        "<button version=""2""" & _
        " controlName=""" & ex_Helpers.fn_EscapeXmlAttr(m_ControlName) & """" & _
        " captionRaw=""" & ex_Helpers.fn_EscapeXmlAttr(m_CaptionRaw) & """" & _
        " captionInlineSource=""" & ex_Helpers.fn_EscapeXmlAttr(m_CaptionInlineSource) & """" & _
        " onClickRaw=""" & ex_Helpers.fn_EscapeXmlAttr(m_OnClickRaw) & """" & _
        " captionText=""" & ex_Helpers.fn_EscapeXmlAttr(m_CaptionText) & """" & _
        " onClickMacroRef=""" & ex_Helpers.fn_EscapeXmlAttr(m_OnClickMacroRef) & """" & _
        " runtimeKey=""" & ex_Helpers.fn_EscapeXmlAttr(m_RuntimeControlKey) & """" & _
        " sheet=""" & ex_Helpers.fn_EscapeXmlAttr(sheetName) & """" & _
        " rowStart=""" & VBA.CStr(m_ControlLayout.RowStart) & """" & _
        " colStart=""" & VBA.CStr(m_ControlLayout.ColStart) & """" & _
        " rowEnd=""" & VBA.CStr(m_ControlLayout.RowEnd) & """" & _
        " colEnd=""" & VBA.CStr(m_ControlLayout.ColEnd) & """" & _
        " style=""" & ex_Helpers.fn_EscapeXmlAttr(styleName) & """" & _
        " isConfigured=""" & isConfiguredText & """" & _
        " shape=""" & ex_Helpers.fn_EscapeXmlAttr(shapeName) & """" & _
        "/>"

    TrySerializeSnapshot = True
End Function

Public Function TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
    Dim dom As Object
    Dim root As Object
    Dim ws As Worksheet
    Dim shp As Shape
    Dim shapeName As String
    Dim layoutSheetName As String
    Dim layoutRowStart As Long
    Dim layoutColStart As Long
    Dim layoutRowEnd As Long
    Dim layoutColEnd As Long
    Dim layoutStyle As String
    Dim onClickMacroRef As String
    Dim isConfiguredAttr As String
    Dim versionText As String

    snapshotXml = VBA.Trim$(snapshotXml)
    If VBA.Len(snapshotXml) = 0 Then Exit Function

    If Not ex_Core.fn_CustomXmlPartStore_TryLoadDomFromXml(snapshotXml, dom) Then Exit Function
    Set root = dom.DocumentElement
    If root Is Nothing Then Exit Function
    If VBA.LCase$(VBA.CStr(root.baseName)) <> "button" Then Exit Function
    versionText = VBA.Trim$(VBA.CStr(root.getAttribute("version")))

    m_ControlName = VBA.Trim$(VBA.CStr(root.getAttribute("controlName")))
    m_CaptionRaw = VBA.CStr(root.getAttribute("captionRaw"))
    m_CaptionInlineSource = VBA.CStr(root.getAttribute("captionInlineSource"))
    m_OnClickRaw = VBA.CStr(root.getAttribute("onClickRaw"))
    m_CaptionText = VBA.CStr(root.getAttribute("captionText"))
    onClickMacroRef = VBA.Trim$(VBA.CStr(root.getAttribute("onClickMacroRef")))
    If VBA.Len(onClickMacroRef) = 0 Then onClickMacroRef = VBA.Trim$(VBA.CStr(root.getAttribute("onClick")))
    m_OnClickMacroRef = VBA.Trim$(onClickMacroRef)
    m_RuntimeControlKey = VBA.LCase$(VBA.Trim$(VBA.CStr(root.getAttribute("runtimeKey"))))
    shapeName = VBA.Trim$(VBA.CStr(root.getAttribute("shape")))
    layoutSheetName = VBA.Trim$(VBA.CStr(root.getAttribute("sheet")))
    layoutRowStart = ex_Helpers.fn_ReadSnapshotLongAttr(root, "rowStart", 1)
    layoutColStart = ex_Helpers.fn_ReadSnapshotLongAttr(root, "colStart", 1)
    layoutRowEnd = ex_Helpers.fn_ReadSnapshotLongAttr(root, "rowEnd", layoutRowStart)
    layoutColEnd = ex_Helpers.fn_ReadSnapshotLongAttr(root, "colEnd", layoutColStart)
    layoutStyle = VBA.Trim$(VBA.CStr(root.getAttribute("style")))
    isConfiguredAttr = VBA.LCase$(VBA.Trim$(VBA.CStr(root.getAttribute("isConfigured"))))

    If VBA.Len(m_ControlName) = 0 Then Exit Function
    If VBA.Len(m_OnClickMacroRef) = 0 Then Exit Function
    If VBA.Len(layoutSheetName) = 0 Then Exit Function
    If VBA.Len(m_CaptionRaw) = 0 Then m_CaptionRaw = m_CaptionText
    If VBA.Len(m_CaptionRaw) = 0 Then m_CaptionRaw = DEFAULT_CAPTION
    Set m_CaptionInlineTextPart = New obj_InlineTextPart
    If VBA.Len(m_CaptionInlineSource) = 0 Then
        m_CaptionInlineSource = m_CaptionRaw
        If VBA.InStr(1, m_CaptionInlineSource, "{Binding ", VBA.vbTextCompare) > 0 Then
            m_CaptionInlineSource = m_CaptionText
        End If
    End If
    If VBA.Len(m_RuntimeControlKey) = 0 Then
        m_RuntimeControlKey = "button|" & VBA.LCase$(VBA.Trim$(layoutSheetName & "|" & m_ControlName))
    End If

    If VBA.Len(shapeName) = 0 Then shapeName = "btn_" & m_ControlName

    ' 1) Восстанавливаем объект layout (лист/границы/style) из snapshot-атрибутов.
    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromRuntimeValues( _
        "Button", _
        m_ControlName, _
        layoutSheetName, _
        layoutRowStart, _
        layoutColStart, _
        layoutRowEnd, _
        layoutColEnd, _
        layoutStyle) Then Exit Function

    ' 2) Ищем worksheet по имени листа из snapshot.
    Set ws = ex_HelpersSheet.fn_GetRuntimeWorksheetByName(layoutSheetName)
    If ws Is Nothing Then Exit Function

    If Not private_TryRestoreCallbackContextFromPage() Then Exit Function

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Function

    If Not private_TryBindRuntimeRoute(shp) Then Exit Function

    ' В restore-фазе страница уже отрисована и inline-стили caption уже применены.
    ' Не накладываем inline-runs повторно из snapshot, иначе при рассинхроне
    ' snapshot/state (например закрытие книги без сохранения) получаем смешанную подсветку.

    If isConfiguredAttr = "false" Or isConfiguredAttr = "0" Then
        m_IsConfigured = False
    Else
        m_IsConfigured = True
    End If
    TryDeserializeSnapshot = True
End Function

Private Function private_TryRestoreCallbackContextFromPage() As Boolean
    Dim callbackContext As Object

    Set m_OnClickCallbackContext = Nothing
    If Not m_Page.TryGetController(callbackContext) Then Exit Function
    If callbackContext Is Nothing Then
        private_TryRestoreCallbackContextFromPage = True
        Exit Function
    End If

    Set m_OnClickCallbackContext = callbackContext
    private_TryRestoreCallbackContextFromPage = True
End Function

' //
' // Internal
' //
' Привязываем shape к централизованному обработчику клика
' и регистрируем routing в PageBase (shape -> control key -> method).
Private Function private_TryBindRuntimeRoute(ByVal shp As Shape) As Boolean
    Dim callbackMacroRef As String
    Dim pageBase As obj_PageBase

    If shp Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "button:bind-route shape-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
#End If
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_RuntimeControlKey)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "button:bind-route runtime-control-key-empty control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
#End If
        Exit Function
    End If

    callbackMacroRef = private_GetRuntimeCallbackMacroRef()
    If VBA.Len(callbackMacroRef) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "button:bind-route callback-macro-empty control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
#End If
        Exit Function
    End If

    If Not private_TryAssignShapeOnActionIfChanged(shp, callbackMacroRef) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "button:bind-route set-onaction-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(shp.Name), "'", "''") & "'"
#End If
        Exit Function
    End If

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    If Not pageBase.RegisterControl(m_RuntimeControlKey, Me) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "button:bind-route register-control-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' key='" & VBA.Replace$(VBA.Trim$(m_RuntimeControlKey), "'", "''") & "'"
#End If
        Exit Function
    End If
    If Not pageBase.RegisterShapeRoute(shp.Name, m_RuntimeControlKey, "RuntimeHandleClick", False) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "button:bind-route register-route-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(shp.Name), "'", "''") & "'"
#End If
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "button:bind-route ok control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(shp.Name), "'", "''") & "' macro='" & VBA.Replace$(callbackMacroRef, "'", "''") & "'"
#End If

    private_TryBindRuntimeRoute = True
End Function

Private Function private_RegisterCaptionInlineRuns(ByVal page As obj_PageBase, ByVal shp As Shape) As Boolean
    If page Is Nothing Then Exit Function
    If shp Is Nothing Then Exit Function
    If m_CaptionInlineTextPart Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: caption inline part is not initialized for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If
    private_RegisterCaptionInlineRuns = m_CaptionInlineTextPart.RegisterForShape(page, shp)
End Function

Private Function private_GetUiShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    If ws Is Nothing Then Exit Function
    shapeName = VBA.Trim$(shapeName)
    If VBA.Len(shapeName) = 0 Then Exit Function

    On Error Resume Next
    Set private_GetUiShapeByName = ws.Shapes(shapeName)
    On Error GoTo 0
End Function

Private Function private_TryAssignShapeOnActionIfChanged(ByVal shp As Shape, ByVal macroRef As String) As Boolean
    Dim currentMacroRef As String

    If shp Is Nothing Then Exit Function
    macroRef = VBA.Trim$(macroRef)
    If VBA.Len(macroRef) = 0 Then
        private_TryAssignShapeOnActionIfChanged = True
        Exit Function
    End If

    On Error Resume Next
    currentMacroRef = VBA.Trim$(VBA.CStr(shp.OnAction))
    If Err.Number <> 0 Then
        Err.Clear
        currentMacroRef = VBA.vbNullString
    End If
    On Error GoTo 0

    If VBA.StrComp(currentMacroRef, macroRef, VBA.vbBinaryCompare) <> 0 Then
        On Error GoTo EH_SET
        shp.OnAction = macroRef
        On Error GoTo 0
    End If

    private_TryAssignShapeOnActionIfChanged = True
    Exit Function

EH_SET:
    On Error GoTo 0
End Function

Private Function private_TryResolveCaptionInlineText( _
    ByVal page As obj_PageBase, _
    ByVal rawCaptionText As String _
) As Boolean
    Dim inlineTextProfile As obj_InlineTextProfile

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: page is not specified for inline caption resolve in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    If m_CaptionInlineTextPart Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Button: caption inline part is not initialized for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    ' Для caption используем тот же pipeline, что и для ViewItem:
    ' profile(part) -> inline part -> resolve -> register.
    If Not page.TryGetInlineTextProfile(INLINE_PART_BUTTON, inlineTextProfile) Then Exit Function
    Set m_CaptionInlineTextPart.InlineProfile = inlineTextProfile

    If Not m_CaptionInlineTextPart.Resolve(rawCaptionText) Then Exit Function
    m_CaptionText = m_CaptionInlineTextPart.ResolvedText
    private_TryResolveCaptionInlineText = True
End Function

Private Function private_GetRuntimeCallbackMacroRef() As String
    private_GetRuntimeCallbackMacroRef = private_QualifyMacroName("rt_Bridge.fn_OnShapeClick")
End Function

Private Function private_QualifyMacroName(ByVal macroName As String) As String
    Dim wbName As String

    macroName = VBA.Trim$(macroName)
    If VBA.Len(macroName) = 0 Then Exit Function
    If VBA.InStr(1, macroName, "!", VBA.vbBinaryCompare) > 0 Then
        private_QualifyMacroName = macroName
        Exit Function
    End If

    wbName = ThisWorkbook.Name
    wbName = VBA.Replace$(wbName, "'", "''")
    private_QualifyMacroName = "'" & wbName & "'!" & macroName
End Function

Public Property Get DefaultCaption() As String
    DefaultCaption = DEFAULT_CAPTION
End Property
