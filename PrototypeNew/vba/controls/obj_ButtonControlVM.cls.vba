VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ButtonControlVM"
Option Explicit
Implements obj_IControl
Implements obj_ISerializable

Private Const DEFAULT_CAPTION As String = "Update Code"

Private m_Base As obj_ControlBase
Private m_ControlName As String
Private m_CaptionRaw As String
Private m_OnClickRaw As String
Private m_Layout As obj_ControlLayout
Private m_CaptionText As String
Private m_OnClickMacroRef As String
Private m_RuntimeControlKey As String
Private m_IsConfigured As Boolean
Private m_Page As obj_PageBase

' //
' // Interface
' //
' Configure поднимает контракт контрола из XML:
' читаем attrs -> резолвим биндинги -> нормализуем layout -> готовим runtime key.
Private Sub obj_IControl_Configure(ByVal page As obj_PageBase, ByVal controlNode As Object)
    m_IsConfigured = False
    Set m_Layout = Nothing
    Set m_Base = Nothing
    Set m_Page = Nothing
    m_RuntimeControlKey = VBA.vbNullString

    Set m_Base = New obj_ControlBase
    If Not m_Base.Configure(page, controlNode, "Button", "button", m_ControlName) Then Exit Sub
    Set m_Page = m_Base.PageBase

    m_CaptionRaw = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "caption"))
    If VBA.Len(m_CaptionRaw) = 0 Then m_CaptionRaw = DEFAULT_CAPTION

    m_OnClickRaw = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "onClick"))
    If VBA.Len(VBA.Trim$(m_OnClickRaw)) = 0 Then
        VBA.MsgBox "Button: onClick is required for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If Not ex_BindingRuntime.m_TryResolveTextBinding(m_CaptionRaw, Me, m_CaptionText) Then Exit Sub
    If Not ex_BindingRuntime.m_TryResolveMacroBinding(m_OnClickRaw, Me, m_OnClickMacroRef) Then Exit Sub
    m_OnClickMacroRef = private_NormalizeLegacyOnClickMacroRef(m_OnClickMacroRef)

    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.TryReadFromNode(controlNode, "Button", m_ControlName, "style") Then Exit Sub
    m_RuntimeControlKey = "button|" & VBA.LCase$(VBA.Trim$(m_Layout.LayoutSheetName & "|" & m_ControlName))

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
    Dim page As obj_PageBase

    If Not m_IsConfigured Then
        VBA.MsgBox "Button: control '" & m_ControlName & "' is not configured.", VBA.vbExclamation
        Exit Sub
    End If

    Set page = Nothing
    If Not m_Base Is Nothing Then Set page = m_Base.PageBase
    If page Is Nothing Then Set page = m_Page
    If page Is Nothing Then
        VBA.MsgBox "Button: page is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If
    Set m_Page = page

    Set ws = page.Worksheet
    If ws Is Nothing Then
        VBA.MsgBox "Button: page worksheet is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(m_Layout.RowStart, m_Layout.ColStart), ws.Cells(m_Layout.RowEnd, m_Layout.ColEnd))
    On Error GoTo 0

    If targetRange Is Nothing Then
        VBA.MsgBox "Button: failed to resolve target range for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If
    If targetRange.Width <= 0# Or targetRange.Height <= 0# Then
        VBA.MsgBox "Button: target range has non-positive width/height for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    buttonName = "btn_" & m_ControlName

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo EH_BUTTON

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    shp.Name = buttonName
    shp.Placement = xlMoveAndSize
    If Not private_TryBindRuntimeRoute(shp) Then Exit Sub

    On Error Resume Next
    shp.TextFrame2.TextRange.Text = m_CaptionText
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame.Characters.Text = m_CaptionText
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter

    Set metaMap = VBA.CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.control") = m_ControlName
    If VBA.Len(VBA.Trim$(m_Layout.StyleName)) > 0 Then
        metaMap("pn.style") = m_Layout.StyleName
    Else
        metaMap("pn.style") = VBA.vbNullString
    End If
    If Not ex_ShapeMetaRuntime.m_TrySetShapeMetaValues(shp, metaMap) Then Exit Sub
    On Error GoTo EH_BUTTON

    Exit Sub

EH_BUTTON:
    VBA.MsgBox "Button: failed to render control '" & m_ControlName & "': " & Err.Description, VBA.vbExclamation
    Exit Sub

EH_RANGE:
    VBA.MsgBox "Button: failed to resolve target range for control '" & m_ControlName & "': " & Err.Description, VBA.vbExclamation
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "caption", "onclick"
            obj_IControl_SupportsAttribute = True
    End Select
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

' //
' // API
' //
Public Function RuntimeHandleClick() As Boolean
    If Not rt_Bridge.m_RunMacro(m_OnClickMacroRef) Then Exit Function
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
    If m_Layout Is Nothing Then Exit Function

    sheetName = VBA.Trim$(m_Layout.LayoutSheetName)
    If VBA.Len(sheetName) = 0 Then Exit Function
    shapeName = "btn_" & m_ControlName
    styleName = VBA.Trim$(m_Layout.StyleName)
    isConfiguredText = VBA.IIf(m_IsConfigured, "true", "false")

    outSnapshotXml = _
        "<button version=""2""" & _
        " controlName=""" & ex_Helpers.m_EscapeXmlAttr(m_ControlName) & """" & _
        " captionRaw=""" & ex_Helpers.m_EscapeXmlAttr(m_CaptionRaw) & """" & _
        " onClickRaw=""" & ex_Helpers.m_EscapeXmlAttr(m_OnClickRaw) & """" & _
        " captionText=""" & ex_Helpers.m_EscapeXmlAttr(m_CaptionText) & """" & _
        " onClickMacroRef=""" & ex_Helpers.m_EscapeXmlAttr(m_OnClickMacroRef) & """" & _
        " runtimeKey=""" & ex_Helpers.m_EscapeXmlAttr(m_RuntimeControlKey) & """" & _
        " sheet=""" & ex_Helpers.m_EscapeXmlAttr(sheetName) & """" & _
        " rowStart=""" & VBA.CStr(m_Layout.RowStart) & """" & _
        " colStart=""" & VBA.CStr(m_Layout.ColStart) & """" & _
        " rowEnd=""" & VBA.CStr(m_Layout.RowEnd) & """" & _
        " colEnd=""" & VBA.CStr(m_Layout.ColEnd) & """" & _
        " style=""" & ex_Helpers.m_EscapeXmlAttr(styleName) & """" & _
        " isConfigured=""" & isConfiguredText & """" & _
        " shape=""" & ex_Helpers.m_EscapeXmlAttr(shapeName) & """" & _
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

    If Not ex_CustomXmlPartStore.m_TryLoadDomFromXml(snapshotXml, dom) Then Exit Function
    Set root = dom.DocumentElement
    If root Is Nothing Then Exit Function
    If VBA.LCase$(VBA.CStr(root.baseName)) <> "button" Then Exit Function
    versionText = VBA.Trim$(VBA.CStr(root.getAttribute("version")))

    m_ControlName = VBA.Trim$(VBA.CStr(root.getAttribute("controlName")))
    m_CaptionRaw = VBA.CStr(root.getAttribute("captionRaw"))
    m_OnClickRaw = VBA.CStr(root.getAttribute("onClickRaw"))
    m_CaptionText = VBA.CStr(root.getAttribute("captionText"))
    onClickMacroRef = VBA.Trim$(VBA.CStr(root.getAttribute("onClickMacroRef")))
    If VBA.Len(onClickMacroRef) = 0 Then onClickMacroRef = VBA.Trim$(VBA.CStr(root.getAttribute("onClick")))
    m_OnClickMacroRef = private_NormalizeLegacyOnClickMacroRef(onClickMacroRef)
    m_RuntimeControlKey = VBA.LCase$(VBA.Trim$(VBA.CStr(root.getAttribute("runtimeKey"))))
    shapeName = VBA.Trim$(VBA.CStr(root.getAttribute("shape")))
    layoutSheetName = VBA.Trim$(VBA.CStr(root.getAttribute("sheet")))
    layoutRowStart = ex_Helpers.m_ReadSnapshotLongAttr(root, "rowStart", 1)
    layoutColStart = ex_Helpers.m_ReadSnapshotLongAttr(root, "colStart", 1)
    layoutRowEnd = ex_Helpers.m_ReadSnapshotLongAttr(root, "rowEnd", layoutRowStart)
    layoutColEnd = ex_Helpers.m_ReadSnapshotLongAttr(root, "colEnd", layoutColStart)
    layoutStyle = VBA.Trim$(VBA.CStr(root.getAttribute("style")))
    isConfiguredAttr = VBA.LCase$(VBA.Trim$(VBA.CStr(root.getAttribute("isConfigured"))))

    If VBA.Len(m_ControlName) = 0 Then Exit Function
    If VBA.Len(m_OnClickMacroRef) = 0 Then Exit Function
    If VBA.Len(layoutSheetName) = 0 Then Exit Function
    If VBA.Len(m_CaptionRaw) = 0 Then m_CaptionRaw = m_CaptionText
    If VBA.Len(m_CaptionRaw) = 0 Then m_CaptionRaw = DEFAULT_CAPTION
    If VBA.Len(m_CaptionText) = 0 Then m_CaptionText = m_CaptionRaw
    If VBA.Len(m_RuntimeControlKey) = 0 Then
        m_RuntimeControlKey = "button|" & VBA.LCase$(VBA.Trim$(layoutSheetName & "|" & m_ControlName))
    End If

    If VBA.Len(shapeName) = 0 Then shapeName = "btn_" & m_ControlName

    ' 1) Восстанавливаем объект layout (лист/границы/style) из snapshot-атрибутов.
    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.TryReadFromRuntimeValues( _
        "Button", _
        m_ControlName, _
        layoutSheetName, _
        layoutRowStart, _
        layoutColStart, _
        layoutRowEnd, _
        layoutColEnd, _
        layoutStyle) Then Exit Function

    ' 2) Ищем worksheet по имени листа из snapshot.
    Set ws = ex_HelpersSheet.m_GetRuntimeWorksheetByName(layoutSheetName)
    If ws Is Nothing Then Exit Function

    ' 3) Восстанавливаем page-контекст (PageBase) для этого листа,
    '    чтобы затем зарегистрировать control/route обратно в runtime-реестры страницы.
    If Not ex_HelpersSheet.m_TryGetPageBaseByWorksheetName(layoutSheetName, m_Page) Then Exit Function

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Function

    If Not private_TryBindRuntimeRoute(shp) Then Exit Function

    If isConfiguredAttr = "false" Or isConfiguredAttr = "0" Then
        m_IsConfigured = False
    Else
        m_IsConfigured = True
    End If
    TryDeserializeSnapshot = True
End Function

' //
' // Internal
' //
' Привязываем shape к централизованному обработчику клика
' и регистрируем routing в PageBase (shape -> control key -> method).
Private Function private_TryBindRuntimeRoute(ByVal shp As Shape) As Boolean
    Dim callbackMacroRef As String

    If shp Is Nothing Then
        ex_Core.m_LogError "button:bind-route shape-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_RuntimeControlKey)) = 0 Then
        ex_Core.m_LogError "button:bind-route runtime-control-key-empty control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
        Exit Function
    End If

    callbackMacroRef = private_GetRuntimeCallbackMacroRef()
    If VBA.Len(callbackMacroRef) = 0 Then
        ex_Core.m_LogError "button:bind-route callback-macro-empty control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
        Exit Function
    End If

    On Error Resume Next
    shp.OnAction = callbackMacroRef
    If Err.Number <> 0 Then
        ex_Core.m_LogError "button:bind-route set-onaction-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(shp.Name), "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If m_Page Is Nothing Then
        ex_Core.m_LogError "button:bind-route page-base-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
        Exit Function
    End If
    If Not m_Page.RegisterControl(m_RuntimeControlKey, Me) Then
        ex_Core.m_LogError "button:bind-route register-control-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' key='" & VBA.Replace$(VBA.Trim$(m_RuntimeControlKey), "'", "''") & "'"
        Exit Function
    End If
    If Not m_Page.RegisterShapeRoute(shp.Name, m_RuntimeControlKey, "RuntimeHandleClick", False) Then
        ex_Core.m_LogError "button:bind-route register-route-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(shp.Name), "'", "''") & "'"
        Exit Function
    End If

    ex_Core.m_LogInfo "button:bind-route ok control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(shp.Name), "'", "''") & "' macro='" & VBA.Replace$(callbackMacroRef, "'", "''") & "'"

    private_TryBindRuntimeRoute = True
End Function

Private Function private_GetRuntimeCallbackMacroRef() As String
    private_GetRuntimeCallbackMacroRef = private_QualifyMacroName("rt_Bridge.m_OnShapeClick")
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


Private Function private_NormalizeLegacyOnClickMacroRef(ByVal macroRef As String) As String
    Dim bangPos As Long
    Dim workbookPrefix As String
    Dim methodName As String

    macroRef = VBA.Trim$(macroRef)
    If VBA.Len(macroRef) = 0 Then Exit Function

    bangPos = VBA.InStr(1, macroRef, "!", VBA.vbBinaryCompare)
    If bangPos > 0 Then
        workbookPrefix = VBA.Left$(macroRef, bangPos)
        methodName = VBA.Mid$(macroRef, bangPos + 1)
    Else
        workbookPrefix = VBA.vbNullString
        methodName = macroRef
    End If

    Select Case VBA.LCase$(VBA.Trim$(methodName))
        Case "rt_coreactions.m_updatecodefullandrerender", "ex_core.dev_updateallmodulesunsafe"
            methodName = "ex_Core.dev_UpdateAllModules"
        Case "rt_coreactions.m_updatecodedateandrerender", "ex_core.dev_updatecodebydateunsafe"
            methodName = "ex_Core.dev_UpdateCodeByDate"
        Case "ex_core.dev_updatecodebysizeunsafe"
            methodName = "ex_Core.dev_UpdateCodeBySize"
        Case "rt_coreactions.m_updatecodesizeandrerender"
            methodName = "ex_Core.dev_UpdateCodeBySize"
    End Select

    If VBA.Len(workbookPrefix) > 0 Then
        private_NormalizeLegacyOnClickMacroRef = workbookPrefix & methodName
    Else
        private_NormalizeLegacyOnClickMacroRef = methodName
    End If
End Function

Public Property Get DefaultCaption() As String
    DefaultCaption = DEFAULT_CAPTION
End Property
