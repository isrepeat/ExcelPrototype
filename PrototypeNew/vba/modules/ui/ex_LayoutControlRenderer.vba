Attribute VB_Name = "ex_LayoutControlRenderer"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

' Рендерер узлов <control>.
' Поток:
' 1) Берем декларативный узел <control> из страницы.
' 2) Находим/грузим XML-шаблон контрола по его type (obj_<Type>ControlUI.xml).
' 3) Копируем в runtime-узел только валидные атрибуты по контракту контрола.
' 4) Дописываем служебные runtime-атрибуты __layout* (фактическая геометрия на листе).
' 5) Передаем узел в VM контрола (Configure/Render), затем рендерим template children.

Private Const UI_NS As String = "urn:excelprototype:profiles"
' UI-шаблоны контролов теперь лежат в пронумерованной папке.
' Если структура `PrototypeNew/vba` снова изменится, обновить только этот базовый путь.
Private Const CONTROL_UI_BASE_REL_PATH As String = "vba\[4] controls\"
Private Const CONTROL_UI_FILE_PREFIX As String = "obj_"
Private Const CONTROL_UI_FILE_SUFFIX As String = "ControlUI.xml"

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_LayoutControlRenderer.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function fn_Render( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal layoutNode As Object, _
    Optional ByVal rowStart As Long = 0, _
    Optional ByVal colStart As Long = 0, _
    Optional ByVal rowEnd As Long = 0, _
    Optional ByVal colEnd As Long = 0 _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim layoutControlName As String
    Dim controlType As String
    Dim typeRoot As String
    Dim controlUiRelPath As String
    Dim runtimeControlNode As Object
    Dim control As obj_IControl
    Dim pageUiPath As String
    Dim page As obj_IPage
    Dim pageBase As obj_PageBase

    If renderCtx Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: render context is not specified for control render."
#End If
        Exit Function
    End If
    If layoutNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: control node is not specified."
#End If
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.CStr(layoutNode.baseName)), "control", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: ex_LayoutControlRenderer supports only <control> nodes."
#End If
        Exit Function
    End If
    If Not private_TryGetPageRenderContext(renderCtx, wb, ws) Then Exit Function
    Set page = renderCtx.Page
    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page base is not specified in render context."
#End If
        Exit Function
    End If

    ' name/type берем из layout-узла страницы.
    layoutControlName = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(layoutNode, "name"))
    controlType = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(layoutNode, "type"))
    typeRoot = private_NormalizeTypeRoot(controlType)

    If VBA.Len(layoutControlName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page control is missing required attribute 'name'."
#End If
        Exit Function
    End If
    If VBA.Len(controlType) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page control '" & layoutControlName & "' is missing required attribute 'type'."
#End If
        Exit Function
    End If
    If VBA.Len(typeRoot) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page control '" & layoutControlName & "' has invalid type '" & controlType & "'."
#End If
        Exit Function
    End If

    ' Создаем VM контрола по типу (Button, Select, Table...),
    ' сразу прокидывая page-контекст для классов, которые хранят m_Page.
    Set control = ex_ControlFactory.fn_CreateControlByTypeRoot(typeRoot, page)
    If control Is Nothing Then Exit Function

    ' Грузим XML-шаблон контрола и применяем overrides из page layout.
    controlUiRelPath = private_ResolveControlUiRelPathByTypeRoot(typeRoot)
    Set runtimeControlNode = private_LoadControlNodeFromControlUi( _
        wb, controlUiRelPath, layoutNode, control, layoutControlName, typeRoot)
    If runtimeControlNode Is Nothing Then Exit Function

    ' Назначаем служебные runtime-границы (лист + координаты размещения).
    ' Эти атрибуты не пользовательские, они нужны VM в рантайме.
    private_ApplyRuntimeLayoutBounds runtimeControlNode, ws.Name, rowStart, colStart, rowEnd, colEnd

    pageUiPath = VBA.Trim$(pageBase.UiPath)

    ' Регистрируем текущие bounds в runtime-реестре, чтобы refresh-пайплайн
    ' умел переиспользовать координаты и переотрисовывать контрол адресно.
    ex_ControlRefreshRuntime.fn_RegisterControlRenderBounds _
        layoutControlName, _
        typeRoot, _
        ws.Name, _
        pageUiPath, _
        rowStart, _
        colStart, _
        rowEnd, _
        colEnd

    ex_LayoutDebugBoundsRndr.fn_RegisterDebugBounds ws, rowStart, colStart, rowEnd, colEnd, "control", layoutControlName

    ' Передаем итоговый runtime-узел в VM и запускаем render контрола.
    ' На этом шаге VM читает dataContext/objectSource/itemsSource и запускает resolve через RuntimeSourceResolver.
    On Error GoTo EH_CONTROL_PIPELINE
    control.Configure runtimeControlNode
    If Not control.IsConfigured() Then
        ex_LayoutControlFallbackRndr.fn_RegisterControlFallback ws, rowStart, colStart, rowEnd, colEnd, layoutControlName
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogWarning "PrototypeNew: control '" & layoutControlName & "' (type='" & typeRoot & "') is not configured. Fallback was applied."
    #End If
        fn_Render = True
        Exit Function
    End If

    control.Render
    On Error GoTo 0

    ' Если в шаблоне контрола есть дочерний layout (template children),
    ' рендерим его в тех же границах.
    If Not ex_XmlLayoutEngine.fn_RenderTemplateChildren( _
        renderCtx, runtimeControlNode, _
        rowStart, colStart, rowEnd, colEnd) Then Exit Function

    fn_Render = True
    Exit Function

EH_CONTROL_PIPELINE:
    ex_LayoutControlFallbackRndr.fn_RegisterControlFallback ws, rowStart, colStart, rowEnd, colEnd, layoutControlName
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: control '" & layoutControlName & "' (type='" & typeRoot & "') render pipeline failed. Fallback was applied: " & Err.Description
#End If
    fn_Render = True
End Function

' //
' // Internal
' //
Private Function private_LoadControlNodeFromControlUi( _
    ByVal wb As Workbook, _
    ByVal controlUiRelPath As String, _
    ByVal layoutControlNode As Object, _
    ByVal control As obj_IControl, _
    ByVal controlName As String, _
    ByVal typeRoot As String _
) As Object
    Dim uiDoc As Object
    Dim escapedName As String
    Dim xPath As String

    Set uiDoc = ex_XmlCore.fn_LoadDomByRelativePath( _
        wb, _
        controlUiRelPath, _
        "PrototypeNew: control UI file was not found: ", _
        "PrototypeNew: failed to parse control UI file: ", _
        UI_NS)
    If uiDoc Is Nothing Then Exit Function

    escapedName = ex_XmlCore.fn_XPathLiteral(controlName)
    xPath = "/p:uiDefinition/p:layout//p:control[@name=" & escapedName & "]"
    Set private_LoadControlNodeFromControlUi = uiDoc.selectSingleNode(xPath)

    ' Сначала пытаемся взять control с нужным именем.
    ' Если нет — используем первый <control> как дефолтный шаблон.
    If private_LoadControlNodeFromControlUi Is Nothing Then
        Set private_LoadControlNodeFromControlUi = uiDoc.selectSingleNode("/p:uiDefinition/p:layout//p:control[1]")
    End If

    If private_LoadControlNodeFromControlUi Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: control template has no <control> node in UI file '" & controlUiRelPath & "'."
#End If
        Exit Function
    End If

    ' Накладываем атрибуты из page layout на template control по контракту:
    ' только поддерживаемые control-атрибуты, без layout-атрибутов позиции.
    ' Сюда попадают и source-атрибуты (dataContext/itemsSource/objectSource) как исходный текст.
    ' Их фактическое разрешение происходит позже в VM в runtime-контексте страницы.
    If Not private_ApplyLayoutControlOverridesByContract( _
        private_LoadControlNodeFromControlUi, layoutControlNode, control, controlName, typeRoot) Then
        Set private_LoadControlNodeFromControlUi = Nothing
        Exit Function
    End If

    ' Финальное имя всегда берем из page layout (единый источник имени).
    On Error Resume Next
    private_LoadControlNodeFromControlUi.setAttribute "name", controlName
    On Error GoTo 0
End Function


Private Function private_ApplyLayoutControlOverridesByContract( _
    ByVal runtimeControlNode As Object, _
    ByVal layoutControlNode As Object, _
    ByVal control As obj_IControl, _
    ByVal controlName As String, _
    ByVal typeRoot As String _
) As Boolean
    Dim layoutAttrs As Object
    Dim attrNode As Object
    Dim attrName As String

    If runtimeControlNode Is Nothing Then Exit Function
    If layoutControlNode Is Nothing Then Exit Function
    If control Is Nothing Then Exit Function

    Set layoutAttrs = layoutControlNode.selectNodes("@*")
    If layoutAttrs Is Nothing Then Exit Function

    ' Переносим только "бизнес"-атрибуты контрола.
    ' Layout-атрибуты (at/span*/visibility) игнорируются здесь, т.к. они уже
    ' отработаны layout-рендерерами и не должны переписывать runtime bounds.
    For Each attrNode In layoutAttrs
        attrName = VBA.CStr(attrNode.nodeName)

        If private_IsLayoutAttribute(attrName) Then GoTo ContinueLoop

        ' Строгая валидация: атрибут должен входить в контракт конкретного VM.
        If Not ex_ControlAttributeContracts.fn_IsSupportedControlAttribute(control, attrName) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: attribute '" & attrName & "' is not supported by control '" & controlName & "' of type '" & typeRoot & "'."
#End If
            Exit Function
        End If

        On Error Resume Next
        runtimeControlNode.setAttribute attrName, VBA.CStr(attrNode.Text)
        If Err.Number <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to apply attribute '" & attrName & "' to control '" & controlName & "': " & Err.Description
#End If
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0

ContinueLoop:
    Next attrNode

    private_ApplyLayoutControlOverridesByContract = True
End Function


Private Function private_IsLayoutAttribute(ByVal attrName As String) As Boolean
    ' Атрибуты раскладки страницы. Они управляют размещением в grid/stack/list,
    ' но не являются "настройками VM контрола".
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "at", "spancolls", "spanrows", "visibility"
            private_IsLayoutAttribute = True
    End Select
End Function


Private Sub private_ApplyRuntimeLayoutBounds( _
    ByVal runtimeControlNode As Object, _
    ByVal sheetName As String, _
    ByVal layoutRowStart As Long, _
    ByVal layoutColStart As Long, _
    ByVal layoutRowEnd As Long, _
    ByVal layoutColEnd As Long _
) 
    ' Служебные runtime-атрибуты геометрии контрола:
    ' __layoutSheetName          - имя листа рендера
    ' __layoutRowStart/End   - верхняя/нижняя строка диапазона
    ' __layoutColStart/End   - левая/правая колонка диапазона
    ' Эти значения читаются VM/refresh-модулями после рендера и при восстановлении.
    runtimeControlNode.setAttribute "__layoutSheetName", sheetName

    If layoutRowStart > 0 Then runtimeControlNode.setAttribute "__layoutRowStart", VBA.CStr(layoutRowStart)
    If layoutColStart > 0 Then runtimeControlNode.setAttribute "__layoutColStart", VBA.CStr(layoutColStart)
    If layoutRowEnd > 0 Then runtimeControlNode.setAttribute "__layoutRowEnd", VBA.CStr(layoutRowEnd)
    If layoutColEnd > 0 Then runtimeControlNode.setAttribute "__layoutColEnd", VBA.CStr(layoutColEnd)
End Sub


Private Function private_NormalizeTypeRoot(ByVal controlType As String) As String
    private_NormalizeTypeRoot = VBA.Trim$(controlType)
End Function


Private Function private_ResolveControlUiRelPathByTypeRoot(ByVal typeRoot As String) As String
    private_ResolveControlUiRelPathByTypeRoot = _
        CONTROL_UI_BASE_REL_PATH & CONTROL_UI_FILE_PREFIX & typeRoot & CONTROL_UI_FILE_SUFFIX
End Function


Private Function private_TryGetPageRenderContext( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByRef outWb As Workbook, _
    ByRef outWs As Worksheet _
) As Boolean
    Set outWb = Nothing
    Set outWs = Nothing

    If renderCtx Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: render context is not specified."
#End If
        Exit Function
    End If

    Set outWs = renderCtx.Worksheet
    If outWs Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified."
#End If
        Exit Function
    End If

    Set outWb = renderCtx.Workbook
    If outWb Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: workbook is not specified."
#End If
        Exit Function
    End If

    private_TryGetPageRenderContext = True
End Function
