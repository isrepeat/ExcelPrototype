VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_EntityLookupCandidatesControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const DEFAULT_CANDIDATES_AT As String = "r1c1"

Implements obj_IControl

Private m_Page As obj_IPage
Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_TableList As obj_IControl
Private m_IsConfigured As Boolean
Private m_IsDisposed As Boolean

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
    m_IsDisposed = False
    m_IsConfigured = False
    Set m_Page = page
    obj_IControl_Initialize = True
End Function

Private Sub obj_IControl_Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_TableList Is Nothing Then m_TableList.Dispose
    If Not m_ControlBase Is Nothing Then m_ControlBase.Dispose
    Set m_TableList = Nothing
    Set m_ControlBase = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim pageBase As obj_PageBase
    Dim tableNode As Object
    Dim adjustedRowStart As Long
    Dim adjustedColStart As Long
    Dim adjustedRowEnd As Long
    Dim adjustedColEnd As Long

    m_IsConfigured = False
    Set m_TableList = Nothing
    Set m_ControlBase = Nothing

    If m_Page Is Nothing Then Exit Sub
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Sub

    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "EntityLookupCandidates", "entitylookupcandidates", m_ControlName) Then Exit Sub

    ' XML задает только внешний слот кандидатов. Сам контрол может сдвинуть
    ' внутренний TableList внутри этого слота так, чтобы search column совпала
    ' с колонкой активного lookup input.
    If Not private_TryResolveAdjustedBounds( _
        controlNode, _
        pageBase, _
        adjustedRowStart, _
        adjustedColStart, _
        adjustedRowEnd, _
        adjustedColEnd) Then Exit Sub

    Set tableNode = controlNode.cloneNode(True)
    If tableNode Is Nothing Then Exit Sub

    ' Реальный рендер, стили и itemVisibility уже умеет TableList.
    ' Этот контрол только пересчитывает runtime bounds и делегирует дальше.
    tableNode.setAttribute "type", "TableList"
    If Not private_TrySetLayoutLongAttr(tableNode, "__layoutRowStart", adjustedRowStart) Then Exit Sub
    If Not private_TrySetLayoutLongAttr(tableNode, "__layoutColStart", adjustedColStart) Then Exit Sub
    If Not private_TrySetLayoutLongAttr(tableNode, "__layoutRowEnd", adjustedRowEnd) Then Exit Sub
    If Not private_TrySetLayoutLongAttr(tableNode, "__layoutColEnd", adjustedColEnd) Then Exit Sub

    Set m_TableList = New obj_TableListControlVM
    If Not m_TableList.Initialize(m_Page) Then Exit Sub
    m_TableList.Configure tableNode
    If Not m_TableList.IsConfigured() Then Exit Sub

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "EntityLookupCandidates: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If
    If m_TableList Is Nothing Then Exit Sub

    m_TableList.Render
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource", "itemvisibility", "lookupfeature"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

' //
' // Internal
' //
Private Function private_TryResolveAdjustedBounds( _
    ByVal controlNode As Object, _
    ByVal pageBase As obj_PageBase, _
    ByRef outRowStart As Long, _
    ByRef outColStart As Long, _
    ByRef outRowEnd As Long, _
    ByRef outColEnd As Long _
) As Boolean
    Dim lookupFeature As obj_EntityLookupFeature
    Dim cfgParser As obj_EntityLookupCfgParser
    Dim lookupKey As String
    Dim candidateTable As obj_TableDynamic
    Dim searchColumnAlias As String
    Dim candidateAt As String
    Dim relativeRow As Long
    Dim relativeCol As Long
    Dim rowSpan As Long
    Dim colSpan As Long
    Dim candidateSpanCols As Long

    If controlNode Is Nothing Then Exit Function
    If pageBase Is Nothing Then Exit Function

    ' Эти bounds уже посчитаны XML layout engine-ом из обычных spanRows/spanColls.
    ' Для EntityLookupCandidates это "контейнер", а не финальная позиция таблицы.
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", outRowStart, True) Then Exit Function
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColStart", outColStart, True) Then Exit Function
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", outRowEnd, True) Then Exit Function
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", outColEnd, True) Then Exit Function

    rowSpan = outRowEnd - outRowStart + 1
    colSpan = outColEnd - outColStart + 1
    If rowSpan <= 0 Or colSpan <= 0 Then Exit Function

    ' lookupFeature хранит текущий активный lookup и таблицу кандидатов.
    ' Если активного поиска нет, TableList остается в начале контейнера.
    Set lookupFeature = Nothing
    If Not private_TryResolveLookupFeature(controlNode, pageBase, lookupFeature) Then Exit Function
    If lookupFeature Is Nothing Then
        private_TryResolveAdjustedBounds = True
        Exit Function
    End If

    ' Контекст появляется только после SearchCandidates:
    ' lookupKey - какой input искал пользователь,
    ' candidateTable - результат SQL,
    ' searchColumnAlias - колонка результата, которую надо выровнять под input.
    Set cfgParser = Nothing
    Set candidateTable = Nothing
    If Not lookupFeature.TryGetActiveCandidatesContext(cfgParser, lookupKey, candidateTable, searchColumnAlias) Then
        private_TryResolveAdjustedBounds = True
        Exit Function
    End If

    ' Старый ActiveCandidatesAt теперь вычисляется локально здесь. PresentationState
    ' возвращает относительный адрес внутри контейнера, например r1c2.
    candidateSpanCols = colSpan
    If candidateSpanCols <= 0 Then candidateSpanCols = 1
    If Not private_TryBuildCandidateAt(cfgParser, lookupKey, candidateTable, searchColumnAlias, candidateSpanCols, candidateAt) Then Exit Function
    If Not private_TryParseAtAddress(candidateAt, relativeRow, relativeCol) Then Exit Function

    If relativeRow <= 0 Then relativeRow = 1
    If relativeCol <= 0 Then relativeCol = 1

    ' Переносим относительный rNcM в абсолютные координаты листа. Именно эти
    ' координаты затем увидит внутренний TableList в своих __layout* атрибутах.
    outRowStart = outRowStart + relativeRow - 1
    outColStart = outColStart + relativeCol - 1
    outRowEnd = outRowStart + rowSpan - 1
    outColEnd = outColStart + colSpan - 1

    private_TryResolveAdjustedBounds = True
End Function

Private Function private_TryResolveLookupFeature( _
    ByVal controlNode As Object, _
    ByVal pageBase As obj_PageBase, _
    ByRef outLookupFeature As obj_EntityLookupFeature _
) As Boolean
    Dim lookupFeatureRaw As String
    Dim resolvedObject As Object
    Dim dataContext As Object

    Set outLookupFeature = Nothing
    If controlNode Is Nothing Then Exit Function
    If pageBase Is Nothing Then Exit Function

    lookupFeatureRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "lookupFeature")))
    If VBA.Len(lookupFeatureRaw) > 0 Then
        ' Предпочтительный путь: XML привязывает lookupFeature к свойству
        ' dataContext-а, обычно controller.LookupFeature. Если передан runtime
        ' object source, он тоже поддерживается для более прямых сценариев.
        If Not private_TryResolveLookupFeatureRaw(lookupFeatureRaw, pageBase, resolvedObject) Then Exit Function
        If resolvedObject Is Nothing Then Exit Function
        If Not TypeOf resolvedObject Is obj_EntityLookupFeature Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "EntityLookupCandidates: lookupFeature for control '" & m_ControlName & "' resolved to unexpected type '" & VBA.TypeName(resolvedObject) & "'."
#End If
            Exit Function
        End If

        Set outLookupFeature = resolvedObject
        private_TryResolveLookupFeature = True
        Exit Function
    End If

    Set dataContext = Nothing
    If Not m_ControlBase Is Nothing Then Set dataContext = m_ControlBase.DataContext
    ' Fallback для простых сценариев: сам dataContext может быть feature-ом.
    ' В текущих страницах используется явный lookupFeature.
    If dataContext Is Nothing Then
        private_TryResolveLookupFeature = True
        Exit Function
    End If
    If Not TypeOf dataContext Is obj_EntityLookupFeature Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "EntityLookupCandidates: dataContext for control '" & m_ControlName & "' is not obj_EntityLookupFeature and lookupFeature is not specified."
#End If
        Exit Function
    End If

    Set outLookupFeature = dataContext
    private_TryResolveLookupFeature = True
End Function

Private Function private_TryResolveLookupFeatureRaw( _
    ByVal lookupFeatureRaw As String, _
    ByVal pageBase As obj_PageBase, _
    ByRef outResolvedObject As Object _
) As Boolean
    Dim dataContext As Object
    Dim resolvedValue As Variant

    Set outResolvedObject = Nothing
    lookupFeatureRaw = VBA.Trim$(lookupFeatureRaw)
    If VBA.Len(lookupFeatureRaw) = 0 Then
        private_TryResolveLookupFeatureRaw = True
        Exit Function
    End If

    Set dataContext = Nothing
    If Not m_ControlBase Is Nothing Then Set dataContext = m_ControlBase.DataContext
    If Not dataContext Is Nothing Then
        If Not ex_BindingRuntime.fn_TryResolveValueBinding(lookupFeatureRaw, dataContext, resolvedValue) Then Exit Function
        If VBA.IsObject(resolvedValue) Then
            Set outResolvedObject = resolvedValue
            private_TryResolveLookupFeatureRaw = True
            Exit Function
        End If
    End If

    If pageBase Is Nothing Then Exit Function
    If pageBase.RuntimeSources Is Nothing Then Exit Function
    If Not ex_RuntimeSourceResolver.fn_TryResolveObjectSource(pageBase.RuntimeSources, lookupFeatureRaw, outResolvedObject, False) Then Exit Function

    private_TryResolveLookupFeatureRaw = True
End Function

Private Function private_TryBuildCandidateAt( _
    ByVal cfgParser As obj_EntityLookupCfgParser, _
    ByVal lookupKey As String, _
    ByVal candidateTable As obj_TableDynamic, _
    ByVal searchColumnAlias As String, _
    ByVal candidateSpanCols As Long, _
    ByRef outCandidateAt As String _
) As Boolean
    Dim presentationState As obj_EntityLookupPresentationState

    outCandidateAt = DEFAULT_CANDIDATES_AT
    Set presentationState = New obj_EntityLookupPresentationState
    If Not presentationState.Initialize(candidateSpanCols) Then Exit Function
    ' Формула внутри PresentationState:
    ' startGridCol = inputGridCol - searchColumnIndex + 1.
    ' Так search column из candidateTable встает под lookup input.
    If Not presentationState.TrySetActiveCandidateLayout(cfgParser, lookupKey, candidateTable, searchColumnAlias) Then Exit Function

    outCandidateAt = presentationState.ActiveCandidatesAt
    presentationState.Dispose
    private_TryBuildCandidateAt = True
End Function

Private Function private_TryParseAtAddress( _
    ByVal atText As String, _
    ByRef outRow As Long, _
    ByRef outCol As Long _
) As Boolean
    Dim cPos As Long
    Dim rowText As String
    Dim colText As String

    outRow = 0
    outCol = 0
    atText = VBA.LCase$(VBA.Trim$(atText))
    If VBA.Len(atText) < 4 Then Exit Function
    If VBA.Left$(atText, 1) <> "r" Then Exit Function

    cPos = VBA.InStr(2, atText, "c", VBA.vbBinaryCompare)
    If cPos <= 2 Then Exit Function

    rowText = VBA.Mid$(atText, 2, cPos - 2)
    colText = VBA.Mid$(atText, cPos + 1)
    If Not VBA.IsNumeric(rowText) Then Exit Function
    If Not VBA.IsNumeric(colText) Then Exit Function

    outRow = VBA.CLng(rowText)
    outCol = VBA.CLng(colText)
    private_TryParseAtAddress = (outRow > 0 And outCol > 0)
End Function

Private Function private_TryReadLayoutLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByRef outValue As Long, _
    ByVal isRequired As Boolean _
) As Boolean
    Dim rawText As String

    rawText = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(controlNode, attrName))
    If VBA.Len(rawText) = 0 Then
        If isRequired Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "EntityLookupCandidates: runtime layout attribute '" & attrName & "' is missing for control '" & m_ControlName & "'."
#End If
            Exit Function
        End If
        outValue = 0
        private_TryReadLayoutLongAttr = True
        Exit Function
    End If
    If Not VBA.IsNumeric(rawText) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "EntityLookupCandidates: runtime layout attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    outValue = VBA.CLng(rawText)
    private_TryReadLayoutLongAttr = True
End Function

Private Function private_TrySetLayoutLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByVal attrValue As Long _
) As Boolean
    If controlNode Is Nothing Then Exit Function
    If attrValue <= 0 Then Exit Function

    controlNode.setAttribute attrName, VBA.CStr(attrValue)
    private_TrySetLayoutLongAttr = True
End Function
