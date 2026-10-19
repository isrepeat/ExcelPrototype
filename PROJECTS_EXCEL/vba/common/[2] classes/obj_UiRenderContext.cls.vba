VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiRenderContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const ENABLE_LOGGING_VERBOSE = False

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_rendering As Boolean
Private m_hasRendered As Boolean
Private m_needsMeasure As Boolean
Private m_rows As Long
Private m_columns As Long
Private m_styles As obj_UiStyleCatalog
Private m_root As obj_IUiElement
Private m_router As obj_UiEventRouter
Private m_forms As Object
Private m_elements As Object
Private m_tablePlans As Collection
Private m_shapeLayouts As Object
Private m_targetWorksheet As Worksheet
Private m_uiPageDefinition As obj_UiPageDefinition
Private m_uiFolderPath As String
Private m_uiBindingContext As obj_UiBindingContext
Private m_dirty As Object
Private m_pending As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get Styles() As obj_UiStyleCatalog
    Set Styles = m_styles
End Property

Public Property Get Router() As obj_UiEventRouter
    Set Router = m_router
End Property

Public Property Get TargetWorksheet() As Worksheet
    Set TargetWorksheet = m_targetWorksheet
End Property

Public Property Get PageDefinition() As obj_UiPageDefinition
    Set PageDefinition = m_uiPageDefinition
End Property

Public Property Get BindingContext() As obj_UiBindingContext
    Set BindingContext = m_uiBindingContext
End Property

Public Property Get UiFolderPath() As String
    UiFolderPath = m_uiFolderPath
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal targetWorksheet As Worksheet, _
    ByVal uiPageDefinition As obj_UiPageDefinition, _
    ByVal uiFolderPath As String, _
    ByVal uiBindingContext As obj_UiBindingContext _
) As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If targetWorksheet Is Nothing Or uiPageDefinition Is Nothing Or uiBindingContext Is Nothing Then
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(uiFolderPath)) = 0 Then
        Exit Function
    End If
    Set m_targetWorksheet = targetWorksheet
    Set m_uiPageDefinition = uiPageDefinition
    m_uiFolderPath = uiFolderPath
    Set m_uiBindingContext = uiBindingContext
    Set m_styles = New obj_UiStyleCatalog
    If Not m_styles.Initialize() Then
        Exit Function
    End If
    Set m_router = New obj_UiEventRouter
    m_router.Initialize
    Set m_forms = VBA.CreateObject("Scripting.Dictionary")
    m_forms.CompareMode = VBA.vbTextCompare
    Set m_elements = VBA.CreateObject("Scripting.Dictionary")
    m_elements.CompareMode = VBA.vbTextCompare
    Set m_dirty = VBA.CreateObject("Scripting.Dictionary")
    m_dirty.CompareMode = VBA.vbTextCompare
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    Dim key As Variant
    Dim element As obj_IUiElement

    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False

#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_CONTEXT_DISPOSE_STARTED"
#End If
    If Not m_router Is Nothing Then
#If ENABLE_LOGGING_VERBOSE Then
        ex_Core.fn_Diagnostic_WriteLog "UI_ROUTER_DISPOSE_STARTED"
#End If
        m_router.Dispose
#If ENABLE_LOGGING_VERBOSE Then
        ex_Core.fn_Diagnostic_WriteLog "UI_ROUTER_DISPOSE_COMPLETED"
#End If
    End If
    If Not m_root Is Nothing Then
#If ENABLE_LOGGING_VERBOSE Then
        ex_Core.fn_Diagnostic_WriteLog "UI_ROOT_DISPOSE_STARTED"
#End If
        m_root.Dispose
#If ENABLE_LOGGING_VERBOSE Then
        ex_Core.fn_Diagnostic_WriteLog "UI_ROOT_DISPOSE_COMPLETED"
#End If
    End If
    If Not m_elements Is Nothing Then
        For Each key In m_elements.Keys
            If m_elements.Exists(key) Then
                Set element = m_elements(key)
                element.Dispose
            End If
        Next key
    End If
    If Not m_styles Is Nothing Then
        m_styles.Dispose
    End If
    Set m_styles = Nothing
    Set m_root = Nothing
    Set m_router = Nothing
    Set m_forms = Nothing
    Set m_elements = Nothing
    Set m_dirty = Nothing
    Set m_pending = Nothing
    Set m_tablePlans = Nothing
    Set m_shapeLayouts = Nothing
    Set m_uiBindingContext = Nothing
    If Not m_uiPageDefinition Is Nothing Then
        m_uiPageDefinition.Dispose
    End If
    Set m_uiPageDefinition = Nothing
    Set m_targetWorksheet = Nothing
    m_uiFolderPath = VBA.vbNullString
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_CONTEXT_DISPOSE_COMPLETED"
#End If
End Sub

Public Sub RegisterElement(ByVal name As String, ByVal element As obj_IUiElement)
    If VBA.Len(name) = 0 Then Exit Sub
    If m_elements.Exists(name) Then VBA.Err.Raise VBA.vbObjectError + 2232, , "Duplicate element: " & name
    Set m_elements(name) = element
End Sub

Public Function TryGetItemArea( _
    ByVal name As String, _
    ByVal index As Long, _
    ByRef area As Range _
) As Boolean
    Dim registered As obj_IUiElement
    Dim element As obj_UiControlElement

    Set area = Nothing
    If m_elements Is Nothing Then Exit Function
    If Not m_elements.Exists(name) Then Exit Function
    Set registered = m_elements(name)
    If Not TypeOf registered Is obj_UiControlElement Then Exit Function
    Set element = registered
    TryGetItemArea = element.TryGetItemArea(index, area)
End Function

Public Sub UnregisterElement(ByVal name As String, ByVal element As obj_IUiElement)
    Dim registered As obj_IUiElement

    If m_elements Is Nothing Then Exit Sub
    If Not m_elements.Exists(name) Then Exit Sub
    Set registered = m_elements(name)
    If Not registered Is element Then Exit Sub
    m_elements.Remove name
    If Not m_dirty Is Nothing Then
        If m_dirty.Exists(name) Then m_dirty.Remove name
    End If
    If Not m_pending Is Nothing Then
        If m_pending.Exists(name) Then m_pending.Remove name
    End If
    m_styles.ForgetControl name
End Sub

Public Function TryRefreshTableRow( _
    ByVal name As String, _
    ByVal rowIndex As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim registered As obj_IUiElement
    Dim element As obj_UiControlElement
    Dim previousEvents As Boolean
    Dim previousScreen As Boolean

    If m_isDisposed Or rowIndex < 1 Then
        diagnostic = "Invalid table row refresh request."
        Exit Function
    End If
    If Not m_elements.Exists(name) Then
        diagnostic = "Table control was not found: " & name
        Exit Function
    End If
    Set registered = m_elements(name)
    If Not TypeOf registered Is obj_UiControlElement Then
        diagnostic = "The element is not a table control: " & name
        Exit Function
    End If
    Set element = registered
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    On Error GoTo FailedRow
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    TryRefreshTableRow = element.TryRefreshTableRow(rowIndex, diagnostic)
CleanRow:
    Application.ScreenUpdating = previousScreen
    Application.EnableEvents = previousEvents
    Exit Function
FailedRow:
    diagnostic = "Table.RefreshRow: " & VBA.Err.Description
    Resume CleanRow
End Function

Public Sub UnregisterForm(ByVal name As String)
    If m_forms Is Nothing Then Exit Sub
    If m_forms.Exists(name) Then m_forms.Remove name
End Sub

Public Sub UnregisterShapeLayout(ByVal name As String)
    m_styles.ForgetShape name
    If m_shapeLayouts Is Nothing Then Exit Sub
    If m_shapeLayouts.Exists(name) Then m_shapeLayouts.Remove name
End Sub

Public Sub MarkDirty(ByVal name As String, ByVal layoutChanged As Boolean)
    Dim key As Variant
    Dim registered As obj_IUiElement
    Dim controlElement As obj_UiControlElement

    If m_isDisposed Or Not m_hasRendered Or m_rendering Then Exit Sub
    If Not m_elements.Exists(name) Then VBA.Err.Raise 5, , "Element was not found: " & name
    If m_dirty.Exists(name) Then
        If m_dirty(name) Then Exit Sub
        m_dirty(name) = VBA.CBool(m_dirty(name)) Or layoutChanged
    Else
        m_dirty.Add name, layoutChanged
    End If
    If layoutChanged Then
        For Each key In m_elements.Keys
            Set registered = m_elements(key)
            If TypeOf registered Is obj_UiControlElement Then
                Set controlElement = registered
                If controlElement.DependsOnLayout(name) Then Me.MarkDirty VBA.CStr(key), True
            End If
        Next key
    End If
End Sub

Public Function InvalidateVisual(ByVal name As String, ByRef diagnostic As String) As Boolean
    If Not m_isInitialized Or m_isDisposed Or Not m_hasRendered Then Exit Function
    If Not m_elements.Exists(name) Then
        diagnostic = "Element was not found: " & name
        Exit Function
    End If
    Me.MarkDirty name, False
    InvalidateVisual = Me.FlushLayout(diagnostic)
End Function

Public Sub RegisterForm(ByVal name As String, ByVal form As obj_IUiControl)
    If m_forms.Exists(name) Then VBA.Err.Raise VBA.vbObjectError + 2223, , "Duplicate form: " & name
    Set m_forms(name) = form
End Sub

Public Function ValidateForm(ByVal name As String, ByVal errors As Collection) As Boolean
    Dim form As obj_IUiControl

    If Not m_forms.Exists(name) Then
        errors.Add "Form was not found: " & name
        Exit Function
    End If
    Set form = m_forms(name)
    ValidateForm = form.Validate(errors)
End Function

Public Function TryGetFormDefinition( _
    ByVal name As String, _
    ByRef definition As Object _
) As Boolean
    Dim form As obj_UiFormControl
    Dim registeredForm As obj_IUiControl

    Set definition = Nothing
    If m_forms Is Nothing Then Exit Function
    If Not m_forms.Exists(name) Then Exit Function
    Set registeredForm = m_forms(name)
    If Not TypeOf registeredForm Is obj_UiFormControl Then Exit Function
    Set form = registeredForm
    Set definition = form.GetDefinition()
    TryGetFormDefinition = Not definition Is Nothing
End Function

Public Function Build(ByRef diagnostic As String) As Boolean
    Dim validator As New obj_UiMarkupValidator
    Dim errors As New Collection
    Dim markupError As obj_UiMarkupDiagnostic

#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_VALIDATOR_INITIALIZE_STARTED"
#End If
    If Not validator.Initialize() Then
        Err.Raise VBA.vbObjectError + 2167, , "Schema/validator initialization failed."
    End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_VALIDATOR_INITIALIZE_COMPLETED"
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_MARKUP_VALIDATE_STARTED"
#End If
    If Not validator.Validate(m_uiPageDefinition.Document.documentElement, errors) Then
        ex_Core.fn_Diagnostic_WriteLog "UI_MARKUP_VALIDATE_FAILED"
        diagnostic = VBA.vbNullString
        For Each markupError In errors
            If VBA.Len(diagnostic) > 0 Then diagnostic = diagnostic & VBA.vbCrLf
            diagnostic = diagnostic & markupError.Describe()
        Next markupError
        Exit Function
    End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_MARKUP_VALIDATE_COMPLETED"
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_ROOT_CREATE_STARTED"
#End If
    If Not ex_UiFormLayouts.fn_Validate(m_uiPageDefinition.Document.documentElement, diagnostic) Then Exit Function
    Set m_root = ex_UiElementFactory.fn_Create(m_uiPageDefinition.Document.documentElement, _
        Me, VBA.vbNullString, diagnostic)
    Build = Not m_root Is Nothing
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_ROOT_CREATE_COMPLETED | Success=" & VBA.CStr(Build)
#End If
End Function

Public Sub InvalidateMeasure()
    m_needsMeasure = True
End Sub

Public Sub RegisterTablePlan(ByVal plan As Object)
    If m_tablePlans Is Nothing Then Set m_tablePlans = New Collection
    m_tablePlans.Add plan
End Sub

Public Sub RegisterShapeLayout( _
    ByVal shape As Shape, _
    ByVal area As Range, _
    ByVal node As Object, _
    Optional ByVal defaultWidth As String = "1", _
    Optional ByVal defaultHorizontal As String = "left", _
    Optional ByVal wrapCaption As Boolean = False _
)
    Dim entry As Collection

    If m_shapeLayouts Is Nothing Then
        Set m_shapeLayouts = VBA.CreateObject("Scripting.Dictionary")
    End If
    Set entry = New Collection
    entry.Add shape
    entry.Add area
    entry.Add private_ShapeFraction(node, "shapeWidth", defaultWidth)
    entry.Add private_ShapeFraction(node, "shapeHeight", "1")
    entry.Add private_ShapeAlignment(node, "shapeHorizontalAlignment", defaultHorizontal, "left|center|right")
    entry.Add private_ShapeAlignment(node, "shapeVerticalAlignment", "top", "top|center|bottom")
    entry.Add wrapCaption
    entry.Add private_ShapeAlignment(node, "shapePlacement", "freeFloating", "freeFloating|move|moveAndSize")
    Set m_shapeLayouts(shape.Name) = entry
    private_ArrangeShape entry
End Sub

Public Function TryEnsureSmart( _
    ByVal controlName As String, _
    ByVal tableIndex As Long, _
    ByRef diagnostic As String, _
    Optional ByVal excelTableName As String _
) As Boolean
    Dim plan As Object
    Dim area As Range

    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Or m_rendering Or Not m_hasRendered Or m_tablePlans Is Nothing Then
        diagnostic = "Table conversion requires a rendered page."
        Exit Function
    End If
    For Each plan In m_tablePlans
        If plan("owner") = controlName & "|" & VBA.CStr(tableIndex) Then
            If plan("selectable") Then
                diagnostic = "Smart table does not support positional selection bindings."
                Exit Function
            End If
            Set area = plan("conversionRange")
            If area Is Nothing Then
                diagnostic = "Table conversion requires visible headers and at least one data row."
                Exit Function
            End If
            TryEnsureSmart = ex_UiTables.fn_TryEnsureSmart(area, plan("owner"), diagnostic, excelTableName)
            Exit Function
        End If
    Next plan
    diagnostic = "Rendered table was not found: " & controlName & "[" & VBA.CStr(tableIndex) & "]"
End Function

Public Function FlushLayout(ByRef diagnostic As String) As Boolean
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    ex_Core.fn_Diagnostic_WriteLog "UI_LAYOUT_FLUSH | TraceVersion=2 | NeedsMeasure=" & VBA.CStr(m_needsMeasure) & _
        " | Rendering=" & VBA.CStr(m_rendering) & " | HasRendered=" & VBA.CStr(m_hasRendered)
    If Not m_needsMeasure Then
        FlushLayout = Me.FlushUpdates(diagnostic)
        Exit Function
    End If
    If m_rendering Then
        FlushLayout = True
        Exit Function
    End If
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH_LAYOUT
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    m_needsMeasure = False
    m_router.Dispose
    Set m_router = New obj_UiEventRouter
    m_router.Initialize
    If m_rows > 0 And m_columns > 0 Then
        ex_UiTables.fn_ClearRange m_targetWorksheet.Cells(1, 1).Resize(m_rows, m_columns), True
    End If
    FlushLayout = Me.RenderTree(diagnostic)
CleanLayout:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Exit Function
EH_LAYOUT:
    diagnostic = VBA.Err.Description
    Resume CleanLayout
End Function

Public Function FlushUpdates(ByRef diagnostic As String) As Boolean
    Dim previousEvents As Boolean
    Dim previousScreen As Boolean
    Dim previousExpand As Boolean
    Dim previousFill As Boolean
    Dim pending As Object
    Dim key As Variant
    Dim registered As obj_IUiElement
    Dim controlElement As obj_UiControlElement
    Dim scope As Range
    Dim area As Range
    Dim rows As Long
    Dim columns As Long
    Dim needsLayout As Boolean
    Dim startedAt As Double
    Dim part As Range
    Dim oldAreas As Object
    Dim oldArea As Range
    Dim oldElements As Object
    Dim oldElement As obj_IUiElement
    Dim layoutChanged As Variant
    Dim refreshed As Boolean
    Dim tableSettingsCaptured As Boolean

    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then
        diagnostic = "Cannot flush updates: the UI render context is not initialized or has been disposed."
        Exit Function
    End If
    ' Полный рендер сам прочитает актуальные bindings; промежуточный flush не является ошибкой.
    If m_rendering Or Not m_hasRendered Then
        FlushUpdates = True
        Exit Function
    End If
    If m_dirty.Count = 0 Then
        FlushUpdates = True
        Exit Function
    End If
    For Each layoutChanged In m_dirty.Items
        If layoutChanged Then
            needsLayout = True
            Exit For
        End If
    Next layoutChanged
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    If needsLayout Then
        previousExpand = Application.AutoCorrect.AutoExpandListRange
        previousFill = Application.AutoCorrect.AutoFillFormulasInLists
        tableSettingsCaptured = True
    End If
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If needsLayout Then
        Application.AutoCorrect.AutoExpandListRange = False
        Application.AutoCorrect.AutoFillFormulasInLists = False
    End If
    startedAt = VBA.Timer
    m_rendering = True
    Set pending = m_dirty
    Set m_pending = pending
    Set m_dirty = VBA.CreateObject("Scripting.Dictionary")
    m_dirty.CompareMode = VBA.vbTextCompare
    If pending.Count > 0 Then
        ' Consume value-only updates and compatible source replacements before rendering.
        For Each key In pending.Keys
            If m_elements.Exists(key) Then
                Set registered = m_elements(key)
                If TypeOf registered Is obj_UiControlElement Then
                    Set controlElement = registered
                    If Not controlElement.TryRefreshValues(pending, refreshed, diagnostic) Then GoTo CleanExit
                    If refreshed Then pending.Remove key
                End If
            End If
        Next key
        needsLayout = False
        For Each layoutChanged In pending.Items
            If layoutChanged Then
                needsLayout = True
                Exit For
            End If
        Next layoutChanged
    End If
    If needsLayout Then
        Set oldAreas = VBA.CreateObject("Scripting.Dictionary")
        oldAreas.CompareMode = VBA.vbTextCompare
        Set oldElements = VBA.CreateObject("Scripting.Dictionary")
        oldElements.CompareMode = VBA.vbTextCompare
        For Each key In m_elements.Keys
            Set registered = m_elements(key)
            If TypeOf registered Is obj_UiControlElement Then
                Set controlElement = registered
                Set oldArea = controlElement.RenderArea
                If Not oldArea Is Nothing Then
                    Set oldAreas(key) = oldArea
                    Set oldElements(key) = registered
                End If
            End If
        Next key

        ' Rebuild structural owners before measuring dependent controls.
        For Each key In pending.Keys
            If Not pending.Exists(key) Then GoTo NextStructure
            If pending(key) And m_elements.Exists(key) Then
                needsLayout = True
                Set registered = m_elements(key)
                If TypeOf registered Is obj_UiControlElement Then
                    Set controlElement = registered
                    m_styles.ForgetControl VBA.CStr(key)
                    Set area = controlElement.RenderArea
                    private_UnionArea scope, area
                    If Not controlElement.RebuildStructure(diagnostic) Then GoTo CleanExit
                End If
            End If
NextStructure:
        Next key
    End If
    If needsLayout Then
        Set m_tablePlans = New Collection
        If Not m_root.Measure(rows, columns, diagnostic) Then GoTo CleanExit
        If Not m_root.Arrange(1, 1, diagnostic) Then GoTo CleanExit
        ' Newly visible controls have no previous area and still require rendering.
        For Each key In m_elements.Keys
            Set registered = m_elements(key)
            If TypeOf registered Is obj_UiControlElement Then
                Set controlElement = registered
                Set area = controlElement.RenderArea
                If Not area Is Nothing Then
                    If Not oldAreas.Exists(key) Then pending(key) = True
                End If
            End If
        Next key
        ' Изменение размеров может сдвинуть соседей в stackPanel или grid.
        For Each key In oldAreas.Keys
            If m_elements.Exists(key) Then
                Set registered = m_elements(key)
                Set oldElement = oldElements(key)
                If registered Is oldElement Then
                    Set controlElement = registered
                    Set oldArea = oldAreas(key)
                    Set area = controlElement.RenderArea
                    If Not area Is Nothing Then
                        If area.Address <> oldArea.Address Then
                            pending(key) = True
                            private_UnionArea scope, oldArea
                        End If
                    End If
                End If
            End If
        Next key
        m_rows = rows
        m_columns = columns
        For Each key In pending.Keys
            If pending(key) And m_elements.Exists(key) Then
                Set registered = m_elements(key)
                If TypeOf registered Is obj_UiControlElement Then
                    Set controlElement = registered
                    private_UnionArea scope, controlElement.RenderArea
                End If
            End If
        Next key
        If Not scope Is Nothing Then
            ex_UiTables.fn_Prepare m_targetWorksheet, m_tablePlans, scope
            For Each part In scope.Areas
                ex_UiTables.fn_ClearRange part
            Next part
            m_styles.BeginPartialRender scope
        End If
    End If

    ' Обновляем только dirty-контролы, сохраняя остальные объекты листа.
    For Each key In pending.Keys
        If m_elements.Exists(key) Then
            Set registered = m_elements(key)
            If TypeOf registered Is obj_UiControlElement Then
                Set controlElement = registered
                If pending(key) Then
                    m_styles.ForgetControl VBA.CStr(key)
                    If Not registered.Render(diagnostic) Then GoTo CleanExit
                Else
                    If Not controlElement.RefreshVisual(diagnostic) Then GoTo CleanExit
                End If
            Else
                If Not registered.Render(diagnostic) Then GoTo CleanExit
            End If
        End If
    Next key
    If Not scope Is Nothing Then m_styles.CompletePartialRender
    If needsLayout Then private_ArrangeShapes
    FlushUpdates = True
    ex_Core.fn_Diagnostic_WritePerf "Page.FlushDirty | Sheet=" & m_targetWorksheet.Name & " | Count=" & VBA.CStr(pending.Count), startedAt
CleanExit:
    m_styles.EndPartialRender
    If Not FlushUpdates Then
        ' Сохраняем неудавшиеся обновления для повторной попытки.
        If Not pending Is Nothing Then
            For Each key In pending.Keys
                If m_elements.Exists(key) Then m_dirty(key) = pending(key)
            Next key
        End If
    End If
    m_rendering = False
    Set m_pending = Nothing
    If tableSettingsCaptured Then
        Application.AutoCorrect.AutoExpandListRange = previousExpand
        Application.AutoCorrect.AutoFillFormulasInLists = previousFill
    End If
    Application.ScreenUpdating = previousScreen
    Application.EnableEvents = previousEvents
    Exit Function
Failed:
    diagnostic = VBA.Err.Source & ": " & VBA.Err.Description
    Resume CleanExit
End Function

Public Function RenderTree(ByRef diagnostic As String) As Boolean
    Dim pageElement As obj_UiPageElement
    Dim key As Variant
    Dim element As obj_IUiElement
    Dim controlElement As obj_UiControlElement
    Dim rows As Long
    Dim columns As Long
    Dim stageStartedAt As Double
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousExpand As Boolean
    Dim previousFill As Boolean

    If m_root Is Nothing Then Exit Function
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousExpand = Application.AutoCorrect.AutoExpandListRange
    previousFill = Application.AutoCorrect.AutoFillFormulasInLists
    On Error GoTo EH_RENDER
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AutoCorrect.AutoExpandListRange = False
    Application.AutoCorrect.AutoFillFormulasInLists = False
    m_rendering = True
    m_hasRendered = False
    Set m_tablePlans = New Collection
    Set m_shapeLayouts = VBA.CreateObject("Scripting.Dictionary")
    stageStartedAt = VBA.Timer
    If Not m_root.Measure(rows, columns, diagnostic) Then GoTo CleanExit
    ex_Core.fn_Diagnostic_WritePerf "Page.Measure | Sheet=" & m_targetWorksheet.Name, stageStartedAt
    stageStartedAt = VBA.Timer
    If Not m_root.Arrange(1, 1, diagnostic) Then GoTo CleanExit
    ex_Core.fn_Diagnostic_WritePerf "Page.Arrange | Sheet=" & m_targetWorksheet.Name, stageStartedAt
    stageStartedAt = VBA.Timer
    ex_UiTables.fn_Prepare m_targetWorksheet, m_tablePlans
    ex_Core.fn_Diagnostic_WritePerf "Page.PrepareTables | Sheet=" & m_targetWorksheet.Name, stageStartedAt
    m_rows = rows
    m_columns = columns
    stageStartedAt = VBA.Timer
    m_styles.BeginRender rows, columns
    ex_Core.fn_Diagnostic_WritePerf "Page.BeginRender | Sheet=" & m_targetWorksheet.Name, stageStartedAt
    stageStartedAt = VBA.Timer
    RenderTree = m_root.Render(diagnostic)
    ex_Core.fn_Diagnostic_WritePerf "Page.RenderElements | Sheet=" & m_targetWorksheet.Name, stageStartedAt
    If RenderTree Then
        stageStartedAt = VBA.Timer
        m_styles.CompleteRender
        private_ArrangeShapes
        For Each key In m_elements.Keys
            Set element = m_elements(key)
            If TypeOf element Is obj_UiControlElement Then
                Set controlElement = element
                controlElement.TraceVisibility "Page.CompleteRender"
            End If
        Next key
        If TypeOf m_root Is obj_UiPageElement Then
            Set pageElement = m_root
            pageElement.ApplyZoom
        End If
        ex_Core.fn_Diagnostic_WritePerf "Page.CompleteRender | Sheet=" & m_targetWorksheet.Name, stageStartedAt
    End If
    m_hasRendered = RenderTree
    If RenderTree Then m_dirty.RemoveAll
CleanExit:
    m_rendering = False
    Application.AutoCorrect.AutoExpandListRange = previousExpand
    Application.AutoCorrect.AutoFillFormulasInLists = previousFill
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Exit Function
EH_RENDER:
    RenderTree = False
    diagnostic = VBA.Err.Source & ": " & VBA.Err.Description
    Resume CleanExit
End Function

' //
' // Private
' //
Private Sub private_ArrangeShapes()
    Dim entry As Collection
    Dim key As Variant

    ' Recalculate geometry after the final column widths and row heights are applied.
    For Each key In m_shapeLayouts.Keys
        Set entry = m_shapeLayouts(key)
        private_ArrangeShape entry
    Next key
End Sub

Private Sub private_UnionArea(ByRef scope As Range, ByVal area As Range)
    If area Is Nothing Then Exit Sub
    If scope Is Nothing Then
        Set scope = area
    Else
        Set scope = Application.Union(scope, area)
    End If
End Sub

Private Sub private_ArrangeShape(ByVal entry As Collection)
    Dim shape As Shape
    Dim area As Range
    Dim horizontalOffset As Double
    Dim verticalOffset As Double

    Set shape = entry(1)
    Set area = entry(2)
    shape.LockAspectRatio = msoFalse
    Select Case entry(8)
        Case "freeFloating"
            shape.Placement = xlFreeFloating
        Case "move"
            shape.Placement = xlMove
        Case "moveAndSize"
            shape.Placement = xlMoveAndSize
    End Select
    shape.Width = area.Width * entry(3)
    shape.Height = area.Height * entry(4)
    Select Case entry(5)
        Case "center"
            horizontalOffset = (area.Width - shape.Width) / 2
        Case "right"
            horizontalOffset = area.Width - shape.Width
    End Select
    Select Case entry(6)
        Case "center"
            verticalOffset = (area.Height - shape.Height) / 2
        Case "bottom"
            verticalOffset = area.Height - shape.Height
    End Select
    shape.Left = area.Left + horizontalOffset
    shape.Top = area.Top + verticalOffset
    If entry(7) Then ex_UiInlineText.fn_WrapCaption shape
End Sub

Private Function private_ShapeFraction( _
    ByVal node As Object, _
    ByVal attributeName As String, _
    ByVal defaultValue As String _
) As Double
    Dim raw As String
    Dim parts As Variant
    Dim part As Variant
    Dim numerator As Double
    Dim denominator As Double

    raw = ex_UiElementFactory.fn_Attribute(node, attributeName)
    If node.getAttributeNode(attributeName) Is Nothing Then
        raw = defaultValue
    End If
    If VBA.Len(raw) = 0 Then GoTo InvalidFraction
    parts = VBA.Split(raw, "/")
    If UBound(parts) > 1 Then GoTo InvalidFraction
    For Each part In parts
        If VBA.Len(part) = 0 Or VBA.Len(part) > 9 Then GoTo InvalidFraction
        If part Like "*[!0-9]*" Then GoTo InvalidFraction
    Next part
    numerator = VBA.CDbl(parts(0))
    denominator = 1
    If UBound(parts) = 1 Then
        denominator = VBA.CDbl(parts(1))
    End If
    If numerator <= 0 Or denominator <= 0 Then GoTo InvalidFraction
    If numerator > denominator Then GoTo InvalidFraction
    private_ShapeFraction = numerator / denominator
    Exit Function
InvalidFraction:
    VBA.Err.Raise 5, "ShapeLayout", "Invalid " & attributeName & "=" & raw & _
        " for " & ex_UiElementFactory.fn_Attribute(node, "name") & "; expected a fraction in (0, 1]."
End Function

Private Function private_ShapeAlignment( _
    ByVal node As Object, _
    ByVal attributeName As String, _
    ByVal defaultValue As String, _
    ByVal allowedValues As String _
) As String
    Dim raw As String

    raw = ex_UiElementFactory.fn_Attribute(node, attributeName)
    If node.getAttributeNode(attributeName) Is Nothing Then
        raw = defaultValue
    End If
    If VBA.InStr(1, "|" & allowedValues & "|", "|" & raw & "|", VBA.vbBinaryCompare) = 0 Or VBA.Len(raw) = 0 Then
        VBA.Err.Raise 5, "ShapeLayout", "Invalid " & attributeName & "=" & raw
    End If
    private_ShapeAlignment = raw
End Function