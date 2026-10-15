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

' //
' // API
' //
Public Sub RegisterElement(ByVal name As String, ByVal element As obj_IUiElement)
    If VBA.Len(name) = 0 Then Exit Sub
    If m_elements.Exists(name) Then VBA.Err.Raise VBA.vbObjectError + 2232, , "Duplicate element: " & name
    Set m_elements(name) = element
End Sub

Public Function InvalidateVisual(ByVal name As String, ByRef diagnostic As String) As Boolean
    Dim element As obj_IUiElement
    Dim previousEvents As Boolean
    Dim previousExpand As Boolean
    Dim previousFill As Boolean

    If Not m_elements.Exists(name) Then
        diagnostic = "Element was not found: " & name
        Exit Function
    End If
    Set element = m_elements(name)
    previousEvents = Application.EnableEvents
    previousExpand = Application.AutoCorrect.AutoExpandListRange
    previousFill = Application.AutoCorrect.AutoFillFormulasInLists
    On Error GoTo EH_VISUAL
    Application.EnableEvents = False
    Application.AutoCorrect.AutoExpandListRange = False
    Application.AutoCorrect.AutoFillFormulasInLists = False
    InvalidateVisual = element.Render(diagnostic)
CleanVisual:
    Application.AutoCorrect.AutoExpandListRange = previousExpand
    Application.AutoCorrect.AutoFillFormulasInLists = previousFill
    Application.EnableEvents = previousEvents
    Exit Function
EH_VISUAL:
    diagnostic = VBA.Err.Description
    Resume CleanVisual
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
    Optional ByVal defaultHorizontal As String = "left" _
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

    If m_rendering Or Not m_needsMeasure Then
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

Public Function RenderTree(ByRef diagnostic As String) As Boolean
    Dim pageElement As obj_UiPageElement
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
        If TypeOf m_root Is obj_UiPageElement Then
            Set pageElement = m_root
            pageElement.ApplyZoom
        End If
        ex_Core.fn_Diagnostic_WritePerf "Page.CompleteRender | Sheet=" & m_targetWorksheet.Name, stageStartedAt
    End If
    m_hasRendered = RenderTree
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
' // Properties
' //
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
    If targetWorksheet Is Nothing Or uiPageDefinition Is Nothing Or uiBindingContext Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(uiFolderPath)) = 0 Then Exit Function
    Set m_targetWorksheet = targetWorksheet
    Set m_uiPageDefinition = uiPageDefinition
    m_uiFolderPath = uiFolderPath
    Set m_uiBindingContext = uiBindingContext
    Set m_styles = New obj_UiStyleCatalog
    If Not m_styles.Initialize() Then Exit Function
    Set m_router = New obj_UiEventRouter
    m_router.Initialize
    Set m_forms = VBA.CreateObject("Scripting.Dictionary")
    m_forms.CompareMode = VBA.vbTextCompare
    Set m_elements = VBA.CreateObject("Scripting.Dictionary")
    m_elements.CompareMode = VBA.vbTextCompare
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
            Set element = m_elements(key)
            element.Dispose
        Next key
    End If
    If Not m_styles Is Nothing Then m_styles.Dispose
    Set m_styles = Nothing
    Set m_root = Nothing
    Set m_router = Nothing
    Set m_forms = Nothing
    Set m_elements = Nothing
    Set m_tablePlans = Nothing
    Set m_shapeLayouts = Nothing
    Set m_uiBindingContext = Nothing
    If Not m_uiPageDefinition Is Nothing Then m_uiPageDefinition.Dispose
    Set m_uiPageDefinition = Nothing
    Set m_targetWorksheet = Nothing
    m_uiFolderPath = VBA.vbNullString
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_CONTEXT_DISPOSE_COMPLETED"
#End If
End Sub

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

Private Sub private_ArrangeShape(ByVal entry As Collection)
    Dim shape As Shape
    Dim area As Range
    Dim horizontalOffset As Double
    Dim verticalOffset As Double

    Set shape = entry(1)
    Set area = entry(2)
    shape.LockAspectRatio = msoFalse
    shape.Placement = xlFreeFloating
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