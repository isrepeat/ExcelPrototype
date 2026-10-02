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

Private m_rendering As Boolean
Private m_needsMeasure As Boolean
Private m_rows As Long
Private m_columns As Long
Private m_styles As obj_UiStyleCatalog
Private m_root As obj_IUiElement
Private m_router As obj_UiEventRouter
Private m_forms As Object
Private m_elements As Object
Private m_targetWorksheet As Worksheet
Private m_uiPageDefinition As obj_UiPageDefinition
Private m_uiFolderPath As String
Private m_uiBindingContext As obj_UiBindingContext
Private m_isDisposed As Boolean

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

    If Not m_elements.Exists(name) Then
        diagnostic = "Element was not found: " & name
        Exit Function
    End If
    Set element = m_elements(name)
    previousEvents = Application.EnableEvents
    On Error GoTo EH_VISUAL
    Application.EnableEvents = False
    InvalidateVisual = element.Render(diagnostic)
CleanVisual:
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

Public Function Build(ByRef diagnostic As String) As Boolean
    Dim validator As New obj_UiMarkupValidator
    Dim errors As New Collection
    Dim markupError As obj_UiMarkupDiagnostic

    If Not validator.Validate(m_uiPageDefinition.Document.documentElement, errors) Then
        diagnostic = VBA.vbNullString
        For Each markupError In errors
            If VBA.Len(diagnostic) > 0 Then diagnostic = diagnostic & VBA.vbCrLf
            diagnostic = diagnostic & markupError.Describe()
        Next markupError
        Exit Function
    End If
    Set m_root = ex_UiElementFactory.fn_Create(m_uiPageDefinition.Document.documentElement, _
        Me, VBA.vbNullString, diagnostic)
    Build = Not m_root Is Nothing
End Function

Public Sub InvalidateMeasure()
    m_needsMeasure = True
End Sub

Public Function FlushLayout(ByRef diagnostic As String) As Boolean
    Dim previousEvents As Boolean

    If m_rendering Or Not m_needsMeasure Then
        FlushLayout = True
        Exit Function
    End If
    previousEvents = Application.EnableEvents
    On Error GoTo EH_LAYOUT
    Application.EnableEvents = False
    m_needsMeasure = False
    m_router.Dispose
    m_router.Initialize
    If m_rows > 0 And m_columns > 0 Then
        With m_targetWorksheet.Cells(1, 1).Resize(m_rows, m_columns)
            .UnMerge
            .ClearContents
        End With
    End If
    FlushLayout = Me.RenderTree(diagnostic)
CleanLayout:
    Application.EnableEvents = previousEvents
    Exit Function
EH_LAYOUT:
    diagnostic = VBA.Err.Description
    Resume CleanLayout
End Function

Public Function RenderTree(ByRef diagnostic As String) As Boolean
    Dim rows As Long
    Dim columns As Long
    Dim previousEvents As Boolean

    If m_root Is Nothing Then Exit Function
    If Not m_root.Measure(rows, columns, diagnostic) Then Exit Function
    If Not m_root.Arrange(1, 1, diagnostic) Then Exit Function
    previousEvents = Application.EnableEvents
    On Error GoTo EH_RENDER
    Application.EnableEvents = False
    m_rendering = True
    m_rows = rows
    m_columns = columns
    RenderTree = m_root.Render(diagnostic)
CleanExit:
    m_rendering = False
    Application.EnableEvents = previousEvents
    Exit Function
EH_RENDER:
    diagnostic = VBA.Err.Description
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
    m_isDisposed = False
    If targetWorksheet Is Nothing Or uiPageDefinition Is Nothing Or uiBindingContext Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(uiFolderPath)) = 0 Then Exit Function
    Set m_targetWorksheet = targetWorksheet
    Set m_uiPageDefinition = uiPageDefinition
    m_uiFolderPath = uiFolderPath
    Set m_uiBindingContext = uiBindingContext
    Set m_styles = New obj_UiStyleCatalog
    Set m_router = New obj_UiEventRouter
    m_router.Initialize
    Set m_forms = VBA.CreateObject("Scripting.Dictionary")
    m_forms.CompareMode = VBA.vbTextCompare
    Set m_elements = VBA.CreateObject("Scripting.Dictionary")
    m_elements.CompareMode = VBA.vbTextCompare
    Initialize = True
End Function

Public Sub Dispose()
    Dim key As Variant
    Dim element As obj_IUiElement

    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    If Not m_router Is Nothing Then m_router.Dispose
    If Not m_root Is Nothing Then m_root.Dispose
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
    Set m_uiBindingContext = Nothing
    If Not m_uiPageDefinition Is Nothing Then m_uiPageDefinition.Dispose
    Set m_uiPageDefinition = Nothing
    Set m_targetWorksheet = Nothing
    m_uiFolderPath = VBA.vbNullString
End Sub