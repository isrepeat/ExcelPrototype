VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiMarkupSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_attributes As Object
Private m_requiredAny As Collection
Private m_children As Object
Private m_visualChildren As Boolean
Private m_dispatchControls As Boolean

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_requiredAny = New Collection
    Set m_attributes = VBA.CreateObject("Scripting.Dictionary")
    Set m_children = VBA.CreateObject("Scripting.Dictionary")
    m_children.CompareMode = VBA.vbTextCompare
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get RequiredAny() As Collection
    Set RequiredAny = m_requiredAny
End Property

Public Property Get Attributes() As Object
    Set Attributes = m_attributes
End Property

Public Property Get Children() As Object
    Set Children = m_children
End Property

Public Property Get VisualChildren() As Boolean
    VisualChildren = m_visualChildren
End Property

Public Property Get DispatchControls() As Boolean
    DispatchControls = m_dispatchControls
End Property

' //
' // API
' //
Public Sub Dispose()
    Set m_requiredAny = Nothing
    Set m_attributes = Nothing
    Set m_children = Nothing
End Sub

Public Sub AddAttribute( _
    ByVal name As String, _
    Optional ByVal kind As String = "string", _
    Optional ByVal required As Boolean = False, _
    Optional ByVal values As String = "", _
    Optional ByVal binding As Boolean = False, _
    Optional ByVal maxLength As Long = 0 _
)
    m_attributes(name) = VBA.Array(kind, required, values, binding, maxLength)
End Sub

Public Sub AddChild( _
    ByVal tag As String, _
    ByVal schema As obj_UiMarkupSchema, _
    Optional ByVal minimum As Long = 0, _
    Optional ByVal maximum As Long = -1 _
)
    m_children(tag) = VBA.Array(schema, minimum, maximum)
End Sub

Public Sub AllowVisualChildren()
    m_visualChildren = True
End Sub

Public Sub ResolveControlType()
    m_dispatchControls = True
End Sub

Public Sub RequireAnyAttribute(ByVal names As String)
    m_requiredAny.Add names
End Sub