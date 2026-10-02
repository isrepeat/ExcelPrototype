VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiGridAxis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_specs As Collection
Private m_sizes() As Long
Private m_available As Long
Private m_explicit As Boolean

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
Public Property Get HasDefinitions() As Boolean
    HasDefinitions = m_explicit
End Property

Public Property Get Total() As Long
    Dim index As Long
    For index = 1 To m_specs.Count
        Total = Total + m_sizes(index)
    Next index
End Property

' //
' // API
' //
Public Function Initialize(ByVal definition As Object, ByVal groupName As String, ByVal itemName As String, ByVal extentName As String, ByRef diagnostic As String) As Boolean
    Dim group As Object, node As Object
    Dim spec As String
    If m_isDisposed Or m_isInitialized Then
        diagnostic = "Object is already initialized or disposed."
        Exit Function
    End If
    Set m_specs = New Collection
    m_available = 0
    spec = ex_UiElementFactory.fn_Attribute(definition, extentName)
    If VBA.Len(spec) > 0 Then m_available = VBA.CLng(spec)
    For Each group In definition.ChildNodes
        If group.NodeType = 1 Then
            If VBA.CStr(group.namespaceURI) = "urn:excelprototype:profiles" And VBA.CStr(group.baseName) = groupName Then
                m_explicit = True
                For Each node In group.ChildNodes
                    If node.NodeType = 1 Then
                        spec = ex_UiElementFactory.fn_Attribute(node, "size")
                        m_specs.Add spec
                    End If
                Next node
            End If
        End If
    Next group
    If m_explicit And m_specs.Count = 0 Then
        diagnostic = "Grid definitions must contain at least one track: " & groupName
        Exit Function
    End If
    If Not m_explicit Then m_specs.Add "auto"
    Me.Reset
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_specs = Nothing
    Erase m_sizes
    m_available = 0
    m_explicit = False
End Sub

Public Sub Reset()
    Dim index As Long, spec As String
    ReDim m_sizes(1 To m_specs.Count)
    For index = 1 To m_specs.Count
        spec = m_specs(index)
        If spec = "auto" Or VBA.Right$(spec, 1) = "*" Then
            m_sizes(index) = 1
        Else
            m_sizes(index) = VBA.CLng(spec)
        End If
    Next index
End Sub

Public Function Contains(ByVal first As Long, ByVal count As Long) As Boolean
    Contains = first > 0 And count > 0 And first + count - 1 <= m_specs.Count
End Function

Public Function Offset(ByVal first As Long) As Long
    Dim index As Long
    For index = 1 To first - 1
        Offset = Offset + m_sizes(index)
    Next index
End Function

Public Function Extent(ByVal first As Long, ByVal count As Long) As Long
    Dim index As Long
    For index = first To first + count - 1
        Extent = Extent + m_sizes(index)
    Next index
End Function

Public Sub Grow(ByVal first As Long, ByVal count As Long, ByVal desired As Long)
    Dim index As Long, deficit As Long, tracks As Long, increment As Long
    deficit = desired - Me.Extent(first, count)
    If deficit <= 0 Then Exit Sub
    For index = first To first + count - 1
        If m_specs(index) = "auto" Then tracks = tracks + 1
    Next index
    For index = first To first + count - 1
        If m_specs(index) = "auto" Then
            increment = (deficit + tracks - 1) \ tracks
            m_sizes(index) = m_sizes(index) + increment
            deficit = deficit - increment
            tracks = tracks - 1
        End If
    Next index
End Sub

Public Function Resolve(ByRef diagnostic As String) As Boolean
    Dim index As Long, fixed As Long, remaining As Long, assigned As Long, stars As Long
    Dim weight As Double, totalWeight As Double, cumulative As Double
    For index = 1 To m_specs.Count
        If VBA.Right$(m_specs(index), 1) = "*" Then
            stars = stars + 1
            totalWeight = totalWeight + private_Weight(m_specs(index))
        Else
            fixed = fixed + m_sizes(index)
        End If
    Next index
    If stars = 0 Then
        Resolve = True
        Exit Function
    End If
    If m_available = 0 Then
        diagnostic = "Star grid tracks require an explicit rowSpan or columnSpan on the grid."
        Exit Function
    End If
    remaining = m_available - fixed
    If remaining < stars Then
        diagnostic = "Grid extent is too small for its track definitions."
        Exit Function
    End If
    For index = 1 To m_specs.Count
        If VBA.Right$(m_specs(index), 1) = "*" Then
            cumulative = cumulative + private_Weight(m_specs(index))
            weight = VBA.Fix((remaining - stars) * cumulative / totalWeight)
            m_sizes(index) = 1 + VBA.CLng(weight) - assigned
            assigned = VBA.CLng(weight)
        End If
    Next index
    Resolve = True
End Function

' //
' // Private
' //
Private Function private_Weight(ByVal spec As String) As Double
    private_Weight = 1
    If VBA.Len(spec) > 1 Then private_Weight = VBA.CDbl(VBA.Left$(spec, VBA.Len(spec) - 1))
End Function