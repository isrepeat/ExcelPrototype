VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_StyleRule"
Option Explicit

Private m_RuleId As String
Private m_Target As String
Private m_Selector As String
Private m_Declarations As Object

Public Sub Initialize( _
    ByVal ruleId As String, _
    ByVal target As String, _
    ByVal selector As String, _
    Optional ByVal declarations As Object = Nothing _
)
    m_RuleId = Trim$(ruleId)
    m_Target = LCase$(Trim$(target))
    m_Selector = Trim$(selector)
    Set m_Declarations = mp_CreateDeclarationsDictionary()
    If Not declarations Is Nothing Then
        mp_MergeDeclarations m_Declarations, declarations
    End If
End Sub

Public Property Get RuleId() As String
    RuleId = m_RuleId
End Property

Public Property Let RuleId(ByVal value As String)
    m_RuleId = Trim$(value)
End Property

Public Property Get Target() As String
    Target = m_Target
End Property

Public Property Let Target(ByVal value As String)
    m_Target = LCase$(Trim$(value))
End Property

Public Property Get Selector() As String
    Selector = m_Selector
End Property

Public Property Let Selector(ByVal value As String)
    m_Selector = Trim$(value)
End Property

Public Property Get Declarations() As Object
    If m_Declarations Is Nothing Then Set m_Declarations = mp_CreateDeclarationsDictionary()
    Set Declarations = m_Declarations
End Property

Public Property Set Declarations(ByVal value As Object)
    Set m_Declarations = mp_CreateDeclarationsDictionary()
    If Not value Is Nothing Then
        mp_MergeDeclarations m_Declarations, value
    End If
End Property

Public Function HasDeclaration(ByVal propertyName As String) As Boolean
    If m_Declarations Is Nothing Then Exit Function
    HasDeclaration = m_Declarations.Exists(LCase$(Trim$(propertyName)))
End Function

Public Function GetDeclarationText( _
    ByVal propertyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim normalizedName As String

    normalizedName = LCase$(Trim$(propertyName))
    If Len(normalizedName) = 0 Then
        GetDeclarationText = defaultValue
        Exit Function
    End If

    If m_Declarations Is Nothing Then
        GetDeclarationText = defaultValue
        Exit Function
    End If

    If m_Declarations.Exists(normalizedName) Then
        GetDeclarationText = CStr(m_Declarations(normalizedName))
    Else
        GetDeclarationText = defaultValue
    End If
End Function

Public Sub SetDeclaration(ByVal propertyName As String, ByVal propertyValue As String)
    Dim normalizedName As String

    normalizedName = LCase$(Trim$(propertyName))
    If Len(normalizedName) = 0 Then Exit Sub
    If m_Declarations Is Nothing Then Set m_Declarations = mp_CreateDeclarationsDictionary()
    m_Declarations(normalizedName) = Trim$(propertyValue)
End Sub

Private Function mp_CreateDeclarationsDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Set mp_CreateDeclarationsDictionary = dict
End Function

Private Sub mp_MergeDeclarations(ByVal targetDict As Object, ByVal sourceDict As Object)
    Dim key As Variant

    If targetDict Is Nothing Then Exit Sub
    If sourceDict Is Nothing Then Exit Sub

    For Each key In sourceDict.Keys
        targetDict(CStr(key)) = CStr(sourceDict(key))
    Next key
End Sub
