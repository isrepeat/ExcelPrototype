VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_StyleLayer"
Option Explicit

Private m_LayerId As String
Private m_Priority As Long
Private m_Source As String
Private m_IsEnabled As Boolean
Private m_Stage As String
Private m_Rules As Collection

Public Sub Initialize( _
    ByVal layerId As String, _
    ByVal priority As Long, _
    ByVal sourceName As String, _
    Optional ByVal isEnabled As Boolean = True, _
    Optional ByVal stageName As String = vbNullString _
)
    m_LayerId = Trim$(layerId)
    m_Priority = priority
    m_Source = Trim$(sourceName)
    m_IsEnabled = isEnabled
    m_Stage = LCase$(Trim$(stageName))
    Set m_Rules = New Collection
End Sub

Public Property Get LayerId() As String
    LayerId = m_LayerId
End Property

Public Property Let LayerId(ByVal value As String)
    m_LayerId = Trim$(value)
End Property

Public Property Get Priority() As Long
    Priority = m_Priority
End Property

Public Property Let Priority(ByVal value As Long)
    m_Priority = value
End Property

Public Property Get SourceName() As String
    SourceName = m_Source
End Property

Public Property Let SourceName(ByVal value As String)
    m_Source = Trim$(value)
End Property

Public Property Get IsEnabled() As Boolean
    IsEnabled = m_IsEnabled
End Property

Public Property Let IsEnabled(ByVal value As Boolean)
    m_IsEnabled = value
End Property

Public Property Get StageName() As String
    StageName = m_Stage
End Property

Public Property Let StageName(ByVal value As String)
    m_Stage = LCase$(Trim$(value))
End Property

Public Property Get Rules() As Collection
    If m_Rules Is Nothing Then Set m_Rules = New Collection
    Set Rules = m_Rules
End Property

Public Property Get RuleCount() As Long
    If m_Rules Is Nothing Then
        RuleCount = 0
    Else
        RuleCount = m_Rules.Count
    End If
End Property

Public Sub AddRule(ByVal rule As obj_StyleRule)
    If rule Is Nothing Then Exit Sub
    If m_Rules Is Nothing Then Set m_Rules = New Collection
    m_Rules.Add rule
End Sub

Public Sub InsertRule(ByVal atIndex As Long, ByVal rule As obj_StyleRule)
    If rule Is Nothing Then Exit Sub
    If m_Rules Is Nothing Then Set m_Rules = New Collection

    If m_Rules.Count = 0 Then
        m_Rules.Add rule
    ElseIf atIndex <= 1 Then
        m_Rules.Add rule, Before:=1
    ElseIf atIndex > m_Rules.Count Then
        m_Rules.Add rule
    Else
        m_Rules.Add rule, Before:=atIndex
    End If
End Sub

Public Function GetRuleById(ByVal ruleId As String) As obj_StyleRule
    Dim i As Long
    Dim candidate As obj_StyleRule
    Dim normalizedId As String

    normalizedId = Trim$(ruleId)
    If Len(normalizedId) = 0 Then Exit Function
    If m_Rules Is Nothing Then Exit Function

    For i = 1 To m_Rules.Count
        Set candidate = m_Rules(i)
        If Not candidate Is Nothing Then
            If StrComp(candidate.RuleId, normalizedId, vbTextCompare) = 0 Then
                Set GetRuleById = candidate
                Exit Function
            End If
        End If
    Next i
End Function

Public Function RemoveRuleById(ByVal ruleId As String) As Boolean
    Dim i As Long
    Dim candidate As obj_StyleRule
    Dim normalizedId As String

    normalizedId = Trim$(ruleId)
    If Len(normalizedId) = 0 Then Exit Function
    If m_Rules Is Nothing Then Exit Function

    For i = m_Rules.Count To 1 Step -1
        Set candidate = m_Rules(i)
        If Not candidate Is Nothing Then
            If StrComp(candidate.RuleId, normalizedId, vbTextCompare) = 0 Then
                m_Rules.Remove i
                RemoveRuleById = True
                Exit Function
            End If
        End If
    Next i
End Function
