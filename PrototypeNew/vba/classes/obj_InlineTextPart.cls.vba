VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_InlineTextPart"
Option Explicit

Private m_InlineProfile As obj_InlineTextProfile
Private m_ResolvedText As String
Private m_Runs As Collection

Private Sub Class_Initialize()
    m_ResolvedText = VBA.vbNullString
    Set m_Runs = Nothing
End Sub

Public Property Get InlineProfile() As obj_InlineTextProfile
    Set InlineProfile = m_InlineProfile
End Property

Public Property Set InlineProfile(ByVal value As obj_InlineTextProfile)
    Set m_InlineProfile = value
End Property

Public Property Get ResolvedText() As String
    ResolvedText = m_ResolvedText
End Property

Public Property Get Runs() As Collection
    Set Runs = m_Runs
End Property

Public Function Resolve(ByVal rawText As String) As Boolean
    ' InlinePart описывает конкретное текстовое поле (caption/header/message).
    ' Он не знает правил сам по себе, а делегирует их в назначенный профиль.
    If m_InlineProfile Is Nothing Then
        m_ResolvedText = rawText
        Set m_Runs = Nothing
        Resolve = True
        Exit Function
    End If

    If Not m_InlineProfile.TryResolveInlineText(rawText, m_ResolvedText, m_Runs) Then Exit Function
    Resolve = True
End Function

Public Function RegisterForRange(ByVal page As obj_PageBase, ByVal targetRange As Range) As Boolean
    If page Is Nothing Then Exit Function
    If m_Runs Is Nothing Then
        RegisterForRange = True
        Exit Function
    End If
    If m_InlineProfile Is Nothing Then Exit Function

    ' Регистрируем runs в оркестраторе страницы; само применение будет позже (ApplyInlineRuns).
    RegisterForRange = page.RegisterInlineRuns(targetRange, m_Runs, m_InlineProfile)
End Function

Public Function RegisterForShape(ByVal page As obj_PageBase, ByVal targetShape As Shape) As Boolean
    If page Is Nothing Then Exit Function
    If m_Runs Is Nothing Then
        RegisterForShape = True
        Exit Function
    End If
    If m_InlineProfile Is Nothing Then Exit Function

    ' Аналогично для shape-целей.
    RegisterForShape = page.RegisterInlineRunsForShape(targetShape, m_Runs, m_InlineProfile)
End Function
