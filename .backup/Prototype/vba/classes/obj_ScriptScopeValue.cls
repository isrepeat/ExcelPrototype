VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ScriptScopeValue"
Option Explicit

Private m_Kind As String
Private m_TextValue As String
Private m_ObjectValue As Object
Private m_HasObjectValue As Boolean
Private m_HasTextValue As Boolean

Public Property Get Kind() As String
    Kind = m_Kind
End Property

Public Property Let Kind(ByVal valueText As String)
    m_Kind = LCase$(Trim$(valueText))
End Property

Public Property Get HasObjectValue() As Boolean
    HasObjectValue = m_HasObjectValue
End Property

Public Property Get HasTextValue() As Boolean
    HasTextValue = m_HasTextValue
End Property

Public Property Get TextValue() As String
    TextValue = m_TextValue
End Property

Public Property Let TextValue(ByVal valueText As String)
    m_TextValue = valueText
    m_HasTextValue = True
End Property

Public Property Get ObjectValue() As Object
    If Not m_HasObjectValue Then Exit Property
    Set ObjectValue = m_ObjectValue
End Property

Public Property Set ObjectValue(ByVal valueObject As Object)
    Set m_ObjectValue = valueObject
    m_HasObjectValue = Not (valueObject Is Nothing)
End Property

Public Sub InitializeString(ByVal kindText As String, ByVal valueText As String)
    m_Kind = LCase$(Trim$(kindText))
    m_TextValue = valueText
    m_HasTextValue = True
    Set m_ObjectValue = Nothing
    m_HasObjectValue = False
End Sub

Public Sub InitializeObject(ByVal kindText As String, ByVal valueObject As Object)
    m_Kind = LCase$(Trim$(kindText))
    Set m_ObjectValue = valueObject
    m_HasObjectValue = Not (valueObject Is Nothing)
    m_TextValue = vbNullString
    m_HasTextValue = False
End Sub
