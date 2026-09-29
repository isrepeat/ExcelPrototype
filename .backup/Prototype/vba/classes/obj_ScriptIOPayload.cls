VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ScriptIOPayload"
Option Explicit

Private m_Values As Object

Private Sub Class_Initialize()
    Set m_Values = CreateObject("Scripting.Dictionary")
    m_Values.CompareMode = 1
End Sub

Public Sub m_SetString(ByVal keyName As String, ByVal valueText As String)
    Dim scopeValue As obj_ScriptScopeValue

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Sub

    Set scopeValue = ex_ScriptScopeValue.m_CreateStringValue(CStr(valueText))
    Set m_Values(keyName) = scopeValue
End Sub

Public Sub m_SetObject(ByVal keyName As String, ByVal valueObject As Object)
    Dim scopeValue As obj_ScriptScopeValue

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Sub
    If valueObject Is Nothing Then Exit Sub

    Set scopeValue = ex_ScriptScopeValue.m_CreateObjectValue(valueObject)
    Set m_Values(keyName) = scopeValue
End Sub

Public Function m_Has(ByVal keyName As String) As Boolean
    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    m_Has = m_Values.Exists(keyName)
End Function

Public Function m_TryGetScopeValue(ByVal keyName As String, ByRef outScopeValue As obj_ScriptScopeValue) As Boolean
    Dim rawObject As Object

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If Not m_Values.Exists(keyName) Then Exit Function
    If Not IsObject(m_Values(keyName)) Then Exit Function

    Set rawObject = m_Values(keyName)
    If Not TypeOf rawObject Is obj_ScriptScopeValue Then Exit Function

    Set outScopeValue = rawObject
    m_TryGetScopeValue = True
End Function

Public Function m_TryGetString(ByVal keyName As String, ByRef outValue As String) As Boolean
    Dim scopeValue As obj_ScriptScopeValue

    If Not m_TryGetScopeValue(keyName, scopeValue) Then Exit Function
    If scopeValue.HasObjectValue Then Exit Function

    outValue = scopeValue.TextValue
    m_TryGetString = True
End Function

Public Function m_TryGetObject(ByVal keyName As String, ByRef outObject As Object) As Boolean
    Dim scopeValue As obj_ScriptScopeValue

    If Not m_TryGetScopeValue(keyName, scopeValue) Then Exit Function
    If Not scopeValue.HasObjectValue Then Exit Function

    Set outObject = scopeValue.ObjectValue
    m_TryGetObject = Not (outObject Is Nothing)
End Function
