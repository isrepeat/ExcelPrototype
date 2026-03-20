Attribute VB_Name = "ex_ScriptScopeValue"
Option Explicit

Public Const KIND_STRING As String = "string"
Public Const KIND_ROW As String = "row"
Public Const KIND_COLUMN As String = "column"
Public Const KIND_DICTIONARY As String = "dictionary"
Public Const KIND_RESULT_TABLE As String = "resulttable"
Public Const KIND_COLLECTION As String = "collection"
Public Const KIND_OBJECT As String = "object"
Public Const KIND_TABLEREF As String = "tableref"

Public Function m_CreateStringValue(ByVal valueText As String) As obj_ScriptScopeValue
    Dim result As obj_ScriptScopeValue

    Set result = New obj_ScriptScopeValue
    result.InitializeString KIND_STRING, valueText
    Set m_CreateStringValue = result
End Function

Public Function m_CreateObjectValue(ByVal valueObject As Object) As obj_ScriptScopeValue
    Dim result As obj_ScriptScopeValue

    Set result = New obj_ScriptScopeValue
    result.InitializeObject m_DetectObjectKind(valueObject), valueObject
    Set m_CreateObjectValue = result
End Function

Public Function m_CreateTableRefValue(ByVal tableRef As String) As obj_ScriptScopeValue
    Dim result As obj_ScriptScopeValue

    Set result = New obj_ScriptScopeValue
    result.InitializeString KIND_TABLEREF, tableRef
    Set m_CreateTableRefValue = result
End Function

Public Function m_DetectObjectKind(ByVal valueObject As Object) As String
    If valueObject Is Nothing Then
        m_DetectObjectKind = KIND_OBJECT
        Exit Function
    End If

    If TypeOf valueObject Is obj_ResultRow Then
        m_DetectObjectKind = KIND_ROW
        Exit Function
    End If
    If TypeOf valueObject Is obj_ResultColumn Then
        m_DetectObjectKind = KIND_COLUMN
        Exit Function
    End If
    If TypeOf valueObject Is obj_ResultTable Then
        m_DetectObjectKind = KIND_RESULT_TABLE
        Exit Function
    End If
    If TypeName(valueObject) = "Dictionary" Then
        m_DetectObjectKind = KIND_DICTIONARY
        Exit Function
    End If
    If TypeName(valueObject) = "Collection" Then
        m_DetectObjectKind = KIND_COLLECTION
        Exit Function
    End If

    m_DetectObjectKind = KIND_OBJECT
End Function

Public Function m_TryGetScopeValue( _
    ByVal scopeRef As Object, _
    ByVal variableName As String, _
    ByRef outScopeValue As obj_ScriptScopeValue, _
    ByRef outErrorText As String _
) As Boolean
    Dim rawObject As Object

    variableName = Trim$(variableName)
    If Len(variableName) = 0 Then
        outErrorText = "Variable name is empty."
        Exit Function
    End If
    If scopeRef Is Nothing Or Not scopeRef.Exists(variableName) Then
        outErrorText = "Unknown variable '" & variableName & "'."
        Exit Function
    End If

    If Not IsObject(scopeRef(variableName)) Then
        outErrorText = "Variable '" & variableName & "' is stored in legacy runtime format. Scope entries must be obj_ScriptScopeValue."
        Exit Function
    End If
    Set rawObject = scopeRef(variableName)
    If Not TypeOf rawObject Is obj_ScriptScopeValue Then
        outErrorText = "Variable '" & variableName & "' has unsupported scope container type '" & TypeName(rawObject) & "'."
        Exit Function
    End If

    Set outScopeValue = rawObject
    m_TryGetScopeValue = True
End Function

Public Function m_TryGetObjectValue( _
    ByVal scopeValue As obj_ScriptScopeValue, _
    ByRef outObject As Object, _
    ByRef outErrorText As String _
) As Boolean
    If scopeValue Is Nothing Then
        outErrorText = "Scope value container is missing."
        Exit Function
    End If
    If Not scopeValue.HasObjectValue Then
        outErrorText = "Scope value kind '" & scopeValue.Kind & "' does not contain object payload."
        Exit Function
    End If

    Set outObject = scopeValue.ObjectValue
    If outObject Is Nothing Then
        outErrorText = "Scope value kind '" & scopeValue.Kind & "' has empty object payload."
        Exit Function
    End If

    m_TryGetObjectValue = True
End Function

Public Function m_TryGetStringValue( _
    ByVal scopeValue As obj_ScriptScopeValue, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    If scopeValue Is Nothing Then
        outErrorText = "Scope value container is missing."
        Exit Function
    End If
    If scopeValue.HasObjectValue Then
        outErrorText = "Scope value kind '" & scopeValue.Kind & "' is object and cannot be rendered as string."
        Exit Function
    End If

    outValue = scopeValue.TextValue
    m_TryGetStringValue = True
End Function
