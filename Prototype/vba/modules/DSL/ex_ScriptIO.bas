Attribute VB_Name = "ex_ScriptIO"
Option Explicit

Private g_Input As Object
Private g_Output As Object

Public Sub m_ResetContext(Optional ByVal inputObject As Object = Nothing)
    m_SetInput inputObject
    Set g_Output = Nothing
End Sub

Public Sub m_SetInput(ByVal inputObject As Object)
    If inputObject Is Nothing Then
        Set g_Input = New obj_ScriptIOPayload
    Else
        Set g_Input = inputObject
    End If
End Sub

Public Function m_GetInput() As Object
    If g_Input Is Nothing Then Set g_Input = New obj_ScriptIOPayload
    Set m_GetInput = g_Input
End Function

Public Function m_CreateOutput() As Object
    Set g_Output = New obj_ScriptIOPayload
    Set m_CreateOutput = g_Output
End Function

Public Function m_GetLastOutput() As Object
    Set m_GetLastOutput = g_Output
End Function

Public Function m_CreateCollection() As Object
    Dim result As Collection

    Set result = New Collection
    Set m_CreateCollection = result
End Function

Public Function m_CreateDictionary() As Object
    Dim result As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    Set m_CreateDictionary = result
End Function

Public Function m_CreateObjectBag() As Object
    Set m_CreateObjectBag = m_CreateDictionary()
End Function

Public Function m_CollectionAddString(ByVal targetCollection As Object, ByVal valueText As String) As String
    If targetCollection Is Nothing Then Exit Function
    If TypeName(targetCollection) <> "Collection" Then
        Err.Raise vbObjectError + 6270, "ex_ScriptIO", "Target must be Collection for m_CollectionAddString."
    End If

    targetCollection.Add CStr(valueText)
    m_CollectionAddString = CStr(valueText)
End Function

Public Function m_CollectionAddObject(ByVal targetCollection As Object, ByVal valueObject As Object) As String
    If targetCollection Is Nothing Then Exit Function
    If valueObject Is Nothing Then Exit Function
    If TypeName(targetCollection) <> "Collection" Then
        Err.Raise vbObjectError + 6271, "ex_ScriptIO", "Target must be Collection for m_CollectionAddObject."
    End If

    targetCollection.Add valueObject
    m_CollectionAddObject = TypeName(valueObject)
End Function

Public Function m_DictionarySetString(ByVal targetDictionary As Object, ByVal keyName As String, ByVal valueText As String) As String
    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If targetDictionary Is Nothing Then Exit Function
    If Not mp_IsDictionary(targetDictionary) Then
        Err.Raise vbObjectError + 6272, "ex_ScriptIO", "Target must be Dictionary for m_DictionarySetString."
    End If

    targetDictionary(keyName) = CStr(valueText)
    m_DictionarySetString = CStr(valueText)
End Function

Public Function m_DictionarySetObject(ByVal targetDictionary As Object, ByVal keyName As String, ByVal valueObject As Object) As String
    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If targetDictionary Is Nothing Then Exit Function
    If valueObject Is Nothing Then Exit Function
    If Not mp_IsDictionary(targetDictionary) Then
        Err.Raise vbObjectError + 6273, "ex_ScriptIO", "Target must be Dictionary for m_DictionarySetObject."
    End If

    Set targetDictionary(keyName) = valueObject
    m_DictionarySetObject = TypeName(valueObject)
End Function

Public Function m_SetString(ByVal target As Object, ByVal keyName As String, ByVal valueText As String) As String
    Dim payload As obj_ScriptIOPayload

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If target Is Nothing Then Exit Function

    If mp_TryAsPayload(target, payload) Then
        payload.m_SetString keyName, CStr(valueText)
    Else
        target(keyName) = CStr(valueText)
    End If

    m_SetString = CStr(valueText)
End Function

Public Function m_SetObject(ByVal target As Object, ByVal keyName As String, ByVal valueObject As Object) As String
    Dim payload As obj_ScriptIOPayload

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If target Is Nothing Then Exit Function
    If valueObject Is Nothing Then Exit Function

    If mp_TryAsPayload(target, payload) Then
        payload.m_SetObject keyName, valueObject
    Else
        Set target(keyName) = valueObject
    End If
End Function

Public Function m_GetStringOrDefault( _
    ByVal sourceContainer As Object, _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim payload As obj_ScriptIOPayload
    Dim scopeValue As obj_ScriptScopeValue
    Dim rawValue As Variant
    Dim rawObject As Object

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then
        m_GetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If
    If sourceContainer Is Nothing Then
        m_GetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If

    If mp_TryAsPayload(sourceContainer, payload) Then
        If payload.m_TryGetString(keyName, m_GetStringOrDefault) Then Exit Function
        m_GetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If

    If Not mp_IsDictionary(sourceContainer) Then
        m_GetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If

    If Not sourceContainer.Exists(keyName) Then
        m_GetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If

    Set rawObject = Nothing
    On Error Resume Next
    Set rawObject = sourceContainer(keyName)
    On Error GoTo 0

    If Not rawObject Is Nothing Then
        If TypeOf rawObject Is obj_ScriptScopeValue Then
            Set scopeValue = rawObject
            If scopeValue.HasObjectValue Then
                m_GetStringOrDefault = CStr(defaultValue)
            Else
                m_GetStringOrDefault = CStr(scopeValue.TextValue)
            End If
            Exit Function
        End If

        m_GetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If

    rawValue = sourceContainer(keyName)

    m_GetStringOrDefault = CStr(rawValue)
End Function

Public Function m_GetInputStringOrDefault( _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    m_GetInputStringOrDefault = m_GetStringOrDefault(m_GetInput(), keyName, defaultValue)
End Function

Public Function m_GetValue( _
    ByVal sourceContainer As Object, _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    m_GetValue = m_GetStringOrDefault(sourceContainer, keyName, defaultValue)
End Function

Public Function m_TryGetValue( _
    ByVal sourceContainer As Object, _
    ByVal keyName As String, _
    ByRef outValue As String _
) As Boolean
    Dim payload As obj_ScriptIOPayload
    Dim scopeValue As obj_ScriptScopeValue
    Dim rawValue As Variant
    Dim rawObject As Object

    outValue = vbNullString
    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If sourceContainer Is Nothing Then Exit Function

    If mp_TryAsPayload(sourceContainer, payload) Then
        m_TryGetValue = payload.m_TryGetString(keyName, outValue)
        Exit Function
    End If

    If Not mp_IsDictionary(sourceContainer) Then Exit Function

    If Not sourceContainer.Exists(keyName) Then Exit Function
    Set rawObject = Nothing
    On Error Resume Next
    Set rawObject = sourceContainer(keyName)
    On Error GoTo 0

    If Not rawObject Is Nothing Then
        If TypeOf rawObject Is obj_ScriptScopeValue Then
            Set scopeValue = rawObject
            If scopeValue.HasObjectValue Then Exit Function
            outValue = CStr(scopeValue.TextValue)
            m_TryGetValue = True
            Exit Function
        End If
        Exit Function
    End If

    rawValue = sourceContainer(keyName)

    outValue = CStr(rawValue)
    m_TryGetValue = True
End Function

Public Function m_GetByPath( _
    ByVal sourceContainer As Object, _
    ByVal pathText As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim parts() As String
    Dim i As Long
    Dim segment As String
    Dim currentObject As Object
    Dim nextObject As Object
    Dim valueText As String
    Dim indexValue As Long

    pathText = Trim$(pathText)
    If Len(pathText) = 0 Then
        m_GetByPath = CStr(defaultValue)
        Exit Function
    End If
    If sourceContainer Is Nothing Then
        m_GetByPath = CStr(defaultValue)
        Exit Function
    End If

    parts = Split(pathText, ".")
    Set currentObject = sourceContainer

    For i = LBound(parts) To UBound(parts)
        segment = Trim$(CStr(parts(i)))
        If Len(segment) = 0 Then
            m_GetByPath = CStr(defaultValue)
            Exit Function
        End If

        If i = UBound(parts) Then
            If TypeName(currentObject) = "Collection" And mp_TryParseCollectionIndex(segment, indexValue) Then
                m_GetByPath = m_CollectionGetStringOrDefault(currentObject, indexValue, defaultValue)
                Exit Function
            End If

            If m_TryGetValue(currentObject, segment, valueText) Then
                m_GetByPath = valueText
            Else
                m_GetByPath = CStr(defaultValue)
            End If
            Exit Function
        End If

        Set nextObject = Nothing
        If TypeName(currentObject) = "Collection" And mp_TryParseCollectionIndex(segment, indexValue) Then
            If Not mp_TryCollectionGetObject(currentObject, indexValue, nextObject) Then
                m_GetByPath = CStr(defaultValue)
                Exit Function
            End If
        Else
            If Not m_TryGetObject(currentObject, segment, nextObject) Then
                m_GetByPath = CStr(defaultValue)
                Exit Function
            End If
        End If

        Set currentObject = nextObject
    Next i

    m_GetByPath = CStr(defaultValue)
End Function

Public Function m_GetObject(ByVal sourceContainer As Object, ByVal keyName As String) As Object
    Dim valueObject As Object

    If Not m_TryGetObject(sourceContainer, keyName, valueObject) Then
        Err.Raise vbObjectError + 6281, "ex_ScriptIO", "Object field '" & keyName & "' was not found or is not an object."
    End If

    Set m_GetObject = valueObject
End Function

Public Function m_GetInputObject(ByVal keyName As String) As Object
    Dim inputObject As Object

    If Not m_TryGetObject(m_GetInput(), keyName, inputObject) Then
        Err.Raise vbObjectError + 6274, "ex_ScriptIO", "Input object field '" & keyName & "' was not found or is not an object."
    End If

    Set m_GetInputObject = inputObject
End Function

Public Function m_CollectionCount(ByVal sourceCollection As Object) As Long
    If sourceCollection Is Nothing Then Exit Function
    If TypeName(sourceCollection) <> "Collection" Then
        Err.Raise vbObjectError + 6282, "ex_ScriptIO", "Target must be Collection for m_CollectionCount."
    End If

    m_CollectionCount = CLng(sourceCollection.Count)
End Function

Public Function m_CollectionGetObject(ByVal sourceCollection As Object, ByVal indexValue As Long) As Object
    Dim valueObject As Object

    If Not mp_TryCollectionGetObject(sourceCollection, indexValue, valueObject) Then
        Err.Raise vbObjectError + 6283, "ex_ScriptIO", "Collection index " & CStr(indexValue) & " is missing or does not contain an object."
    End If

    Set m_CollectionGetObject = valueObject
End Function

Public Function m_CollectionGetStringOrDefault( _
    ByVal sourceCollection As Object, _
    ByVal indexValue As Long, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim valueRef As Variant

    If sourceCollection Is Nothing Then
        m_CollectionGetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If
    If TypeName(sourceCollection) <> "Collection" Then
        Err.Raise vbObjectError + 6284, "ex_ScriptIO", "Target must be Collection for m_CollectionGetStringOrDefault."
    End If
    If indexValue < 1 Or indexValue > sourceCollection.Count Then
        m_CollectionGetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If

    valueRef = sourceCollection(indexValue)
    If IsObject(valueRef) Then
        m_CollectionGetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If

    m_CollectionGetStringOrDefault = CStr(valueRef)
End Function

Public Function m_TryGetObject( _
    ByVal sourceContainer As Object, _
    ByVal keyName As String, _
    ByRef outObject As Object _
) As Boolean
    Dim payload As obj_ScriptIOPayload
    Dim scopeValue As obj_ScriptScopeValue
    Dim rawValue As Variant
    Dim rawObject As Object

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If sourceContainer Is Nothing Then Exit Function

    If mp_TryAsPayload(sourceContainer, payload) Then
        m_TryGetObject = payload.m_TryGetObject(keyName, outObject)
        Exit Function
    End If

    If Not mp_IsDictionary(sourceContainer) Then Exit Function

    If Not sourceContainer.Exists(keyName) Then Exit Function

    Set rawObject = Nothing
    On Error Resume Next
    Set rawObject = sourceContainer(keyName)
    On Error GoTo 0
    If rawObject Is Nothing Then Exit Function

    If TypeOf rawObject Is obj_ScriptScopeValue Then
        Set scopeValue = rawObject
        If Not scopeValue.HasObjectValue Then Exit Function
        Set outObject = scopeValue.ObjectValue
    Else
        Set outObject = rawObject
    End If

    m_TryGetObject = Not (outObject Is Nothing)
End Function

Public Function m_DictionaryGetStringOrDefault( _
    ByVal sourceDictionary As Object, _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim rawObject As Object
    Dim scopeValue As obj_ScriptScopeValue

    If sourceDictionary Is Nothing Then
        m_DictionaryGetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If
    If Not mp_IsDictionary(sourceDictionary) Then
        Err.Raise vbObjectError + 6275, "ex_ScriptIO", "Target must be Dictionary for m_DictionaryGetStringOrDefault."
    End If

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then
        m_DictionaryGetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If
    If Not sourceDictionary.Exists(keyName) Then
        m_DictionaryGetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If
    Set rawObject = Nothing
    On Error Resume Next
    Set rawObject = sourceDictionary(keyName)
    On Error GoTo 0
    If Not rawObject Is Nothing Then
        If TypeOf rawObject Is obj_ScriptScopeValue Then
            Set scopeValue = rawObject
            If scopeValue.HasObjectValue Then
                m_DictionaryGetStringOrDefault = CStr(defaultValue)
            Else
                m_DictionaryGetStringOrDefault = CStr(scopeValue.TextValue)
            End If
            Exit Function
        End If

        m_DictionaryGetStringOrDefault = CStr(defaultValue)
        Exit Function
    End If

    m_DictionaryGetStringOrDefault = CStr(sourceDictionary(keyName))
End Function

Public Function m_DictionaryGetObject(ByVal sourceDictionary As Object, ByVal keyName As String) As Object
    Dim valueRef As Variant
    Dim valueObject As Object

    If sourceDictionary Is Nothing Then
        Err.Raise vbObjectError + 6276, "ex_ScriptIO", "Dictionary is Nothing in m_DictionaryGetObject."
    End If
    If Not mp_IsDictionary(sourceDictionary) Then
        Err.Raise vbObjectError + 6277, "ex_ScriptIO", "Target must be Dictionary for m_DictionaryGetObject."
    End If

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then
        Err.Raise vbObjectError + 6278, "ex_ScriptIO", "Key name is empty in m_DictionaryGetObject."
    End If
    If Not sourceDictionary.Exists(keyName) Then
        Err.Raise vbObjectError + 6279, "ex_ScriptIO", "Dictionary key '" & keyName & "' was not found."
    End If

    Set valueObject = Nothing
    On Error Resume Next
    Set valueObject = sourceDictionary(keyName)
    On Error GoTo 0
    If valueObject Is Nothing Then
        valueRef = sourceDictionary(keyName)
        Err.Raise vbObjectError + 6280, "ex_ScriptIO", "Dictionary key '" & keyName & "' is not an object."
    End If

    Set m_DictionaryGetObject = valueObject
End Function

Public Function m_SetScopeString(ByVal target As Object, ByVal keyName As String, ByVal valueText As String) As String
    Dim scopeValue As obj_ScriptScopeValue
    Dim payload As obj_ScriptIOPayload

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If target Is Nothing Then Exit Function

    If mp_TryAsPayload(target, payload) Then
        payload.m_SetString keyName, CStr(valueText)
    Else
        Set scopeValue = ex_ScriptScopeValue.m_CreateStringValue(CStr(valueText))
        If mp_IsDictionary(target) Then
            If target.Exists(keyName) Then
                target(keyName) = scopeValue
            Else
                target.Add keyName, scopeValue
            End If
        Else
            Set target(keyName) = scopeValue
        End If
    End If

    m_SetScopeString = CStr(valueText)
End Function

Public Function m_SetScopeObject(ByVal target As Object, ByVal keyName As String, ByVal valueObject As Object) As String
    Dim scopeValue As obj_ScriptScopeValue
    Dim payload As obj_ScriptIOPayload

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Function
    If target Is Nothing Then Exit Function
    If valueObject Is Nothing Then Exit Function

    If mp_TryAsPayload(target, payload) Then
        payload.m_SetObject keyName, valueObject
    Else
        Set scopeValue = ex_ScriptScopeValue.m_CreateObjectValue(valueObject)
        If mp_IsDictionary(target) Then
            If target.Exists(keyName) Then
                target(keyName) = scopeValue
            Else
                target.Add keyName, scopeValue
            End If
        Else
            Set target(keyName) = scopeValue
        End If
    End If
End Function


Private Function mp_TryAsPayload(ByVal valueObject As Object, ByRef outPayload As obj_ScriptIOPayload) As Boolean
    If valueObject Is Nothing Then Exit Function
    If Not TypeOf valueObject Is obj_ScriptIOPayload Then Exit Function

    Set outPayload = valueObject
    mp_TryAsPayload = True
End Function

Private Function mp_IsDictionary(ByVal valueObject As Object) As Boolean
    Dim typeNameText As String

    If valueObject Is Nothing Then Exit Function
    typeNameText = TypeName(valueObject)
    mp_IsDictionary = (StrComp(typeNameText, "Dictionary", vbTextCompare) = 0 Or _
                       StrComp(typeNameText, "Scripting.Dictionary", vbTextCompare) = 0)
End Function

Private Function mp_TryCollectionGetObject( _
    ByVal sourceCollection As Object, _
    ByVal indexValue As Long, _
    ByRef outObject As Object _
) As Boolean
    Dim valueObject As Object

    If sourceCollection Is Nothing Then Exit Function
    If TypeName(sourceCollection) <> "Collection" Then Exit Function
    If indexValue < 1 Or indexValue > sourceCollection.Count Then Exit Function

    Set valueObject = Nothing
    On Error Resume Next
    Set valueObject = sourceCollection(indexValue)
    On Error GoTo 0
    If valueObject Is Nothing Then Exit Function

    Set outObject = valueObject
    mp_TryCollectionGetObject = Not (outObject Is Nothing)
End Function

Private Function mp_TryParseCollectionIndex(ByVal segmentText As String, ByRef outIndex As Long) As Boolean
    segmentText = Trim$(segmentText)
    If Len(segmentText) = 0 Then Exit Function
    If Not IsNumeric(segmentText) Then Exit Function

    On Error GoTo FailParse
    outIndex = CLng(segmentText)
    If outIndex < 1 Then Exit Function
    mp_TryParseCollectionIndex = True
    Exit Function

FailParse:
    outIndex = 0
End Function
