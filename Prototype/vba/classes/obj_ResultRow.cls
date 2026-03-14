VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ResultRow"
Option Explicit

Private m_RowIndex As Long
Private m_RowAnchorName As String
Private m_Kind As String
Private m_ValuesByAlias As Object
Private m_ValuesByMapKey As Object
Private m_MapKeyByAlias As Object

Public Sub Initialize(ByVal rowIndex As Long)
    m_RowIndex = rowIndex
    Set m_ValuesByAlias = CreateObject("Scripting.Dictionary")
    m_ValuesByAlias.CompareMode = 1
    Set m_ValuesByMapKey = CreateObject("Scripting.Dictionary")
    m_ValuesByMapKey.CompareMode = 1
    Set m_MapKeyByAlias = CreateObject("Scripting.Dictionary")
    m_MapKeyByAlias.CompareMode = 1
End Sub

Public Property Get RowIndex() As Long
    RowIndex = m_RowIndex
End Property

Public Property Get RowAnchorName() As String
    RowAnchorName = m_RowAnchorName
End Property

Public Property Let RowAnchorName(ByVal valueText As String)
    m_RowAnchorName = Trim$(valueText)
End Property

Public Property Get Kind() As String
    Kind = m_Kind
End Property

Public Property Let Kind(ByVal valueText As String)
    m_Kind = Trim$(valueText)
End Property

Public Sub SetValue(ByVal fieldAlias As String, ByVal mapKey As String, ByVal valueText As String)
    If m_ValuesByAlias Is Nothing Then Initialize m_RowIndex
    m_ValuesByAlias(Trim$(fieldAlias)) = valueText
    m_ValuesByMapKey(Trim$(mapKey)) = valueText
    m_MapKeyByAlias(Trim$(fieldAlias)) = Trim$(mapKey)
End Sub

Public Function HasAlias(ByVal fieldAlias As String) As Boolean
    If m_ValuesByAlias Is Nothing Then Exit Function
    HasAlias = m_ValuesByAlias.Exists(Trim$(fieldAlias))
End Function

Public Function GetByAlias(ByVal fieldAlias As String) As String
    Dim availableAliases As String

    fieldAlias = Trim$(fieldAlias)
    If Len(fieldAlias) = 0 Then
        Err.Raise vbObjectError + 1710, "obj_ResultRow", "Field alias is empty."
    End If
    If m_ValuesByAlias Is Nothing Then
        Err.Raise vbObjectError + 1711, "obj_ResultRow", "Row storage is not initialized."
    End If
    If Not m_ValuesByAlias.Exists(fieldAlias) Then
        availableAliases = mp_JoinDictionaryKeys(m_ValuesByAlias)
        Err.Raise vbObjectError + 1712, "obj_ResultRow", _
            "Unknown field alias '" & fieldAlias & "' at row index " & CStr(m_RowIndex) & _
            ". Available aliases: " & availableAliases
    End If
    GetByAlias = CStr(m_ValuesByAlias(fieldAlias))
End Function

Public Function MapKeyByAlias(ByVal fieldAlias As String) As String
    Dim availableAliases As String

    fieldAlias = Trim$(fieldAlias)
    If Len(fieldAlias) = 0 Then
        Err.Raise vbObjectError + 1713, "obj_ResultRow", "Field alias is empty for map key resolve."
    End If
    If m_MapKeyByAlias Is Nothing Then
        Err.Raise vbObjectError + 1714, "obj_ResultRow", "Row map key storage is not initialized."
    End If
    If Not m_MapKeyByAlias.Exists(fieldAlias) Then
        availableAliases = mp_JoinDictionaryKeys(m_MapKeyByAlias)
        Err.Raise vbObjectError + 1715, "obj_ResultRow", _
            "Unknown field alias '" & fieldAlias & "' for map key resolve at row index " & CStr(m_RowIndex) & _
            ". Available aliases: " & availableAliases
    End If
    MapKeyByAlias = CStr(m_MapKeyByAlias(fieldAlias))
End Function

Public Property Get Column(ByVal fieldAlias As String) As String
    Column = GetByAlias(fieldAlias)
End Property

Public Function HasMapKey(ByVal mapKey As String) As Boolean
    If m_ValuesByMapKey Is Nothing Then Exit Function
    HasMapKey = m_ValuesByMapKey.Exists(Trim$(mapKey))
End Function

Public Function GetByMapKey(ByVal mapKey As String) As String
    Dim availableMapKeys As String

    mapKey = Trim$(mapKey)
    If Len(mapKey) = 0 Then
        Err.Raise vbObjectError + 1716, "obj_ResultRow", "Map key is empty."
    End If
    If m_ValuesByMapKey Is Nothing Then
        Err.Raise vbObjectError + 1717, "obj_ResultRow", "Row map key storage is not initialized."
    End If
    If Not m_ValuesByMapKey.Exists(mapKey) Then
        availableMapKeys = mp_JoinDictionaryKeys(m_ValuesByMapKey)
        Err.Raise vbObjectError + 1718, "obj_ResultRow", _
            "Unknown map key '" & mapKey & "' at row index " & CStr(m_RowIndex) & _
            ". Available map keys: " & availableMapKeys
    End If
    GetByMapKey = CStr(m_ValuesByMapKey(mapKey))
End Function

Public Property Get Columns() As Collection
    Dim result As Collection
    Dim fieldAlias As Variant
    Dim columnObj As obj_ResultColumn
    Dim aliasText As String

    Set result = New Collection
    If m_ValuesByAlias Is Nothing Then
        Set Columns = result
        Exit Property
    End If

    For Each fieldAlias In m_ValuesByAlias.Keys
        aliasText = CStr(fieldAlias)
        Set columnObj = New obj_ResultColumn
        columnObj.Initialize aliasText, MapKeyByAlias(aliasText), CStr(m_ValuesByAlias(fieldAlias))
        result.Add columnObj
    Next fieldAlias

    Set Columns = result
End Property

Private Function mp_JoinDictionaryKeys(ByVal dict As Object) As String
    Dim key As Variant
    Dim result As String

    If dict Is Nothing Then
        mp_JoinDictionaryKeys = "(none)"
        Exit Function
    End If
    If dict.Count = 0 Then
        mp_JoinDictionaryKeys = "(none)"
        Exit Function
    End If

    For Each key In dict.Keys
        If Len(result) > 0 Then result = result & ", "
        result = result & CStr(key)
    Next key

    mp_JoinDictionaryKeys = result
End Function
