VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ResultTable"
Option Explicit

Private m_TableRef As String
Private m_Rows As Collection
Private m_RowsByIndex As Object
Private m_FieldMapByAlias As Object
Private m_MetaInfo As Object

Public Sub Initialize(ByVal tableRef As String)
    m_TableRef = Trim$(tableRef)
    Set m_Rows = New Collection
    Set m_RowsByIndex = CreateObject("Scripting.Dictionary")
    m_RowsByIndex.CompareMode = 1
    Set m_FieldMapByAlias = CreateObject("Scripting.Dictionary")
    m_FieldMapByAlias.CompareMode = 1
    Set m_MetaInfo = CreateObject("Scripting.Dictionary")
    m_MetaInfo.CompareMode = 1
    m_MetaInfo("Name") = m_TableRef
End Sub

Public Property Get TableRef() As String
    TableRef = m_TableRef
End Property

Public Property Get Rows() As Collection
    Set Rows = m_Rows
End Property

Public Property Get Count() As Long
    If m_Rows Is Nothing Then Exit Property
    Count = m_Rows.Count
End Property

Public Property Get RowCount() As Long
    RowCount = Count
End Property

Public Property Get LastRow() As obj_ResultRow
    If m_Rows Is Nothing Then Exit Property
    If m_Rows.Count = 0 Then Exit Property
    Set LastRow = m_Rows(m_Rows.Count)
End Property

Public Property Get PrevRow() As obj_ResultRow
    If m_Rows Is Nothing Then Exit Property
    If m_Rows.Count < 2 Then Exit Property
    Set PrevRow = m_Rows(m_Rows.Count - 1)
End Property

Public Sub AddFieldMap(ByVal fieldAlias As String, ByVal mapKey As String)
    m_FieldMapByAlias(Trim$(fieldAlias)) = Trim$(mapKey)
End Sub

Public Function HasFieldAlias(ByVal fieldAlias As String) As Boolean
    HasFieldAlias = m_FieldMapByAlias.Exists(Trim$(fieldAlias))
End Function

Public Function MapKeyByAlias(ByVal fieldAlias As String) As String
    MapKeyByAlias = CStr(m_FieldMapByAlias(Trim$(fieldAlias)))
End Function

Public Property Get FieldMapByAlias() As Object
    Set FieldMapByAlias = m_FieldMapByAlias
End Property

Public Property Get MetaInfo() As Object
    If m_MetaInfo Is Nothing Then
        Set m_MetaInfo = CreateObject("Scripting.Dictionary")
        m_MetaInfo.CompareMode = 1
        m_MetaInfo("Name") = m_TableRef
    End If
    Set MetaInfo = m_MetaInfo
End Property

Public Sub SetMetaInfoValue(ByVal keyName As String, ByVal value As Variant)
    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then Exit Sub
    If m_MetaInfo Is Nothing Then
        Set m_MetaInfo = CreateObject("Scripting.Dictionary")
        m_MetaInfo.CompareMode = 1
    End If
    m_MetaInfo(keyName) = value
End Sub

Public Function GetMetaInfoValue(ByVal keyName As String, Optional ByVal defaultValue As String = vbNullString) As String
    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then
        GetMetaInfoValue = defaultValue
        Exit Function
    End If
    If m_MetaInfo Is Nothing Then
        GetMetaInfoValue = defaultValue
        Exit Function
    End If
    If Not m_MetaInfo.Exists(keyName) Then
        GetMetaInfoValue = defaultValue
        Exit Function
    End If
    GetMetaInfoValue = CStr(m_MetaInfo(keyName))
End Function

Public Function HasRow(ByVal rowIndex As Long) As Boolean
    HasRow = m_RowsByIndex.Exists(CStr(rowIndex))
End Function

Public Function RowByIndex(ByVal rowIndex As Long) As obj_ResultRow
    Set RowByIndex = m_RowsByIndex(CStr(rowIndex))
End Function

Public Property Get Row(ByVal rowIndex As Long) As obj_ResultRow
    If rowIndex < 0 Then Exit Property
    If m_Rows Is Nothing Then Exit Property
    If rowIndex + 1 > m_Rows.Count Then Exit Property
    Set Row = m_Rows(rowIndex + 1)
End Property

Public Function EnsureRow(ByVal rowIndex As Long) As obj_ResultRow
    Dim rowObj As obj_ResultRow
    If Not m_RowsByIndex.Exists(CStr(rowIndex)) Then
        Set rowObj = New obj_ResultRow
        rowObj.Initialize rowIndex
        m_RowsByIndex.Add CStr(rowIndex), rowObj
        m_Rows.Add rowObj
    End If
    Set EnsureRow = m_RowsByIndex(CStr(rowIndex))
End Function

Public Sub SetRowValue(ByVal rowIndex As Long, ByVal fieldAlias As String, ByVal mapKey As String, ByVal valueText As String)
    Dim rowObj As obj_ResultRow
    Set rowObj = EnsureRow(rowIndex)
    rowObj.SetValue fieldAlias, mapKey, valueText
End Sub
