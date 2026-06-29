VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableData"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_Values As Variant
Private m_RowCount As Long
Private m_ColumnCount As Long
Private m_AliasToIndex As Object
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_AliasToIndex = VBA.CreateObject("Scripting.Dictionary")
    m_AliasToIndex.CompareMode = 1
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

Public Property Get RowCount() As Long
    RowCount = m_RowCount
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = m_ColumnCount
End Property

Public Function Initialize( _
    ByRef values As Variant, _
    ByVal rowCount As Long, _
    ByVal columnCount As Long, _
    ByVal columnAliases As Collection _
) As Boolean
    Dim colIndex As Long
    Dim aliasText As String

    m_Values = Empty
    m_RowCount = 0
    m_ColumnCount = 0
    Set m_AliasToIndex = VBA.CreateObject("Scripting.Dictionary")
    m_AliasToIndex.CompareMode = 1

    If rowCount < 0 Or columnCount < 0 Then Exit Function
    If rowCount > 0 Then
        If IsEmpty(values) Then Exit Function
    End If

    m_Values = values
    m_RowCount = rowCount
    m_ColumnCount = columnCount

    If Not columnAliases Is Nothing Then
        For colIndex = 1 To columnAliases.Count
            aliasText = VBA.Trim$(VBA.CStr(columnAliases.Item(colIndex)))
            If VBA.Len(aliasText) > 0 Then
                If Not m_AliasToIndex.Exists(aliasText) Then m_AliasToIndex.Add aliasText, colIndex
            End If
        Next colIndex
    End If

    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    m_Values = Empty
    Set m_AliasToIndex = Nothing
    m_RowCount = 0
    m_ColumnCount = 0
    On Error GoTo 0
End Sub

Public Function ValueAt(ByVal rowIndex As Long, ByVal colIndex As Long) As String
    If rowIndex <= 0 Or rowIndex > m_RowCount Then Exit Function
    If colIndex <= 0 Or colIndex > m_ColumnCount Then Exit Function
    ValueAt = VBA.CStr(m_Values(rowIndex, colIndex))
End Function

Public Function TryGetColumnIndexByAlias(ByVal aliasText As String, ByRef outColumnIndex As Long) As Boolean
    outColumnIndex = 0
    If m_AliasToIndex Is Nothing Then Exit Function
    aliasText = VBA.Trim$(VBA.CStr(aliasText))
    If VBA.Len(aliasText) = 0 Then Exit Function
    If Not m_AliasToIndex.Exists(aliasText) Then Exit Function

    outColumnIndex = CLng(m_AliasToIndex(aliasText))
    TryGetColumnIndexByAlias = (outColumnIndex > 0)
End Function
