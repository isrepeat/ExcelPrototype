VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_DynamicTableSchema"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

' Lightweight shared schema for obj_TableDynamic rows. Every row references the
' same object, so column metadata is not copied per record and no row -> table
' reference cycle is created.
Private m_Columns As list__obj_Column
Private m_AliasToIndex As Object
Private m_NameToIndex As Object
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
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

Public Function Initialize(ByVal columns As list__obj_Column) As Boolean
    m_IsDisposed = False
    Initialize = Me.BindColumns(columns)
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    Set m_Columns = Nothing
    Set m_AliasToIndex = Nothing
    Set m_NameToIndex = Nothing
End Sub

' Rebind is needed only when InsertColumnAt replaces the columns collection.
' Individual column aliases remain live because schema and table share objects.
Public Function BindColumns(ByVal columns As list__obj_Column) As Boolean
    If columns Is Nothing Then Exit Function
    Set m_Columns = columns
    m_IsDisposed = False
    If Not private_RebuildIndexes() Then Exit Function
    BindColumns = True
End Function

Public Function TryGetColumnIndex( _
    ByVal aliasOrName As String, _
    ByRef outIndex As Long _
) As Boolean
    Dim columnIndex As Long
    Dim columnObj As obj_Column
    Dim normalizedName As String
    Dim normalizedAlias As String

    outIndex = 0
    aliasOrName = VBA.Trim$(VBA.CStr(aliasOrName))
    If VBA.Len(aliasOrName) = 0 Then Exit Function
    If m_Columns Is Nothing Then Exit Function

    normalizedAlias = private_NormalizeAliasKey(aliasOrName)
    If Not m_AliasToIndex Is Nothing Then
        If m_AliasToIndex.Exists(normalizedAlias) Then
            outIndex = VBA.CLng(m_AliasToIndex(normalizedAlias))
            TryGetColumnIndex = True
            Exit Function
        End If
    End If

    normalizedName = private_NormalizeText(aliasOrName)
    If Not m_NameToIndex Is Nothing Then
        If m_NameToIndex.Exists(normalizedName) Then
            outIndex = VBA.CLng(m_NameToIndex(normalizedName))
            TryGetColumnIndex = True
            Exit Function
        End If
    End If

    ' Fallback preserves correctness if caller added an alias directly through
    ' table.Columns after the last schema rebuild. A hit is memoized.
    ' Alias has priority, matching obj_TableDynamic lookup semantics.
    For columnIndex = 1 To m_Columns.Count
        Set columnObj = m_Columns.Item(columnIndex)
        If Not columnObj Is Nothing Then
            If columnObj.HasAlias(aliasOrName) Then
                outIndex = columnIndex
                If Not m_AliasToIndex.Exists(normalizedAlias) Then m_AliasToIndex.Add normalizedAlias, columnIndex
                TryGetColumnIndex = True
                Exit Function
            End If
        End If
    Next columnIndex

    For columnIndex = 1 To m_Columns.Count
        Set columnObj = m_Columns.Item(columnIndex)
        If Not columnObj Is Nothing Then
            If VBA.StrComp(private_NormalizeText(columnObj.Name), normalizedName, VBA.vbTextCompare) = 0 Then
                outIndex = columnIndex
                If Not m_NameToIndex.Exists(normalizedName) Then m_NameToIndex.Add normalizedName, columnIndex
                TryGetColumnIndex = True
                Exit Function
            End If
        End If
    Next columnIndex
End Function

Private Function private_RebuildIndexes() As Boolean
    Dim columnIndex As Long
    Dim columnObj As obj_Column
    Dim aliases As Collection
    Dim aliasItem As Variant
    Dim normalizedKey As String

    Set m_AliasToIndex = VBA.CreateObject("Scripting.Dictionary")
    m_AliasToIndex.CompareMode = 1
    Set m_NameToIndex = VBA.CreateObject("Scripting.Dictionary")
    m_NameToIndex.CompareMode = 1

    If m_Columns Is Nothing Then Exit Function
    For columnIndex = 1 To m_Columns.Count
        Set columnObj = m_Columns.Item(columnIndex)
        If columnObj Is Nothing Then GoTo ContinueColumn

        normalizedKey = private_NormalizeText(columnObj.Name)
        If VBA.Len(normalizedKey) > 0 Then
            If Not m_NameToIndex.Exists(normalizedKey) Then m_NameToIndex.Add normalizedKey, columnIndex
        End If

        Set aliases = columnObj.Aliases
        If Not aliases Is Nothing Then
            For Each aliasItem In aliases
                normalizedKey = private_NormalizeAliasKey(VBA.CStr(aliasItem))
                If VBA.Len(normalizedKey) > 0 Then
                    If Not m_AliasToIndex.Exists(normalizedKey) Then m_AliasToIndex.Add normalizedKey, columnIndex
                End If
            Next aliasItem
        End If
ContinueColumn:
    Next columnIndex

    private_RebuildIndexes = True
End Function

Private Function private_NormalizeAliasKey(ByVal valueText As String) As String
    private_NormalizeAliasKey = VBA.LCase$(VBA.Trim$(VBA.CStr(valueText)))
End Function

Private Function private_NormalizeText(ByVal valueText As String) As String
    valueText = VBA.CStr(valueText)
    valueText = VBA.Replace$(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace$(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace$(valueText, VBA.vbTab, " ")
    valueText = VBA.Replace$(valueText, VBA.ChrW$(160), " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace$(valueText, "  ", " ")
    Loop
    private_NormalizeText = VBA.LCase$(VBA.Trim$(valueText))
End Function
