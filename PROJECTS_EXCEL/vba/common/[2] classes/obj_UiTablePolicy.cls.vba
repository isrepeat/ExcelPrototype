VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiTablePolicy"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_rules As Collection
Private m_default As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByVal definition As Object) As Boolean
    Dim policy As Object
    Dim node As Object
    Dim rule As Object
    Dim intervals As Collection
    Dim interval As Variant
    Dim mode As String
    Dim tableName As String

    If m_isInitialized Or m_isDisposed Then Exit Function
    Set m_rules = New Collection
    m_default = ex_UiElementFactory.fn_Attribute(definition, "representation")
    If VBA.Len(m_default) = 0 Then m_default = "Raw"
    Set policy = definition.SelectSingleNode("*[local-name()='tableList.tablePolicy']/*[local-name()='tablePolicy']")
    If Not policy Is Nothing Then
        mode = ex_UiElementFactory.fn_Attribute(policy, "defaultRepresentation")
        If VBA.Len(mode) > 0 Then m_default = mode
        For Each node In policy.SelectNodes("*[local-name()='rule']")
            mode = ex_UiElementFactory.fn_Attribute(node, "representation")
            private_ValidateMode mode
            tableName = ex_UiElementFactory.fn_Attribute(node, "excelTableName")
            Set intervals = Me.ParseIndices(ex_UiElementFactory.fn_Attribute(node, "index"))
            If VBA.Len(tableName) > 0 Then
                If mode <> "Smart" Or intervals.Count <> 1 Then
                    Err.Raise VBA.vbObjectError + 2280, , "excelTableName requires one Smart table index."
                End If
                interval = intervals(1)
                If interval(0) <> interval(1) Then
                    Err.Raise VBA.vbObjectError + 2280, , "excelTableName requires one Smart table index."
                End If
            End If
            Set rule = VBA.CreateObject("Scripting.Dictionary")
            rule.Add "intervals", Empty
            Set rule("intervals") = intervals
            rule.Add "representation", mode
            rule.Add "excelTableName", tableName
            m_rules.Add rule
        Next node
    End If
    private_ValidateMode m_default
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    Set m_rules = Nothing
End Sub

Public Sub Resolve( _
    ByVal index As Long, _
    ByRef representation As String, _
    ByRef excelTableName As String _
)
    Dim rule As Object
    Dim interval As Variant
    Dim intervals As Collection

    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise VBA.vbObjectError + 2280, , "Table policy is not initialized."
    End If
    representation = m_default
    excelTableName = VBA.vbNullString
    For Each rule In m_rules
        Set intervals = rule("intervals")
        For Each interval In intervals
            If index >= interval(0) And index <= interval(1) Then
                representation = rule("representation")
                If representation = "Raw" Then excelTableName = VBA.vbNullString
                If VBA.Len(rule("excelTableName")) > 0 Then excelTableName = rule("excelTableName")
                Exit For
            End If
        Next interval
    Next rule
End Sub

Public Function ParseIndices(ByVal expression As String) As Collection
    Dim result As New Collection
    Dim token As Variant
    Dim bounds As Variant
    Dim first As Long
    Dim last As Long

    expression = VBA.Trim$(expression)
    If VBA.Len(expression) = 0 Then GoTo InvalidIndex
    For Each token In VBA.Split(expression, ",")
        bounds = VBA.Split(VBA.Trim$(VBA.CStr(token)), "-")
        If UBound(bounds) > 1 Then GoTo InvalidIndex
        first = private_PositiveIndex(VBA.Trim$(bounds(0)))
        last = first
        If UBound(bounds) = 1 Then last = private_PositiveIndex(VBA.Trim$(bounds(1)))
        If last < first Then GoTo InvalidIndex
        result.Add VBA.Array(first, last)
    Next token
    Set ParseIndices = result
    Exit Function
InvalidIndex:
    Err.Raise VBA.vbObjectError + 2280, , "Invalid table index selector: " & expression
End Function

' //
' // Private
' //
Private Function private_PositiveIndex(ByVal value As String) As Long
    Dim position As Long
    Dim character As String
    Dim number As Double

    If VBA.Len(value) = 0 Or VBA.Len(value) > 10 Then GoTo InvalidIndex
    For position = 1 To VBA.Len(value)
        character = VBA.Mid$(value, position, 1)
        If character < "0" Or character > "9" Then GoTo InvalidIndex
    Next position
    number = VBA.CDbl(value)
    If number < 1 Or number > 2147483647# Then GoTo InvalidIndex
    private_PositiveIndex = VBA.CLng(number)
    Exit Function
InvalidIndex:
    Err.Raise VBA.vbObjectError + 2280, , "Invalid positive table index: " & value
End Function

Private Sub private_ValidateMode(ByVal value As String)
    If value <> "Raw" And value <> "Smart" Then
        Err.Raise VBA.vbObjectError + 2280, , "Invalid table representation: " & value
    End If
End Sub