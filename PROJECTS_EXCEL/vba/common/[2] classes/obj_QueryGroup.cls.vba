VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_QueryGroup"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_items As Collection
Public MatchAny As Boolean

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_items = New Collection
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_items = New Collection
    MatchAny = False
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_items = Nothing
    MatchAny = False
End Sub

Public Sub AddCondition(ByVal condition As obj_QueryCondition)
    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    If condition Is Nothing Then
        Err.Raise VBA.vbObjectError + 2115, , "Condition is required."
    End If
    m_items.Add condition
End Sub

Public Sub AddGroup(ByVal group As obj_QueryGroup)
    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    If group Is Nothing Then
        Err.Raise VBA.vbObjectError + 2116, , "Group is required."
    End If
    If group.ContainsGroup(Me) Then
        Err.Raise VBA.vbObjectError + 2117, , "Cyclic query group."
    End If
    m_items.Add group
End Sub

Public Function ContainsGroup(ByVal target As obj_QueryGroup) As Boolean
    Dim item As Object

    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    If Me Is target Then
        ContainsGroup = True
        Exit Function
    End If
    For Each item In m_items
        If TypeOf item Is obj_QueryGroup Then
            If item.ContainsGroup(target) Then
                ContainsGroup = True
                Exit Function
            End If
        End If
    Next item
End Function

Public Sub Validate(ByVal headers As Variant)
    Dim item As Object

    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    For Each item In m_items
        item.Validate headers
    Next item
End Sub

Public Sub CollectColumns(ByVal names As Collection)
    Dim item As Object

    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    For Each item In m_items
        If TypeOf item Is obj_QueryGroup Then
            item.CollectColumns names
        Else
            names.Add item.ColumnName
        End If
    Next item
End Sub

Public Function Matches( _
    ByVal values As Variant, _
    ByVal row As Long, _
    ByVal headers As Variant _
) As Boolean
    Dim item As Object
    Dim matched As Boolean

    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    Matches = Not MatchAny
    For Each item In m_items
        If TypeOf item Is obj_QueryGroup Then
            matched = item.Matches(values, row, headers)
        Else
            matched = item.MatchesRow(values, row)
        End If
        If MatchAny And matched Then
            Matches = True
            Exit Function
        End If
        If Not MatchAny And Not matched Then
            Matches = False
            Exit Function
        End If
    Next item
End Function