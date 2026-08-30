VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ItemIdentity"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IClonable

Private m_IsDisposed As Boolean
Private m_KeyValues As Object

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_KeyValues = ex_Helpers.fn_CreateDictionaryTextCompare()
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IClonable_Clone(Optional ByVal targetColumnCount As Long = 0) As Object
    Set obj_IClonable_Clone = Me.Clone(targetColumnCount)
End Function

' //
' // Properties
' //
Public Property Get Count() As Long
    If m_KeyValues Is Nothing Then Exit Property
    Count = m_KeyValues.Count
End Property

Public Property Get Keys() As list__obj_String
    Dim result As list__obj_String
    Dim keyObj As Variant

    Set result = New list__obj_String
    If m_KeyValues Is Nothing Then
        Set Keys = result
        Exit Property
    End If

    For Each keyObj In m_KeyValues.Keys
        If Not result.Add(VBA.CStr(keyObj)) Then Exit Property
    Next keyObj

    Set Keys = result
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    Set m_KeyValues = ex_Helpers.fn_CreateDictionaryTextCompare()
    Initialize = Not m_KeyValues Is Nothing
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    Set m_KeyValues = Nothing
    On Error GoTo 0
End Sub

Public Function Add(ByVal identityKey As String, ByVal identityValue As Variant) As Boolean
    identityKey = private_NormalizeKey(identityKey)
    If VBA.Len(identityKey) = 0 Then Exit Function
    If Not private_EnsureDictionary() Then Exit Function
    If m_KeyValues.Exists(identityKey) Then Exit Function

    Add = private_WriteValue(identityKey, identityValue, True)
End Function

Public Function SetValue(ByVal identityKey As String, ByVal identityValue As Variant) As Boolean
    identityKey = private_NormalizeKey(identityKey)
    If VBA.Len(identityKey) = 0 Then Exit Function
    If Not private_EnsureDictionary() Then Exit Function

    SetValue = private_WriteValue(identityKey, identityValue, False)
End Function

Public Function HasKey(ByVal identityKey As String) As Boolean
    identityKey = private_NormalizeKey(identityKey)
    If VBA.Len(identityKey) = 0 Then Exit Function
    If m_KeyValues Is Nothing Then Exit Function

    HasKey = m_KeyValues.Exists(identityKey)
End Function

Public Function TryGetValue(ByVal identityKey As String, ByRef outValue As Variant) As Boolean
    identityKey = private_NormalizeKey(identityKey)
    If VBA.Len(identityKey) = 0 Then Exit Function
    If m_KeyValues Is Nothing Then Exit Function
    If Not m_KeyValues.Exists(identityKey) Then Exit Function

    If VBA.IsObject(m_KeyValues(identityKey)) Then
        Set outValue = m_KeyValues(identityKey)
    Else
        outValue = m_KeyValues(identityKey)
    End If

    TryGetValue = True
End Function

Public Function TryGetString(ByVal identityKey As String, ByRef outValue As String) As Boolean
    Dim valueCandidate As Variant

    outValue = VBA.vbNullString
    If Not Me.TryGetValue(identityKey, valueCandidate) Then Exit Function
    If VBA.IsObject(valueCandidate) Then Exit Function

    outValue = VBA.Trim$(VBA.CStr(valueCandidate))
    TryGetString = True
End Function

Public Function TryGetLong(ByVal identityKey As String, ByRef outValue As Long) As Boolean
    Dim valueCandidate As Variant

    outValue = 0
    If Not Me.TryGetValue(identityKey, valueCandidate) Then Exit Function
    If VBA.IsObject(valueCandidate) Then Exit Function
    If Not VBA.IsNumeric(valueCandidate) Then Exit Function

    outValue = VBA.CLng(valueCandidate)
    TryGetLong = True
End Function

Public Function TryGetBoolean(ByVal identityKey As String, ByRef outValue As Boolean) As Boolean
    Dim valueCandidate As Variant

    outValue = False
    If Not Me.TryGetValue(identityKey, valueCandidate) Then Exit Function

    TryGetBoolean = ex_Helpers.fn_TryGetBooleanFromVariant(valueCandidate, outValue)
End Function

Public Function Remove(ByVal identityKey As String) As Boolean
    identityKey = private_NormalizeKey(identityKey)
    If VBA.Len(identityKey) = 0 Then Exit Function
    If m_KeyValues Is Nothing Then Exit Function
    If Not m_KeyValues.Exists(identityKey) Then Exit Function

    m_KeyValues.Remove identityKey
    Remove = True
End Function

Public Sub Clear()
    If Not private_EnsureDictionary() Then Exit Sub
    m_KeyValues.RemoveAll
End Sub

Public Function Clone(Optional ByVal targetColumnCount As Long = 0) As obj_ItemIdentity
    Dim result As obj_ItemIdentity
    Dim keyObj As Variant
    Dim valueCandidate As Variant
    Dim objectCandidate As Object

    Set result = New obj_ItemIdentity
    If result Is Nothing Then Exit Function

    If Not m_KeyValues Is Nothing Then
        For Each keyObj In m_KeyValues.Keys
            If VBA.IsObject(m_KeyValues(keyObj)) Then
                Set objectCandidate = m_KeyValues(keyObj)
                If objectCandidate Is Nothing Then
                    Call result.SetValue(VBA.CStr(keyObj), Empty)
                Else
                    Call result.SetValue(VBA.CStr(keyObj), objectCandidate)
                End If
            Else
                valueCandidate = m_KeyValues(keyObj)
                Call result.SetValue(VBA.CStr(keyObj), valueCandidate)
            End If
        Next keyObj
    End If

    Set Clone = result
End Function

' //
' // Internal
' //
Private Function private_EnsureDictionary() As Boolean
    If m_KeyValues Is Nothing Then
        Set m_KeyValues = ex_Helpers.fn_CreateDictionaryTextCompare()
    End If

    private_EnsureDictionary = Not m_KeyValues Is Nothing
End Function

Private Function private_NormalizeKey(ByVal identityKey As String) As String
    private_NormalizeKey = VBA.Trim$(VBA.CStr(identityKey))
End Function

Private Function private_WriteValue( _
    ByVal identityKey As String, _
    ByVal identityValue As Variant, _
    ByVal failIfExists As Boolean _
) As Boolean
    Dim exists As Boolean

    If Not private_EnsureDictionary() Then Exit Function

    exists = m_KeyValues.Exists(identityKey)
    If failIfExists And exists Then Exit Function

    If VBA.IsObject(identityValue) Then
        If identityValue Is Nothing Then
            If exists Then
                m_KeyValues(identityKey) = Empty
            Else
                m_KeyValues.Add identityKey, Empty
            End If
            private_WriteValue = True
            Exit Function
        End If

        If exists Then
            Set m_KeyValues(identityKey) = identityValue
        Else
            m_KeyValues.Add identityKey, Empty
            Set m_KeyValues(identityKey) = identityValue
        End If

        private_WriteValue = True
        Exit Function
    End If

    If exists Then
        m_KeyValues(identityKey) = identityValue
    Else
        m_KeyValues.Add identityKey, identityValue
    End If

    private_WriteValue = True
End Function
