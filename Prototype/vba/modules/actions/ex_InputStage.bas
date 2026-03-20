Attribute VB_Name = "ex_InputStage"
Option Explicit

' Stage 1 of mode pipeline: resolve input payload (keys, sources) before mode logic starts.

Public Function m_ResolveExplicitKeys( _
    ByVal rawKeysText As String, _
    Optional ByVal configKeyName As String = "MultiKeys" _
) As Collection
    Dim parts() As String
    Dim i As Long
    Dim keyValue As String
    Dim seen As Object
    Dim result As Collection

    rawKeysText = Trim$(CStr(rawKeysText))
    If Len(rawKeysText) = 0 Then
        Err.Raise vbObjectError + 6101, "ex_InputStage", _
            "Config key '" & Trim$(configKeyName) & "' is empty. Provide keys separated by ';'."
    End If

    parts = Split(rawKeysText, ";")
    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    For i = LBound(parts) To UBound(parts)
        keyValue = Trim$(parts(i))
        If Len(keyValue) > 0 Then
            If Not seen.Exists(keyValue) Then
                seen.Add keyValue, True
                result.Add keyValue
            End If
        End If
    Next i

    If result.Count = 0 Then
        Err.Raise vbObjectError + 6102, "ex_InputStage", _
            "Config key '" & Trim$(configKeyName) & "' contains no valid keys after parsing."
    End If

    Set m_ResolveExplicitKeys = result
End Function
