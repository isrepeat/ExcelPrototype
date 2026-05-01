VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ObjectCollectionBase"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_Items As Collection

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_Items = New Collection
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Set m_Items = Nothing
    On Error GoTo 0
End Sub

Public Property Get Count() As Long
    On Error GoTo EH_COUNT
    Count = m_Items.Count
    Exit Property

EH_COUNT:
    private_LogCollectionError "Count", Err.Number, Err.Description
End Property

Public Property Get IsEmpty() As Boolean
    On Error GoTo EH_ISEMPTY
    IsEmpty = (m_Items.Count = 0)
    Exit Property

EH_ISEMPTY:
    private_LogCollectionError "IsEmpty", Err.Number, Err.Description
End Property

Public Sub AddObject(ByVal item As Object)
    On Error GoTo EH_ADD
    m_Items.Add item
    Exit Sub

EH_ADD:
    private_LogCollectionError "AddObject", Err.Number, Err.Description
End Sub

Public Property Get ItemObject(ByVal oneBasedIndex As Long) As Object
    On Error GoTo EH_ITEM

    If oneBasedIndex <= 0 Or oneBasedIndex > m_Items.Count Then
        private_LogCollectionError "ItemObject", 9, "Index out of range.", "index=" & VBA.CStr(oneBasedIndex) & ", count=" & VBA.CStr(m_Items.Count)
        Exit Property
    End If

    Set ItemObject = m_Items(oneBasedIndex)
    Exit Property

EH_ITEM:
    private_LogCollectionError "ItemObject", Err.Number, Err.Description, "index=" & VBA.CStr(oneBasedIndex)
End Property

Public Sub RemoveAt(ByVal oneBasedIndex As Long)
    On Error GoTo EH_REMOVE

    If oneBasedIndex <= 0 Or oneBasedIndex > m_Items.Count Then
        private_LogCollectionError "RemoveAt", 9, "Index out of range.", "index=" & VBA.CStr(oneBasedIndex) & ", count=" & VBA.CStr(m_Items.Count)
        Exit Sub
    End If

    m_Items.Remove oneBasedIndex
    Exit Sub

EH_REMOVE:
    private_LogCollectionError "RemoveAt", Err.Number, Err.Description, "index=" & VBA.CStr(oneBasedIndex)
End Sub

Public Sub Clear()
    Set m_Items = New Collection
End Sub

Public Property Get AsCollection() As Collection
    On Error GoTo EH_ASCOLL
    Set AsCollection = m_Items
    Exit Property

EH_ASCOLL:
    private_LogCollectionError "AsCollection", Err.Number, Err.Description
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    On Error GoTo EH_ENUM
    Set NewEnum = m_Items.[_NewEnum]
    Exit Property

EH_ENUM:
    private_LogCollectionError "NewEnum", Err.Number, Err.Description
End Property

Private Sub private_LogCollectionError( _
    ByVal memberName As String, _
    ByVal errNumber As Long, _
    ByVal errDescription As String, _
    Optional ByVal details As String = "" _
)
    Dim msg As String

    msg = "CollectionBase." & memberName & " failed (Err " & VBA.CStr(errNumber) & "): " & errDescription
    If VBA.Len(VBA.Trim$(details)) > 0 Then
        msg = msg & " [" & details & "]"
    End If

    Debug.Print msg
End Sub

