Attribute VB_Name = "ex_PEB_WordState"
Option Explicit

Private m_wordSession As obj_PEB_WordSession

#If VBA7 Then
Private Declare PtrSafe Function private_ReplaceFile Lib "kernel32" Alias "ReplaceFileW" (ByVal target As LongPtr, ByVal replacement As LongPtr, ByVal backup As LongPtr, ByVal flags As Long, ByVal exclude As LongPtr, ByVal reserved As LongPtr) As Long
#Else
Private Declare Function private_ReplaceFile Lib "kernel32" Alias "ReplaceFileW" (ByVal target As Long, ByVal replacement As Long, ByVal backup As Long, ByVal flags As Long, ByVal exclude As Long, ByVal reserved As Long) As Long
#End If

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_AcquireWord() As Object
    If m_wordSession Is Nothing Then
        Set m_wordSession = New obj_PEB_WordSession
        If Not m_wordSession.Initialize() Then
            Err.Raise 5, , "Cannot initialize Word session."
        End If
    End If
    Set fn_AcquireWord = m_wordSession.Acquire()
End Function

Public Sub fn_CloseWord()
    If Not m_wordSession Is Nothing Then
        m_wordSession.Dispose
        Set m_wordSession = Nothing
    End If
End Sub

Public Sub fn_Module_Dispose()
    fn_CloseWord
End Sub

Public Sub fn_Apply( _
    ByVal document As Object, _
    ByVal values As Object, _
    Optional ByRef resultCode As String = "" _
)
    Dim uid As String

    ex_PEB_WordDocument.fn_Change document, values, "INSERT", uid, resultCode
    values("WordEventUid") = uid
End Sub

Public Sub fn_Publish( _
    ByVal temporary As String, _
    ByVal output As String _
)
    Dim files As Object
    Dim backup As String
    Dim errorNumber As Long

    Set files = VBA.CreateObject("Scripting.FileSystemObject")
    backup = files.BuildPath(files.GetParentFolderName(output), files.GetTempName & ".bak")
    If private_ReplaceFile(VBA.StrPtr(output), VBA.StrPtr(temporary), VBA.StrPtr(backup), 0, 0, 0) = 0 Then
        errorNumber = VBA.Err.LastDllError
        If Not files.FileExists(output) And files.FileExists(backup) Then
            files.MoveFile backup, output
        End If
        Err.Raise 5, , "Cannot atomically replace Word output; Windows error " & VBA.CStr(errorNumber) & "; recovery backup: " & backup
    End If
    private_DeleteBackup files, backup
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Sub private_DeleteBackup( _
    ByVal files As Object, _
    ByVal path As String _
)
    On Error GoTo EH
    If files.FileExists(path) Then
        files.DeleteFile path
    End If
    Exit Sub
EH:
    ex_Core.fn_Diagnostic_WriteError "Word.Publish.BackupCleanup", VBA.Err.Number, VBA.Err.Source, VBA.Err.Description
End Sub