Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function private_MessageBoxW Lib "user32" Alias "MessageBoxW" ( _
    ByVal windowHandle As LongPtr, _
    ByVal messagePointer As LongPtr, _
    ByVal captionPointer As LongPtr, _
    ByVal messageBoxType As Long _
) As Long
#Else
Private Declare Function private_MessageBoxW Lib "user32" Alias "MessageBoxW" ( _
    ByVal windowHandle As Long, _
    ByVal messagePointer As Long, _
    ByVal captionPointer As Long, _
    ByVal messageBoxType As Long _
) As Long
#End If

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_ShowMessage( _
    ByVal messageText As String, _
    Optional ByVal buttonsStyle As VbMsgBoxStyle = VBA.vbOKOnly, _
    Optional ByVal captionText As String = VBA.vbNullString _
) As VbMsgBoxResult
#If VBA7 Then
    fn_ShowMessage = private_MessageBoxW( _
        0, StrPtr(messageText), StrPtr(captionText), VBA.CLng(buttonsStyle))
#Else
    fn_ShowMessage = private_MessageBoxW( _
        0, StrPtr(messageText), StrPtr(captionText), VBA.CLng(buttonsStyle))
#End If
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------