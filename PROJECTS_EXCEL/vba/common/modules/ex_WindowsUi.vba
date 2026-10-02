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

Private Const MESSAGE_BOX_OK As Long = &H0&
Private Const MESSAGE_BOX_INFORMATION As Long = &H40&

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
Public Function fn_ShowInformation( _
    ByVal messageText As String, _
    ByVal captionText As String _
) As Long
#If VBA7 Then
    fn_ShowInformation = private_MessageBoxW( _
        0, StrPtr(messageText), StrPtr(captionText), _
        MESSAGE_BOX_OK Or MESSAGE_BOX_INFORMATION)
#Else
    fn_ShowInformation = private_MessageBoxW( _
        0, StrPtr(messageText), StrPtr(captionText), _
        MESSAGE_BOX_OK Or MESSAGE_BOX_INFORMATION)
#End If
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------