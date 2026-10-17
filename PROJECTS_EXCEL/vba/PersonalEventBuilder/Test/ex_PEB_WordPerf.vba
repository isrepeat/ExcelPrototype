Attribute VB_Name = "ex_PEB_WordPerf"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private m_frequency As Currency

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    m_frequency = 0
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Now() As Double
    Dim counter As Currency

    ' Currency provides an eight-byte buffer in both 32-bit and 64-bit VBA.
    If m_frequency = 0 Then
        If QueryPerformanceFrequency(m_frequency) = 0 Or m_frequency = 0 Then
            Err.Raise 5, , "High-resolution timer frequency is unavailable."
        End If
    End If
    If QueryPerformanceCounter(counter) = 0 Then Err.Raise 5, , "High-resolution timer is unavailable."
    fn_Now = VBA.CDbl(counter) / VBA.CDbl(m_frequency)
End Function

Public Function fn_ElapsedMs(ByVal started As Double) As Double
    fn_ElapsedMs = (fn_Now() - started) * 1000#
End Function

Public Sub fn_WritePerf( _
    ByVal stage As String, _
    ByVal started As Double _
)
    ex_Core.fn_Diagnostic_WriteLog "PERF | Stage=Tests.Word." & stage & _
        " | ElapsedMs=" & VBA.Format$(fn_ElapsedMs(started), "0.000")
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------