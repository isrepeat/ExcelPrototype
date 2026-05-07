Attribute VB_Name = "rt_Enums"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:rt_Enums.fn_Module_Dispose"
#End If
End Sub
