Attribute VB_Name = "rt_Enums"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:rt_Enums.m_Module_Dispose"
#End If
End Sub
