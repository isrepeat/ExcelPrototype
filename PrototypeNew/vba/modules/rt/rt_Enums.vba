Attribute VB_Name = "rt_Enums"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Public Enum PageTypeEnum
    PageTypeMain = 1
    PageTypeGenerated = 2
End Enum


Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:rt_Enums.m_Module_Dispose"
#End If
End Sub
