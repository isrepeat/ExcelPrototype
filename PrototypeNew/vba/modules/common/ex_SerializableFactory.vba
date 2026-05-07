Attribute VB_Name = "ex_SerializableFactory"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_SerializableFactory.m_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function m_TryCreatePageByTypeRoot( _
    ByVal typeRoot As String, _
    ByRef outPage As obj_IPage _
) As Boolean
    typeRoot = VBA.LCase$(VBA.Trim$(typeRoot))
    Set outPage = Nothing

    Select Case typeRoot
        Case "page.personalcard"
            Set outPage = New obj_PagePersonalCard
            m_TryCreatePageByTypeRoot = True
            Exit Function

        Case "page.main"
            Set outPage = New obj_PageMain
            m_TryCreatePageByTypeRoot = True
            Exit Function
    End Select

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "SerializableFactory: unsupported page type root '" & VBA.Replace$(typeRoot, "'", "''") & "'."
#End If
End Function
