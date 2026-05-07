Attribute VB_Name = "ex_SerializableFactory"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_SerializableFactory.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function fn_TryCreatePageByTypeRoot( _
    ByVal typeRoot As String, _
    ByRef outPage As obj_IPage _
) As Boolean
    typeRoot = VBA.LCase$(VBA.Trim$(typeRoot))
    Set outPage = Nothing

    Select Case typeRoot
        Case "page.personalcard"
            Set outPage = New obj_PagePersonalCard
            fn_TryCreatePageByTypeRoot = True
            Exit Function

        Case "page.main"
            Set outPage = New obj_PageMain
            fn_TryCreatePageByTypeRoot = True
            Exit Function
    End Select

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "SerializableFactory: unsupported page type root '" & VBA.Replace$(typeRoot, "'", "''") & "'."
#End If
End Function
