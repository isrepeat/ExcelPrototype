Attribute VB_Name = "ex_SqlRowProcessorFactory"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_SqlRowProcessorFactory.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function fn_TryCreateByClassName( _
    ByVal className As String, _
    ByRef outProcessor As obj_ISqlRowProcessor _
) As Boolean
    Set outProcessor = Nothing
    className = VBA.Trim$(className)
    If VBA.Len(className) = 0 Then Exit Function

    Select Case VBA.LCase$(className)
        Case VBA.LCase$("obj_PersonalCardSqlRowPcsr")
            Set outProcessor = New obj_PersonalCardSqlRowPcsr
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo "SqlRowProcessorFactory: created row processor class='" & className & "' type='" & VBA.TypeName(outProcessor) & "'"
#End If

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "SqlRowProcessorFactory: unsupported row processor class '" & className & "'."
#End If
            Exit Function
    End Select

    fn_TryCreateByClassName = Not outProcessor Is Nothing
End Function
