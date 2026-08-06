Attribute VB_Name = "ex_SerializableFactory"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:ex_SerializableFactory.fn_Module_Dispose"
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
        Case "page.entitylookup"
            Set outPage = New obj_PageEntityLookup
            fn_TryCreatePageByTypeRoot = True
            Exit Function

        Case "page.prsnlevntbuilder"
            Set outPage = New obj_PagePrsnlEvntBuilder
            fn_TryCreatePageByTypeRoot = True
            Exit Function

        Case "page.supportingdocumentbuilder"
            Set outPage = New obj_PageSDB
            fn_TryCreatePageByTypeRoot = True
            Exit Function

        Case "page.comparing"
            Set outPage = New obj_PageComparing
            fn_TryCreatePageByTypeRoot = True
            Exit Function

        Case "page.multisourcesview"
            Set outPage = New obj_PageMultiSourcesView
            fn_TryCreatePageByTypeRoot = True
            Exit Function

        Case "page.worddataextractor"
            Set outPage = New obj_PageWordDataExtractor
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
