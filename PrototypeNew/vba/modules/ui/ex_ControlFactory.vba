Attribute VB_Name = "ex_ControlFactory"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_ControlFactory.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function fn_CreateControlByTypeRoot( _
    ByVal controlTypeRoot As String, _
    ByVal page As obj_IPage _
) As obj_IControl
    Dim control As obj_IControl

    controlTypeRoot = VBA.LCase$(VBA.Trim$(controlTypeRoot))

    Select Case controlTypeRoot
        Case "button"
            Set control = New obj_ButtonControlVM

        Case "label"
            Set control = New obj_LabelControlVM

        Case "input"
            Set control = New obj_InputControlVM

        Case "banner"
            Set control = New obj_BannerControlVM

        Case "config"
            Set control = New obj_ConfigControlVM

        Case "select"
            Set control = New obj_SelectControlVM

        Case "tablelist"
            Set control = New obj_TableListControlVM

        Case "tablesingle"
            Set control = New obj_TableSingleControlVM

        Case "tabletpl"
            Set control = New obj_TableTplControlVM

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Control type '" & controlTypeRoot & "' is not supported in PrototypeNew runtime."
#End If
            Exit Function
    End Select

    If control Is Nothing Then Exit Function
    If Not control.Initialize(page) Then Exit Function
    Set fn_CreateControlByTypeRoot = control
End Function
