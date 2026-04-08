Attribute VB_Name = "ex_ControlFactory"
Option Explicit

Public Function m_CreateControlByTypeRoot(ByVal controlTypeRoot As String) As obj_IControl
    controlTypeRoot = LCase$(Trim$(controlTypeRoot))

    Select Case controlTypeRoot
        Case "button"
            Dim buttonVm As obj_ButtonControlViewModel
            Set buttonVm = New obj_ButtonControlViewModel
            Set m_CreateControlByTypeRoot = buttonVm

        Case "table"
            Dim tableVm As obj_TableControlViewModel
            Set tableVm = New obj_TableControlViewModel
            Set m_CreateControlByTypeRoot = tableVm

        Case Else
            MsgBox "Control type '" & controlTypeRoot & "' is not supported in PrototypeNew runtime.", vbExclamation
    End Select
End Function
