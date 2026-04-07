Attribute VB_Name = "ex_ControlFactory"
Option Explicit

Public Function m_CreateControlByTypeRoot(ByVal controlTypeRoot As String) As obj_IControl
    controlTypeRoot = LCase$(Trim$(controlTypeRoot))

    Select Case controlTypeRoot
        Case "helloworld"
            Dim helloVm As obj_HelloWorldControlViewModel
            Set helloVm = New obj_HelloWorldControlViewModel
            Set m_CreateControlByTypeRoot = helloVm

        Case Else
            MsgBox "Control type '" & controlTypeRoot & "' is not supported in PrototypeNew runtime.", vbExclamation
    End Select
End Function
