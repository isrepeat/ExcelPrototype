Attribute VB_Name = "ex_ControlFactory"
Option Explicit

' //
' // API
' //
Public Function m_CreateControlByTypeRoot(ByVal controlTypeRoot As String) As obj_IControl
    controlTypeRoot = VBA.LCase$(VBA.Trim$(controlTypeRoot))

    Select Case controlTypeRoot
        Case "button"
            Dim buttonVm As obj_ButtonControlVM
            Set buttonVm = New obj_ButtonControlVM
            Set m_CreateControlByTypeRoot = buttonVm

        Case "label"
            Dim labelVm As obj_LabelControlVM
            Set labelVm = New obj_LabelControlVM
            Set m_CreateControlByTypeRoot = labelVm

        Case "banner"
            Dim bannerVm As obj_BannerControlVM
            Set bannerVm = New obj_BannerControlVM
            Set m_CreateControlByTypeRoot = bannerVm

        Case "config"
            Dim configVm As obj_ConfigControlVM
            Set configVm = New obj_ConfigControlVM
            Set m_CreateControlByTypeRoot = configVm

        Case "select"
            Dim selectVm As obj_SelectControlVM
            Set selectVm = New obj_SelectControlVM
            Set m_CreateControlByTypeRoot = selectVm

        Case "tablelist"
            Dim tableListVm As obj_TableListControlVM
            Set tableListVm = New obj_TableListControlVM
            Set m_CreateControlByTypeRoot = tableListVm

        Case "tablesingle"
            Dim tableSingleVm As obj_TableSingleControlVM
            Set tableSingleVm = New obj_TableSingleControlVM
            Set m_CreateControlByTypeRoot = tableSingleVm

        Case "tabletpl"
            Dim tableTemplateVm As obj_TableTplControlVM
            Set tableTemplateVm = New obj_TableTplControlVM
            Set m_CreateControlByTypeRoot = tableTemplateVm

        Case Else
            VBA.MsgBox "Control type '" & controlTypeRoot & "' is not supported in PrototypeNew runtime.", VBA.vbExclamation
    End Select
End Function
