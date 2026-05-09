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
    controlTypeRoot = VBA.LCase$(VBA.Trim$(controlTypeRoot))

    Select Case controlTypeRoot
        Case "button"
            Dim buttonVm As obj_ButtonControlVM
            Set buttonVm = New obj_ButtonControlVM
            If Not buttonVm.Initialize(page) Then Exit Function
            Set fn_CreateControlByTypeRoot = buttonVm

        Case "label"
            Dim labelVm As obj_LabelControlVM
            Set labelVm = New obj_LabelControlVM
            If Not labelVm.Initialize(page) Then Exit Function
            Set fn_CreateControlByTypeRoot = labelVm

        Case "banner"
            Dim bannerVm As obj_BannerControlVM
            Set bannerVm = New obj_BannerControlVM
            If Not bannerVm.Initialize(page) Then Exit Function
            Set fn_CreateControlByTypeRoot = bannerVm

        Case "config"
            Dim configVm As obj_ConfigControlVM
            Set configVm = New obj_ConfigControlVM
            If Not configVm.Initialize(page) Then Exit Function
            Set fn_CreateControlByTypeRoot = configVm

        Case "select"
            Dim selectVm As obj_SelectControlVM
            Set selectVm = New obj_SelectControlVM
            If Not selectVm.Initialize(page) Then Exit Function
            Set fn_CreateControlByTypeRoot = selectVm

        Case "tablelist"
            Dim tableListVm As obj_TableListControlVM
            Set tableListVm = New obj_TableListControlVM
            If Not tableListVm.Initialize(page) Then Exit Function
            Set fn_CreateControlByTypeRoot = tableListVm

        Case "tablesingle"
            Dim tableSingleVm As obj_TableSingleControlVM
            Set tableSingleVm = New obj_TableSingleControlVM
            If Not tableSingleVm.Initialize(page) Then Exit Function
            Set fn_CreateControlByTypeRoot = tableSingleVm

        Case "tabletpl"
            Dim tableTemplateVm As obj_TableTplControlVM
            Set tableTemplateVm = New obj_TableTplControlVM
            If Not tableTemplateVm.Initialize(page) Then Exit Function
            Set fn_CreateControlByTypeRoot = tableTemplateVm

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Control type '" & controlTypeRoot & "' is not supported in PrototypeNew runtime."
#End If
    End Select
End Function
