Attribute VB_Name = "ex_UIActions"
Option Explicit

' UI entrypoints layer: keeps user-triggered callbacks in actions/*
' and delegates work to domain/config modules.

Public Sub m_DeleteResultSheets()
    ex_SheetStylesXmlProvider.m_DeleteResultSheets
End Sub

Public Sub m_SwitchMode_OnClick()
    ex_Settings.m_SwitchMode_OnClick
End Sub

Public Sub m_OnProfileChanged()
    ex_ConfigProfilesManager.m_OnProfileChanged
End Sub

Public Sub m_OnModeChanged()
    ex_ConfigProfilesManager.m_OnModeChanged
End Sub

Public Sub m_HelloWorld()
    ex_Startup.m_HelloWorld
End Sub
