Attribute VB_Name = "ex_Actions"
Option Explicit

Public Sub m_LoadPrototypeNewUi()
    ex_UiLoader.m_LoadPrototypeNewUi ThisWorkbook
End Sub

Public Sub m_SetStateSimpleTestDefault()
    ex_State.m_SetText ex_State.STATE_ACTIVE_MODE, "SimpleTest"
    ex_State.m_SetText ex_State.STATE_ACTIVE_PROFILE, "Default"
End Sub
