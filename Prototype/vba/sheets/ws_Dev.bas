Option Explicit

Private Sub Worksheet_Activate()

    ' dynamic profile refresh disabled

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    ' dynamic profile updates and autosave disabled

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    ex_CustomDropdown.m_HideDevTestDropdown Me

End Sub

Private Sub Worksheet_Deactivate()

    ex_CustomDropdown.m_HideDevTestDropdown Me

End Sub
