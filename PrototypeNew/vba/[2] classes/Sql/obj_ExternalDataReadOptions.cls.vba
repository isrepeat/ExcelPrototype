VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ExternalDataReadOptions"
Option Explicit

Private m_LongValuesMode As en_AdoLongValuesMode

Private Sub Class_Initialize()
    m_LongValuesMode = AdoLongValuesMarkCandidates
End Sub

Public Property Get LongValuesMode() As en_AdoLongValuesMode
    LongValuesMode = m_LongValuesMode
End Property

Public Property Let LongValuesMode(ByVal value As en_AdoLongValuesMode)
    Select Case value
        Case AdoLongValuesMarkCandidates, AdoLongValuesHydrate
            m_LongValuesMode = value
        Case Else
            Err.Raise VBA.vbObjectError + 7360, _
                "obj_ExternalDataReadOptions.LongValuesMode", _
                "Unsupported ADO long-values mode: " & VBA.CStr(value)
    End Select
End Property

Public Property Get AdoSupportLongValues() As Boolean
    AdoSupportLongValues = (m_LongValuesMode = AdoLongValuesHydrate)
End Property

Public Function ToggleLongValuesMode() As Boolean
    If m_LongValuesMode = AdoLongValuesHydrate Then
        m_LongValuesMode = AdoLongValuesMarkCandidates
    Else
        m_LongValuesMode = AdoLongValuesHydrate
    End If
    ToggleLongValuesMode = True
End Function
