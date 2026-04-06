Attribute VB_Name = "ex_SqlAdoHelpers"
Option Explicit

Private Const ADO_TYPE_DATE As Long = 7
Private Const ADO_TYPE_FILETIME As Long = 64
Private Const ADO_TYPE_DBDATE As Long = 133
Private Const ADO_TYPE_DBTIME As Long = 134
Private Const ADO_TYPE_DBTIMESTAMP As Long = 135

Public Function m_ToNormalizedCellValue(ByVal valueIn As Variant, Optional ByVal adoFieldType As Long = -1) As Variant
    If IsError(valueIn) Then
        m_ToNormalizedCellValue = vbNullString
        Exit Function
    End If

    If IsNull(valueIn) Then
        m_ToNormalizedCellValue = vbNullString
        Exit Function
    End If

    If m_IsAdoDateFieldType(adoFieldType) Or VarType(valueIn) = vbDate Then
        m_ToNormalizedCellValue = mp_FormatAdoDateText(valueIn)
        Exit Function
    End If

    m_ToNormalizedCellValue = valueIn
End Function

Public Function m_ToNormalizedText(ByVal valueIn As Variant, Optional ByVal adoFieldType As Long = -1) As String
    Dim normalized As Variant

    normalized = m_ToNormalizedCellValue(valueIn, adoFieldType)
    If IsNull(normalized) Then
        m_ToNormalizedText = vbNullString
    Else
        m_ToNormalizedText = CStr(normalized)
    End If
End Function

Public Function m_IsAdoDateFieldType(ByVal fieldType As Long) As Boolean
    ' ADO Field.Type codes as returned by ACE/OLEDB Recordset.Fields(i).Type.
    Select Case fieldType
        Case ADO_TYPE_DATE, ADO_TYPE_FILETIME, ADO_TYPE_DBDATE, ADO_TYPE_DBTIME, ADO_TYPE_DBTIMESTAMP
            m_IsAdoDateFieldType = True
    End Select
End Function

Private Function mp_FormatAdoDateText(ByVal valueIn As Variant) As String
    Dim dtValue As Date
    Dim numericDate As Double

    On Error GoTo Fallback
    dtValue = CDate(valueIn)
    numericDate = CDbl(dtValue)

    If numericDate = Int(numericDate) Then
        mp_FormatAdoDateText = Format$(dtValue, "dd.mm.yyyy")
    Else
        mp_FormatAdoDateText = Format$(dtValue, "dd.mm.yyyy hh:nn:ss")
    End If
    Exit Function

Fallback:
    mp_FormatAdoDateText = CStr(valueIn)
End Function
