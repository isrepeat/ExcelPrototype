Option Explicit

Private Const SHPO_RELATIVE_PATH As String = "ШПО.xlsx"
Private Const ALF_TABLE_REF As String = "[АЛФ$A1:J12000]"
Private Const OS_TABLE_REF As String = "[ОС$A1:AB12000]"
Private Const RANKS_TABLE_REF As String = "[Звання$A1:E12000]"
Private Const POSITIONS_TABLE_REF As String = "[Посади$A1:E12000]"
Private Const POSITION_ROZP_PREFIX As String = "A1A"
Private Const POSITION_ROZP_TEXT_PREFIX As String = "у розпорядженні командира військової частини "
Private Const POSITION_ROZP_OFFICER_UNIT As String = "А3369"
Private Const POSITION_ROZP_OTHER_UNIT As String = "А7383"

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function ex_PersonnelData_TryResolveIpn( _
    ByVal personLookup As String, _
    ByRef outIpnText As String _
) As Boolean
    outIpnText = VBA.vbNullString
    If ex_Helpers.private_Text_IsIpn(personLookup) Then
        outIpnText = personLookup
        ex_PersonnelData_TryResolveIpn = True
        Exit Function
    End If

    ex_PersonnelData_TryResolveIpn = private_TryLookupShpoValue( _
        ALF_TABLE_REF, "ПІБ", personLookup, "ІПН", outIpnText)
End Function

Public Function ex_PersonnelData_TryResolveFioNominative( _
    ByVal ipnText As String, _
    ByRef outFioText As String _
) As Boolean
    ex_PersonnelData_TryResolveFioNominative = private_TryResolveFioCase( _
        ipnText, "ПІБ", outFioText)
End Function

Public Function ex_PersonnelData_TryResolveFioGenitive( _
    ByVal ipnText As String, _
    ByRef outFioText As String _
) As Boolean
    ex_PersonnelData_TryResolveFioGenitive = private_TryResolveFioCase( _
        ipnText, "Родовий", outFioText)
End Function

Public Function ex_PersonnelData_TryResolveFioDative( _
    ByVal ipnText As String, _
    ByRef outFioText As String _
) As Boolean
    ex_PersonnelData_TryResolveFioDative = private_TryResolveFioCase( _
        ipnText, "Давальний", outFioText)
End Function

Public Function ex_PersonnelData_TryResolveFioAccusative( _
    ByVal ipnText As String, _
    ByRef outFioText As String _
) As Boolean
    ex_PersonnelData_TryResolveFioAccusative = private_TryResolveFioCase( _
        ipnText, "Знахідний", outFioText)
End Function

Public Function ex_PersonnelData_TryResolveRankNominative( _
    ByVal ipnText As String, _
    ByRef outRankText As String _
) As Boolean
    ex_PersonnelData_TryResolveRankNominative = private_TryResolveRankCase( _
        ipnText, "Звання", outRankText)
End Function

Public Function ex_PersonnelData_TryResolveRankGenitive( _
    ByVal ipnText As String, _
    ByRef outRankText As String _
) As Boolean
    ex_PersonnelData_TryResolveRankGenitive = private_TryResolveRankCase( _
        ipnText, "Родовий", outRankText)
End Function

Public Function ex_PersonnelData_TryResolveRankDative( _
    ByVal ipnText As String, _
    ByRef outRankText As String _
) As Boolean
    ex_PersonnelData_TryResolveRankDative = private_TryResolveRankCase( _
        ipnText, "Давальний", outRankText)
End Function

Public Function ex_PersonnelData_TryResolveRankAccusative( _
    ByVal ipnText As String, _
    ByRef outRankText As String _
) As Boolean
    ex_PersonnelData_TryResolveRankAccusative = private_TryResolveRankCase( _
        ipnText, "Знахідний", outRankText)
End Function

Public Function ex_PersonnelData_TryResolvePositionNominative( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByRef outPositionText As String _
) As Boolean
    ex_PersonnelData_TryResolvePositionNominative = private_TryResolvePositionCase( _
        positionCode, rankText, "Назва", outPositionText)
End Function

Public Function ex_PersonnelData_TryResolvePositionGenitive( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByRef outPositionText As String _
) As Boolean
    ex_PersonnelData_TryResolvePositionGenitive = private_TryResolvePositionCase( _
        positionCode, rankText, "Родовий", outPositionText)
End Function

Public Function ex_PersonnelData_TryResolvePositionDative( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByRef outPositionText As String _
) As Boolean
    ex_PersonnelData_TryResolvePositionDative = private_TryResolvePositionCase( _
        positionCode, rankText, "Давальний", outPositionText)
End Function

Public Function ex_PersonnelData_TryResolvePositionAccusative( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByRef outPositionText As String _
) As Boolean
    ex_PersonnelData_TryResolvePositionAccusative = private_TryResolvePositionCase( _
        positionCode, rankText, "Знахідний", outPositionText)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_TryLookupShpoValue( _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal keyValue As String, _
    ByVal resultHeader As String, _
    ByRef outValue As String _
) As Boolean
    Dim shpoPath As String
    Dim connection As Object

    On Error GoTo EH
    outValue = VBA.vbNullString
    shpoPath = ex_Helpers.private_Path_ResolveFromWorkbook(SHPO_RELATIVE_PATH)
    If VBA.Len(VBA.Dir$(shpoPath)) = 0 Then
        ex_Helpers.LogError "SHPO file was not found: " & shpoPath
        VBA.MsgBox "SHPO file was not found: " & shpoPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    If Not ex_ExternalTables.ex_ExternalTables_TryOpenConnection( _
        shpoPath, "SHPO", connection) Then GoTo CleanExit
    private_TryLookupShpoValue = ex_ExternalTables.ex_ExternalTables_TryLookup( _
        connection, "SHPO", tableRef, keyHeader, keyValue, resultHeader, outValue)

CleanExit:
    On Error Resume Next
    If Not connection Is Nothing Then connection.Close
    Set connection = Nothing
    On Error GoTo 0
    Exit Function

EH:
    ex_Helpers.LogError "SHPO lookup failed | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    VBA.MsgBox "SHPO lookup failed: [" & VBA.CStr(Err.Number) & "] " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Function

Private Function private_TryResolveFioCase( _
    ByVal ipnText As String, _
    ByVal resultHeader As String, _
    ByRef outFioText As String _
) As Boolean
    private_TryResolveFioCase = private_TryLookupShpoValue( _
        ALF_TABLE_REF, "ІПН", ipnText, resultHeader, outFioText)
End Function

Private Function private_TryResolveRankCase( _
    ByVal ipnText As String, _
    ByVal resultHeader As String, _
    ByRef outRankText As String _
) As Boolean
    Dim rankNominative As String

    If Not private_TryLookupShpoValue( _
        OS_TABLE_REF, "ІПН", ipnText, "Військове звання фактично", _
        rankNominative) Then Exit Function

    private_TryResolveRankCase = private_TryLookupShpoValue( _
        RANKS_TABLE_REF, "Звання", rankNominative, resultHeader, outRankText)
End Function

Private Function private_TryResolvePositionCase( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByVal resultHeader As String, _
    ByRef outPositionText As String _
) As Boolean
    Dim normalizedCode As String

    normalizedCode = VBA.UCase$(ex_Helpers.private_Text_Normalize(positionCode))
    If private_IsRozpPositionCode(normalizedCode) Then
        If private_IsOfficerRank(rankText) Then
            outPositionText = POSITION_ROZP_TEXT_PREFIX & _
                POSITION_ROZP_OFFICER_UNIT
        Else
            outPositionText = POSITION_ROZP_TEXT_PREFIX & _
                POSITION_ROZP_OTHER_UNIT
        End If
        private_TryResolvePositionCase = True
        Exit Function
    End If

    private_TryResolvePositionCase = private_TryLookupShpoValue( _
        POSITIONS_TABLE_REF, "Код", normalizedCode, resultHeader, outPositionText)
End Function

Private Function private_IsRozpPositionCode( _
    ByVal positionCode As String _
) As Boolean
    positionCode = VBA.Replace$(positionCode, " ", VBA.vbNullString)
    positionCode = VBA.Replace$(positionCode, "А", "A")
    private_IsRozpPositionCode = ( _
        VBA.Left$(positionCode, VBA.Len(POSITION_ROZP_PREFIX)) = _
        POSITION_ROZP_PREFIX)
End Function

Private Function private_IsOfficerRank(ByVal rankText As String) As Boolean
    Select Case VBA.LCase$(ex_Helpers.private_Text_Normalize(rankText))
        Case "молодший лейтенант", "лейтенант", "старший лейтенант", _
             "капітан", "майор", "підполковник", "полковник"
            private_IsOfficerRank = True
    End Select
End Function