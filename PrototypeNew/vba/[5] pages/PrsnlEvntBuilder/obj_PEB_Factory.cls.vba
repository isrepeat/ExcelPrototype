VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_Factory"
Option Explicit

Public Function TryCreatePageController( _
    ByVal className As String, _
    ByRef outController As obj_PagePrsnlEvntBuilderCtrl _
) As Boolean
    Set outController = Nothing
    Select Case VBA.LCase$(VBA.Trim$(className))
        Case VBA.LCase$("obj_PagePrsnlEvntBuilderCtrl")
            Set outController = New obj_PagePrsnlEvntBuilderCtrl
        Case Else
            private_ShowUnsupportedClass "page controller", className
            Exit Function
    End Select
    TryCreatePageController = Not outController Is Nothing
End Function

Public Function TryCreateProfilesProvider( _
    ByVal className As String, _
    ByRef outProvider As obj_PrsnlEvntBuilderData _
) As Boolean
    Set outProvider = Nothing
    Select Case VBA.LCase$(VBA.Trim$(className))
        Case VBA.LCase$("obj_PrsnlEvntBuilderData")
            Set outProvider = New obj_PrsnlEvntBuilderData
        Case Else
            private_ShowUnsupportedClass "profiles provider", className
            Exit Function
    End Select
    TryCreateProfilesProvider = Not outProvider Is Nothing
End Function

Public Function TryCreateDataExporter( _
    ByVal className As String, _
    ByVal exportConfigTable As obj_ConfigTable, _
    ByVal profileConfigTable As obj_ConfigTable, _
    ByVal exporterCfgDataProvider As obj_PEB_ExptrCfgDataPrvdr, _
    ByRef outExporter As obj_IDataExporter _
) As Boolean
    Dim movementExporter As obj_PEB_ExptrMovement
    Dim wordExporter As obj_PEB_ExptrWord

    Set outExporter = Nothing
    Select Case VBA.LCase$(VBA.Trim$(className))
        Case VBA.LCase$("obj_PEB_ExptrMovement")
            If exporterCfgDataProvider Is Nothing Then
                VBA.MsgBox "PrsnlEvntBuilder: Movement exporter requires config data provider.", _
                    VBA.vbExclamation, "PrsnlEvntBuilder / Factory"
                Exit Function
            End If
            Set movementExporter = New obj_PEB_ExptrMovement
            If Not movementExporter.Initialize(exportConfigTable, profileConfigTable, _
                exporterCfgDataProvider) Then Exit Function
            Set outExporter = movementExporter
        Case VBA.LCase$("obj_PEB_ExptrWord")
            If exporterCfgDataProvider Is Nothing Then
                VBA.MsgBox "PrsnlEvntBuilder: WORD exporter requires config data provider.", _
                    VBA.vbExclamation, "PrsnlEvntBuilder / Factory"
                Exit Function
            End If
            Set wordExporter = New obj_PEB_ExptrWord
            If Not wordExporter.Initialize(exportConfigTable, profileConfigTable, _
                exporterCfgDataProvider) Then Exit Function
            Set outExporter = wordExporter
        Case Else
            private_ShowUnsupportedClass "data exporter", className
            Exit Function
    End Select
    TryCreateDataExporter = Not outExporter Is Nothing
End Function

Private Sub private_ShowUnsupportedClass( _
    ByVal roleName As String, _
    ByVal className As String _
)
    VBA.MsgBox "PrsnlEvntBuilder: unsupported " & roleName & _
        " class '" & VBA.Trim$(className) & "'.", VBA.vbExclamation, _
        "PrsnlEvntBuilder / Factory"
End Sub
