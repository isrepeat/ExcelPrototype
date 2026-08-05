VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SDB_Factory"
Option Explicit

Public Function TryCreatePageController( _
    ByVal className As String, _
    ByRef outController As obj_PageSDBCtrl _
) As Boolean
    Set outController = Nothing
    Select Case VBA.LCase$(VBA.Trim$(className))
        Case VBA.LCase$("obj_PageSDBCtrl")
            Set outController = New obj_PageSDBCtrl
        Case Else
            private_ShowUnsupportedClass "page controller", className
            Exit Function
    End Select
    TryCreatePageController = Not outController Is Nothing
End Function

Public Function TryCreateProfilesProvider( _
    ByVal className As String, _
    ByRef outProvider As obj_SDB_Data _
) As Boolean
    Set outProvider = Nothing
    Select Case VBA.LCase$(VBA.Trim$(className))
        Case VBA.LCase$("obj_SDB_Data")
            Set outProvider = New obj_SDB_Data
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
    ByRef outExporter As obj_IDataExporter _
) As Boolean
    Dim wordExporter As obj_SDB_ExptrWord

    Set outExporter = Nothing
    Select Case VBA.LCase$(VBA.Trim$(className))
        Case VBA.LCase$("obj_SDB_ExptrWord")
            Set wordExporter = New obj_SDB_ExptrWord
            If Not wordExporter.Initialize(exportConfigTable) Then Exit Function
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
    VBA.MsgBox "SupportingDocumentBuilder: unsupported " & roleName & _
        " class '" & VBA.Trim$(className) & "'.", VBA.vbExclamation, _
        "Supporting Document Builder / Factory"
End Sub
