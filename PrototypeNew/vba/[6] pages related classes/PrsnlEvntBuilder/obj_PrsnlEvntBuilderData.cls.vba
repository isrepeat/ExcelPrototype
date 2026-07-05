VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PrsnlEvntBuilderData"
Option Explicit

Private Const SECTION_TYPE_CLOSE_FROM_TREATMENT As String = "З лікування"
Private Const SECTION_TYPE_CLOSE_FROM_TREATMENT_VACATION As String = "З відпустки для лікування"
Private Const SECTION_TYPE_CLOSE_FROM_ANNUAL_VACATION As String = "З щорічної основної відпустки"
Private Const SECTION_TYPE_CLOSE_FROM_FAMILY_VACATION As String = "З відпустки за сімейними обставинами"
Private Const SECTION_TYPE_CLOSE_FROM_TREATMENT_MEDICAL_COMPANY As String = "З лікування медична рота"
Private Const SECTION_TYPE_CLOSE_FROM_AMBULATORY_VLK As String = "З амбулаторного обстеження влк"
Private Const SECTION_TYPE_TO_TREATMENT As String = "На лікування"
Private Const SECTION_TYPE_TO_ANNUAL_VACATION_PART As String = "У частину щорічної основної відпустки"
Private Const SECTION_TYPE_TO_FAMILY_VACATION As String = "У відпустку за сімейними обставинами"
Private Const SECTION_TYPE_TO_TREATMENT_VACATION As String = "У відпустку для лікування"
Private Const SECTION_TYPE_TO_TREATMENT_MEDICAL_COMPANY As String = "На лікування медична рота"
Private Const SECTION_TYPE_TO_AMBULATORY_VLK As String = "На амбулаторне обстеження влк"
Private Const SECTION_TYPE_TRANSFER_TREATMENT_TO_TREATMENT_VACATION As String = "Зміна місця перебування лікування => відпустка для лік"
Private Const SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION As String = "Зміна місця перебування відпустка для лік => відпустка для лік"
Private Const SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT As String = "Зміна місця перебування відпустка для лік => лікування"
Private Const SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_VLK As String = "Зміна місця перебування відпустка для лік => влк"
Private Const SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT_VACATION As String = "Зміна місця перебування влк => відпустка для лік"
Private Const SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT As String = "Зміна місця перебування влк => лікування"
Private Const SECTION_TYPE_TO_BUSINESS_TRIP As String = "У відрядження"
Private Const SECTION_TYPE_TO_BUSINESS_TRIP_SZCH As String = "У відрядження сзч"

Private Const MOVEMENT_EVENT_STATIONARY_TREATMENT As String = "Стаціонарне лікування"
Private Const MOVEMENT_EVENT_ANNUAL_VACATION As String = "Щорічна відпустка"
Private Const MOVEMENT_EVENT_FAMILY_VACATION As String = "Відпустка за сімейними обставинами"
Private Const MOVEMENT_EVENT_TREATMENT_VACATION As String = "Відпустка для лікування"
Private Const MOVEMENT_EVENT_AMBULATORY_VLK As String = "Амбулаторне ВЛК"
Private Const MOVEMENT_EVENT_STATIONARY_VLK As String = "Стаціонарне ВЛК"

Private m_SectionTypeNames As Collection

Private Sub Class_Initialize()
    Set m_SectionTypeNames = private_BuildSectionTypeNames()
End Sub

Public Property Get SectionTypeNames() As Collection
    Set SectionTypeNames = private_CopyCollection(m_SectionTypeNames)
End Property

Public Function IsMovementClosingSectionType(ByVal sectionTypeText As String) As Boolean
    Select Case private_NormalizeText(sectionTypeText)
        Case private_NormalizeText(SECTION_TYPE_CLOSE_FROM_TREATMENT), _
             private_NormalizeText(SECTION_TYPE_CLOSE_FROM_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_CLOSE_FROM_ANNUAL_VACATION), _
             private_NormalizeText(SECTION_TYPE_CLOSE_FROM_FAMILY_VACATION), _
             private_NormalizeText(SECTION_TYPE_CLOSE_FROM_TREATMENT_MEDICAL_COMPANY), _
             private_NormalizeText(SECTION_TYPE_CLOSE_FROM_AMBULATORY_VLK)
            IsMovementClosingSectionType = True
    End Select
End Function

Public Function IsMovementMirrorTransferSectionType(ByVal sectionTypeText As String) As Boolean
    Dim normalizedSectionType As String
    Dim transferToken As String

    normalizedSectionType = private_NormalizeText(sectionTypeText)
    transferToken = private_NormalizeText("зміна місця перебування")

    If VBA.InStr(1, normalizedSectionType, transferToken, VBA.vbTextCompare) > 0 Then
        IsMovementMirrorTransferSectionType = True
        Exit Function
    End If

    Select Case normalizedSectionType
        Case private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_VLK), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT)
            IsMovementMirrorTransferSectionType = True
    End Select
End Function

Public Function ShouldWriteMovementSpecialOpeningFields(ByVal sectionTypeText As String) As Boolean
    Select Case private_NormalizeText(sectionTypeText)
        Case private_NormalizeText(SECTION_TYPE_TO_ANNUAL_VACATION_PART), _
             private_NormalizeText(SECTION_TYPE_TO_FAMILY_VACATION), _
             private_NormalizeText(SECTION_TYPE_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT_VACATION)
            ShouldWriteMovementSpecialOpeningFields = True
    End Select
End Function

Public Function UsesMovementVacationDestination(ByVal sectionTypeText As String) As Boolean
    Select Case private_NormalizeText(sectionTypeText)
        Case private_NormalizeText(SECTION_TYPE_TO_ANNUAL_VACATION_PART), _
             private_NormalizeText(SECTION_TYPE_TO_FAMILY_VACATION), _
             private_NormalizeText(SECTION_TYPE_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT_VACATION)
            UsesMovementVacationDestination = True
    End Select
End Function

Public Function TryMapMovementSectionTypeToEventText( _
    ByVal sectionTypeText As String, _
    ByRef outEventText As String _
) As Boolean
    outEventText = VBA.vbNullString

    Select Case private_NormalizeText(sectionTypeText)
        Case private_NormalizeText(SECTION_TYPE_TO_TREATMENT), _
             private_NormalizeText(SECTION_TYPE_TO_TREATMENT_MEDICAL_COMPANY), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT)
            outEventText = MOVEMENT_EVENT_STATIONARY_TREATMENT

        Case private_NormalizeText(SECTION_TYPE_TO_ANNUAL_VACATION_PART)
            outEventText = MOVEMENT_EVENT_ANNUAL_VACATION

        Case private_NormalizeText(SECTION_TYPE_TO_FAMILY_VACATION)
            outEventText = MOVEMENT_EVENT_FAMILY_VACATION

        Case private_NormalizeText(SECTION_TYPE_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION), _
             private_NormalizeText(SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT_VACATION)
            outEventText = MOVEMENT_EVENT_TREATMENT_VACATION

        Case private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_VLK)
            outEventText = MOVEMENT_EVENT_STATIONARY_VLK

        Case private_NormalizeText(SECTION_TYPE_TO_AMBULATORY_VLK)
            outEventText = MOVEMENT_EVENT_AMBULATORY_VLK
    End Select

    TryMapMovementSectionTypeToEventText = (VBA.Len(outEventText) > 0)
End Function

Private Function private_BuildSectionTypeNames() As Collection
    Dim sectionTypes As Collection

    Set sectionTypes = New Collection
    sectionTypes.Add SECTION_TYPE_CLOSE_FROM_TREATMENT
    sectionTypes.Add SECTION_TYPE_CLOSE_FROM_TREATMENT_VACATION
    sectionTypes.Add SECTION_TYPE_CLOSE_FROM_ANNUAL_VACATION
    sectionTypes.Add SECTION_TYPE_CLOSE_FROM_FAMILY_VACATION
    sectionTypes.Add SECTION_TYPE_CLOSE_FROM_TREATMENT_MEDICAL_COMPANY
    sectionTypes.Add SECTION_TYPE_CLOSE_FROM_AMBULATORY_VLK
    sectionTypes.Add SECTION_TYPE_TO_TREATMENT
    sectionTypes.Add SECTION_TYPE_TO_ANNUAL_VACATION_PART
    sectionTypes.Add SECTION_TYPE_TO_FAMILY_VACATION
    sectionTypes.Add SECTION_TYPE_TO_TREATMENT_VACATION
    sectionTypes.Add SECTION_TYPE_TO_TREATMENT_MEDICAL_COMPANY
    sectionTypes.Add SECTION_TYPE_TO_AMBULATORY_VLK
    sectionTypes.Add SECTION_TYPE_TRANSFER_TREATMENT_TO_TREATMENT_VACATION
    sectionTypes.Add SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION
    sectionTypes.Add SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT
    sectionTypes.Add SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_VLK
    sectionTypes.Add SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT_VACATION
    sectionTypes.Add SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT
    sectionTypes.Add SECTION_TYPE_TO_BUSINESS_TRIP
    sectionTypes.Add SECTION_TYPE_TO_BUSINESS_TRIP_SZCH

    Set private_BuildSectionTypeNames = sectionTypes
End Function

Private Function private_NormalizeText(ByVal valueText As String) As String
    valueText = VBA.LCase$(VBA.Trim$(VBA.CStr(valueText)))
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    valueText = VBA.Replace(valueText, ":", VBA.vbNullString)
    valueText = VBA.Replace(valueText, ".", VBA.vbNullString)
    valueText = VBA.Replace(valueText, "/", " ")
    valueText = VBA.Replace(valueText, "(", " ")
    valueText = VBA.Replace(valueText, ")", " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    private_NormalizeText = VBA.Trim$(valueText)
End Function

Private Function private_CopyCollection(ByVal sourceItems As Collection) As Collection
    Dim copyItems As Collection
    Dim itemValue As Variant

    Set copyItems = New Collection
    If Not sourceItems Is Nothing Then
        For Each itemValue In sourceItems
            copyItems.Add itemValue
        Next itemValue
    End If

    Set private_CopyCollection = copyItems
End Function
