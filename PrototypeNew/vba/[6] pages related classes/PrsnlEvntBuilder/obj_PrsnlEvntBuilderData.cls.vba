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
Private Const META_SECTION_TYPE_TVO As String = "Мета: ТВО"
Private Const META_SECTION_TYPE_DOCUMENT As String = "Мета: документ"

Private Const PROFILE_TAG_PREFIX As String = "profile."
Private Const PROFILE_TAG_ALL_FIELDS As String = "profile.allFields"
Private Const PROFILE_TAG_CORE As String = "profile.core"
Private Const PROFILE_TAG_CLOSE_TREATMENT As String = "profile.closeTreatment"
Private Const PROFILE_TAG_CLOSE_TREATMENT_VACATION As String = "profile.closeTreatmentVacation"
Private Const PROFILE_TAG_CLOSE_ANNUAL_VACATION As String = "profile.closeAnnualVacation"
Private Const PROFILE_TAG_CLOSE_FAMILY_VACATION As String = "profile.closeFamilyVacation"
Private Const PROFILE_TAG_CLOSE_TREATMENT_MEDICAL_COMPANY As String = "profile.closeTreatmentMedicalCompany"
Private Const PROFILE_TAG_CLOSE_AMBULATORY_VLK As String = "profile.closeAmbulatoryVlk"
Private Const PROFILE_TAG_TO_TREATMENT As String = "profile.toTreatment"
Private Const PROFILE_TAG_TO_ANNUAL_VACATION_PART As String = "profile.toAnnualVacationPart"
Private Const PROFILE_TAG_TO_FAMILY_VACATION As String = "profile.toFamilyVacation"
Private Const PROFILE_TAG_TO_TREATMENT_VACATION As String = "profile.toTreatmentVacation"
Private Const PROFILE_TAG_TO_TREATMENT_MEDICAL_COMPANY As String = "profile.toTreatmentMedicalCompany"
Private Const PROFILE_TAG_TO_AMBULATORY_VLK As String = "profile.toAmbulatoryVlk"
Private Const PROFILE_TAG_TRANSFER_TREATMENT_TO_TREATMENT_VACATION As String = "profile.transferTreatmentToTreatmentVacation"
Private Const PROFILE_TAG_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION As String = "profile.transferTreatmentVacationToTreatmentVacation"
Private Const PROFILE_TAG_TRANSFER_TREATMENT_VACATION_TO_TREATMENT As String = "profile.transferTreatmentVacationToTreatment"
Private Const PROFILE_TAG_TRANSFER_TREATMENT_VACATION_TO_VLK As String = "profile.transferTreatmentVacationToVlk"
Private Const PROFILE_TAG_TRANSFER_VLK_TO_TREATMENT_VACATION As String = "profile.transferVlkToTreatmentVacation"
Private Const PROFILE_TAG_TRANSFER_VLK_TO_TREATMENT As String = "profile.transferVlkToTreatment"
Private Const PROFILE_TAG_TO_BUSINESS_TRIP As String = "profile.toBusinessTrip"
Private Const PROFILE_TAG_TO_BUSINESS_TRIP_SZCH As String = "profile.toBusinessTripSzch"
Private Const PROFILE_TAG_META_TVO As String = "profile.metaTvo"
Private Const PROFILE_TAG_META_DOCUMENT As String = "profile.metaDocument"

Private Const MOVEMENT_EVENT_STATIONARY_TREATMENT As String = "Стаціонарне лікування"
Private Const MOVEMENT_EVENT_ANNUAL_VACATION As String = "Щорічна відпустка"
Private Const MOVEMENT_EVENT_FAMILY_VACATION As String = "Відпустка за сімейними обставинами"
Private Const MOVEMENT_EVENT_TREATMENT_VACATION As String = "Відпустка для лікування"
Private Const MOVEMENT_EVENT_AMBULATORY_VLK As String = "Амбулаторне ВЛК"
Private Const MOVEMENT_EVENT_STATIONARY_VLK As String = "Стаціонарне ВЛК"

Private m_ProfileNames As Collection
Private m_MetaProfileNames As Collection
Private m_ProfileTagBySectionType As Object

Private Sub Class_Initialize()
    Set m_ProfileNames = private_BuildProfileNames()
    Set m_MetaProfileNames = private_BuildMetaProfileNames()
    Set m_ProfileTagBySectionType = private_BuildProfileTagMap()
End Sub

Public Property Get SectionTypeNames() As Collection
    Set SectionTypeNames = private_CopyCollection(m_ProfileNames)
End Property

Public Property Get ProfileNames() As Collection
    Set ProfileNames = private_CopyCollection(m_ProfileNames)
End Property

Public Property Get MetaProfileNames() As Collection
    Set MetaProfileNames = private_CopyCollection(m_MetaProfileNames)
End Property

Public Function ResolveProfileVisibilityState( _
    ByVal profileText As String, _
    ByVal tagsText As String _
) As String
    ' UI перечисляет profile.* теги на колонках формы.
    ' Provider выбирает теги активного профиля, а layout engine скрывает все,
    ' где нет пересечения. Контроллер при этом не знает ни алиасы колонок, ни теги.
    If private_ShouldShowTaggedControlForProfile(profileText, tagsText) Then
        ResolveProfileVisibilityState = "visible"
    Else
        ResolveProfileVisibilityState = "collapsed"
    End If
End Function

Public Property Get SectionTypeCloseFromTreatment() As String
    SectionTypeCloseFromTreatment = SECTION_TYPE_CLOSE_FROM_TREATMENT
End Property

Public Property Get SectionTypeCloseFromTreatmentVacation() As String
    SectionTypeCloseFromTreatmentVacation = SECTION_TYPE_CLOSE_FROM_TREATMENT_VACATION
End Property

Public Property Get SectionTypeCloseFromAnnualVacation() As String
    SectionTypeCloseFromAnnualVacation = SECTION_TYPE_CLOSE_FROM_ANNUAL_VACATION
End Property

Public Property Get SectionTypeCloseFromFamilyVacation() As String
    SectionTypeCloseFromFamilyVacation = SECTION_TYPE_CLOSE_FROM_FAMILY_VACATION
End Property

Public Property Get SectionTypeCloseFromTreatmentMedicalCompany() As String
    SectionTypeCloseFromTreatmentMedicalCompany = SECTION_TYPE_CLOSE_FROM_TREATMENT_MEDICAL_COMPANY
End Property

Public Property Get SectionTypeCloseFromAmbulatoryVlk() As String
    SectionTypeCloseFromAmbulatoryVlk = SECTION_TYPE_CLOSE_FROM_AMBULATORY_VLK
End Property

Public Property Get SectionTypeToTreatment() As String
    SectionTypeToTreatment = SECTION_TYPE_TO_TREATMENT
End Property

Public Property Get SectionTypeToAnnualVacationPart() As String
    SectionTypeToAnnualVacationPart = SECTION_TYPE_TO_ANNUAL_VACATION_PART
End Property

Public Property Get SectionTypeToFamilyVacation() As String
    SectionTypeToFamilyVacation = SECTION_TYPE_TO_FAMILY_VACATION
End Property

Public Property Get SectionTypeToTreatmentVacation() As String
    SectionTypeToTreatmentVacation = SECTION_TYPE_TO_TREATMENT_VACATION
End Property

Public Property Get SectionTypeToTreatmentMedicalCompany() As String
    SectionTypeToTreatmentMedicalCompany = SECTION_TYPE_TO_TREATMENT_MEDICAL_COMPANY
End Property

Public Property Get SectionTypeToAmbulatoryVlk() As String
    SectionTypeToAmbulatoryVlk = SECTION_TYPE_TO_AMBULATORY_VLK
End Property

Public Property Get SectionTypeTransferTreatmentToTreatmentVacation() As String
    SectionTypeTransferTreatmentToTreatmentVacation = SECTION_TYPE_TRANSFER_TREATMENT_TO_TREATMENT_VACATION
End Property

Public Property Get SectionTypeTransferTreatmentVacationToTreatmentVacation() As String
    SectionTypeTransferTreatmentVacationToTreatmentVacation = SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION
End Property

Public Property Get SectionTypeTransferTreatmentVacationToTreatment() As String
    SectionTypeTransferTreatmentVacationToTreatment = SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT
End Property

Public Property Get SectionTypeTransferTreatmentVacationToVlk() As String
    SectionTypeTransferTreatmentVacationToVlk = SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_VLK
End Property

Public Property Get SectionTypeTransferVlkToTreatmentVacation() As String
    SectionTypeTransferVlkToTreatmentVacation = SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT_VACATION
End Property

Public Property Get SectionTypeTransferVlkToTreatment() As String
    SectionTypeTransferVlkToTreatment = SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT
End Property

Public Property Get SectionTypeToBusinessTrip() As String
    SectionTypeToBusinessTrip = SECTION_TYPE_TO_BUSINESS_TRIP
End Property

Public Property Get SectionTypeToBusinessTripSzch() As String
    SectionTypeToBusinessTripSzch = SECTION_TYPE_TO_BUSINESS_TRIP_SZCH
End Property

Public Property Get MetaSectionTypeTvo() As String
    MetaSectionTypeTvo = META_SECTION_TYPE_TVO
End Property

Public Property Get MetaSectionTypeDocument() As String
    MetaSectionTypeDocument = META_SECTION_TYPE_DOCUMENT
End Property

Public Function IsMetaProfileName(ByVal profileText As String) As Boolean
    Select Case private_NormalizeText(profileText)
        Case private_NormalizeText(META_SECTION_TYPE_TVO), _
             private_NormalizeText(META_SECTION_TYPE_DOCUMENT)
            IsMetaProfileName = True
    End Select
End Function

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

Private Function private_BuildProfileNames() As Collection
    Dim profileNames As Collection

    Set profileNames = New Collection
    profileNames.Add SECTION_TYPE_CLOSE_FROM_TREATMENT
    profileNames.Add SECTION_TYPE_CLOSE_FROM_TREATMENT_VACATION
    profileNames.Add SECTION_TYPE_CLOSE_FROM_ANNUAL_VACATION
    profileNames.Add SECTION_TYPE_CLOSE_FROM_FAMILY_VACATION
    profileNames.Add SECTION_TYPE_CLOSE_FROM_TREATMENT_MEDICAL_COMPANY
    profileNames.Add SECTION_TYPE_CLOSE_FROM_AMBULATORY_VLK
    profileNames.Add SECTION_TYPE_TO_TREATMENT
    profileNames.Add SECTION_TYPE_TO_ANNUAL_VACATION_PART
    profileNames.Add SECTION_TYPE_TO_FAMILY_VACATION
    profileNames.Add SECTION_TYPE_TO_TREATMENT_VACATION
    profileNames.Add SECTION_TYPE_TO_TREATMENT_MEDICAL_COMPANY
    profileNames.Add SECTION_TYPE_TO_AMBULATORY_VLK
    profileNames.Add SECTION_TYPE_TRANSFER_TREATMENT_TO_TREATMENT_VACATION
    profileNames.Add SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION
    profileNames.Add SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT
    profileNames.Add SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_VLK
    profileNames.Add SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT_VACATION
    profileNames.Add SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT
    profileNames.Add SECTION_TYPE_TO_BUSINESS_TRIP
    profileNames.Add SECTION_TYPE_TO_BUSINESS_TRIP_SZCH

    Set private_BuildProfileNames = profileNames
End Function

Private Function private_BuildMetaProfileNames() As Collection
    Dim profileNames As Collection

    Set profileNames = New Collection
    profileNames.Add META_SECTION_TYPE_TVO
    profileNames.Add META_SECTION_TYPE_DOCUMENT

    Set private_BuildMetaProfileNames = profileNames
End Function

Private Function private_BuildProfileTagMap() As Object
    Dim tagMap As Object

    Set tagMap = VBA.CreateObject("Scripting.Dictionary")
    tagMap.CompareMode = 1

    tagMap(private_NormalizeText(SECTION_TYPE_CLOSE_FROM_TREATMENT)) = PROFILE_TAG_CLOSE_TREATMENT
    tagMap(private_NormalizeText(SECTION_TYPE_CLOSE_FROM_TREATMENT_VACATION)) = PROFILE_TAG_CLOSE_TREATMENT_VACATION
    tagMap(private_NormalizeText(SECTION_TYPE_CLOSE_FROM_ANNUAL_VACATION)) = PROFILE_TAG_CLOSE_ANNUAL_VACATION
    tagMap(private_NormalizeText(SECTION_TYPE_CLOSE_FROM_FAMILY_VACATION)) = PROFILE_TAG_CLOSE_FAMILY_VACATION
    tagMap(private_NormalizeText(SECTION_TYPE_CLOSE_FROM_TREATMENT_MEDICAL_COMPANY)) = PROFILE_TAG_CLOSE_TREATMENT_MEDICAL_COMPANY
    tagMap(private_NormalizeText(SECTION_TYPE_CLOSE_FROM_AMBULATORY_VLK)) = PROFILE_TAG_CLOSE_AMBULATORY_VLK
    tagMap(private_NormalizeText(SECTION_TYPE_TO_TREATMENT)) = PROFILE_TAG_TO_TREATMENT
    tagMap(private_NormalizeText(SECTION_TYPE_TO_ANNUAL_VACATION_PART)) = PROFILE_TAG_TO_ANNUAL_VACATION_PART
    tagMap(private_NormalizeText(SECTION_TYPE_TO_FAMILY_VACATION)) = PROFILE_TAG_TO_FAMILY_VACATION
    tagMap(private_NormalizeText(SECTION_TYPE_TO_TREATMENT_VACATION)) = PROFILE_TAG_TO_TREATMENT_VACATION
    tagMap(private_NormalizeText(SECTION_TYPE_TO_TREATMENT_MEDICAL_COMPANY)) = PROFILE_TAG_TO_TREATMENT_MEDICAL_COMPANY
    tagMap(private_NormalizeText(SECTION_TYPE_TO_AMBULATORY_VLK)) = PROFILE_TAG_TO_AMBULATORY_VLK
    tagMap(private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_TO_TREATMENT_VACATION)) = PROFILE_TAG_TRANSFER_TREATMENT_TO_TREATMENT_VACATION
    tagMap(private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION)) = PROFILE_TAG_TRANSFER_TREATMENT_VACATION_TO_TREATMENT_VACATION
    tagMap(private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_TREATMENT)) = PROFILE_TAG_TRANSFER_TREATMENT_VACATION_TO_TREATMENT
    tagMap(private_NormalizeText(SECTION_TYPE_TRANSFER_TREATMENT_VACATION_TO_VLK)) = PROFILE_TAG_TRANSFER_TREATMENT_VACATION_TO_VLK
    tagMap(private_NormalizeText(SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT_VACATION)) = PROFILE_TAG_TRANSFER_VLK_TO_TREATMENT_VACATION
    tagMap(private_NormalizeText(SECTION_TYPE_TRANSFER_VLK_TO_TREATMENT)) = PROFILE_TAG_TRANSFER_VLK_TO_TREATMENT
    tagMap(private_NormalizeText(SECTION_TYPE_TO_BUSINESS_TRIP)) = PROFILE_TAG_TO_BUSINESS_TRIP
    tagMap(private_NormalizeText(SECTION_TYPE_TO_BUSINESS_TRIP_SZCH)) = PROFILE_TAG_TO_BUSINESS_TRIP_SZCH
    tagMap(private_NormalizeText(META_SECTION_TYPE_TVO)) = PROFILE_TAG_META_TVO
    tagMap(private_NormalizeText(META_SECTION_TYPE_DOCUMENT)) = PROFILE_TAG_META_DOCUMENT

    Set private_BuildProfileTagMap = tagMap
End Function

Private Function private_ShouldShowTaggedControlForProfile( _
    ByVal profileText As String, _
    ByVal tagsText As String _
) As Boolean
    Dim visibleTags As Object
    Dim tagObj As Variant
    Dim tagText As String
    Dim hasProfileTags As Boolean

    tagsText = VBA.Trim$(tagsText)
    If VBA.Len(tagsText) = 0 Then
        private_ShouldShowTaggedControlForProfile = True
        Exit Function
    End If

    Set visibleTags = private_ProfileVisibleTags(profileText)
    If visibleTags Is Nothing Then Exit Function

    For Each tagObj In VBA.Split(tagsText, ";")
        tagText = private_NormalizeTagText(VBA.CStr(tagObj))
        If VBA.Len(tagText) = 0 Then GoTo ContinueTag
        If Not private_IsProfileTag(tagText) Then GoTo ContinueTag

        hasProfileTags = True
        If visibleTags.Exists(tagText) Then
            private_ShouldShowTaggedControlForProfile = True
            Exit Function
        End If

ContinueTag:
    Next tagObj

    ' Обычные layout tags без profile.* не участвуют в фильтрации видимости.
    private_ShouldShowTaggedControlForProfile = Not hasProfileTags
End Function

Private Function private_ProfileVisibleTags(ByVal profileText As String) As Object
    Dim result As Object
    Dim profileKey As String
    Dim profileTag As String

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    profileKey = private_NormalizeText(profileText)
    If m_ProfileTagBySectionType Is Nothing Then Set m_ProfileTagBySectionType = private_BuildProfileTagMap()
    If Not m_ProfileTagBySectionType.Exists(profileKey) Then Exit Function

    profileTag = VBA.Trim$(VBA.CStr(m_ProfileTagBySectionType(profileKey)))
    private_AddProfileTag result, profileTag
    If Not private_IsMetaProfileTag(profileTag) Then private_AddProfileTag result, PROFILE_TAG_CORE

    ' Для профилей без отдельной настройки формы показываем все колонки,
    ' помеченные wildcard-тегом profile.allFields.
    Select Case profileTag
        Case PROFILE_TAG_TO_BUSINESS_TRIP, PROFILE_TAG_TO_BUSINESS_TRIP_SZCH
            private_AddProfileTag result, PROFILE_TAG_ALL_FIELDS
    End Select

    Set private_ProfileVisibleTags = result
End Function

Private Function private_IsMetaProfileTag(ByVal profileTag As String) As Boolean
    Select Case private_NormalizeTagText(profileTag)
        Case private_NormalizeTagText(PROFILE_TAG_META_TVO), _
             private_NormalizeTagText(PROFILE_TAG_META_DOCUMENT)
            private_IsMetaProfileTag = True
    End Select
End Function

Private Function private_IsProfileTag(ByVal tagText As String) As Boolean
    tagText = private_NormalizeTagText(tagText)
    private_IsProfileTag = (VBA.Left$(tagText, VBA.Len(PROFILE_TAG_PREFIX)) = PROFILE_TAG_PREFIX)
End Function

Private Sub private_AddProfileTag(ByVal tags As Object, ByVal tagText As String)
    tagText = private_NormalizeTagText(tagText)
    If tags Is Nothing Then Exit Sub
    If VBA.Len(tagText) = 0 Then Exit Sub
    tags(tagText) = True
End Sub

Private Function private_NormalizeTagText(ByVal tagText As String) As String
    tagText = VBA.CStr(tagText)
    tagText = VBA.Replace(tagText, VBA.vbCr, " ")
    tagText = VBA.Replace(tagText, VBA.vbLf, " ")
    tagText = VBA.Replace(tagText, VBA.vbTab, " ")
    private_NormalizeTagText = VBA.LCase$(VBA.Trim$(tagText))
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
