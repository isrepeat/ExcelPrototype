VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SDB_Data"
Option Explicit

Private Const SECTION_TYPE_TO_BUSINESS_TRIP As String = "У відрядження"
Private Const PROFILE_TAG_PREFIX As String = "profile."
Private Const PROFILE_TAG_TO_BUSINESS_TRIP As String = "profile.toBusinessTrip"

Private m_SectionNames As Collection
Private m_ProfileTagBySection As Object

Private Sub Class_Initialize()
    ' Data-класс владеет каталогом секций и их UI-тегами. Controller не должен
    ' знать текстовые имена секций: добавление нового документа ограничивается
    ' регистрацией секции здесь и разметкой соответствующих profile.* полей.
    Set m_SectionNames = New Collection
    m_SectionNames.Add SECTION_TYPE_TO_BUSINESS_TRIP

    Set m_ProfileTagBySection = ex_Helpers.fn_CreateDictionaryTextCompare()
    m_ProfileTagBySection( _
        ex_Helpers.fn_NormalizeText(SECTION_TYPE_TO_BUSINESS_TRIP)) = _
        PROFILE_TAG_TO_BUSINESS_TRIP
End Sub

Public Property Get SectionNames() As Collection
    Dim result As Collection
    Dim sectionObj As Variant

    Set result = New Collection
    For Each sectionObj In m_SectionNames
        result.Add VBA.CStr(sectionObj)
    Next sectionObj
    Set SectionNames = result
End Property

Public Property Get DefaultSectionName() As String
    If m_SectionNames Is Nothing Then Exit Property
    If m_SectionNames.Count = 0 Then Exit Property
    DefaultSectionName = VBA.CStr(m_SectionNames.Item(1))
End Property

Public Function IsSectionName(ByVal sectionText As String) As Boolean
    If m_ProfileTagBySection Is Nothing Then Exit Function
    IsSectionName = m_ProfileTagBySection.Exists( _
        ex_Helpers.fn_NormalizeText(sectionText))
End Function

Public Function ResolveProfileVisibilityState( _
    ByVal sectionText As String, _
    ByVal tagsText As String _
) As String
    If private_ShouldShowControl(sectionText, tagsText) Then
        ResolveProfileVisibilityState = "visible"
    Else
        ResolveProfileVisibilityState = "collapsed"
    End If
End Function

Private Function private_ShouldShowControl( _
    ByVal sectionText As String, _
    ByVal tagsText As String _
) As Boolean
    Dim sectionKey As String
    Dim sectionProfileTag As String
    Dim tagObj As Variant
    Dim tagText As String
    Dim hasProfileTag As Boolean

    If VBA.Len(VBA.Trim$(tagsText)) = 0 Then
        private_ShouldShowControl = True
        Exit Function
    End If

    sectionKey = ex_Helpers.fn_NormalizeText(sectionText)
    If m_ProfileTagBySection Is Nothing Then Exit Function
    If Not m_ProfileTagBySection.Exists(sectionKey) Then Exit Function
    sectionProfileTag = ex_Helpers.fn_NormalizeText( _
        VBA.CStr(m_ProfileTagBySection(sectionKey)))

    For Each tagObj In VBA.Split(tagsText, ";")
        tagText = ex_Helpers.fn_NormalizeText(VBA.CStr(tagObj))
        If VBA.Left$(tagText, VBA.Len(PROFILE_TAG_PREFIX)) = PROFILE_TAG_PREFIX Then
            hasProfileTag = True
            If VBA.StrComp( _
                tagText, sectionProfileTag, VBA.vbTextCompare) = 0 Then
                private_ShouldShowControl = True
                Exit Function
            End If
        End If
    Next tagObj

    ' Обычные layout-теги не участвуют в фильтрации по секции.
    private_ShouldShowControl = Not hasProfileTag
End Function
