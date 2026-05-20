VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SIP_ModeProfilesXml"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Implements obj_ISelectItemsSourceProvider

Private m_ProviderKey As String
Private m_ModesRootRelativePath As String
Private m_ModeProfilesFileSuffix As String
Private m_OnSelectMacro As String
Private m_CurrentModeId As String
Private m_CurrentProfilesFilePath As String

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_ISelectItemsSourceProvider_GetProviderKey() As String
    obj_ISelectItemsSourceProvider_GetProviderKey = m_ProviderKey
End Function

Private Function obj_ISelectItemsSourceProvider_TryGetCurrentStamp(ByRef outStamp As String) As Boolean
    Dim filePath As String
    Dim fso As Object
    Dim fileObj As Object
    Dim lastModified As Date
    Dim fileSize As Double

    outStamp = VBA.vbNullString

    If VBA.Len(VBA.Trim$(m_CurrentModeId)) = 0 Then
        private_ReportError "PrototypeNew: current mode id is not specified for profiles provider."
        Exit Function
    End If
    If Not private_TryResolveProfilesFilePathByModeId(m_CurrentModeId, filePath) Then Exit Function
    If VBA.Len(VBA.Trim$(Dir$(filePath))) = 0 Then
        private_ReportError "PrototypeNew: profiles file was not found for mode '" & m_CurrentModeId & "': " & filePath
        Exit Function
    End If

    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then
        private_ReportError "PrototypeNew: failed to create FileSystemObject for profiles provider."
        Exit Function
    End If
    Set fileObj = fso.GetFile(filePath)
    If fileObj Is Nothing Then
        private_ReportError "PrototypeNew: failed to access profiles file '" & filePath & "'."
        Exit Function
    End If

    On Error GoTo EH_STAMP
    lastModified = VBA.CDate(fileObj.DateLastModified)
    fileSize = VBA.CDbl(fileObj.Size)
    On Error GoTo 0

    m_CurrentProfilesFilePath = filePath
    ' Stamp привязан и к режиму, и к файлу профилей.
    ' Если modeId другой или файл изменился (дата/размер), будет cache-miss.
    outStamp = VBA.LCase$(m_CurrentModeId) & "|" & VBA.Format$(lastModified, "yyyy-mm-dd hh:nn:ss") & "|" & VBA.CStr(fileSize)
    obj_ISelectItemsSourceProvider_TryGetCurrentStamp = True
    Exit Function

EH_STAMP:
    private_ReportError "PrototypeNew: failed to read profiles file stamp '" & filePath & "': " & Err.Description
End Function

Private Function obj_ISelectItemsSourceProvider_TryBuildItems(ByRef outItems As Collection) As Boolean
    Dim filePath As String
    Dim dom As Object

    Set outItems = Nothing

    If VBA.Len(VBA.Trim$(m_CurrentModeId)) = 0 Then
        private_ReportError "PrototypeNew: current mode id is not specified for profiles provider."
        Exit Function
    End If
    If Not private_TryResolveProfilesFilePathByModeId(m_CurrentModeId, filePath) Then Exit Function
    If VBA.Len(VBA.Trim$(Dir$(filePath))) = 0 Then
        private_ReportError "PrototypeNew: profiles file was not found for mode '" & m_CurrentModeId & "': " & filePath
        Exit Function
    End If

    ' Загружаем XML выбранного режима и превращаем профили в коллекцию SelectOption.
    Set dom = ex_XmlCore.fn_LoadDomByFilePath( _
        filePath, _
        "PrototypeNew: profiles file was not found: ", _
        "PrototypeNew: failed to parse profiles file: ", _
        VBA.vbNullString)
    If dom Is Nothing Then
        private_ReportError "PrototypeNew: failed to load profiles file '" & filePath & "'."
        Exit Function
    End If

    If Not private_TryCollectProfileSelectOptionsFromDom(dom, outItems) Then Exit Function
    If outItems Is Nothing Then Exit Function
    If outItems.Count = 0 Then
        private_ReportError "PrototypeNew: profiles file '" & filePath & "' does not contain selectable profiles."
        Exit Function
    End If

    m_CurrentProfilesFilePath = filePath
    obj_ISelectItemsSourceProvider_TryBuildItems = True
End Function

' //
' // API
' //
Public Function Initialize( _
    ByVal providerKey As String, _
    ByVal modesRootRelativePath As String, _
    ByVal modeProfilesFileSuffix As String, _
    ByVal onSelectMacro As String _
) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    providerKey = VBA.LCase$(VBA.Trim$(providerKey))
    modesRootRelativePath = VBA.Trim$(modesRootRelativePath)
    modeProfilesFileSuffix = VBA.Trim$(modeProfilesFileSuffix)
    onSelectMacro = VBA.Trim$(onSelectMacro)

    If VBA.Len(providerKey) = 0 Then
        private_ReportError "PrototypeNew: profile source provider key is empty."
        Exit Function
    End If
    If VBA.Len(modesRootRelativePath) = 0 Then
        private_ReportError "PrototypeNew: profile source root path is empty."
        Exit Function
    End If
    If VBA.Len(modeProfilesFileSuffix) = 0 Then
        private_ReportError "PrototypeNew: profile source file suffix is empty."
        Exit Function
    End If
    If VBA.Len(onSelectMacro) = 0 Then
        private_ReportError "PrototypeNew: profile source onSelect macro is empty."
        Exit Function
    End If

    m_ProviderKey = providerKey
    m_ModesRootRelativePath = modesRootRelativePath
    m_ModeProfilesFileSuffix = modeProfilesFileSuffix
    m_OnSelectMacro = onSelectMacro
    m_CurrentModeId = VBA.vbNullString
    m_CurrentProfilesFilePath = VBA.vbNullString
    Initialize = True
End Function

Public Function SetCurrentModeId(ByVal modeId As String) As Boolean
    modeId = VBA.Trim$(modeId)
    If Not private_IsSafeModeId(modeId) Then
        private_ReportError "PrototypeNew: unsafe mode id '" & modeId & "'."
        Exit Function
    End If

    m_CurrentModeId = modeId
    m_CurrentProfilesFilePath = VBA.vbNullString
    SetCurrentModeId = True
End Function

Public Property Get CurrentProfilesFilePath() As String
    CurrentProfilesFilePath = m_CurrentProfilesFilePath
End Property

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    m_ProviderKey = VBA.vbNullString
    m_ModesRootRelativePath = VBA.vbNullString
    m_ModeProfilesFileSuffix = VBA.vbNullString
    m_OnSelectMacro = VBA.vbNullString
    m_CurrentModeId = VBA.vbNullString
    m_CurrentProfilesFilePath = VBA.vbNullString
    On Error GoTo 0
End Sub

' //
' // Internal
' //
Private Function private_IsSafeModeId(ByVal modeId As String) As Boolean
    modeId = VBA.Trim$(modeId)
    If VBA.Len(modeId) = 0 Then Exit Function
    If VBA.InStr(1, modeId, "\", VBA.vbBinaryCompare) > 0 Then Exit Function
    If VBA.InStr(1, modeId, "/", VBA.vbBinaryCompare) > 0 Then Exit Function
    If VBA.InStr(1, modeId, ":", VBA.vbBinaryCompare) > 0 Then Exit Function
    If VBA.InStr(1, modeId, "..", VBA.vbBinaryCompare) > 0 Then Exit Function

    private_IsSafeModeId = True
End Function

Private Function private_TryResolveProfilesFilePathByModeId(ByVal modeId As String, ByRef outFilePath As String) As Boolean
    Dim modesRootPath As String

    outFilePath = VBA.vbNullString
    modeId = VBA.Trim$(modeId)

    If Not private_IsSafeModeId(modeId) Then
        private_ReportError "PrototypeNew: unsafe mode id '" & modeId & "'."
        Exit Function
    End If

    modesRootPath = VBA.Trim$(ex_XmlCore.fn_CombineBasePath(ThisWorkbook, m_ModesRootRelativePath))
    If VBA.Len(modesRootPath) = 0 Then
        private_ReportError "PrototypeNew: failed to resolve modes root path from '" & m_ModesRootRelativePath & "'."
        Exit Function
    End If

    ' Формат файла профилей: modes\<ModeId>\<ModeId>Profiles.xml
    outFilePath = modesRootPath & "\" & modeId & "\" & modeId & m_ModeProfilesFileSuffix
    private_TryResolveProfilesFilePathByModeId = True
End Function

Private Function private_TryCollectProfileSelectOptionsFromDom(ByVal dom As Object, ByRef outOptions As Collection) As Boolean
    Dim profileNodes As Object
    Dim profileNode As Object
    Dim seenIds As Object
    Dim optionItem As obj_SelectOption
    Dim optionId As String

    Set outOptions = New Collection
    Set seenIds = VBA.CreateObject("Scripting.Dictionary")
    seenIds.CompareMode = 1

    If dom Is Nothing Then
        private_ReportError "PrototypeNew: profiles DOM is not specified."
        Exit Function
    End If

    On Error GoTo EH_XML
    Set profileNodes = dom.selectNodes("//*[local-name()='profile' or local-name()='configProfile' or local-name()='preset' or local-name()='variant']")
    On Error GoTo 0

    If Not profileNodes Is Nothing Then
        For Each profileNode In profileNodes
            Set optionItem = Nothing
            If Not private_TryCreateProfileSelectOptionFromNode(profileNode, optionItem) Then Exit Function
            If optionItem Is Nothing Then GoTo ContinueProfileNodePrimary

            optionId = VBA.LCase$(VBA.Trim$(optionItem.Id))
            If VBA.Len(optionId) = 0 Then GoTo ContinueProfileNodePrimary
            If seenIds.Exists(optionId) Then GoTo ContinueProfileNodePrimary

            seenIds(optionId) = True
            outOptions.Add optionItem
ContinueProfileNodePrimary:
        Next profileNode
    End If

    If outOptions.Count > 0 Then
        private_TryCollectProfileSelectOptionsFromDom = True
        Exit Function
    End If

    On Error GoTo EH_XML
    Set profileNodes = dom.selectNodes("//*[" & _
                                      "(@id or @name or @key)" & _
                                      " and " & _
                                      "(.//*[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config']" & _
                                      " or *[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config'])" & _
                                      "]")
    On Error GoTo 0

    If Not profileNodes Is Nothing Then
        For Each profileNode In profileNodes
            Set optionItem = Nothing
            If Not private_TryCreateProfileSelectOptionFromNode(profileNode, optionItem) Then Exit Function
            If optionItem Is Nothing Then GoTo ContinueProfileNodeFallback

            optionId = VBA.LCase$(VBA.Trim$(optionItem.Id))
            If VBA.Len(optionId) = 0 Then GoTo ContinueProfileNodeFallback
            If seenIds.Exists(optionId) Then GoTo ContinueProfileNodeFallback

            seenIds(optionId) = True
            outOptions.Add optionItem
ContinueProfileNodeFallback:
        Next profileNode
    End If

    private_TryCollectProfileSelectOptionsFromDom = True
    Exit Function

EH_XML:
    private_ReportError "PrototypeNew: failed to read profile list from XML: " & Err.Description
End Function

Private Function private_TryCreateProfileSelectOptionFromNode( _
    ByVal profileNode As Object, _
    ByRef outOption As obj_SelectOption _
) As Boolean
    Dim profileId As String
    Dim captionText As String

    Set outOption = Nothing
    If profileNode Is Nothing Then
        private_TryCreateProfileSelectOptionFromNode = True
        Exit Function
    End If

    profileId = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(profileNode, "id")))
    If VBA.Len(profileId) = 0 Then profileId = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(profileNode, "name")))
    If VBA.Len(profileId) = 0 Then profileId = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(profileNode, "key")))
    If VBA.Len(profileId) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "id", profileId) Then Exit Function
    End If
    If VBA.Len(profileId) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "name", profileId) Then Exit Function
    End If
    If VBA.Len(profileId) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "key", profileId) Then Exit Function
    End If
    If VBA.Len(profileId) = 0 Then
        private_TryCreateProfileSelectOptionFromNode = True
        Exit Function
    End If

    captionText = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(profileNode, "caption")))
    If VBA.Len(captionText) = 0 Then captionText = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(profileNode, "title")))
    If VBA.Len(captionText) = 0 Then captionText = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(profileNode, "display")))
    If VBA.Len(captionText) = 0 Then captionText = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(profileNode, "name")))
    If VBA.Len(captionText) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "caption", captionText) Then Exit Function
    End If
    If VBA.Len(captionText) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "title", captionText) Then Exit Function
    End If
    If VBA.Len(captionText) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "display", captionText) Then Exit Function
    End If
    If VBA.Len(captionText) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "name", captionText) Then Exit Function
    End If
    If VBA.Len(captionText) = 0 Then captionText = profileId

    Set outOption = private_CreateSelectOption(captionText, profileId, m_OnSelectMacro)
    private_TryCreateProfileSelectOptionFromNode = True
End Function

Private Function private_TryReadChildNodeText( _
    ByVal parentNode As Object, _
    ByVal childLocalName As String, _
    ByRef outText As String _
) As Boolean
    Dim childNode As Object

    outText = VBA.vbNullString
    If parentNode Is Nothing Then
        private_TryReadChildNodeText = True
        Exit Function
    End If

    On Error GoTo EH_XML
    Set childNode = parentNode.selectSingleNode("./*[local-name()='" & childLocalName & "']")
    On Error GoTo 0

    If Not childNode Is Nothing Then outText = VBA.Trim$(VBA.CStr(childNode.Text))
    private_TryReadChildNodeText = True
    Exit Function

EH_XML:
    private_ReportError "PrototypeNew: failed to read child node '" & childLocalName & "': " & Err.Description
End Function

Private Function private_CreateSelectOption( _
    ByVal captionText As String, _
    ByVal idText As String, _
    ByVal onSelectMacro As String _
) As obj_SelectOption
    Dim selectOption As obj_SelectOption

    Set selectOption = New obj_SelectOption
    selectOption.Caption = VBA.CStr(captionText)
    selectOption.Id = VBA.CStr(idText)
    selectOption.OnSelect = VBA.CStr(onSelectMacro)
    Set private_CreateSelectOption = selectOption
End Function

Private Sub private_ReportError(ByVal messageText As String)
    messageText = VBA.Trim$(messageText)
    If VBA.Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError messageText
#End If
    MsgBox messageText, vbExclamation, "PrototypeNew / Select provider"
End Sub
