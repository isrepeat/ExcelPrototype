VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SIP_ModeFolders"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Implements obj_ISelectItemsSourceProvider

Private m_ProviderKey As String
Private m_ModesRootRelativePath As String
Private m_OnSelectMacro As String

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
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
    Dim modesRootPath As String
    Dim modesRootFolder As Object
    Dim lastModified As Date
    Dim folderCount As Long

    outStamp = VBA.vbNullString

    If Not private_TryResolveModesRoot(modesRootPath, modesRootFolder) Then Exit Function

    On Error GoTo EH_STAMP
    lastModified = VBA.CDate(modesRootFolder.DateLastModified)
    folderCount = VBA.CLng(modesRootFolder.SubFolders.Count)
    On Error GoTo 0

    ' Stamp режима = modified-time корневой папки + число подпапок.
    ' Этого достаточно, чтобы дешево детектить добавление/удаление mode folders.
    outStamp = VBA.Format$(lastModified, "yyyy-mm-dd hh:nn:ss") & "|" & VBA.CStr(folderCount)
    obj_ISelectItemsSourceProvider_TryGetCurrentStamp = True
    Exit Function

EH_STAMP:
    private_ReportError "PrototypeNew: failed to read mode folders stamp under '" & modesRootPath & "': " & Err.Description
End Function

Private Function obj_ISelectItemsSourceProvider_TryBuildItems(ByRef outItems As Collection) As Boolean
    Dim modesRootPath As String
    Dim modesRootFolder As Object
    Dim modeFolder As Object
    Dim modeName As String
    Dim modeNames() As String
    Dim modeCount As Long
    Dim i As Long

    Set outItems = New Collection

    If Not private_TryResolveModesRoot(modesRootPath, modesRootFolder) Then Exit Function

    ' Строим select options из подпапок modes.
    For Each modeFolder In modesRootFolder.SubFolders
        modeName = VBA.Trim$(VBA.CStr(modeFolder.Name))
        If VBA.Len(modeName) = 0 Then GoTo ContinueModeFolder

        modeCount = modeCount + 1
        ReDim Preserve modeNames(1 To modeCount)
        modeNames(modeCount) = modeName
ContinueModeFolder:
    Next modeFolder

    If modeCount <= 0 Then
        private_ReportError "PrototypeNew: no valid mode folders were found under '" & modesRootPath & "'."
        Exit Function
    End If

    ' Стабильный порядок, чтобы UI не "прыгал" между рендерами.
    private_SortTextInPlace modeNames, 1, modeCount
    For i = 1 To modeCount
        outItems.Add private_CreateSelectOption(modeNames(i), modeNames(i), m_OnSelectMacro)
    Next i

    obj_ISelectItemsSourceProvider_TryBuildItems = True
End Function

' //
' // API
' //
Public Function Initialize( _
    ByVal providerKey As String, _
    ByVal modesRootRelativePath As String, _
    ByVal onSelectMacro As String _
) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    providerKey = VBA.LCase$(VBA.Trim$(providerKey))
    modesRootRelativePath = VBA.Trim$(modesRootRelativePath)
    onSelectMacro = VBA.Trim$(onSelectMacro)

    If VBA.Len(providerKey) = 0 Then
        private_ReportError "PrototypeNew: mode source provider key is empty."
        Exit Function
    End If
    If VBA.Len(modesRootRelativePath) = 0 Then
        private_ReportError "PrototypeNew: mode source root path is empty."
        Exit Function
    End If
    If VBA.Len(onSelectMacro) = 0 Then
        private_ReportError "PrototypeNew: mode source onSelect macro is empty."
        Exit Function
    End If

    m_ProviderKey = providerKey
    m_ModesRootRelativePath = modesRootRelativePath
    m_OnSelectMacro = onSelectMacro
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    m_ProviderKey = VBA.vbNullString
    m_ModesRootRelativePath = VBA.vbNullString
    m_OnSelectMacro = VBA.vbNullString
    On Error GoTo 0
End Sub

Public Function Configure( _
    ByVal providerKey As String, _
    ByVal modesRootRelativePath As String, _
    ByVal onSelectMacro As String _
) As Boolean
    ' Backward-compatible wrapper.
    Configure = Initialize(providerKey, modesRootRelativePath, onSelectMacro)
End Function

' //
' // Internal
' //
Private Function private_TryResolveModesRoot( _
    ByRef outRootPath As String, _
    ByRef outRootFolder As Object _
) As Boolean
    Dim fso As Object

    outRootPath = VBA.Trim$(ex_XmlCore.m_CombineBasePath(ThisWorkbook, m_ModesRootRelativePath))
    Set outRootFolder = Nothing

    If VBA.Len(outRootPath) = 0 Then
        private_ReportError "PrototypeNew: failed to resolve modes root path from '" & m_ModesRootRelativePath & "'."
        Exit Function
    End If

    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then
        private_ReportError "PrototypeNew: failed to create FileSystemObject for modes provider."
        Exit Function
    End If
    If Not fso.FolderExists(outRootPath) Then
        private_ReportError "PrototypeNew: modes folder was not found: " & outRootPath
        Exit Function
    End If

    Set outRootFolder = fso.GetFolder(outRootPath)
    private_TryResolveModesRoot = True
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

Private Sub private_SortTextInPlace( _
    ByRef values() As String, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long _
)
    Dim i As Long
    Dim j As Long
    Dim keyText As String

    If lastIndex <= firstIndex Then Exit Sub

    For i = firstIndex + 1 To lastIndex
        keyText = values(i)
        j = i - 1
        Do While j >= firstIndex
            If VBA.StrComp(values(j), keyText, VBA.vbTextCompare) <= 0 Then Exit Do
            values(j + 1) = values(j)
            j = j - 1
        Loop
        values(j + 1) = keyText
    Next i
End Sub

Private Sub private_ReportError(ByVal messageText As String)
    messageText = VBA.Trim$(messageText)
    If VBA.Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError messageText
#End If
    MsgBox messageText, vbExclamation, "PrototypeNew / Select provider"
End Sub
