' Must be pasted in internal .xlsm module
Option Explicit

Private Const BASE_DIR As String = "vba\\"
Private Const IMPORT_CACHE_FILE As String = ".devtools_import_cache.txt"
Private Const ENABLE_CLASS_IMPORT_VALIDATION As Boolean = False
Private Const MAX_IMPORT_RECURSION_DEPTH As Long = 4
Private Const COMP_TYPE_MODULE As String = "module"
Private Const COMP_TYPE_CLASS As String = "class"
Private Const COMP_TYPE_SHEET As String = "sheet"
Private Const COMP_TYPE_WORKBOOK As String = "workbook"
Private Const MAX_VBA_COMPONENT_NAME_LEN As Long = 31

'==========================
' Public API
'==========================
' Main updater (legacy name preserved).
Public Sub dev_UpdateCode()
    mp_UpdateCodeCore False
End Sub

Public Sub dev_UpdateCodeFast()
    mp_UpdateCodeCore True
End Sub

' Ribbon hook (keeps existing button working if mapped).
Public Sub dev_OnUpdateCodeClicked(ByVal control As Object)
    dev_UpdateCode
End Sub

Public Sub dev_OnUpdateCodeFastClicked(ByVal control As Object)
    dev_UpdateCodeFast
End Sub

Public Sub dev_RemoveAllModulesAndClasses()
    On Error GoTo EH
    Application.ScreenUpdating = False

    mp_RemoveAllModulesAndClasses

    Application.ScreenUpdating = True
    MsgBox "All modules and classes removed (DevTools kept).", vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Remove modules/classes failed: " & Err.Description, vbExclamation
End Sub

Public Sub mp_RemoveAllModulesAndClasses()
    Dim prj As Object
    Dim comp As Object
    Dim names() As String
    Dim n As Long
    Dim i As Long

    Set prj = ThisWorkbook.VBProject

    For Each comp In prj.VBComponents
        Select Case comp.Type
            Case 1, 2 ' vbext_ct_StdModule, vbext_ct_ClassModule
                If StrComp(comp.Name, "DevTools", vbTextCompare) <> 0 Then
                    n = n + 1
                    ReDim Preserve names(1 To n)
                    names(n) = comp.Name
                End If
        End Select
    Next comp

    For i = 1 To n
        On Error GoTo EH_REMOVE
        prj.VBComponents.Remove prj.VBComponents(names(i))
        On Error GoTo 0
    Next i

    Exit Sub

EH_REMOVE:
    Err.Raise vbObjectError + 1008, "mp_RemoveAllModulesAndClasses", _
              "Failed to remove component '" & names(i) & "': " & Err.Description
End Sub

Private Sub mp_UpdateCodeCore(ByVal fastMode As Boolean)
    Dim basePath As String
    Dim cachePath As String
    Dim prevCache As Object
    Dim nextCache As Object

    basePath = ThisWorkbook.Path & "\\" & BASE_DIR
    If Len(Dir(basePath, vbDirectory)) = 0 Then
        MsgBox "Workbook path is empty or vba folder not found. Save the file first.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo EH

    cachePath = basePath & IMPORT_CACHE_FILE
    Set prevCache = mp_LoadImportCache(cachePath)
    Set nextCache = mp_CreateDictionary()

    If Not fastMode Then
        mp_RemoveImportedModules
    End If

    mp_ImportFolder basePath, fastMode, prevCache, nextCache
    If ENABLE_CLASS_IMPORT_VALIDATION Then
        mp_ValidateClassImports basePath
    End If

    If fastMode Then
        mp_RemoveStaleImportedComponents prevCache, nextCache
    End If
    mp_SaveImportCache cachePath, nextCache

    Application.ScreenUpdating = True
    mp_ShowCodeUpdatedNotice
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Update Code failed: " & Err.Description, vbExclamation
End Sub

Private Sub mp_ValidateClassImports(ByVal rootPath As String)
    Dim fso As Object
    Dim failed As String

    If Dir(rootPath, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1006, "mp_ValidateClassImports", "VBA root folder not found: " & rootPath
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    mp_ValidateClassImportsRecursive fso.GetFolder(rootPath), 0, failed

    If Len(failed) > 0 Then
        Err.Raise vbObjectError + 1007, "mp_ValidateClassImports", "Class import validation failed:" & failed
    End If
End Sub

Private Sub mp_ValidateClassImportsRecursive( _
    ByVal folderObj As Object, _
    ByVal depth As Long, _
    ByRef failed As String _
)
    Dim fileObj As Object
    Dim subFolder As Object
    Dim compType As String
    Dim fallbackName As String
    Dim className As String
    Dim vbComp As Object

    If folderObj Is Nothing Then Exit Sub
    If depth > MAX_IMPORT_RECURSION_DEPTH Then Exit Sub

    For Each fileObj In folderObj.Files
        If Not mp_TryResolveFileComponentType(CStr(fileObj.Name), compType, fallbackName) Then GoTo ContinueFile
        If StrComp(compType, COMP_TYPE_CLASS, vbTextCompare) <> 0 Then GoTo ContinueFile

        className = mp_GetComponentNameFromSource(CStr(fileObj.Path))
        Set vbComp = Nothing
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(className)
        On Error GoTo 0

        If vbComp Is Nothing Then
            failed = failed & vbCrLf & "- missing class: " & className
        ElseIf vbComp.Type <> 2 Then ' vbext_ct_ClassModule
            failed = failed & vbCrLf & "- wrong component type for class '" & className & "': " & CStr(vbComp.Type)
        End If

ContinueFile:
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        mp_ValidateClassImportsRecursive subFolder, depth + 1, failed
    Next subFolder
End Sub

Private Sub mp_ShowCodeUpdatedNotice()
    On Error GoTo ShowMsgBox
    Application.Run "ex_Messaging.m_ShowNotice", "Code updated", 2
    Exit Sub

ShowMsgBox:
    MsgBox "Code updated", vbInformation
End Sub

'==========================
' Module management
'==========================
Private Sub mp_RemoveImportedModules()
    Dim prj As Object
    Dim comp As Object
    Dim names() As String
    Dim n As Long
    Dim i As Long

    Set prj = ThisWorkbook.VBProject

    For Each comp In prj.VBComponents
        If comp.Type <> 100 Then ' vbext_ct_Document
            If StrComp(comp.name, "DevTools", vbTextCompare) <> 0 Then
                n = n + 1
                ReDim Preserve names(1 To n)
                names(n) = comp.name
            End If
        End If
    Next comp

    For i = 1 To n
        On Error GoTo EH_REMOVE
        prj.VBComponents.Remove prj.VBComponents(names(i))
        On Error GoTo 0
    Next i

    Exit Sub

EH_REMOVE:
    Err.Raise vbObjectError + 1004, "mp_RemoveImportedModules", _
              "Failed to remove component '" & names(i) & "': " & Err.Description
End Sub

Private Sub mp_ImportFolder( _
    ByVal folderPath As String, _
    ByVal fastMode As Boolean, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object _
)
    Dim fso As Object
    Dim rootFolder As Object
    Dim failed As String

    If Dir(folderPath, vbDirectory) = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(folderPath)
    mp_ImportFolderRecursive rootFolder, 0, failed, fastMode, prevCache, nextCache

    If Len(failed) > 0 Then
        Err.Raise vbObjectError + 1001, "mp_ImportFolder", "Import failed for file(s):" & failed
    End If
End Sub

Private Sub mp_ImportFolderRecursive( _
    ByVal folderObj As Object, _
    ByVal depth As Long, _
    ByRef failed As String, _
    ByVal fastMode As Boolean, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object _
)
    Dim fileObj As Object
    Dim subFolder As Object
    Dim importPath As String
    Dim fileName As String
    Dim normalizedFileName As String
    Dim componentName As String
    Dim errText As String
    Dim sourceText As String
    Dim fileStamp As String
    Dim cacheKey As String
    Dim compType As String
    Dim fallbackName As String
    Dim componentNameForCache As String
    Dim sourceEncodingIsUtf8 As Boolean

    If folderObj Is Nothing Then Exit Sub
    If depth > MAX_IMPORT_RECURSION_DEPTH Then Exit Sub

    For Each fileObj In folderObj.Files
        fileName = CStr(fileObj.Name)
        normalizedFileName = LCase$(fileName)

        If mp_TryResolveFileComponentType(fileName, compType, fallbackName) Then
            If Not mp_IsDevToolsSourceFile(normalizedFileName) Then
                importPath = CStr(fileObj.Path)
                On Error GoTo EH_IMPORT_FILE

                fileStamp = mp_BuildFileStampFromFileObject(fileObj)
                cacheKey = mp_NormalizeCacheKey(importPath)
                sourceText = vbNullString
                sourceEncodingIsUtf8 = mp_HasUtf8MarkerBeforeVba(fileName)

                Select Case LCase$(compType)
                    Case COMP_TYPE_MODULE, COMP_TYPE_CLASS
                        If fastMode Then
                            If mp_TryGetCachedComponentNameByStamp(prevCache, cacheKey, compType, fileStamp, componentName) Then
                                If mp_IsComponentPresentForType(componentName, compType) Then
                                    mp_SetCacheRecord nextCache, cacheKey, compType, componentName, fileStamp
                                    GoTo ContinueNextFile
                                End If
                            End If
                        End If

                        sourceText = mp_ReadAllText(importPath, sourceEncodingIsUtf8)
                        componentName = mp_GetComponentNameFromSourceText(sourceText, fallbackName)
                        mp_EnsureValidComponentNameLength componentName, importPath

                        mp_RemoveComponentIfExists componentName
                        If StrComp(compType, COMP_TYPE_MODULE, vbTextCompare) = 0 Then
                            mp_ImportStandardModuleFromSource componentName, importPath, sourceText
                        Else
                            mp_ImportClassModuleFromSource componentName, importPath, sourceText
                        End If
                        mp_SetCacheRecord nextCache, cacheKey, compType, componentName, fileStamp

                    Case COMP_TYPE_SHEET
                        componentNameForCache = mp_ResolveSheetCodeName(fallbackName)
                        If Len(componentNameForCache) = 0 Then GoTo ContinueNextFile

                        If fastMode Then
                            If mp_IsCacheRecordCurrent(prevCache, cacheKey, COMP_TYPE_SHEET, componentNameForCache, fileStamp) Then
                                If mp_IsComponentPresentForType(componentNameForCache, COMP_TYPE_SHEET) Then
                                    mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_SHEET, componentNameForCache, fileStamp
                                    GoTo ContinueNextFile
                                End If
                            End If
                        End If

                        sourceText = mp_ReadAllText(importPath, sourceEncodingIsUtf8)
                        If mp_UpdateSheetModule(componentNameForCache, importPath, sourceText) Then
                            mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_SHEET, componentNameForCache, fileStamp
                        End If

                    Case COMP_TYPE_WORKBOOK
                        componentNameForCache = mp_FindWorkbookComponentName()
                        If Len(componentNameForCache) = 0 Then GoTo ContinueNextFile

                        If fastMode Then
                            If mp_IsCacheRecordCurrent(prevCache, cacheKey, COMP_TYPE_WORKBOOK, componentNameForCache, fileStamp) Then
                                If mp_IsComponentPresentForType(componentNameForCache, COMP_TYPE_WORKBOOK) Then
                                    mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_WORKBOOK, componentNameForCache, fileStamp
                                    GoTo ContinueNextFile
                                End If
                            End If
                        End If

                        sourceText = mp_ReadAllText(importPath, sourceEncodingIsUtf8)
                        If mp_UpdateWorkbookModuleFromText(componentNameForCache, sourceText) Then
                            mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_WORKBOOK, componentNameForCache, fileStamp
                        End If
                End Select
                On Error GoTo 0
            End If
        End If

ContinueNextFile:
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        mp_ImportFolderRecursive subFolder, depth + 1, failed, fastMode, prevCache, nextCache
    Next subFolder

    Exit Sub

EH_IMPORT_FILE:
    errText = CStr(Err.Number) & ": " & Err.Description
    failed = failed & vbCrLf & "- " & importPath & " (" & errText & ")"
    Err.Clear
    On Error GoTo 0
    GoTo ContinueNextFile
End Sub

Private Sub mp_EnsureValidComponentNameLength(ByVal componentName As String, ByVal importPath As String)
    If Len(componentName) <= MAX_VBA_COMPONENT_NAME_LEN Then Exit Sub
    Err.Raise vbObjectError + 1010, "mp_EnsureValidComponentNameLength", _
              "VBA component name '" & componentName & "' is too long (" & CStr(Len(componentName)) & _
              "). Maximum allowed is " & CStr(MAX_VBA_COMPONENT_NAME_LEN) & ". File: " & importPath
End Sub

Private Sub mp_ImportStandardModuleFromSource( _
    ByVal componentName As String, _
    ByVal importPath As String, _
    Optional ByVal preloadedSourceText As String = vbNullString _
)
    Dim vbComp As Object
    Dim cm As Object
    Dim sourceText As String
    Dim cleanCode As String

    If Len(Trim$(componentName)) = 0 Then
        Err.Raise vbObjectError + 1009, "mp_ImportStandardModuleFromSource", "Standard module name is empty for: " & importPath
    End If

    If Len(preloadedSourceText) > 0 Then
        sourceText = preloadedSourceText
    Else
        sourceText = mp_ReadAllText(importPath)
    End If
    cleanCode = mp_ExtractCodeBody(sourceText)

    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(1) ' vbext_ct_StdModule
    vbComp.Name = componentName
    Set cm = vbComp.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString cleanCode
End Sub

Private Sub mp_ImportClassModuleFromSource( _
    ByVal componentName As String, _
    ByVal importPath As String, _
    Optional ByVal preloadedSourceText As String = vbNullString _
)
    Dim vbComp As Object
    Dim cm As Object
    Dim sourceText As String
    Dim cleanCode As String

    If Len(Trim$(componentName)) = 0 Then
        Err.Raise vbObjectError + 1005, "mp_ImportClassModuleFromSource", "Class module name is empty for: " & importPath
    End If

    If Len(preloadedSourceText) > 0 Then
        sourceText = preloadedSourceText
    Else
        sourceText = mp_ReadAllText(importPath)
    End If
    cleanCode = mp_ExtractCodeBody(sourceText)

    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(2) ' vbext_ct_ClassModule
    vbComp.Name = componentName
    Set cm = vbComp.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString cleanCode
End Sub

Private Function mp_ExtractCodeBody(ByVal sourceText As String) As String
    Dim lines As Variant
    Dim i As Long
    Dim lineText As String
    Dim trimmed As String
    Dim outText As String

    sourceText = Replace(sourceText, vbCrLf, vbLf)
    sourceText = Replace(sourceText, vbCr, vbLf)
    lines = Split(sourceText, vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = CStr(lines(i))
        ' Strip BOM/non-printable prefix if present.
        lineText = Replace(lineText, ChrW$(65279), vbNullString)
        lineText = Replace(lineText, ChrW$(160), " ")
        trimmed = Trim$(lineText)

        If StrComp(Left$(trimmed, 8), "VERSION ", vbTextCompare) = 0 Then GoTo ContinueLine
        If StrComp(trimmed, "BEGIN", vbTextCompare) = 0 Then GoTo ContinueLine
        If StrComp(trimmed, "END", vbTextCompare) = 0 Then GoTo ContinueLine
        If StrComp(Left$(trimmed, 10), "Attribute ", vbTextCompare) = 0 Then GoTo ContinueLine
        ' Class metadata lines from exported .cls header are not valid VBA statements.
        If StrComp(Left$(trimmed, 10), "MultiUse =", vbTextCompare) = 0 Then GoTo ContinueLine
        If StrComp(Left$(trimmed, 13), "Persistable =", vbTextCompare) = 0 Then GoTo ContinueLine
        If StrComp(Left$(trimmed, 20), "DataBindingBehavior =", vbTextCompare) = 0 Then GoTo ContinueLine
        If StrComp(Left$(trimmed, 19), "DataSourceBehavior =", vbTextCompare) = 0 Then GoTo ContinueLine
        If StrComp(Left$(trimmed, 21), "MTSTransactionMode =", vbTextCompare) = 0 Then GoTo ContinueLine

        If Len(outText) > 0 Then outText = outText & vbCrLf
        outText = outText & lineText

ContinueLine:
    Next i

    mp_ExtractCodeBody = outText
End Function

Private Sub mp_RemoveComponentIfExists(ByVal componentName As String)
    Dim vbComp As Object

    If Len(componentName) = 0 Then Exit Sub

    Set vbComp = mp_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then Exit Sub

    ThisWorkbook.VBProject.VBComponents.Remove vbComp
End Sub

Private Function mp_TryGetComponentByName(ByVal componentName As String) As Object
    On Error Resume Next
    Set mp_TryGetComponentByName = ThisWorkbook.VBProject.VBComponents(componentName)
    On Error GoTo 0
End Function

Private Function mp_IsComponentPresentForType(ByVal componentName As String, ByVal compType As String) As Boolean
    Dim vbComp As Object

    Set vbComp = mp_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then Exit Function

    Select Case LCase$(compType)
        Case COMP_TYPE_MODULE
            mp_IsComponentPresentForType = (vbComp.Type = 1) ' vbext_ct_StdModule
        Case COMP_TYPE_CLASS
            mp_IsComponentPresentForType = (vbComp.Type = 2) ' vbext_ct_ClassModule
        Case COMP_TYPE_SHEET, COMP_TYPE_WORKBOOK
            mp_IsComponentPresentForType = (vbComp.Type = 100) ' vbext_ct_Document
    End Select
End Function

Private Function mp_GetComponentNameFromSource(ByVal importPath As String) As String
    Dim fileName As String
    Dim dotPos As Long
    Dim compType As String
    Dim sourceText As String
    Dim fallbackName As String

    fileName = Mid$(importPath, InStrRev(importPath, "\") + 1)
    If Not mp_TryResolveFileComponentType(fileName, compType, fallbackName) Then
        dotPos = InStrRev(fileName, ".")
        If dotPos > 1 Then
            fallbackName = Left$(fileName, dotPos - 1)
        Else
            fallbackName = fileName
        End If
    End If

    sourceText = mp_ReadAllText(importPath)
    mp_GetComponentNameFromSource = mp_GetComponentNameFromSourceText(sourceText, fallbackName)
End Function

Private Function mp_GetComponentNameFromSourceText(ByVal sourceText As String, ByVal fallbackName As String) As String
    Dim attrPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long

    mp_GetComponentNameFromSourceText = fallbackName

    attrPos = InStr(1, sourceText, "Attribute VB_Name", vbTextCompare)
    If attrPos = 0 Then Exit Function

    quoteStart = InStr(attrPos, sourceText, """")
    If quoteStart = 0 Then Exit Function

    quoteEnd = InStr(quoteStart + 1, sourceText, """")
    If quoteEnd <= quoteStart Then Exit Function

    mp_GetComponentNameFromSourceText = Mid$(sourceText, quoteStart + 1, quoteEnd - quoteStart - 1)
End Function

Private Function mp_EndsWith(ByVal value As String, ByVal suffix As String) As Boolean
    mp_EndsWith = (LCase$(Right$(value, Len(suffix))) = LCase$(suffix))
End Function

Private Function mp_IsDevToolsSourceFile(ByVal normalizedFileName As String) As Boolean
    mp_IsDevToolsSourceFile = _
        (normalizedFileName = "devtools.vba") Or _
        (normalizedFileName = "ex_devtools.vba")
End Function

Private Function mp_TryResolveFileComponentType( _
    ByVal fileName As String, _
    ByRef outCompType As String, _
    ByRef outFallbackName As String _
) As Boolean
    Dim normalizedName As String
    Dim baseName As String

    normalizedName = LCase$(Trim$(fileName))
    outCompType = vbNullString
    outFallbackName = vbNullString

    If mp_EndsWith(normalizedName, ".utf8.vba") Then
        baseName = Left$(fileName, Len(fileName) - Len(".utf8.vba"))
        normalizedName = LCase$(Trim$(baseName))
    ElseIf mp_EndsWith(normalizedName, ".vba") Then
        baseName = Left$(fileName, Len(fileName) - Len(".vba"))
        normalizedName = LCase$(Trim$(baseName))
    Else
        Exit Function
    End If

    If StrComp(normalizedName, "thisworkbook", vbTextCompare) = 0 Then
        outCompType = COMP_TYPE_WORKBOOK
        outFallbackName = "ThisWorkbook"
    ElseIf Left$(normalizedName, 3) = "ws_" Then
        outCompType = COMP_TYPE_SHEET
        outFallbackName = Mid$(baseName, 4)
    ElseIf Left$(normalizedName, 3) = "ex_" Then
        outCompType = COMP_TYPE_MODULE
        outFallbackName = baseName
    ElseIf Left$(normalizedName, 4) = "obj_" And mp_EndsWith(normalizedName, ".cls") Then
        outCompType = COMP_TYPE_CLASS
        outFallbackName = Left$(baseName, Len(baseName) - Len(".cls"))
    End If

    mp_TryResolveFileComponentType = (Len(Trim$(outCompType)) > 0 And Len(Trim$(outFallbackName)) > 0)
End Function

Private Function mp_HasUtf8MarkerBeforeVba(ByVal fileName As String) As Boolean
    mp_HasUtf8MarkerBeforeVba = mp_EndsWith(LCase$(Trim$(fileName)), ".utf8.vba")
End Function

Private Function mp_CreateDictionary() As Object
    Set mp_CreateDictionary = CreateObject("Scripting.Dictionary")
    mp_CreateDictionary.CompareMode = 1
End Function

Private Function mp_NormalizeCacheKey(ByVal filePath As String) As String
    mp_NormalizeCacheKey = LCase$(Replace$(CStr(filePath), "/", "\"))
End Function

Private Function mp_BuildFileStamp(ByVal filePath As String) As String
    Dim fso As Object
    Dim fileObj As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    Set fileObj = fso.GetFile(filePath)
    mp_BuildFileStamp = mp_BuildFileStampFromFileObject(fileObj)
End Function

Private Function mp_BuildFileStampFromFileObject(ByVal fileObj As Object) As String
    mp_BuildFileStampFromFileObject = CStr(CDbl(fileObj.DateLastModified)) & ":" & CStr(CLng(fileObj.Size))
End Function

Private Function mp_IsCacheRecordCurrent( _
    ByVal cache As Object, _
    ByVal cacheKey As String, _
    ByVal compType As String, _
    ByVal componentName As String, _
    ByVal fileStamp As String _
) As Boolean
    Dim rec As Object

    If cache Is Nothing Then Exit Function
    If Not cache.Exists(cacheKey) Then Exit Function
    Set rec = cache(cacheKey)
    If rec Is Nothing Then Exit Function

    If StrComp(CStr(rec("Type")), compType, vbTextCompare) <> 0 Then Exit Function
    If StrComp(CStr(rec("Name")), componentName, vbTextCompare) <> 0 Then Exit Function
    If StrComp(CStr(rec("Stamp")), fileStamp, vbBinaryCompare) <> 0 Then Exit Function

    mp_IsCacheRecordCurrent = True
End Function

Private Function mp_TryGetCachedComponentNameByStamp( _
    ByVal cache As Object, _
    ByVal cacheKey As String, _
    ByVal compType As String, _
    ByVal fileStamp As String, _
    ByRef outComponentName As String _
) As Boolean
    Dim rec As Object

    If cache Is Nothing Then Exit Function
    If Not cache.Exists(cacheKey) Then Exit Function

    Set rec = cache(cacheKey)
    If rec Is Nothing Then Exit Function
    If StrComp(CStr(rec("Type")), compType, vbTextCompare) <> 0 Then Exit Function
    If StrComp(CStr(rec("Stamp")), fileStamp, vbBinaryCompare) <> 0 Then Exit Function

    outComponentName = CStr(rec("Name"))
    mp_TryGetCachedComponentNameByStamp = (Len(outComponentName) > 0)
End Function

Private Sub mp_SetCacheRecord( _
    ByVal cache As Object, _
    ByVal cacheKey As String, _
    ByVal compType As String, _
    ByVal componentName As String, _
    ByVal fileStamp As String _
)
    Dim rec As Object

    If cache Is Nothing Then Exit Sub

    Set rec = mp_CreateDictionary()
    rec("Type") = compType
    rec("Name") = componentName
    rec("Stamp") = fileStamp

    If cache.Exists(cacheKey) Then
        cache.Remove cacheKey
    End If
    cache.Add cacheKey, rec
End Sub

Private Function mp_LoadImportCache(ByVal cachePath As String) As Object
    Dim cache As Object
    Dim lineText As String
    Dim parts() As String
    Dim f As Integer
    Dim stampText As String
    Dim i As Long

    Set cache = mp_CreateDictionary()
    If Len(Dir(cachePath)) = 0 Then
        Set mp_LoadImportCache = cache
        Exit Function
    End If

    f = FreeFile
    Open cachePath For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        If Len(Trim$(lineText)) = 0 Then GoTo ContinueLoop
        parts = Split(lineText, "|")
        If UBound(parts) < 3 Then GoTo ContinueLoop
        stampText = CStr(parts(3))
        If UBound(parts) > 3 Then
            For i = 4 To UBound(parts)
                stampText = stampText & "|" & CStr(parts(i))
            Next i
        End If
        mp_SetCacheRecord cache, CStr(parts(0)), CStr(parts(1)), CStr(parts(2)), stampText
ContinueLoop:
    Loop
    Close #f

    Set mp_LoadImportCache = cache
End Function

Private Sub mp_SaveImportCache(ByVal cachePath As String, ByVal cache As Object)
    Dim f As Integer
    Dim key As Variant
    Dim rec As Object

    If cache Is Nothing Then Exit Sub

    f = FreeFile
    Open cachePath For Output As #f
    For Each key In cache.Keys
        Set rec = cache(CStr(key))
        Print #f, CStr(key) & "|" & CStr(rec("Type")) & "|" & CStr(rec("Name")) & "|" & CStr(rec("Stamp"))
    Next key
    Close #f
End Sub

Private Sub mp_RemoveStaleImportedComponents(ByVal prevCache As Object, ByVal nextCache As Object)
    Dim key As Variant
    Dim rec As Object
    Dim compType As String
    Dim componentName As String

    If prevCache Is Nothing Then Exit Sub
    If nextCache Is Nothing Then Exit Sub

    For Each key In prevCache.Keys
        If Not nextCache.Exists(CStr(key)) Then
            Set rec = prevCache(CStr(key))
            If Not rec Is Nothing Then
                compType = CStr(rec("Type"))
                componentName = CStr(rec("Name"))
                If StrComp(componentName, "DevTools", vbTextCompare) <> 0 Then
                    If StrComp(compType, COMP_TYPE_MODULE, vbTextCompare) = 0 Or _
                       StrComp(compType, COMP_TYPE_CLASS, vbTextCompare) = 0 Then
                        mp_RemoveComponentIfExists componentName
                    End If
                End If
            End If
        End If
    Next key
End Sub

'==========================
' Sheet module refresh
'==========================
Private Function mp_UpdateSheetModule( _
    ByVal sheetName As String, _
    ByVal sheetCodePath As String, _
    Optional ByVal preloadedCodeText As String = vbNullString _
) As Boolean
    Dim vbProj As Object
    Dim vbComp As Object
    Dim cm As Object
    Dim codeText As String

    Set vbProj = ThisWorkbook.VBProject
    If Not mp_SheetModuleExists(vbProj, sheetName) Then Exit Function

    If Len(preloadedCodeText) > 0 Then
        codeText = preloadedCodeText
    Else
        If Len(mp_BuildFileStamp(sheetCodePath)) = 0 Then Exit Function
        codeText = mp_ReadAllText(sheetCodePath)
    End If

    Set vbComp = vbProj.VBComponents(sheetName)
    Set cm = vbComp.CodeModule

    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText
    mp_UpdateSheetModule = True
End Function

Private Function mp_ResolveSheetCodeName(ByVal fileStem As String) As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(fileStem)
    On Error GoTo 0

    If Not ws Is Nothing Then
        mp_ResolveSheetCodeName = ws.CodeName
    Else
        mp_ResolveSheetCodeName = fileStem
    End If
End Function

Private Function mp_SheetModuleExists(ByVal vbProj As Object, ByVal sheetName As String) As Boolean
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(sheetName)
    mp_SheetModuleExists = Not vbComp Is Nothing
    On Error GoTo 0
End Function

Private Function mp_FindWorkbookComponentName() As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim nameCandidates(1 To 4) As String
    Dim i As Long

    Set vbProj = ThisWorkbook.VBProject
    nameCandidates(1) = "wb_Host"
    nameCandidates(2) = "ThisWorkbook"
    nameCandidates(3) = "ЭтаКнига"
    nameCandidates(4) = "ЦяКнига"

    For i = LBound(nameCandidates) To UBound(nameCandidates)
        On Error Resume Next
        Set vbComp = vbProj.VBComponents(nameCandidates(i))
        On Error GoTo 0
        If Not vbComp Is Nothing Then
            mp_FindWorkbookComponentName = nameCandidates(i)
            Exit Function
        End If
    Next i
End Function

Private Function mp_UpdateWorkbookModuleFromText( _
    ByVal workbookComponentName As String, _
    ByVal codeText As String _
) As Boolean
    Dim vbProj As Object
    Dim vbComp As Object
    Dim cm As Object

    If Len(Trim$(workbookComponentName)) = 0 Then Exit Function
    If Len(codeText) = 0 Then Exit Function

    Set vbProj = ThisWorkbook.VBProject
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(workbookComponentName)
    On Error GoTo 0
    If vbComp Is Nothing Then Exit Function

    Set cm = vbComp.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText

    mp_UpdateWorkbookModuleFromText = True
End Function

Private Function mp_ReadAllText( _
    ByVal filePath As String, _
    Optional ByVal preferUtf8 As Boolean = False _
) As String
    If preferUtf8 Then
        On Error GoTo FallbackLegacy
        mp_ReadAllText = mp_ReadAllTextByCharset(filePath, "utf-8")
        If Left$(mp_ReadAllText, 1) = ChrW$(65279) Then
            mp_ReadAllText = Mid$(mp_ReadAllText, 2)
        End If
        Exit Function
    End If

    On Error GoTo FallbackLegacy
    mp_ReadAllText = mp_ReadAllTextByCharset(filePath, "utf-8")
    If Left$(mp_ReadAllText, 1) = ChrW$(65279) Then
        mp_ReadAllText = Mid$(mp_ReadAllText, 2)
    End If
    Exit Function

FallbackLegacy:
    Err.Clear
    mp_ReadAllText = mp_ReadAllTextLegacy(filePath)
End Function

Private Function mp_ReadAllTextByCharset(ByVal filePath As String, ByVal charsetName As String) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Mode = 3 ' adModeReadWrite
    stream.Charset = charsetName
    stream.Open
    stream.LoadFromFile filePath
    mp_ReadAllTextByCharset = stream.ReadText(-1)
    stream.Close
End Function

Private Function mp_ReadAllTextLegacy(ByVal filePath As String) As String
    Dim f As Integer
    Dim text As String

    f = FreeFile
    Open filePath For Input As #f
    text = Input$(LOF(f), f)
    Close #f

    mp_ReadAllTextLegacy = text
End Function
