' Must be pasted in internal .xlsm module
Option Explicit

Private Const BASE_DIR As String = "vba\\"
Private Const MODULES_DIR As String = "modules\\"
Private Const CLASSES_DIR As String = "classes\\"
Private Const FORMS_DIR As String = "forms\\"
Private Const SHEETS_DIR As String = "sheets\\"
Private Const IMPORT_CACHE_FILE As String = ".devtools_import_cache.txt"
Private Const ENABLE_CLASS_IMPORT_VALIDATION As Boolean = False
Private Const COMP_TYPE_MODULE As String = "module"
Private Const COMP_TYPE_CLASS As String = "class"
Private Const COMP_TYPE_FORM As String = "form"
Private Const COMP_TYPE_SHEET As String = "sheet"
Private Const COMP_TYPE_WORKBOOK As String = "workbook"
Private Const UTF8_BAS_FILENAME_SUFFIX As String = ".utf8.bas"
' True  -> legacy fast import for .bas via VBComponents.Import.
' False -> current UTF-safe import path via AddFromString.
Private Const USE_LEGACY_FAST_BAS_IMPORT As Boolean = True

' Main updater (legacy name preserved).
Public Sub dev_UpdateCode()
    mp_UpdateCodeCore False
End Sub

Public Sub dev_UpdateCodeFast()
    mp_UpdateCodeCore True
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

    mp_ImportFolder basePath & MODULES_DIR, fastMode, prevCache, nextCache
    mp_ImportFolder basePath & CLASSES_DIR, fastMode, prevCache, nextCache
    If ENABLE_CLASS_IMPORT_VALIDATION Then
        mp_ValidateClassImports basePath & CLASSES_DIR
    End If
    mp_ImportUserFormsFromFolder basePath & FORMS_DIR, fastMode, prevCache, nextCache

    ' Refresh sheet modules from vba\sheets\*.bas (if sheet exists)
    mp_UpdateSheetModulesFromFolder basePath & SHEETS_DIR, fastMode, prevCache, nextCache
    ' Refresh ThisWorkbook module if provided
    mp_UpdateWorkbookModule basePath & "ThisWorkbook.bas", fastMode, prevCache, nextCache

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

Private Sub mp_ValidateClassImports(ByVal classesPath As String)
    Dim fileName As String
    Dim className As String
    Dim vbComp As Object
    Dim failed As String

    If Dir(classesPath, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1006, "mp_ValidateClassImports", "Classes folder not found: " & classesPath
    End If

    fileName = Dir(classesPath & "*.cls")
    Do While fileName <> ""
        className = Left$(fileName, Len(fileName) - 4)
        Set vbComp = Nothing
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(className)
        On Error GoTo 0

        If vbComp Is Nothing Then
            failed = failed & vbCrLf & "- missing class: " & className
        ElseIf vbComp.Type <> 2 Then ' vbext_ct_ClassModule
            failed = failed & vbCrLf & "- wrong component type for class '" & className & "': " & CStr(vbComp.Type)
        End If

        fileName = Dir()
    Loop

    If Len(failed) > 0 Then
        Err.Raise vbObjectError + 1007, "mp_ValidateClassImports", "Class import validation failed:" & failed
    End If
End Sub

Private Sub mp_ShowCodeUpdatedNotice()
    On Error GoTo ShowMsgBox
    Application.Run "ex_Messaging.m_ShowNotice", "Code updated", 2
    Exit Sub

ShowMsgBox:
    MsgBox "Code updated", vbInformation
End Sub

Private Sub mp_ImportUserFormsFromFolder( _
    ByVal formsPath As String, _
    ByVal fastMode As Boolean, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object _
)
    Dim fileName As String
    Dim formName As String
    Dim vbComp As Object
    Dim cm As Object
    Dim codeText As String
    Dim failed As String
    Dim importPath As String
    Dim fileStamp As String
    Dim cacheKey As String

    If Dir(formsPath, vbDirectory) = "" Then Exit Sub

    fileName = Dir(formsPath & "*.bas")
    Do While fileName <> ""
        importPath = formsPath & fileName
        formName = Left$(fileName, Len(fileName) - 4)
        cacheKey = mp_NormalizeCacheKey(importPath)
        fileStamp = mp_BuildFileStamp(importPath)

        If fastMode Then
            If mp_IsCacheRecordCurrent(prevCache, cacheKey, COMP_TYPE_FORM, formName, fileStamp) Then
                If mp_IsComponentPresentForType(formName, COMP_TYPE_FORM) Then
                    mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_FORM, formName, fileStamp
                    fileName = Dir()
                    GoTo ContinueLoop
                End If
            End If
        End If

        codeText = mp_ReadAllText(importPath)

        Set vbComp = Nothing
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(formName)
        On Error GoTo 0

        If vbComp Is Nothing Then
            Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3) ' vbext_ct_MSForm
            vbComp.Name = formName
            If StrComp(vbComp.Name, formName, vbTextCompare) <> 0 Then
                failed = failed & vbCrLf & "- failed to name form '" & formName & "', actual '" & vbComp.Name & "'"
            End If
        End If

        On Error Resume Next
        Set cm = vbComp.CodeModule
        cm.DeleteLines 1, cm.CountOfLines
        cm.AddFromString codeText
        If Err.Number <> 0 Then
            failed = failed & vbCrLf & "- " & importPath & " (" & CStr(Err.Number) & ": " & Err.Description & ")"
            Err.Clear
        Else
            mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_FORM, formName, fileStamp
        End If
        On Error GoTo 0

        fileName = Dir()
ContinueLoop:
    Loop

    If Len(failed) > 0 Then
        Err.Raise vbObjectError + 1003, "mp_ImportUserFormsFromFolder", "Form import failed:" & failed
    End If
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
    mp_ImportFolderRecursive rootFolder, failed, fastMode, prevCache, nextCache

    If Len(failed) > 0 Then
        Err.Raise vbObjectError + 1001, "mp_ImportFolder", "Import failed for file(s):" & failed
    End If
End Sub

Private Sub mp_ImportFolderRecursive( _
    ByVal folderObj As Object, _
    ByRef failed As String, _
    ByVal fastMode As Boolean, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object _
)
    Dim fileObj As Object
    Dim subFolder As Object
    Dim importPath As String
    Dim fileName As String
    Dim fileExt As String
    Dim componentName As String
    Dim errText As String
    Dim importedComp As Object
    Dim sourceText As String
    Dim fileStamp As String
    Dim importStamp As String
    Dim cacheKey As String
    Dim compType As String
    Dim fallbackName As String

    For Each fileObj In folderObj.Files
        fileName = LCase$(CStr(fileObj.Name))
        fileExt = mp_GetFileExtension(fileName)
        If fileExt = ".bas" Or fileExt = ".cls" Or fileExt = ".frm" Then
            If fileName <> "devtools.bas" Then
                importPath = CStr(fileObj.Path)
                On Error GoTo EH_IMPORT_FILE

                fileStamp = mp_BuildFileStampFromFileObject(fileObj)
                importStamp = mp_BuildImportStampByFileType(fileExt, fileStamp, importPath)
                cacheKey = mp_NormalizeCacheKey(importPath)
                fallbackName = mp_GetFileStem(CStr(fileObj.Name))
                sourceText = vbNullString

                If fileExt = ".bas" Then
                    compType = COMP_TYPE_MODULE
                ElseIf fileExt = ".cls" Then
                    compType = COMP_TYPE_CLASS
                Else
                    compType = COMP_TYPE_FORM
                End If

                If fastMode Then
                    If mp_TryGetCachedComponentNameByStamp(prevCache, cacheKey, compType, importStamp, componentName) Then
                        If mp_IsComponentPresentForType(componentName, compType) Then
                            mp_SetCacheRecord nextCache, cacheKey, compType, componentName, importStamp
                            GoTo ContinueNextFile
                        End If
                    End If
                End If

                If fileExt = ".bas" Or fileExt = ".cls" Then
                    sourceText = mp_ReadAllText(importPath)
                    componentName = mp_GetComponentNameFromSourceText(sourceText, fallbackName)
                Else
                    componentName = fallbackName
                End If

                mp_RemoveComponentIfExists componentName
                Set importedComp = Nothing
                Select Case fileExt
                    Case ".bas"
                        If mp_ShouldUseLegacyFastImportForBas(importPath) Then
                            Set importedComp = ThisWorkbook.VBProject.VBComponents.Import(importPath)
                            If importedComp Is Nothing Or importedComp.Type <> 1 Then ' vbext_ct_StdModule
                                mp_RemoveComponentIfExists componentName
                                mp_ImportStandardModuleFromSource componentName, importPath, sourceText
                            End If
                        Else
                            mp_ImportStandardModuleFromSource componentName, importPath, sourceText
                        End If
                    Case ".cls"
                        Set importedComp = ThisWorkbook.VBProject.VBComponents.Import(importPath)
                        If importedComp Is Nothing Or importedComp.Type <> 2 Then ' vbext_ct_ClassModule
                            mp_RemoveComponentIfExists componentName
                            mp_ImportClassModuleFromSource componentName, importPath, sourceText
                        End If
                    Case Else
                        Set importedComp = ThisWorkbook.VBProject.VBComponents.Import(importPath)
                End Select
                mp_SetCacheRecord nextCache, cacheKey, compType, componentName, importStamp
                On Error GoTo 0
            End If
        End If

ContinueNextFile:
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        mp_ImportFolderRecursive subFolder, failed, fastMode, prevCache, nextCache
    Next subFolder

    Exit Sub

EH_IMPORT_FILE:
    errText = CStr(Err.Number) & ": " & Err.Description
    failed = failed & vbCrLf & "- " & importPath & " (" & errText & ")"
    Err.Clear
    On Error GoTo 0
    GoTo ContinueNextFile
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
        Case COMP_TYPE_FORM
            mp_IsComponentPresentForType = (vbComp.Type = 3) ' vbext_ct_MSForm
        Case COMP_TYPE_SHEET, COMP_TYPE_WORKBOOK
            mp_IsComponentPresentForType = (vbComp.Type = 100) ' vbext_ct_Document
    End Select
End Function

Private Function mp_GetComponentNameFromSource(ByVal importPath As String) As String
    Dim fileName As String
    Dim dotPos As Long
    Dim sourceText As String
    Dim fallbackName As String

    fileName = Mid$(importPath, InStrRev(importPath, "\") + 1)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 1 Then
        fallbackName = Left$(fileName, dotPos - 1)
    Else
        fallbackName = fileName
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

Private Function mp_GetFileExtension(ByVal fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos <= 0 Then Exit Function
    mp_GetFileExtension = LCase$(Mid$(fileName, dotPos))
End Function

Private Function mp_GetFileStem(ByVal fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos <= 0 Then
        mp_GetFileStem = fileName
    Else
        mp_GetFileStem = Left$(fileName, dotPos - 1)
    End If
End Function

Private Function mp_BuildImportStampByFileType( _
    ByVal fileExt As String, _
    ByVal fileStamp As String, _
    Optional ByVal importPath As String = vbNullString _
) As String
    mp_BuildImportStampByFileType = fileStamp

    If StrComp(fileExt, ".bas", vbTextCompare) = 0 Then
        mp_BuildImportStampByFileType = fileStamp & "|basMode=" & mp_GetBasImportModeToken(importPath)
    End If
End Function

Private Function mp_GetBasImportModeToken(ByVal importPath As String) As String
    If mp_ShouldUseLegacyFastImportForBas(importPath) Then
        mp_GetBasImportModeToken = "legacyFastImport"
    ElseIf USE_LEGACY_FAST_BAS_IMPORT Then
        mp_GetBasImportModeToken = "utfSafeAddFromString_utf8Marker"
    Else
        mp_GetBasImportModeToken = "utfSafeAddFromString"
    End If
End Function

Private Function mp_ShouldUseLegacyFastImportForBas(ByVal importPath As String) As Boolean
    mp_ShouldUseLegacyFastImportForBas = USE_LEGACY_FAST_BAS_IMPORT
    If Not mp_ShouldUseLegacyFastImportForBas Then Exit Function

    If mp_IsUtf8MarkedBasFile(importPath) Then
        mp_ShouldUseLegacyFastImportForBas = False
    End If
End Function

Private Function mp_IsUtf8MarkedBasFile(ByVal importPath As String) As Boolean
    Dim normalizedPath As String

    normalizedPath = LCase$(Replace$(Trim$(CStr(importPath)), "/", "\"))
    If Len(normalizedPath) < Len(UTF8_BAS_FILENAME_SUFFIX) Then Exit Function

    mp_IsUtf8MarkedBasFile = _
        (Right$(normalizedPath, Len(UTF8_BAS_FILENAME_SUFFIX)) = UTF8_BAS_FILENAME_SUFFIX)
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
                       StrComp(compType, COMP_TYPE_CLASS, vbTextCompare) = 0 Or _
                       StrComp(compType, COMP_TYPE_FORM, vbTextCompare) = 0 Then
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

Private Sub mp_UpdateSheetModulesFromFolder( _
    ByVal sheetsPath As String, _
    ByVal fastMode As Boolean, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object _
)
    Dim fileName As String
    Dim fileStem As String
    Dim sheetCodeName As String
    Dim importPath As String
    Dim fileStamp As String
    Dim cacheKey As String
    Dim codeText As String

    If Dir(sheetsPath, vbDirectory) = "" Then Exit Sub

    fileName = Dir(sheetsPath & "*.bas")
    Do While fileName <> ""
        fileStem = Left$(fileName, Len(fileName) - 4)
        If StrComp(fileStem, "ThisWorkbook", vbTextCompare) <> 0 Then
            sheetCodeName = mp_ResolveSheetCodeName(fileStem)
            If Len(sheetCodeName) > 0 Then
                importPath = sheetsPath & fileName
                cacheKey = mp_NormalizeCacheKey(importPath)
                fileStamp = mp_BuildFileStamp(importPath)

                If fastMode Then
                    If mp_IsCacheRecordCurrent(prevCache, cacheKey, COMP_TYPE_SHEET, sheetCodeName, fileStamp) Then
                        If mp_IsComponentPresentForType(sheetCodeName, COMP_TYPE_SHEET) Then
                            mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_SHEET, sheetCodeName, fileStamp
                            fileName = Dir()
                            GoTo ContinueLoop
                        End If
                    End If
                End If

                codeText = mp_ReadAllText(importPath)
                If mp_UpdateSheetModule(sheetCodeName, importPath, codeText) Then
                    mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_SHEET, sheetCodeName, fileStamp
                End If
            End If
        End If
        fileName = Dir()
ContinueLoop:
    Loop
End Sub

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

Private Sub mp_UpdateWorkbookModule( _
    ByVal workbookCodePath As String, _
    ByVal fastMode As Boolean, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object _
)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim cm As Object
    Dim codeText As String
    Dim fileStamp As String
    Dim cacheKey As String
    Dim componentName As String

    If Dir(workbookCodePath) = vbNullString Then Exit Sub

    fileStamp = mp_BuildFileStamp(workbookCodePath)
    cacheKey = mp_NormalizeCacheKey(workbookCodePath)
    componentName = mp_FindWorkbookComponentName()
    If Len(componentName) = 0 Then Exit Sub

    If fastMode Then
        If mp_IsCacheRecordCurrent(prevCache, cacheKey, COMP_TYPE_WORKBOOK, componentName, fileStamp) Then
            If mp_IsComponentPresentForType(componentName, COMP_TYPE_WORKBOOK) Then
                mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_WORKBOOK, componentName, fileStamp
                Exit Sub
            End If
        End If
    End If

    codeText = mp_ReadAllText(workbookCodePath)
    Set vbProj = ThisWorkbook.VBProject
    Set vbComp = vbProj.VBComponents(componentName)

    Set cm = vbComp.CodeModule

    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText
    mp_SetCacheRecord nextCache, cacheKey, COMP_TYPE_WORKBOOK, componentName, fileStamp
End Sub

Private Function mp_ReadAllText(ByVal filePath As String) As String
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
