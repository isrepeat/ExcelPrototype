' Must be pasted in internal .xlsm module
Option Explicit

Private Const BASE_DIR As String = "vba\\"
Private Const MODULES_DIR As String = "modules\\"
Private Const CLASSES_DIR As String = "classes\\"
Private Const FORMS_DIR As String = "forms\\"
Private Const SHEETS_DIR As String = "sheets\\"

' Main updater (legacy name preserved).
Public Sub dev_UpdateCode()
    Dim basePath As String

    basePath = ThisWorkbook.Path & "\\" & BASE_DIR
    If Len(Dir(basePath, vbDirectory)) = 0 Then
        MsgBox "Workbook path is empty or vba folder not found. Save the file first.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo EH

    mp_RemoveImportedModules

    mp_ImportFolder basePath & MODULES_DIR
    mp_ImportFolder basePath & CLASSES_DIR
    mp_ValidateClassImports basePath & CLASSES_DIR
    mp_ImportUserFormsFromFolder basePath & FORMS_DIR

    ' Refresh sheet modules from vba\sheets\*.bas (if sheet exists)
    mp_UpdateSheetModulesFromFolder basePath & SHEETS_DIR
    ' Refresh ThisWorkbook module if provided
    mp_UpdateWorkbookModule basePath & "ThisWorkbook.bas"

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

Private Sub mp_ImportUserFormsFromFolder(ByVal formsPath As String)
    Dim fileName As String
    Dim formName As String
    Dim vbComp As Object
    Dim cm As Object
    Dim codeText As String
    Dim failed As String

    If Dir(formsPath, vbDirectory) = "" Then Exit Sub

    fileName = Dir(formsPath & "*.bas")
    Do While fileName <> ""
        formName = Left$(fileName, Len(fileName) - 4)
        codeText = mp_ReadAllText(formsPath & fileName)

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
            failed = failed & vbCrLf & "- " & formsPath & fileName & " (" & CStr(Err.Number) & ": " & Err.Description & ")"
            Err.Clear
        End If
        On Error GoTo 0

        fileName = Dir()
    Loop

    If Len(failed) > 0 Then
        Err.Raise vbObjectError + 1003, "mp_ImportUserFormsFromFolder", "Form import failed:" & failed
    End If
End Sub

' Ribbon hook (keeps existing button working if mapped).
Public Sub dev_OnUpdateCodeClicked(ByVal control As Object)
    dev_UpdateCode
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

Private Sub mp_ImportFolder(ByVal folderPath As String)
    Dim fso As Object
    Dim rootFolder As Object
    Dim failed As String

    If Dir(folderPath, vbDirectory) = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(folderPath)
    mp_ImportFolderRecursive rootFolder, failed

    If Len(failed) > 0 Then
        Err.Raise vbObjectError + 1001, "mp_ImportFolder", "Import failed for file(s):" & failed
    End If
End Sub

Private Sub mp_ImportFolderRecursive(ByVal folderObj As Object, ByRef failed As String)
    Dim fileObj As Object
    Dim subFolder As Object
    Dim importPath As String
    Dim fileName As String
    Dim fileExt As String
    Dim componentName As String
    Dim errText As String
    Dim importedComp As Object

    For Each fileObj In folderObj.Files
        fileName = LCase$(CStr(fileObj.Name))
        fileExt = mp_GetFileExtension(fileName)
        If fileExt = ".bas" Or fileExt = ".cls" Or fileExt = ".frm" Then
            If fileName <> "devtools.bas" Then
                importPath = CStr(fileObj.Path)
                componentName = mp_GetComponentNameFromSource(importPath)
                On Error GoTo EH_IMPORT_FILE
                mp_RemoveComponentIfExists componentName
                Set importedComp = ThisWorkbook.VBProject.VBComponents.Import(importPath)
                If fileExt = ".cls" Then
                    If importedComp Is Nothing Or importedComp.Type <> 2 Then ' vbext_ct_ClassModule
                        mp_RemoveComponentIfExists componentName
                        mp_ImportClassModuleFromSource componentName, importPath
                    End If
                End If
                On Error GoTo 0
            End If
        End If

ContinueNextFile:
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        mp_ImportFolderRecursive subFolder, failed
    Next subFolder

    Exit Sub

EH_IMPORT_FILE:
    errText = CStr(Err.Number) & ": " & Err.Description
    failed = failed & vbCrLf & "- " & importPath & " (" & errText & ")"
    Err.Clear
    On Error GoTo 0
    GoTo ContinueNextFile
End Sub

Private Sub mp_ImportClassModuleFromSource(ByVal componentName As String, ByVal importPath As String)
    Dim vbComp As Object
    Dim cm As Object
    Dim sourceText As String
    Dim cleanCode As String

    If Len(Trim$(componentName)) = 0 Then
        Err.Raise vbObjectError + 1005, "mp_ImportClassModuleFromSource", "Class module name is empty for: " & importPath
    End If

    sourceText = mp_ReadAllText(importPath)
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

Private Function mp_GetComponentNameFromSource(ByVal importPath As String) As String
    Dim fileName As String
    Dim dotPos As Long
    Dim sourceText As String
    Dim attrPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long

    fileName = Mid$(importPath, InStrRev(importPath, "\") + 1)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 1 Then
        mp_GetComponentNameFromSource = Left$(fileName, dotPos - 1)
    Else
        mp_GetComponentNameFromSource = fileName
    End If

    sourceText = mp_ReadAllText(importPath)
    attrPos = InStr(1, sourceText, "Attribute VB_Name", vbTextCompare)
    If attrPos = 0 Then Exit Function

    quoteStart = InStr(attrPos, sourceText, """")
    If quoteStart = 0 Then Exit Function

    quoteEnd = InStr(quoteStart + 1, sourceText, """")
    If quoteEnd <= quoteStart Then Exit Function

    mp_GetComponentNameFromSource = Mid$(sourceText, quoteStart + 1, quoteEnd - quoteStart - 1)
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

'==========================
' Sheet module refresh
'==========================
Private Sub mp_UpdateSheetModule(ByVal sheetName As String, ByVal sheetCodePath As String)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim cm As Object
    Dim codeText As String

    If Dir(sheetCodePath) = vbNullString Then Exit Sub

    Set vbProj = ThisWorkbook.VBProject
    If Not mp_SheetModuleExists(vbProj, sheetName) Then Exit Sub

    codeText = mp_ReadAllText(sheetCodePath)

    Set vbComp = vbProj.VBComponents(sheetName)
    Set cm = vbComp.CodeModule

    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText
End Sub

Private Sub mp_UpdateSheetModulesFromFolder(ByVal sheetsPath As String)
    Dim fileName As String
    Dim fileStem As String
    Dim sheetCodeName As String

    If Dir(sheetsPath, vbDirectory) = "" Then Exit Sub

    fileName = Dir(sheetsPath & "*.bas")
    Do While fileName <> ""
        fileStem = Left$(fileName, Len(fileName) - 4)
        If StrComp(fileStem, "ThisWorkbook", vbTextCompare) <> 0 Then
            sheetCodeName = mp_ResolveSheetCodeName(fileStem)
            If Len(sheetCodeName) > 0 Then
                mp_UpdateSheetModule sheetCodeName, sheetsPath & fileName
            End If
        End If
        fileName = Dir()
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

Private Sub mp_UpdateWorkbookModule(ByVal workbookCodePath As String)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim cm As Object
    Dim codeText As String
    Dim nameCandidates(1 To 4) As String
    Dim i As Long

    If Dir(workbookCodePath) = vbNullString Then Exit Sub

    codeText = mp_ReadAllText(workbookCodePath)

    Set vbProj = ThisWorkbook.VBProject
    nameCandidates(1) = "wb_Host"
    nameCandidates(2) = "ThisWorkbook"
    nameCandidates(3) = "ЭтаКнига"
    nameCandidates(4) = "ЦяКнига"

    Set vbComp = Nothing
    For i = LBound(nameCandidates) To UBound(nameCandidates)
        On Error Resume Next
        Set vbComp = vbProj.VBComponents(nameCandidates(i))
        On Error GoTo 0
        If Not vbComp Is Nothing Then Exit For
    Next i
    If vbComp Is Nothing Then Exit Sub

    Set cm = vbComp.CodeModule

    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText
End Sub

Private Function mp_ReadAllText(ByVal filePath As String) As String
    Dim f As Integer
    Dim text As String

    f = FreeFile
    Open filePath For Input As #f
    text = Input$(LOF(f), f)
    Close #f

    mp_ReadAllText = text
End Function
