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
    mp_ImportUserFormsFromFolder basePath & FORMS_DIR

    ' Refresh sheet modules from vba\sheets\*.bas (if sheet exists)
    mp_UpdateSheetModulesFromFolder basePath & SHEETS_DIR
    ' Refresh ThisWorkbook module if provided
    mp_UpdateWorkbookModule basePath & "ThisWorkbook.bas"

    Application.ScreenUpdating = True
    mp_ShowNotice "Code updated", 2
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Update Code failed: " & Err.Description, vbExclamation
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

'==========================
' Status bar notification
'==========================
Public Sub dev_ClearStatusBar()
    Application.StatusBar = False
End Sub

 

Private Sub mp_ShowNotice(ByVal msg As String, Optional ByVal seconds As Double = 2)
    Application.StatusBar = msg
    Application.OnTime Now + TimeSerial(0, 0, seconds), "dev_ClearStatusBar"
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
        On Error Resume Next
        prj.VBComponents.Remove prj.VBComponents(names(i))
        On Error GoTo 0
    Next i
End Sub

Private Sub mp_ImportFolder(ByVal folderPath As String)
    Dim fileName As String
    Dim importPath As String
    Dim failed As String
    Dim errText As String

    If Dir(folderPath, vbDirectory) = "" Then Exit Sub

    fileName = Dir(folderPath & "*.*")
    Do While fileName <> ""
        If mp_EndsWith(fileName, ".bas") _
        Or mp_EndsWith(fileName, ".cls") _
        Or mp_EndsWith(fileName, ".frm") Then
            If LCase$(fileName) <> "devtools.bas" Then
                importPath = folderPath & fileName
                On Error Resume Next
                ThisWorkbook.VBProject.VBComponents.Import importPath
                If Err.Number <> 0 Then
                    errText = CStr(Err.Number) & ": " & Err.Description
                    failed = failed & vbCrLf & "- " & importPath & " (" & errText & ")"
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If

        fileName = Dir()
    Loop

    If Len(failed) > 0 Then
        Err.Raise vbObjectError + 1001, "mp_ImportFolder", "Import failed for file(s):" & failed
    End If
End Sub

Private Function mp_EndsWith(ByVal value As String, ByVal suffix As String) As Boolean
    mp_EndsWith = (LCase$(Right$(value, Len(suffix))) = LCase$(suffix))
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
