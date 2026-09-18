Sub fn_RecalculateWorkbook()
    Application.Calculate
End Sub

Sub fn_RecalculateActiveSheet()
    ActiveSheet.Calculate
End Sub


Public Sub fn_TestHotkey()
    ActiveCell.Interior.Color = RGB(255, 0, 0)
End Sub

Public Sub fn_ToggleFirstTwoRows()
    On Error GoTo ExitPoint

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    With ActiveWindow
        .FreezePanes = False
        .SplitRow = 0
        .SplitColumn = 0
    End With

    If ws.Rows("1:2").Hidden Then
        ws.Rows("1:2").Hidden = False
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ws.Range("A3").Select
        ActiveWindow.FreezePanes = True

    Else
        ws.Rows("1:2").Hidden = True
        ActiveWindow.ScrollRow = 3
        ActiveWindow.ScrollColumn = 1
    End If

ExitPoint:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Sub fn_DatePlusOne()
    If IsDate(ActiveCell.value) Then
        ActiveCell.value = CDate(ActiveCell.value) + 1
    End If
End Sub

Sub fn_DateMinusOne()
    If IsDate(ActiveCell.value) Then
        ActiveCell.value = CDate(ActiveCell.value) - 1
    End If
End Sub


Public Sub fn_ReloadActiveWorkbookVba()
    Const VBEXT_CT_STD_MODULE As Long = 1
    Const VBEXT_CT_CLASS_MODULE As Long = 2
    Const VBEXT_CT_MS_FORM As Long = 3

    Dim targetWorkbook As Workbook
    Dim vbProject As Object
    Dim vbComponent As Object
    Dim componentNames As Collection
    Dim importFiles As Collection
    Dim documentImportFiles As Collection
    Dim vbaFolderPath As String
    Dim targetWorkbookName As String
    Dim importedCount As Long
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim applicationStateChanged As Boolean
    Dim componentName As Variant
    Dim importFile As Variant

    On Error GoTo EH

    If Application.ActiveWorkbook Is Nothing Then
        VBA.MsgBox "There is no active workbook whose VBA modules can be reloaded.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    Set targetWorkbook = Application.ActiveWorkbook
    targetWorkbookName = targetWorkbook.Name
    If targetWorkbook Is ThisWorkbook Then
        VBA.MsgBox _
            "The workbook containing the shortcut handler cannot reload itself. " & _
            "Activate the target workbook and press Ctrl+Alt+R again.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    If VBA.Len(targetWorkbook.Path) = 0 Then
        VBA.MsgBox _
            "Save the active workbook first. The vba folder must be located " & _
            "next to the workbook file.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    If Not private_TryResolveVbaFolder(targetWorkbook, vbaFolderPath) Then Exit Sub

    Set importFiles = New Collection
    Set documentImportFiles = New Collection
    private_CollectVbaFiles vbaFolderPath, importFiles, documentImportFiles
    If importFiles.Count = 0 And documentImportFiles.Count = 0 Then
        VBA.MsgBox _
            "No .bas, .cls, .frm, or .vba files were found in: " & vbaFolderPath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    Set vbProject = targetWorkbook.VBProject
    Set componentNames = New Collection
    For Each vbComponent In vbProject.VBComponents
        Select Case CLng(vbComponent.Type)
            Case VBEXT_CT_STD_MODULE, VBEXT_CT_CLASS_MODULE, VBEXT_CT_MS_FORM
                componentNames.Add VBA.CStr(vbComponent.Name)
        End Select
    Next vbComponent

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    applicationStateChanged = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Reloading VBA modules..."

    For Each componentName In componentNames
        vbProject.VBComponents.Remove vbProject.VBComponents(VBA.CStr(componentName))
    Next componentName

    For Each importFile In importFiles
        private_ImportVbaFile vbProject, VBA.CStr(importFile)
        importedCount = importedCount + 1
    Next importFile

    ' Документные модули нельзя импортировать как обычные: Excel создаст
    ' отдельный Module и события Workbook/Worksheet не будут вызываться.
    For Each importFile In documentImportFiles
        private_ImportDocumentVbaFile targetWorkbook, VBA.CStr(importFile)
        importedCount = importedCount + 1
    Next importFile
    private_InitializeReloadedWorkbook targetWorkbook

    Application.StatusBar = "Imported VBA modules: " & _
        VBA.CStr(importedCount) & "; imported at: " & _
        VBA.Format$(VBA.Now, "dd.mm.yyyy HH:nn:ss")

CleanExit:
    If applicationStateChanged Then
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating
    End If
    Exit Sub

EH:
    Application.StatusBar = False
    VBA.MsgBox "Failed to reload VBA modules in workbook '" & _
        targetWorkbookName & _
        "': [" & VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description & VBA.vbCrLf & _
        "Make sure 'Trust access to the VBA project object model' is enabled.", _
        VBA.vbCritical, "Reload VBA"
    Resume CleanExit
End Sub

' Поддерживает книги с runtime-регистрацией событий: после hot reload
' Workbook_Open не выполняется, поэтому bootstrap вызывается явно.
Private Sub private_InitializeReloadedWorkbook( _
    ByVal targetWorkbook As Workbook _
)
    Dim bootstrapComponent As Object
    Dim macroReference As String

    On Error Resume Next
    Set bootstrapComponent = targetWorkbook.VBProject.VBComponents( _
        "ex_DocumentGenerationBootstrap")
    On Error GoTo EH
    If bootstrapComponent Is Nothing Then Exit Sub

    macroReference = "'" & VBA.Replace$(targetWorkbook.Name, "'", "''") & _
        "'!ex_DocumentGenerationBootstrap.fn_Initialize"
    Application.Run macroReference
    Exit Sub
EH:
    VBA.Err.Raise VBA.Err.Number, "private_InitializeReloadedWorkbook", _
        "Failed to initialize reloaded workbook '" & targetWorkbook.Name & _
        "': " & VBA.Err.Description
End Sub

Private Function private_TryResolveVbaFolder( _
    ByVal targetWorkbook As Workbook, _
    ByRef outVbaFolderPath As String _
) As Boolean
    Dim fileSystem As Object
    Dim candidatePaths As Collection
    Dim workbookFolderPath As String
    Dim workspaceRootPath As String
    Dim candidatePath As Variant

    outVbaFolderPath = VBA.vbNullString
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set candidatePaths = New Collection
    workbookFolderPath = targetWorkbook.Path

    candidatePath = workbookFolderPath & Application.PathSeparator & "vba"
    If private_FolderExists(VBA.CStr(candidatePath)) Then _
        candidatePaths.Add VBA.CStr(candidatePath)

    ' Книга DocumentsGeneration хранится в MacrosExcel/2. DocumentsGeneration,
    ' а редактируемые исходники — в PROJECTS/2. DocumentsGeneration/vba.
    If VBA.StrComp(fileSystem.GetFileName(workbookFolderPath), _
            "2. DocumentsGeneration", VBA.vbTextCompare) = 0 And _
       VBA.StrComp(fileSystem.GetFileName( _
            fileSystem.GetParentFolderName(workbookFolderPath)), _
            "MacrosExcel", VBA.vbTextCompare) = 0 Then
        workspaceRootPath = fileSystem.GetParentFolderName( _
            fileSystem.GetParentFolderName(workbookFolderPath))
        candidatePath = workspaceRootPath & Application.PathSeparator & _
            "PROJECTS" & Application.PathSeparator & _
            "2. DocumentsGeneration" & Application.PathSeparator & "vba"
        If private_FolderExists(VBA.CStr(candidatePath)) Then _
            candidatePaths.Add VBA.CStr(candidatePath)
    End If

    If candidatePaths.Count = 0 Then
        VBA.MsgBox "The VBA modules folder was not found beside the workbook " & _
            "or in PROJECTS\2. DocumentsGeneration\vba.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    If candidatePaths.Count > 1 Then
        VBA.MsgBox "More than one VBA source folder was found. Keep only one " & _
            "source location to avoid importing an unexpected version.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    outVbaFolderPath = VBA.CStr(candidatePaths.Item(1))
    private_TryResolveVbaFolder = True
End Function


Private Sub private_CollectVbaFiles( _
    ByVal folderPath As String, _
    ByRef outFiles As Collection, _
    ByRef outDocumentFiles As Collection _
)
    Dim fileSystem As Object
    Dim folder As Object
    Dim childFolder As Object
    Dim file As Object
    Dim lowerName As String

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)

    For Each file In folder.Files
        lowerName = VBA.LCase$(VBA.CStr(file.Name))
        If private_IsDocumentModuleSource(lowerName) Then
            outDocumentFiles.Add VBA.CStr(file.Path)
        ElseIf VBA.Right$(lowerName, 4) = ".bas" Or _
               VBA.Right$(lowerName, 4) = ".cls" Or _
               VBA.Right$(lowerName, 4) = ".frm" Or _
               VBA.Right$(lowerName, 4) = ".vba" Then
            outFiles.Add VBA.CStr(file.Path)
        End If
    Next file

    For Each childFolder In folder.SubFolders
        private_CollectVbaFiles VBA.CStr(childFolder.Path), outFiles, _
            outDocumentFiles
    Next childFolder
End Sub

Private Function private_IsDocumentModuleSource( _
    ByVal lowerFileName As String _
) As Boolean
    private_IsDocumentModuleSource = ( _
        VBA.StrComp(lowerFileName, "thisworkbook.vba", VBA.vbTextCompare) = 0 Or _
        (VBA.Left$(lowerFileName, 3) = "ws_" And _
         VBA.Right$(lowerFileName, 4) = ".vba"))
End Function


Private Sub private_ImportVbaFile( _
    ByVal vbProject As Object, _
    ByVal sourcePath As String _
)
    Dim lowerPath As String
    Dim componentType As Long
    Dim sourceText As String
    Dim componentName As String
    Dim vbComponent As Object

    On Error GoTo EH

    lowerPath = VBA.LCase$(sourcePath)
    If VBA.Right$(lowerPath, 4) <> ".vba" Then
        vbProject.VBComponents.Import sourcePath
        Exit Sub
    End If

    sourceText = private_ReadTextFile(sourcePath)
    componentName = private_GetComponentName(sourcePath, sourceText)
    componentType = private_GetVbaComponentType(lowerPath, sourceText)

    Set vbComponent = vbProject.VBComponents.Add(componentType)
    vbComponent.Name = componentName
    vbComponent.CodeModule.AddFromString private_RemoveExportMetadata(sourceText)
    Exit Sub

EH:
    VBA.Err.Raise VBA.Err.Number, "private_ImportVbaFile", _
        "Failed to import '" & sourcePath & "': " & VBA.Err.Description
End Sub

' Обновляет исходник в уже существующем модуле книги или листа.
Private Sub private_ImportDocumentVbaFile( _
    ByVal targetWorkbook As Workbook, _
    ByVal sourcePath As String _
)
    Dim fileSystem As Object
    Dim fileName As String
    Dim worksheetName As String
    Dim targetSheet As Worksheet
    Dim componentName As String
    Dim vbComponent As Object
    Dim sourceText As String

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    fileName = VBA.CStr(fileSystem.GetFileName(sourcePath))
    sourceText = private_ReadTextFile(sourcePath)
    If VBA.Len(sourceText) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 1201, _
            "private_ImportDocumentVbaFile", "Source is empty: " & sourcePath
    End If

    If VBA.StrComp(fileName, "ThisWorkbook.vba", VBA.vbTextCompare) = 0 Then
        componentName = targetWorkbook.CodeName
    Else
        worksheetName = VBA.Mid$(fileName, 4, _
            VBA.Len(fileName) - VBA.Len("ws_") - VBA.Len(".vba"))
        On Error Resume Next
        Set targetSheet = targetWorkbook.Worksheets(worksheetName)
        On Error GoTo EH
        If targetSheet Is Nothing Then
            VBA.Err.Raise VBA.vbObjectError + 1202, _
                "private_ImportDocumentVbaFile", _
                "Worksheet '" & worksheetName & "' was not found for " & sourcePath
        End If
        componentName = targetSheet.CodeName
    End If

    Set vbComponent = targetWorkbook.VBProject.VBComponents(componentName)
    With vbComponent.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString private_RemoveExportMetadata(sourceText)
    End With
    Exit Sub
EH:
    VBA.Err.Raise VBA.Err.Number, "private_ImportDocumentVbaFile", _
        "Failed to import document module '" & sourcePath & "': " & _
        VBA.Err.Description
End Sub


Private Function private_GetVbaComponentType( _
    ByVal lowerPath As String, _
    ByVal sourceText As String _
) As Long
    Const VBEXT_CT_STD_MODULE As Long = 1
    Const VBEXT_CT_CLASS_MODULE As Long = 2
    Const VBEXT_CT_MS_FORM As Long = 3

    If VBA.Right$(lowerPath, 8) = ".cls.vba" Or _
        VBA.InStr(1, sourceText, "VERSION 1.0 CLASS", VBA.vbTextCompare) > 0 Then
        private_GetVbaComponentType = VBEXT_CT_CLASS_MODULE
    ElseIf VBA.Right$(lowerPath, 8) = ".frm.vba" Or _
        VBA.InStr(1, sourceText, "BEGIN VB.Form", VBA.vbTextCompare) > 0 Then
        private_GetVbaComponentType = VBEXT_CT_MS_FORM
    Else
        private_GetVbaComponentType = VBEXT_CT_STD_MODULE
    End If
End Function


Private Function private_GetComponentName( _
    ByVal sourcePath As String, _
    ByVal sourceText As String _
) As String
    Dim attributePosition As Long
    Dim nameStart As Long
    Dim nameEnd As Long
    Dim fileSystem As Object
    Dim fileName As String

    attributePosition = VBA.InStr(1, sourceText, "Attribute VB_Name = """, VBA.vbTextCompare)
    If attributePosition > 0 Then
        nameStart = attributePosition + VBA.Len("Attribute VB_Name = """)
        nameEnd = VBA.InStr(nameStart, sourceText, """")
        If nameEnd > nameStart Then
            private_GetComponentName = VBA.Mid$(sourceText, nameStart, nameEnd - nameStart)
            Exit Function
        End If
    End If

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    fileName = fileSystem.GetBaseName(sourcePath)
    If VBA.Right$(VBA.LCase$(fileName), 4) = ".cls" Or _
        VBA.Right$(VBA.LCase$(fileName), 4) = ".bas" Or _
        VBA.Right$(VBA.LCase$(fileName), 4) = ".frm" Then
        fileName = VBA.Left$(fileName, VBA.Len(fileName) - 4)
    End If
    private_GetComponentName = fileName
End Function


Private Function private_RemoveExportMetadata(ByVal sourceText As String) As String
    Dim lines As Variant
    Dim line As Variant
    Dim result As String
    Dim insideHeaderBlock As Boolean
    Dim trimmedLine As String

    lines = VBA.Split(VBA.Replace$(sourceText, VBA.vbCrLf, VBA.vbLf), VBA.vbLf)
    For Each line In lines
        trimmedLine = VBA.Trim$(VBA.CStr(line))
        If VBA.StrComp(trimmedLine, "VERSION 1.0 CLASS", VBA.vbTextCompare) = 0 Then
            insideHeaderBlock = True
        ElseIf insideHeaderBlock And VBA.StrComp(trimmedLine, "END", VBA.vbTextCompare) = 0 Then
            insideHeaderBlock = False
        ElseIf Not insideHeaderBlock And _
            VBA.Left$(trimmedLine, 10) <> "Attribute " Then
            result = result & VBA.CStr(line) & VBA.vbCrLf
        End If
    Next line
    private_RemoveExportMetadata = result
End Function


Private Function private_ReadTextFile(ByVal filePath As String) As String
    Const AD_TYPE_TEXT As Long = 2
    Const AD_READ_ALL As Long = -1

    Dim textStream As Object

    ' Исходники проекта хранятся в UTF-8. OpenTextFile с системной ANSI-
    ' кодировкой превращает кириллицу в последовательности вида "Рџ...".
    Set textStream = VBA.CreateObject("ADODB.Stream")
    textStream.Type = AD_TYPE_TEXT
    textStream.Charset = "utf-8"
    textStream.Open
    textStream.LoadFromFile filePath
    private_ReadTextFile = textStream.ReadText(AD_READ_ALL)
    textStream.Close
End Function


Private Function private_FolderExists(ByVal folderPath As String) As Boolean
    Dim fileSystem As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    private_FolderExists = fileSystem.FolderExists(folderPath)
End Function