Private m_filterInputActive As Boolean
Private m_filterInputCell As Range
Private m_filterInputAddress As String
Private m_filterInputHadFormula As Boolean
Private m_filterInputFormula As Variant
Private m_filterInputValue As Variant
Private m_filterInputInteriorColor As Long
Private m_filterInputInteriorPattern As Long
Private m_filterTable As ListObject
Private m_filterFieldIndex As Long
Private m_filterHeaderCell As Range

' --------------------------------------
' namespace API {
' --------------------------------------
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

    If Not private_VbaReload_TryResolveVbaFolder(targetWorkbook, vbaFolderPath) Then Exit Sub

    Set importFiles = New Collection
    Set documentImportFiles = New Collection
    If Not private_VbaReload_TryCollectConfiguredVbaFiles( _
            vbaFolderPath, targetWorkbook, importFiles, documentImportFiles) Then Exit Sub
    If importFiles.Count = 0 And documentImportFiles.Count = 0 Then
        VBA.MsgBox _
            "No .bas, .cls, .frm, .vba, or .utf8.vba files were found in: " & vbaFolderPath, _
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

    ' Профиль описывает полный набор исходников. Очищаем код книги и листов,
    ' чтобы обработчики событий от предыдущего профиля не оставались активными.
    private_VbaReload_ClearDocumentModules targetWorkbook

    For Each importFile In importFiles
        private_VbaReload_ImportVbaFile vbProject, VBA.CStr(importFile)
        importedCount = importedCount + 1
    Next importFile

    ' Документные модули нельзя импортировать как обычные: Excel создаст
    ' отдельный Module и события Workbook/Worksheet не будут вызываться.
    For Each importFile In documentImportFiles
        private_VbaReload_ImportDocumentVbaFile targetWorkbook, VBA.CStr(importFile)
        importedCount = importedCount + 1
    Next importFile
    private_VbaReload_InitializeReloadedWorkbook targetWorkbook

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

Public Sub fn_RecalculateWorkbook()
    Application.Calculate
End Sub


Public Sub fn_RecalculateActiveSheet()
    ActiveSheet.Calculate
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


Public Sub fn_DatePlusOne()
    If IsDate(ActiveCell.value) Then
        ActiveCell.value = CDate(ActiveCell.value) + 1
    End If
End Sub


Public Sub fn_DateMinusOne()
    If IsDate(ActiveCell.value) Then
        ActiveCell.value = CDate(ActiveCell.value) - 1
    End If
End Sub


Public Sub fn_FilterContainsCurrentColumn()
    Dim tableObj As ListObject
    Dim columnIndex As Long
    Dim headerCell As Range
    Dim inputCell As Range

    private_Logging_WriteFilterLog "START | Workbook=" & ActiveWorkbook.Name & _
        " | Sheet=" & ActiveSheet.Name & " | Cell=" & ActiveCell.Address(False, False)

    ' Use the free cell immediately above the active column header as the input field.
    On Error Resume Next
    Set tableObj = ActiveCell.ListObject
    On Error GoTo EH
    If Not tableObj Is Nothing Then
        columnIndex = ActiveCell.Column - tableObj.Range.Columns(1).Column + 1
        If columnIndex < 1 Or columnIndex > tableObj.ListColumns.Count Then Exit Sub
        Set headerCell = tableObj.HeaderRowRange.Cells(1, columnIndex)
        If headerCell.Row = 1 Then
            Set inputCell = tableObj.Parent.Cells( _
                tableObj.Range.Row + tableObj.Range.Rows.Count, _
                headerCell.Column)
        Else
            Set inputCell = headerCell.Offset(-1, 0)
        End If
        private_Logging_WriteFilterLog "TABLE | Name=" & tableObj.Name & _
            " | Range=" & tableObj.Range.Address(False, False) & _
            " | Field=" & VBA.CStr(columnIndex) & _
            " | Header=" & tableObj.ListColumns(columnIndex).Name
        private_Filter_BeginFilterInput tableObj, columnIndex, headerCell, inputCell
        Exit Sub
    End If

    VBA.MsgBox "Select a cell inside an Excel Table before using Ctrl+Q.", _
        VBA.vbExclamation, "Filter contains"
    Exit Sub
EH:
    private_Logging_WriteFilterLog "ERROR | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    VBA.MsgBox "Filtering failed: [" & VBA.CStr(VBA.Err.Number) & "] " & _
        VBA.Err.Description, VBA.vbExclamation, "Filter contains"
End Sub


Public Sub fn_HandleFilterInputCellChange( _
    ByVal sheetObject As Object, _
    ByVal target As Range _
)
    If Not private_Filter_IsPendingFilterInput(sheetObject, target) Then Exit Sub
    private_Filter_ApplyPendingFilter
End Sub


Public Sub fn_HandleFilterInputSelectionChange( _
    ByVal sheetObject As Object, _
    ByVal target As Range _
)
    If Not private_Filter_IsPendingFilterSheet(sheetObject) Then Exit Sub
    If target.CountLarge <> 1 Then Exit Sub
    If VBA.StrComp(target.Address(False, False), m_filterInputAddress, _
            VBA.vbTextCompare) = 0 Then Exit Sub

    If VBA.Len(private_Filter_NormalizeFilterQuery(m_filterInputCell.Value2)) = 0 Then
        private_Filter_CancelPendingFilterInput
    Else
        private_Filter_ApplyPendingFilter
    End If
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Filter {
' --------------------------------------
Private Sub private_Filter_BeginFilterInput( _
    ByVal tableObj As ListObject, _
    ByVal fieldIndex As Long, _
    ByVal headerCell As Range, _
    ByVal inputCell As Range _
)
    private_Filter_CancelPendingFilterInput

    If inputCell.HasFormula Or Not IsEmpty(inputCell.Value2) Then
        VBA.MsgBox "The cell above the table header must be empty for Ctrl+Q filtering.", _
            VBA.vbExclamation, "Filter contains"
        Exit Sub
    End If

    Set m_filterInputCell = inputCell
    Set m_filterTable = tableObj
    Set m_filterHeaderCell = headerCell
    m_filterFieldIndex = fieldIndex
    m_filterInputAddress = inputCell.Address(False, False)
    m_filterInputHadFormula = inputCell.HasFormula
    If m_filterInputHadFormula Then
        m_filterInputFormula = inputCell.Formula
    Else
        m_filterInputValue = inputCell.Value2
    End If
    m_filterInputInteriorColor = inputCell.Interior.Color
    m_filterInputInteriorPattern = inputCell.Interior.Pattern

    m_filterInputActive = True
    inputCell.Interior.Pattern = xlSolid
    inputCell.Interior.Color = RGB(16, 72, 97)
    ' Goto moves the active cell without cancelling a pending copy operation.
    Application.Goto inputCell, False
    private_Logging_WriteFilterLog "INPUT_STARTED | Address=" & m_filterInputAddress
End Sub


Private Function private_Filter_IsPendingFilterSheet(ByVal sheetObject As Object) As Boolean
    If Not m_filterInputActive Then Exit Function
    If m_filterInputCell Is Nothing Then Exit Function
    private_Filter_IsPendingFilterSheet = sheetObject Is m_filterInputCell.Worksheet
End Function


Private Function private_Filter_IsPendingFilterInput( _
    ByVal sheetObject As Object, _
    ByVal target As Range _
) As Boolean
    If Not private_Filter_IsPendingFilterSheet(sheetObject) Then Exit Function
    If target.CountLarge <> 1 Then Exit Function
    private_Filter_IsPendingFilterInput = VBA.StrComp( _
        target.Address(False, False), m_filterInputAddress, VBA.vbTextCompare) = 0
End Function


Private Sub private_Filter_ApplyPendingFilter()
    Dim filterDate As Date
    Dim query As String
    Dim previousEnableEvents As Boolean

    On Error GoTo EH
    If Not m_filterInputActive Then Exit Sub
    query = private_Filter_NormalizeFilterQuery(VBA.CStr(m_filterInputCell.Value2))
    previousEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
    If VBA.Len(query) > 0 Then
        If private_Filter_IsDateFilterColumn(m_filterTable, m_filterFieldIndex) And _
                (VBA.IsDate(query) Or private_Filter_IsExcelDateSerial(query)) Then
            If private_Filter_IsExcelDateSerial(query) Then
                filterDate = VBA.DateValue(VBA.CDate(VBA.CDbl(query)))
            Else
                filterDate = VBA.DateValue(VBA.CDate(query))
            End If
            m_filterTable.Range.AutoFilter Field:=m_filterFieldIndex, _
                Criteria1:=">=" & VBA.CStr(VBA.CLng(filterDate)), _
                Operator:=xlAnd, _
                Criteria2:="<" & VBA.CStr(VBA.CLng(filterDate) + 1)
            private_Logging_WriteFilterLog "DATE_FILTER_APPLIED | Field=" & _
                VBA.CStr(m_filterFieldIndex) & " | Date=" & _
                VBA.Format$(filterDate, "yyyy-mm-dd")
        Else
            m_filterTable.Range.AutoFilter Field:=m_filterFieldIndex, _
                Criteria1:="*" & query & "*"
            private_Logging_WriteFilterLog "TEXT_FILTER_APPLIED | Field=" & _
                VBA.CStr(m_filterFieldIndex) & " | Query=" & query
        End If
    End If
    private_Filter_RestoreFilterInputCell
    m_filterHeaderCell.Select

CleanExit:
    Application.EnableEvents = previousEnableEvents
    private_Filter_ClearPendingFilterInput
    Exit Sub
EH:
    private_Logging_WriteFilterLog "FILTER_ERROR | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    private_Filter_RestoreFilterInputCell
    Resume CleanExit
End Sub


Private Function private_Filter_IsExcelDateSerial(ByVal query As String) As Boolean
    Dim serialValue As Double

    If Not VBA.IsNumeric(query) Then Exit Function
    serialValue = VBA.CDbl(query)
    private_Filter_IsExcelDateSerial = serialValue >= 1 And serialValue <= 2958465
End Function


Private Function private_Filter_IsDateFilterColumn( _
    ByVal tableObj As ListObject, _
    ByVal fieldIndex As Long _
) As Boolean
    Dim columnCell As Range
    Dim numberFormat As String

    If tableObj.DataBodyRange Is Nothing Then Exit Function
    For Each columnCell In tableObj.ListColumns(fieldIndex).DataBodyRange.Cells
        If Not VBA.IsError(columnCell.Value) Then
            If VBA.Len(VBA.CStr(columnCell.Value2)) > 0 Then
                numberFormat = VBA.Trim$(columnCell.NumberFormat)
                If VBA.StrComp(numberFormat, "General", VBA.vbTextCompare) = 0 Then _
                    Exit Function
                private_Filter_IsDateFilterColumn = VBA.IsDate(columnCell.Value)
                Exit Function
            End If
        End If
    Next columnCell
End Function


Private Sub private_Filter_CancelPendingFilterInput()
    If Not m_filterInputActive Then Exit Sub
    On Error Resume Next
    Application.EnableEvents = False
    private_Filter_RestoreFilterInputCell
    Application.EnableEvents = True
    private_Filter_ClearPendingFilterInput
End Sub


Private Sub private_Filter_RestoreFilterInputCell()
    If m_filterInputCell Is Nothing Then Exit Sub
    If m_filterInputHadFormula Then
        m_filterInputCell.Formula = m_filterInputFormula
    Else
        m_filterInputCell.Value2 = m_filterInputValue
    End If
    m_filterInputCell.Interior.Color = m_filterInputInteriorColor
    m_filterInputCell.Interior.Pattern = m_filterInputInteriorPattern
End Sub


Private Sub private_Filter_ClearPendingFilterInput()
    m_filterInputActive = False
    m_filterInputAddress = VBA.vbNullString
    m_filterInputHadFormula = False
    m_filterInputFormula = Empty
    m_filterInputValue = Empty
    m_filterInputInteriorColor = 0
    m_filterInputInteriorPattern = xlNone
    m_filterFieldIndex = 0
    Set m_filterInputCell = Nothing
    Set m_filterTable = Nothing
    Set m_filterHeaderCell = Nothing
End Sub


Private Function private_Filter_NormalizeFilterQuery(ByVal textValue As String) As String
    ' Normalize line breaks, non-breaking spaces, and repeated spaces.
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, ChrW(160), " ")
    textValue = Trim(textValue)
    Do While InStr(textValue, "  ") > 0
        textValue = Replace(textValue, "  ", " ")
    Loop
    private_Filter_NormalizeFilterQuery = textValue
End Function
' --------------------------------------
' } // namespace Filter
' --------------------------------------

' --------------------------------------
' namespace Logging {
' --------------------------------------
Private Sub private_Logging_WriteFilterLog(ByVal messageText As String)
    Const FOR_APPENDING As Long = 8
    Const TRISTATE_TRUE As Long = -1
    Dim fileSystem As Object
    Dim logFile As Object
    Dim tempPath As String
    Dim logFolderPath As String

    On Error Resume Next
    tempPath = VBA.Environ$("TEMP")
    If VBA.Len(tempPath) = 0 Then Exit Sub
    logFolderPath = tempPath & "\PERSONAL.EXCEL"
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(logFolderPath) Then fileSystem.CreateFolder logFolderPath
    Set logFile = fileSystem.OpenTextFile( _
        logFolderPath & "\filter.log", FOR_APPENDING, True, TRISTATE_TRUE)
    logFile.WriteLine VBA.Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    logFile.Close
End Sub
' --------------------------------------
' } // namespace Logging
' --------------------------------------

' --------------------------------------
' namespace VbaReload {
' --------------------------------------
' Reinitializes runtime event handlers after a hot reload.
Private Sub private_VbaReload_InitializeReloadedWorkbook( _
    ByVal targetWorkbook As Workbook _
)
    Dim bootstrapComponent As Object
    Dim macroReference As String

    On Error Resume Next
    Set bootstrapComponent = targetWorkbook.VBProject.VBComponents( _
        "ex_DocumentGenerationBootstrap")
    On Error GoTo EH
    If bootstrapComponent Is Nothing Then Exit Sub
    If bootstrapComponent.CodeModule.CountOfLines = 0 Then Exit Sub

    macroReference = "'" & VBA.Replace$(targetWorkbook.Name, "'", "''") & _
        "'!ex_DocumentGenerationBootstrap.fn_Initialize"
    Application.Run macroReference
    Exit Sub
EH:
    VBA.Err.Raise VBA.Err.Number, "private_VbaReload_InitializeReloadedWorkbook", _
        "Failed to initialize reloaded workbook '" & targetWorkbook.Name & _
        "': " & VBA.Err.Description
End Sub


' Source folder and profile resolution.
Private Function private_VbaReload_TryResolveVbaFolder( _
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
    If private_VbaReload_FolderExists(VBA.CStr(candidatePath)) Then _
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
        If private_VbaReload_FolderExists(VBA.CStr(candidatePath)) Then _
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
    private_VbaReload_TryResolveVbaFolder = True
End Function


Private Function private_VbaReload_TryCollectConfiguredVbaFiles( _
    ByVal vbaFolderPath As String, _
    ByVal targetWorkbook As Workbook, _
    ByRef outFiles As Collection, _
    ByRef outDocumentFiles As Collection _
) As Boolean
    Const CONFIG_FILE_NAME As String = "modules.json"

    Dim fileSystem As Object
    Dim configPath As String
    Dim profileName As String
    Dim configText As String
    Dim relativeFiles As Collection
    Dim relativeFile As Variant
    Dim sourcePath As String
    Dim lowerName As String
    Dim importedPaths As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    configPath = vbaFolderPath & Application.PathSeparator & CONFIG_FILE_NAME
    If Not fileSystem.FileExists(configPath) Then
        VBA.MsgBox "VBA import configuration was not found: " & configPath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    profileName = fileSystem.GetBaseName(targetWorkbook.Name)
    configText = private_VbaReload_ReadUtf8TextFile(configPath)
    Set relativeFiles = New Collection
    If Not private_VbaReload_TryReadJsonStringArray(configText, profileName, relativeFiles) Then
        VBA.MsgBox "No VBA import profile was found for workbook: " & _
            targetWorkbook.Name, VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    If relativeFiles.Count = 0 Then
        VBA.MsgBox "The VBA import profile is empty for workbook: " & _
            targetWorkbook.Name, VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    Set importedPaths = VBA.CreateObject("Scripting.Dictionary")
    importedPaths.CompareMode = VBA.vbTextCompare
    For Each relativeFile In relativeFiles
        If Not private_VbaReload_TryResolveConfiguredSourcePath( _
                vbaFolderPath, VBA.CStr(relativeFile), sourcePath) Then Exit Function
        If importedPaths.Exists(sourcePath) Then
            VBA.MsgBox "The VBA import profile contains a duplicate module: " & _
                VBA.CStr(relativeFile), VBA.vbExclamation, "Reload VBA"
            Exit Function
        End If
        importedPaths.Add sourcePath, True
        lowerName = VBA.LCase$(fileSystem.GetFileName(sourcePath))
        If private_VbaReload_IsDocumentModuleSource(lowerName) Then
            outDocumentFiles.Add sourcePath
        Else
            outFiles.Add sourcePath
        End If
    Next relativeFile
    private_VbaReload_TryCollectConfiguredVbaFiles = True
End Function


' Допускаются только относительные пути внутри папки vba из профиля книги.
Private Function private_VbaReload_TryResolveConfiguredSourcePath( _
    ByVal vbaFolderPath As String, _
    ByVal relativePath As String, _
    ByRef outSourcePath As String _
) As Boolean
    Dim fileSystem As Object
    Dim normalizedRoot As String
    Dim normalizedPath As String
    Dim lowerName As String

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    relativePath = VBA.Replace$(VBA.Trim$(relativePath), "/", "\")
    If VBA.Len(relativePath) = 0 Or VBA.InStr(relativePath, "..") > 0 Or _
       VBA.InStr(relativePath, ":") > 0 Or VBA.Left$(relativePath, 1) = "\" Then
        VBA.MsgBox "Invalid relative VBA module path: " & relativePath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    normalizedRoot = fileSystem.GetAbsolutePathName(vbaFolderPath)
    normalizedPath = fileSystem.GetAbsolutePathName( _
        normalizedRoot & Application.PathSeparator & relativePath)
    If VBA.StrComp(VBA.Left$(normalizedPath, VBA.Len(normalizedRoot) + 1), _
            normalizedRoot & Application.PathSeparator, VBA.vbTextCompare) <> 0 Then
        VBA.MsgBox "VBA module path is outside the vba folder: " & relativePath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    If Not fileSystem.FileExists(normalizedPath) Then
        VBA.MsgBox "Configured VBA module was not found: " & relativePath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    lowerName = VBA.LCase$(normalizedPath)
    If Not (VBA.Right$(lowerName, 4) = ".bas" Or _
            VBA.Right$(lowerName, 4) = ".cls" Or _
            VBA.Right$(lowerName, 4) = ".frm" Or _
            VBA.Right$(lowerName, 4) = ".vba") Then
        VBA.MsgBox "Configured file is not a VBA module: " & relativePath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    outSourcePath = normalizedPath
    private_VbaReload_TryResolveConfiguredSourcePath = True
End Function


Private Sub private_VbaReload_ClearDocumentModules(ByVal targetWorkbook As Workbook)
    Const VBEXT_CT_DOCUMENT As Long = 100

    Dim vbComponent As Object

    For Each vbComponent In targetWorkbook.VBProject.VBComponents
        If CLng(vbComponent.Type) = VBEXT_CT_DOCUMENT Then
            If vbComponent.CodeModule.CountOfLines > 0 Then _
                vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
        End If
    Next vbComponent
End Sub


' Minimal JSON parsing for modules.json.
Private Function private_VbaReload_TryReadJsonStringArray( _
    ByVal jsonText As String, _
    ByVal profileName As String, _
    ByRef outValues As Collection _
) As Boolean
    Dim keyToken As String
    Dim position As Long
    Dim currentChar As String
    Dim valueText As String

    keyToken = """" & profileName & """"
    position = VBA.InStr(1, jsonText, keyToken, VBA.vbBinaryCompare)
    If position = 0 Then Exit Function
    position = position + VBA.Len(keyToken)
    private_VbaReload_SkipJsonWhitespace jsonText, position
    If VBA.Mid$(jsonText, position, 1) <> ":" Then Exit Function
    position = position + 1
    private_VbaReload_SkipJsonWhitespace jsonText, position
    If VBA.Mid$(jsonText, position, 1) <> "[" Then Exit Function
    position = position + 1

    Do
        private_VbaReload_SkipJsonWhitespace jsonText, position
        currentChar = VBA.Mid$(jsonText, position, 1)
        If currentChar = "]" Then
            private_VbaReload_TryReadJsonStringArray = True
            Exit Function
        End If
        If outValues.Count > 0 Then
            If currentChar <> "," Then Exit Function
            position = position + 1
            private_VbaReload_SkipJsonWhitespace jsonText, position
        End If
        If Not private_VbaReload_TryReadJsonString(jsonText, position, valueText) Then Exit Function
        outValues.Add valueText
    Loop
End Function


Private Sub private_VbaReload_SkipJsonWhitespace(ByVal jsonText As String, ByRef position As Long)
    Do While position <= VBA.Len(jsonText)
        Select Case VBA.Mid$(jsonText, position, 1)
            Case " ", VBA.vbTab, VBA.vbCr, VBA.vbLf
                position = position + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub


Private Function private_VbaReload_TryReadJsonString( _
    ByVal jsonText As String, _
    ByRef position As Long, _
    ByRef outValue As String _
) As Boolean
    Dim currentChar As String
    Dim escapedChar As String

    outValue = VBA.vbNullString
    If VBA.Mid$(jsonText, position, 1) <> """" Then Exit Function
    position = position + 1
    Do While position <= VBA.Len(jsonText)
        currentChar = VBA.Mid$(jsonText, position, 1)
        If currentChar = """" Then
            position = position + 1
            private_VbaReload_TryReadJsonString = True
            Exit Function
        End If
        If currentChar = "\" Then
            position = position + 1
            escapedChar = VBA.Mid$(jsonText, position, 1)
            Select Case escapedChar
                Case """", "\", "/"
                    outValue = outValue & escapedChar
                Case "b"
                    outValue = outValue & VBA.ChrW$(8)
                Case "f"
                    outValue = outValue & VBA.ChrW$(12)
                Case "n"
                    outValue = outValue & VBA.vbLf
                Case "r"
                    outValue = outValue & VBA.vbCr
                Case "t"
                    outValue = outValue & VBA.vbTab
                Case Else
                    Exit Function
            End Select
        Else
            outValue = outValue & currentChar
        End If
        position = position + 1
    Loop
End Function


Private Function private_VbaReload_IsDocumentModuleSource( _
    ByVal lowerFileName As String _
) As Boolean
    private_VbaReload_IsDocumentModuleSource = ( _
        private_VbaReload_IsThisWorkbookModuleSource(lowerFileName) Or _
        (VBA.Left$(lowerFileName, 3) = "ws_" And _
         private_VbaReload_IsVbaSourceFile(lowerFileName)))
End Function


' VBA project import operations.
Private Sub private_VbaReload_ImportVbaFile( _
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

    sourceText = private_VbaReload_ReadUtf8TextFile(sourcePath)
    componentName = private_VbaReload_GetComponentName(sourcePath, sourceText)
    componentType = private_VbaReload_GetVbaComponentType(lowerPath, sourceText)

    Set vbComponent = vbProject.VBComponents.Add(componentType)
    vbComponent.Name = componentName
    vbComponent.CodeModule.AddFromString private_VbaReload_PrepareSourceForVbe(sourceText)
    Exit Sub

EH:
    VBA.Err.Raise VBA.Err.Number, "private_VbaReload_ImportVbaFile", _
        "Failed to import '" & sourcePath & "': " & VBA.Err.Description
End Sub


' Обновляет исходник в уже существующем модуле книги или листа.
Private Sub private_VbaReload_ImportDocumentVbaFile( _
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
    sourceText = private_VbaReload_ReadUtf8TextFile(sourcePath)
    If VBA.Len(sourceText) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 1201, _
            "private_VbaReload_ImportDocumentVbaFile", "Source is empty: " & sourcePath
    End If

    If private_VbaReload_IsThisWorkbookModuleSource(fileName) Then
        componentName = targetWorkbook.CodeName
    Else
        worksheetName = private_VbaReload_GetDocumentModuleSourceName(fileName, "ws_")
        On Error Resume Next
        Set targetSheet = targetWorkbook.Worksheets(worksheetName)
        On Error GoTo EH
        If targetSheet Is Nothing Then
            VBA.Err.Raise VBA.vbObjectError + 1202, _
                "private_VbaReload_ImportDocumentVbaFile", _
                "Worksheet '" & worksheetName & "' was not found for " & sourcePath
        End If
        componentName = targetSheet.CodeName
    End If

    Set vbComponent = targetWorkbook.VBProject.VBComponents(componentName)
    With vbComponent.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString private_VbaReload_PrepareSourceForVbe(sourceText)
    End With
    Exit Sub
EH:
    VBA.Err.Raise VBA.Err.Number, "private_VbaReload_ImportDocumentVbaFile", _
        "Failed to import document module '" & sourcePath & "': " & _
        VBA.Err.Description
End Sub


Private Function private_VbaReload_GetVbaComponentType( _
    ByVal lowerPath As String, _
    ByVal sourceText As String _
) As Long
    Const VBEXT_CT_STD_MODULE As Long = 1
    Const VBEXT_CT_CLASS_MODULE As Long = 2
    Const VBEXT_CT_MS_FORM As Long = 3

    If VBA.Right$(lowerPath, 8) = ".cls.vba" Or _
       VBA.Right$(lowerPath, 13) = ".cls.utf8.vba" Or _
        VBA.InStr(1, sourceText, "VERSION 1.0 CLASS", VBA.vbTextCompare) > 0 Then
        private_VbaReload_GetVbaComponentType = VBEXT_CT_CLASS_MODULE
    ElseIf VBA.Right$(lowerPath, 8) = ".frm.vba" Or _
           VBA.Right$(lowerPath, 13) = ".frm.utf8.vba" Or _
           VBA.InStr(1, sourceText, "BEGIN VB.Form", VBA.vbTextCompare) > 0 Then
        private_VbaReload_GetVbaComponentType = VBEXT_CT_MS_FORM
    Else
        private_VbaReload_GetVbaComponentType = VBEXT_CT_STD_MODULE
    End If
End Function


Private Function private_VbaReload_GetComponentName( _
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
            private_VbaReload_GetComponentName = VBA.Mid$(sourceText, nameStart, nameEnd - nameStart)
            Exit Function
        End If
    End If

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    fileName = fileSystem.GetBaseName(sourcePath)
    If VBA.Right$(VBA.LCase$(fileName), 5) = ".utf8" Then
        fileName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".utf8"))
    End If
    If VBA.Right$(VBA.LCase$(fileName), 4) = ".cls" Or _
        VBA.Right$(VBA.LCase$(fileName), 4) = ".bas" Or _
        VBA.Right$(VBA.LCase$(fileName), 4) = ".frm" Then
        fileName = VBA.Left$(fileName, VBA.Len(fileName) - 4)
    End If
    private_VbaReload_GetComponentName = fileName
End Function


' Source preparation for VBE import.
Private Function private_VbaReload_RemoveExportMetadata(ByVal sourceText As String) As String
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
    private_VbaReload_RemoveExportMetadata = result
End Function


Private Function private_VbaReload_ReadUtf8TextFile(ByVal filePath As String) As String
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
    private_VbaReload_ReadUtf8TextFile = textStream.ReadText(AD_READ_ALL)
    textStream.Close
    If VBA.Left$(private_VbaReload_ReadUtf8TextFile, 1) = VBA.ChrW$(65279) Then
        private_VbaReload_ReadUtf8TextFile = VBA.Mid$(private_VbaReload_ReadUtf8TextFile, 2)
    End If
End Function


' Источник .utf8.vba обязательно читается как UTF-8 и передаётся в
' CodeModule.AddFromString как Unicode String, без VBComponents.Import.
Private Function private_VbaReload_IsVbaSourceFile(ByVal lowerFileName As String) As Boolean
    private_VbaReload_IsVbaSourceFile = (VBA.Right$(lowerFileName, 4) = ".vba")
End Function


Private Function private_VbaReload_IsThisWorkbookModuleSource(ByVal fileName As String) As Boolean
    private_VbaReload_IsThisWorkbookModuleSource = ( _
        VBA.StrComp(fileName, "ThisWorkbook.vba", VBA.vbTextCompare) = 0 Or _
        VBA.StrComp(fileName, "ThisWorkbook.utf8.vba", VBA.vbTextCompare) = 0)
End Function


Private Function private_VbaReload_GetDocumentModuleSourceName( _
    ByVal fileName As String, _
    ByVal requiredPrefix As String _
) As String
    Dim sourceStem As String

    sourceStem = fileName
    If VBA.Right$(VBA.LCase$(sourceStem), 9) = ".utf8.vba" Then
        sourceStem = VBA.Left$(sourceStem, VBA.Len(sourceStem) - VBA.Len(".utf8.vba"))
    ElseIf VBA.Right$(VBA.LCase$(sourceStem), 4) = ".vba" Then
        sourceStem = VBA.Left$(sourceStem, VBA.Len(sourceStem) - VBA.Len(".vba"))
    Else
        VBA.Err.Raise VBA.vbObjectError + 1203, _
            "private_VbaReload_GetDocumentModuleSourceName", _
            "Unsupported document module source extension: " & fileName
    End If
    If VBA.Left$(sourceStem, VBA.Len(requiredPrefix)) <> requiredPrefix Then
        VBA.Err.Raise VBA.vbObjectError + 1204, _
            "private_VbaReload_GetDocumentModuleSourceName", _
            "Document module source has invalid prefix: " & fileName
    End If
    private_VbaReload_GetDocumentModuleSourceName = VBA.Mid$(sourceStem, _
        VBA.Len(requiredPrefix) + 1)
End Function


Private Function private_VbaReload_FolderExists(ByVal folderPath As String) As Boolean
    Dim fileSystem As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    private_VbaReload_FolderExists = fileSystem.FolderExists(folderPath)
End Function


' The classic VBE stores source code using the current system code page.
' Convert only non-ASCII string literals into ChrW$ expressions before adding
' code, so the imported module remains executable on every legacy code page.
Private Function private_VbaReload_PrepareSourceForVbe(ByVal sourceText As String) As String
    private_VbaReload_PrepareSourceForVbe = private_VbaReload_RemoveExportMetadata( _
        private_VbaReload_EncodeUnicodeStringLiterals(sourceText))
End Function


Private Function private_VbaReload_EncodeUnicodeStringLiterals( _
    ByVal sourceText As String _
) As String
    Dim sourceLines As Variant
    Dim sourceLine As Variant
    Dim resultText As String

    sourceLines = VBA.Split(VBA.Replace$(sourceText, VBA.vbCrLf, VBA.vbLf), _
        VBA.vbLf)
    For Each sourceLine In sourceLines
        resultText = resultText & private_VbaReload_EncodeUnicodeStringLiteralsOnLine( _
            VBA.CStr(sourceLine)) & VBA.vbCrLf
    Next sourceLine

    private_VbaReload_EncodeUnicodeStringLiterals = resultText
End Function


Private Function private_VbaReload_EncodeUnicodeStringLiteralsOnLine( _
    ByVal sourceLine As String _
) As String
    Dim charIndex As Long
    Dim literalStartIndex As Long
    Dim literalText As String
    Dim currentCharacter As String
    Dim resultText As String
    Dim literalClosed As Boolean

    charIndex = 1
    Do While charIndex <= VBA.Len(sourceLine)
        currentCharacter = VBA.Mid$(sourceLine, charIndex, 1)

        If currentCharacter = "'" Then
            ' A quote begins a comment outside a string literal.
            resultText = resultText & VBA.Mid$(sourceLine, charIndex)
            Exit Do
        End If

        If currentCharacter <> """" Then
            resultText = resultText & currentCharacter
            charIndex = charIndex + 1
        Else
            literalStartIndex = charIndex
            charIndex = charIndex + 1
            literalText = VBA.vbNullString
            literalClosed = False

            Do While charIndex <= VBA.Len(sourceLine)
                currentCharacter = VBA.Mid$(sourceLine, charIndex, 1)
                If currentCharacter <> """" Then
                    literalText = literalText & currentCharacter
                    charIndex = charIndex + 1
                ElseIf charIndex < VBA.Len(sourceLine) And _
                       VBA.Mid$(sourceLine, charIndex + 1, 1) = """" Then
                    literalText = literalText & """"
                    charIndex = charIndex + 2
                Else
                    charIndex = charIndex + 1
                    literalClosed = True
                    Exit Do
                End If
            Loop

            If Not literalClosed Then
                ' Keep an unterminated literal unchanged and let the VBE report it.
                resultText = resultText & VBA.Mid$(sourceLine, literalStartIndex)
                Exit Do
            End If

            resultText = resultText & private_VbaReload_EncodeUnicodeLiteral(literalText)
        End If
    Loop

    private_VbaReload_EncodeUnicodeStringLiteralsOnLine = resultText
End Function


Private Function private_VbaReload_EncodeUnicodeLiteral( _
    ByVal literalText As String _
) As String
    Dim charIndex As Long
    Dim characterCode As Long
    Dim asciiBuffer As String
    Dim expressionParts As Collection

    Set expressionParts = New Collection

    For charIndex = 1 To VBA.Len(literalText)
        characterCode = VBA.AscW(VBA.Mid$(literalText, charIndex, 1))
        If characterCode >= 0 And characterCode <= 127 Then
            asciiBuffer = asciiBuffer & VBA.Mid$(literalText, charIndex, 1)
        Else
            private_VbaReload_AppendAsciiLiteralPart expressionParts, asciiBuffer
            asciiBuffer = VBA.vbNullString
            expressionParts.Add "VBA.ChrW$(" & VBA.CStr(characterCode) & ")"
        End If
    Next charIndex
    private_VbaReload_AppendAsciiLiteralPart expressionParts, asciiBuffer

    private_VbaReload_EncodeUnicodeLiteral = private_VbaReload_JoinExpressionParts(expressionParts)
End Function


Private Sub private_VbaReload_AppendAsciiLiteralPart( _
    ByVal expressionParts As Collection, _
    ByVal asciiText As String _
)
    If VBA.Len(asciiText) = 0 Then Exit Sub

    expressionParts.Add """" & VBA.Replace$(asciiText, """", """""") & """"
End Sub


Private Function private_VbaReload_JoinExpressionParts( _
    ByVal expressionParts As Collection _
) As String
    Dim partIndex As Long
    Dim resultText As String

    For partIndex = 1 To expressionParts.Count
        If partIndex > 1 Then resultText = resultText & " & "
        resultText = resultText & VBA.CStr(expressionParts(partIndex))
    Next partIndex

    If expressionParts.Count = 0 Then resultText = """"
    private_VbaReload_JoinExpressionParts = resultText
End Function
' --------------------------------------
' } // namespace VbaReload
' --------------------------------------