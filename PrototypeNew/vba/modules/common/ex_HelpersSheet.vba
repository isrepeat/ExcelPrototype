Attribute VB_Name = "ex_HelpersSheet"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_HelpersSheet.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Sub fn_SetBusyCursor(ByVal isBusy As Boolean)
    On Error Resume Next
    If isBusy Then
        Application.Cursor = xlWait
    Else
        Application.Cursor = xlDefault
    End If
    Err.Clear
    On Error GoTo 0
End Sub


Public Function fn_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set fn_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function


Public Function fn_GetRuntimeWorksheetByName(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set fn_GetRuntimeWorksheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function


Public Function fn_TryCastPageBase(ByVal pageRef As Object, ByRef outPageBase As obj_PageBase) As Boolean
    Set outPageBase = Nothing
    If pageRef Is Nothing Then Exit Function

    If TypeOf pageRef Is obj_PageBase Then
        Set outPageBase = pageRef
        fn_TryCastPageBase = True
        Exit Function
    End If

    If TypeOf pageRef Is obj_IPage Then
        On Error Resume Next
        Set outPageBase = pageRef.GetPageBase()
        If Err.Number <> 0 Then
            Err.Clear
            Set outPageBase = Nothing
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0

        If outPageBase Is Nothing Then Exit Function
        fn_TryCastPageBase = True
    End If
End Function


Public Function fn_TryGetPageBaseByWorksheetName(ByVal worksheetName As String, ByRef outPageBase As obj_PageBase) As Boolean
    Dim pageRef As obj_IPage

    Set outPageBase = Nothing
    If Not rt_PageManager.fn_TryGetPageByWorksheetName(worksheetName, pageRef) Then Exit Function
    If Not fn_TryCastPageBase(pageRef, outPageBase) Then Exit Function

    fn_TryGetPageBaseByWorksheetName = True
End Function


Public Function fn_TryGetPageBaseByWorksheet(ByVal ws As Worksheet, ByRef outPageBase As obj_PageBase) As Boolean
    Dim pageRef As obj_IPage

    Set outPageBase = Nothing
    If ws Is Nothing Then Exit Function
    If Not rt_PageManager.fn_TryGetPageByWorksheet(ws, pageRef) Then Exit Function
    If Not fn_TryCastPageBase(pageRef, outPageBase) Then Exit Function

    fn_TryGetPageBaseByWorksheet = True
End Function


Public Function fn_TryGetActivePageBase(ByRef outPageBase As obj_PageBase) As Boolean
    Dim activeSheetObj As Object
    Dim ws As Worksheet

    Set outPageBase = Nothing

    On Error Resume Next
    Set activeSheetObj = Application.ActiveSheet
    On Error GoTo 0
    If Not TypeOf activeSheetObj Is Worksheet Then Exit Function

    Set ws = activeSheetObj
    If Not fn_TryGetPageBaseByWorksheet(ws, outPageBase) Then Exit Function
    fn_TryGetActivePageBase = True
End Function


Public Function fn_TryRerenderActivePage(Optional ByVal reason As String = VBA.vbNullString) As Boolean
    Dim activeSheetObj As Object
    Dim ws As Worksheet
    Dim page As obj_IPage

    On Error Resume Next
    Set activeSheetObj = Application.ActiveSheet
    On Error GoTo 0

    If Not TypeOf activeSheetObj Is Worksheet Then Exit Function
    Set ws = activeSheetObj
    If Not rt_PageManager.fn_TryGetPageByWorksheet(ws, page) Then Exit Function

    fn_TryRerenderActivePage = rt_PageManager.fn_RenderPage(page, reason)
End Function


Public Sub fn_RemovePageByWorksheet(ByVal ws As Worksheet)
    Dim page As obj_IPage

    If ws Is Nothing Then Exit Sub
    If Not rt_PageManager.fn_TryGetPageByWorksheet(ws, page) Then Exit Sub

    Call rt_PageManager.fn_RemovePage(page, False)
End Sub
