Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_ResolveChoice(ByVal bindingContext As obj_UiBindingContext) As Object
    Set fn_ResolveChoice = private_Resolve(private_FormFields(bindingContext))
End Function

Public Function fn_Capture( _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal personName As String, _
    ByVal required As Boolean _
) As Object
    Dim fields As Object
    Dim signatory As Object
    Dim field As Variant

    Set fields = private_FormFields(bindingContext)
    Set signatory = private_Resolve(fields)
    If required And signatory Is Nothing Then private_Fail "SignerMissing", personName
    For Each field In VBA.Array("SignatoryPosition", "SignatoryRank", "SignatoryName")
        fields.Add field, VBA.vbNullString
        If Not signatory Is Nothing Then fields(field) = signatory(field)
    Next field
    Set fn_Capture = fields
End Function

Public Function fn_FromHistory(ByVal row As Object) As Object
    Dim signatory As Object
    Dim field As Variant

    Set signatory = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Array("SignatoryPosition", "SignatoryRank", "SignatoryName")
        If VBA.Len(VBA.Trim$(VBA.CStr(row(field)))) = 0 Then Exit Function
        signatory.Add field, row(field)
    Next field
    Set fn_FromHistory = signatory
End Function

' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_FormFields(ByVal bindingContext As obj_UiBindingContext) As Object
    Dim fields As Object
    Dim mapping As Variant
    Dim pair As Variant

    Set fields = VBA.CreateObject("Scripting.Dictionary")
    For Each mapping In VBA.Split(ex_TDB_History.fn_Config("signerFields"), "|")
        pair = VBA.Split(mapping, "=")
        fields.Add pair(1), private_Setting(bindingContext, VBA.CStr(pair(0)))
    Next mapping
    Set private_FormFields = fields
End Function

Private Function private_Resolve(ByVal fields As Object) As Object
    Dim result As Object
    Dim name As String
    Dim rank As String
    Dim position As String
    Dim code As String
    Dim replaced As String
    Dim replacedCode As String
    Dim signerId As String
    Dim surnameEnd As Long
    Dim givenNames As Variant

    name = VBA.Trim$(VBA.CStr(fields("SignerFullName")))
    rank = VBA.Trim$(VBA.CStr(fields("SignerOwnRank")))
    position = VBA.Trim$(VBA.CStr(fields("SignerOwnPosition")))
    code = VBA.Trim$(VBA.CStr(fields("SignerCode")))
    signerId = VBA.Trim$(VBA.CStr(fields("SignerId")))
    replaced = VBA.Trim$(VBA.CStr(fields("SignerReplaced")))
    replacedCode = VBA.Trim$(VBA.CStr(fields("SignerReplacedCode")))
    If VBA.Len(name & rank & position & code & signerId & replaced & replacedCode) = 0 Then Exit Function
    If VBA.Len(name) = 0 Then private_Fail "SignerManual", ex_TravelDocumentsBuilder.fn_Config("text.PersonNameLabel")
    If VBA.Len(rank) = 0 Then private_Fail "SignerManual", ex_TravelDocumentsBuilder.fn_Config("text.PersonRankLabel")
    If VBA.Len(replaced) > 0 And VBA.Len(replacedCode) = 0 Then private_Fail "SignerLookup", replaced
    If VBA.Len(replacedCode) > 0 Then
        If (VBA.Len(signerId) > 0 And signerId = VBA.Trim$(VBA.CStr(fields("SignerReplacedId")))) Or _
                (VBA.Len(code) > 0 And VBA.StrComp(code, replacedCode, VBA.vbTextCompare) = 0) Then
            private_Fail "SignerSamePost", replaced
        End If
        position = private_Position(replacedCode, "declensions.lookups.6.valueHeader")
        position = VBA.LCase$(VBA.Left$(position, 1)) & VBA.Mid$(position, 2)
        position = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("word.signatory.actingTemplate"), "{position}", position)
    Else
        If VBA.Len(position) = 0 And VBA.Len(code) > 0 Then
            position = private_Position(code, "word.signatory.positionHeader")
        End If
        If VBA.Len(position) = 0 Then private_Fail "SignerManual", ex_TravelDocumentsBuilder.fn_Config("text.PersonPositionLabel")
    End If
    If VBA.Len(signerId) > 0 Then
        name = ex_TableQuery.fn_NormalizeText(name)
        surnameEnd = VBA.InStr(name, " ")
        If surnameEnd > 0 Then
            givenNames = VBA.Split(VBA.LTrim$(VBA.Mid$(name, surnameEnd + 1)), " ")
            name = givenNames(0) & " " & VBA.Left$(name, surnameEnd - 1)
        End If
    End If
    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.Add "SignatoryName", name
    result.Add "SignatoryRank", rank
    result.Add "SignatoryPosition", position
    Set private_Resolve = result
End Function

Private Function private_Setting( _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal field As String _
) As String
    Dim value As Variant
    Dim objectValue As Object
    Dim isObject As Boolean

    If Not bindingContext.TryGetValue("Settings", field, value, objectValue, isObject) Then private_Fail "SignerLookup", field
    If isObject Or VBA.IsError(value) Or VBA.IsNull(value) Then private_Fail "SignerLookup", field
    private_Setting = VBA.Trim$(VBA.CStr(value))
End Function

Private Function private_Position( _
    ByVal code As String, _
    ByVal headerKey As String _
) As String
    Dim table As obj_DataTable
    Dim keyHeader As String
    Dim positionHeader As String

    keyHeader = ex_TravelDocumentsBuilder.fn_Config("declensions.lookups.6.keyHeader")
    positionHeader = ex_TravelDocumentsBuilder.fn_Config(headerKey)
    Set table = private_Query("declensions.sources.Positions", keyHeader, code, VBA.Array(positionHeader))
    If table.RowCount <> 1 Then private_Fail "SignerLookup", code
    private_Position = private_TextAt(table, 1, positionHeader)
    If VBA.Len(private_Position) = 0 Then private_Fail "SignerLookup", code
    table.Dispose
End Function

Private Function private_Query( _
    ByVal prefix As String, _
    ByVal keyHeader As String, _
    ByVal key As String, _
    ByVal headers As Variant _
) As obj_DataTable
    Dim source As New obj_TableSource
    Dim query As New obj_TableQuery
    Dim service As New obj_TableQueryService
    Dim table As obj_DataTable
    Dim files As Object
    Dim path As String
    Dim header As Variant
    Dim diagnostic As String

    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then private_Fail "SignerLookup", prefix
    Set files = VBA.CreateObject("Scripting.FileSystemObject")
    path = ex_TravelDocumentsBuilder.fn_Config(prefix & ".path")
    If files.GetDriveName(path) = VBA.vbNullString And VBA.Left$(path, 2) <> "\\" Then path = files.BuildPath(ThisWorkbook.Path, path)
    source.WorkbookPath = files.GetAbsolutePathName(path)
    source.SheetName = ex_TravelDocumentsBuilder.fn_Config(prefix & ".sheet")
    source.RangeAddress = ex_TravelDocumentsBuilder.fn_Config(prefix & ".range")
    query.AddColumn keyHeader
    For Each header In headers
        query.AddColumn VBA.CStr(header)
    Next header
    query.AddCondition keyHeader, QueryEquals, key, QueryText, True
    query.Limit = 2
    If Not service.TryExecute(source, query, table, diagnostic) Then private_Fail "SignerLookup", diagnostic
    Set private_Query = table
End Function

Private Function private_TextAt( _
    ByVal table As obj_DataTable, _
    ByVal row As Long, _
    ByVal header As String _
) As String
    Dim value As Variant

    value = table.ValueAt(row, ex_TableQuery.fn_HeaderIndex(table.Headers, header))
    If VBA.IsError(value) Then private_Fail "SignerLookup", header
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    private_TextAt = ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
End Function


Private Sub private_Fail( _
    ByVal key As String, _
    ByVal details As String _
)
    VBA.Err.Raise VBA.vbObjectError + 2480, "TicketsWord.Signatory", _
        VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("word.text." & key), "{details}", details)
End Sub
' --------------------------------------
' } // namespace Private