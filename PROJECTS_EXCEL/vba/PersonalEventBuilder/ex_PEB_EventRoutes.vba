Attribute VB_Name = "ex_PEB_EventRoutes"
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
Public Function fn_Apply( _
    ByVal values As Object, _
    ByVal exporter As String _
) As String
    Dim eventId As String
    Dim prefix As String
    Dim policy As String
    Dim identifier As Variant
    Dim registered As Boolean

    eventId = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, ex_PEB_Export.fn_Config("events.bindingField")))
    For Each identifier In VBA.Split(ex_PEB_Export.fn_Config("events.ids"), "|")
        If VBA.CStr(identifier) = eventId Then
            registered = True
        End If
    Next identifier
    If Not registered Then
        Err.Raise 5, , "Unknown business event: " & eventId
    End If
    prefix = "events." & eventId
    values("EventId") = eventId
    values("EventName") = ex_PEB_Export.fn_Config(prefix & ".movement.event")
    values("MovementPlanId") = ex_PEB_Export.fn_Config(prefix & ".movement.plan")
    values("SourceEventName") = ex_PEB_Export.fn_Config(prefix & ".movement.sourceEvent")
    values("SourceDestination") = ex_PEB_Export.fn_Config(prefix & ".movement.sourceDestination")
    policy = ex_PEB_Export.fn_Config(prefix & ".word.policy")
    If policy <> "Render" And policy <> "Omit" Then
        Err.Raise 5, , "Unknown Word policy: " & policy
    End If
    values("WordPolicy") = policy
    values("WordAnchorId") = ex_PEB_Export.fn_Config(prefix & ".word.anchor")
    values("RoutedWordTemplate") = ex_PEB_Export.fn_Config(prefix & ".word.template")
    If exporter = "Word" Then
        fn_Apply = ex_PEB_Export.fn_Config(prefix & ".word.profile")
    ElseIf exporter = "Movement" Then
        policy = ex_PEB_Export.fn_Config(prefix & ".movement.destinationPolicy")
        If policy = "Constant" Then
            values("Destination") = ex_PEB_Export.fn_Config(prefix & ".movement.destinationValue")
        ElseIf policy <> "Field" Then
            Err.Raise 5, , "Unknown Movement destination policy: " & policy
        End If
        fn_Apply = eventId
    Else
        Err.Raise 5, , "Unknown routed exporter: " & exporter
    End If
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------