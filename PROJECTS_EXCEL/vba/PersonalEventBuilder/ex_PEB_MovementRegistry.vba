Attribute VB_Name = "ex_PEB_MovementRegistry"
Option Explicit

Private m_values As Object
Private m_events As Object
Private m_statuses As Object
Private m_formFields As String

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_values = Nothing
    Set m_events = Nothing
    Set m_statuses = Nothing
    m_formFields = VBA.vbNullString
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Value(ByVal key As String) As String
    private_EnsureLoaded
    If Not m_values.Exists(key) Then
        Err.Raise 5, , "Unknown native Movement mapping: " & key
    End If
    fn_Value = m_values(key)
End Function

Public Function fn_Count(ByVal key As String) As Long
    fn_Count = VBA.CLng(fn_Value(key))
End Function

Public Function fn_Event(ByVal eventId As String) As Object
    private_EnsureLoaded
    If Not m_events.Exists(eventId) Then
        Err.Raise 5, , "Unknown native Movement event: " & eventId
    End If
    Set fn_Event = m_events(eventId)
End Function

Public Function fn_TryEvent( _
    ByVal eventId As String, _
    ByRef eventPlan As Object _
) As Boolean
    private_EnsureLoaded
    Set eventPlan = Nothing
    If m_events.Exists(eventId) Then
        Set eventPlan = m_events(eventId)
        fn_TryEvent = True
    End If
End Function

Public Function fn_EventsByOperation(ByVal operation As String) As Collection
    Dim result As New Collection
    Dim eventId As Variant

    private_EnsureLoaded
    If operation <> "Departure" And operation <> "Arrival" And operation <> "Transition" Then
        Err.Raise 5, , "Unknown event operation: " & operation
    End If
    For Each eventId In m_events.Keys
        If m_events(eventId)("operation") = operation Then
            result.Add eventId
        End If
    Next eventId
    Set fn_EventsByOperation = result
End Function

Public Function fn_Statuses() As Object
    private_EnsureLoaded
    Set fn_Statuses = m_statuses
End Function

Public Function fn_WriteValue( _
    ByVal values As Object, _
    ByVal field As String, _
    ByVal header As String _
) As Variant
    Dim texts As Object

    If values.Exists("MovementColumnTexts") Then
        Set texts = values("MovementColumnTexts")
        If texts.Exists(header) Then
            fn_WriteValue = private_ExportValue(texts(header))
            Exit Function
        End If
    End If
    If field = "MovementBasis" Then
        Err.Raise 5, , "Missing Movement column text: " & header
    End If
    fn_WriteValue = private_ExportValue(ex_PEB_Export.fn_Required(values, field))
End Function

Public Sub fn_ValidateTexts(ByVal values As Object)
    Dim prefix As String
    Dim mappingPrefix As Variant
    Dim prefixes As Collection
    Dim texts As Object
    Dim allowed As Object
    Dim header As Variant
    Dim field As String
    Dim index As Long

    private_EnsureLoaded
    prefix = "profiles." & ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, "MovementPlanId"))
    Set prefixes = New Collection
    prefixes.Add prefix & ".fields"
    If fn_Value(prefix & ".operation") = "transition" Then
        prefixes.Add "profiles." & fn_Value(prefix & ".closeProfile") & ".fields"
    End If
    If Not values.Exists("MovementColumnTexts") Then
        Err.Raise 5, , "Movement column texts are not prepared."
    End If
    If VBA.TypeName(values("MovementColumnTexts")) <> "Dictionary" Then
        Err.Raise 5, , "Movement column texts must be a dictionary."
    End If
    Set texts = values("MovementColumnTexts")
    Set allowed = VBA.CreateObject("Scripting.Dictionary")
    allowed.CompareMode = VBA.vbTextCompare
    For Each mappingPrefix In prefixes
        For index = 1 To fn_Count(mappingPrefix & ".count")
            header = fn_Value(mappingPrefix & "." & index & ".header")
            field = fn_Value(mappingPrefix & "." & index & ".field")
            allowed(header) = True
            If field = "MovementBasis" And Not texts.Exists(header) Then
                Err.Raise 5, , "Missing required Movement column text: " & header
            End If
        Next index
    Next mappingPrefix
    For Each header In texts.Keys
        If Not allowed.Exists(header) Then
            Err.Raise 5, , "Text column is not writable by this Movement operation: " & header
        End If
    Next header
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_EnsureLoaded()
    Dim identifier As Variant
    Dim knownNames As Object
    Dim statusName As String
    Dim message As String

    If Not m_values Is Nothing Then
        Exit Sub
    End If
    On Error GoTo EH
    Set m_values = VBA.CreateObject("Scripting.Dictionary")
    Set m_events = VBA.CreateObject("Scripting.Dictionary")
    Set m_statuses = VBA.CreateObject("Scripting.Dictionary")
    Set knownNames = VBA.CreateObject("Scripting.Dictionary")
    m_formFields = "|" & ex_PEB_Export.fn_Config("export.fields") & "|"
    For Each identifier In VBA.Array("HospitalTreatment", "AnnualVacation", "MedicalLeave", "FamilyVacation", "MaternityLeave", "BusinessTrip", "OutpatientVlk", "ExternalVlk", "Awol")
        statusName = ex_PEB_Export.fn_Config("export.Movement.status." & identifier)
        If knownNames.Exists(statusName) Then
            Err.Raise 5, , "Duplicate Movement status name: " & statusName
        End If
        knownNames.Add statusName, True
        m_statuses.Add VBA.CStr(identifier), statusName
    Next identifier
    private_BuildProfiles
    private_BuildEvents
    Exit Sub
EH:
    message = VBA.Err.Description
    fn_Module_Dispose
    Err.Raise 5, , "Native Movement plan: " & message
End Sub

Private Sub private_BuildProfiles()
    Dim fields As Variant
    Dim headers As Variant

    private_AddProfile "Arrival", "close", ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalDate"), "", "EventDate", ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate")
    private_AddMappings "profiles.Arrival.join", _
        VBA.Array( _
            "PersonId", _
            "EventName" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName") _
        )
    private_AddMappings "profiles.Arrival.fields", _
        VBA.Array( _
            "EventDate", _
            "AcceptedOrderNumber", _
            "MovementBasis", _
            "FoodSupportDate" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalOrder"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalBasis"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalFoodSupport") _
        )
    private_AddMappings "profiles.Arrival.identity", _
        VBA.Array( _
            "PersonId", _
            "EventName", _
            "EventDate", _
            "AcceptedOrderNumber" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalOrder") _
        )
    private_AddProfile "Departure", "append", "", "", "", ""
    fields = VBA.Array( _
            "PersonId", _
            "PersonName", _
            "PersonRank", _
            "PositionCode", _
            "MovementDurationDays", _
            "MovementRoadDays", _
            "EventName", _
            "EventDate", _
            "AcceptedOrderNumber", _
            "Destination", _
            "MovementBasis", _
            "TicketDocument", _
            "FoodSupportDate" _
        )
    headers = VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonRank"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.PositionCode"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DurationDays"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.RoadDays"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureOrder"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.Destination"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureBasis"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.TicketDocument"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureFoodSupport") _
        )
    private_AddMappings "profiles.Departure.fields", fields, headers
    private_AddMappings "profiles.Departure.identity", _
        VBA.Array( _
            "PersonId", _
            "EventName", _
            "EventDate", _
            "AcceptedOrderNumber" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureOrder") _
        )
    m_values.Add "profiles.Departure.guard.openHeader", ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalDate")
    private_AddMappings "profiles.Departure.guard.join", _
        VBA.Array("PersonId"), _
        VBA.Array(ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"))
    private_AddProfile "Transition", "transition", ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalDate"), "Arrival", "EventDate", ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate")
    private_AddMappings "profiles.Transition.join", _
        VBA.Array( _
            "PersonId", _
            "SourceEventName" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName") _
        )
    fields = VBA.Array( _
            "PersonId", _
            "PersonName", _
            "PersonRank", _
            "PositionCode", _
            "MovementDurationDays", _
            "MovementRoadDays", _
            "EventName", _
            "EventDate", _
            "AcceptedOrderNumber", _
            "Destination", _
            "MovementBasis", _
            "TicketDocument", _
            "FoodSupportDate" _
        )
    headers = VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonRank"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.PositionCode"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DurationDays"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.RoadDays"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureOrder"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.Destination"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureBasis"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.TicketDocument"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureFoodSupport") _
        )
    private_AddMappings "profiles.Transition.fields", fields, headers
    private_AddMappings "profiles.Transition.identity", _
        VBA.Array( _
            "PersonId", _
            "EventName", _
            "EventDate", _
            "AcceptedOrderNumber" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureOrder") _
        )
    private_AddProfile "ArrivalMedicalCompany", "close", ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalDate"), "", "EventDate", ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate")
    private_AddMappings "profiles.ArrivalMedicalCompany.join", _
        VBA.Array( _
            "PersonId", _
            "EventName", _
            "SourceDestination" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.Destination") _
        )
    private_AddMappings "profiles.ArrivalMedicalCompany.fields", _
        VBA.Array( _
            "EventDate", _
            "AcceptedOrderNumber", _
            "MovementBasis", _
            "FoodSupportDate" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalOrder"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalBasis"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalFoodSupport") _
        )
    private_AddMappings "profiles.ArrivalMedicalCompany.identity", _
        VBA.Array( _
            "PersonId", _
            "EventName", _
            "EventDate", _
            "AcceptedOrderNumber" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalOrder") _
        )
    private_AddProfile "TransitionMedicalCompany", "transition", ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalDate"), "Arrival", "EventDate", ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate")
    private_AddMappings "profiles.TransitionMedicalCompany.join", _
        VBA.Array( _
            "PersonId", _
            "SourceEventName", _
            "SourceDestination" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.Destination") _
        )
    fields = VBA.Array( _
            "PersonId", _
            "PersonName", _
            "PersonRank", _
            "PositionCode", _
            "MovementDurationDays", _
            "MovementRoadDays", _
            "EventName", _
            "EventDate", _
            "AcceptedOrderNumber", _
            "Destination", _
            "MovementBasis", _
            "TicketDocument", _
            "FoodSupportDate" _
        )
    headers = VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonRank"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.PositionCode"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DurationDays"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.RoadDays"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureOrder"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.Destination"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureBasis"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.TicketDocument"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureFoodSupport") _
        )
    private_AddMappings "profiles.TransitionMedicalCompany.fields", fields, headers
    private_AddMappings "profiles.TransitionMedicalCompany.identity", _
        VBA.Array( _
            "PersonId", _
            "EventName", _
            "EventDate", _
            "AcceptedOrderNumber" _
        ), _
        VBA.Array( _
            ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.EventName"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureOrder") _
        )
End Sub

Private Sub private_BuildEvents()
    private_AddEvent _
        "FromBusinessTrip", _
        "Arrival", _
        "Arrival", _
        "BusinessTrip", _
        "", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "FromHospital", _
        "Arrival", _
        "Arrival", _
        "HospitalTreatment", _
        "", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "FromVacationForCuring", _
        "Arrival", _
        "Arrival", _
        "MedicalLeave", _
        "", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "FromAnnualMainVacation", _
        "Arrival", _
        "Arrival", _
        "AnnualVacation", _
        "", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "FromVacationForFamily", _
        "Arrival", _
        "Arrival", _
        "FamilyVacation", _
        "", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "FromOutpatientExaminationOrVlk", _
        "Arrival", _
        "Arrival", _
        "OutpatientVlk", _
        "", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "FromExternalVlk", _
        "Arrival", _
        "Arrival", _
        "ExternalVlk", _
        "", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "FromMedicalCompanyForTreatment", _
        "Arrival", _
        "ArrivalMedicalCompany", _
        "HospitalTreatment", _
        "", _
        "Destination", _
        ex_PEB_Export.fn_Config("export.Movement.destinations.MedicalCompany"), _
        ex_PEB_Export.fn_Config("export.Movement.destinations.MedicalCompany"), _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "ToHospital", _
        "Departure", _
        "Departure", _
        "", _
        "HospitalTreatment", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "ToAnnualMainVacation", _
        "Departure", _
        "Departure", _
        "", _
        "AnnualVacation", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "ToVacationForFamily", _
        "Departure", _
        "Departure", _
        "", _
        "FamilyVacation", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "ToMaternityLeave", _
        "Departure", _
        "Departure", _
        "", _
        "MaternityLeave", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "ToVacationForCuring", _
        "Departure", _
        "Departure", _
        "", _
        "MedicalLeave", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "ToOutpatientExaminationOrVlk", _
        "Departure", _
        "Departure", _
        "", _
        "OutpatientVlk", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "ToMedicalCompanyForTreatment", _
        "Departure", _
        "Departure", _
        "", _
        "HospitalTreatment", _
        "Destination", _
        ex_PEB_Export.fn_Config("export.Movement.destinations.MedicalCompany"), _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromAwolToBusinessTrip", _
        "Transition", _
        "Transition", _
        "Awol", _
        "BusinessTrip", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|Destination|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromAnnualVacationToVacationForCuring", _
        "Transition", _
        "Transition", _
        "AnnualVacation", _
        "MedicalLeave", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|Destination|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromFamilyVacationToVacationForCuring", _
        "Transition", _
        "Transition", _
        "FamilyVacation", _
        "MedicalLeave", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|Destination|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromVacationForCuringToAnnualVacation", _
        "Transition", _
        "Transition", _
        "MedicalLeave", _
        "AnnualVacation", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|Destination|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromHospitalToVacationForCuring", _
        "Transition", _
        "Transition", _
        "HospitalTreatment", _
        "MedicalLeave", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromVacationForCuringToVacationForCuring", _
        "Transition", _
        "Transition", _
        "MedicalLeave", _
        "MedicalLeave", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromVacationForCuringToHospital", _
        "Transition", _
        "Transition", _
        "MedicalLeave", _
        "HospitalTreatment", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromAnnualVacationToHospital", _
        "Transition", _
        "Transition", _
        "AnnualVacation", _
        "HospitalTreatment", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromVacationForCuringToVlk", _
        "Transition", _
        "Transition", _
        "MedicalLeave", _
        "ExternalVlk", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromAnnualVacationToVlk", _
        "Transition", _
        "Transition", _
        "AnnualVacation", _
        "ExternalVlk", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromFamilyVacationToVlk", _
        "Transition", _
        "Transition", _
        "FamilyVacation", _
        "ExternalVlk", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromVlkToVacationForCuring", _
        "Transition", _
        "Transition", _
        "ExternalVlk", _
        "MedicalLeave", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromVlkToHospital", _
        "Transition", _
        "Transition", _
        "ExternalVlk", _
        "HospitalTreatment", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromMedicalCompanyForTreatmentToHospital", _
        "Transition", _
        "TransitionMedicalCompany", _
        "HospitalTreatment", _
        "HospitalTreatment", _
        "HospitalShortName", _
        "", _
        ex_PEB_Export.fn_Config("export.Movement.destinations.MedicalCompany"), _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromMedicalCompanyForTreatmentToVacationForCuring", _
        "Transition", _
        "TransitionMedicalCompany", _
        "HospitalTreatment", _
        "MedicalLeave", _
        "Destination", _
        "", _
        ex_PEB_Export.fn_Config("export.Movement.destinations.MedicalCompany"), _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromMedicalCompanyVacationForCuringToHospital", _
        "Transition", _
        "TransitionMedicalCompany", _
        "MedicalLeave", _
        "HospitalTreatment", _
        "HospitalShortName", _
        "", _
        ex_PEB_Export.fn_Config("export.Movement.destinations.MedicalCompany"), _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromMedicalCompanyVacationForCuringToVacationForCuring", _
        "Transition", _
        "TransitionMedicalCompany", _
        "MedicalLeave", _
        "MedicalLeave", _
        "Destination", _
        "", _
        ex_PEB_Export.fn_Config("export.Movement.destinations.MedicalCompany"), _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromAnnualVacationToFamilyVacation", _
        "Transition", _
        "Transition", _
        "AnnualVacation", _
        "FamilyVacation", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromFamilyVacationToAnnualVacation", _
        "Transition", _
        "Transition", _
        "FamilyVacation", _
        "AnnualVacation", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromHospitalToExternalVlk", _
        "Transition", _
        "Transition", _
        "HospitalTreatment", _
        "ExternalVlk", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromFamilyVacationToHospital", _
        "Transition", _
        "Transition", _
        "FamilyVacation", _
        "HospitalTreatment", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromExternalVlkToVacationForCuring", _
        "Transition", _
        "Transition", _
        "ExternalVlk", _
        "MedicalLeave", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromExternalVlkToHospital", _
        "Transition", _
        "Transition", _
        "ExternalVlk", _
        "HospitalTreatment", _
        "HospitalShortName", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|HospitalShortName|PersonRank|PositionCode|IncomingNumber|IncomingDate"
    private_AddEvent _
        "MoveFromOutpatientVlkToVacationForCuring", _
        "Transition", _
        "Transition", _
        "OutpatientVlk", _
        "MedicalLeave", _
        "Destination", _
        "", _
        "", _
        "PersonId|PersonName|EventDate|AcceptedOrderNumber|AcceptedOrderDate|PersonRank|PositionCode|IncomingNumber|IncomingDate"
End Sub

Private Sub private_AddProfile( _
    ByVal name As String, _
    ByVal operation As String, _
    ByVal openHeader As String, _
    ByVal closeProfile As String, _
    ByVal arrivalDateField As String, _
    ByVal departureDateHeader As String _
)
    Dim prefix As String

    prefix = "profiles." & name & "."
    m_values.Add prefix & "operation", operation
    If VBA.Len(openHeader) > 0 Then
        m_values.Add prefix & "openHeader", openHeader
    End If
    If VBA.Len(closeProfile) > 0 Then
        m_values.Add prefix & "closeProfile", closeProfile
    End If
    If VBA.Len(arrivalDateField) > 0 Then
        m_values.Add prefix & "arrivalDateField", arrivalDateField
    End If
    If VBA.Len(departureDateHeader) > 0 Then
        m_values.Add prefix & "departureDateHeader", departureDateHeader
    End If
End Sub

Private Function private_ExportValue(ByVal value As Variant) As Variant
    private_ExportValue = VBA.vbNullString
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    Select Case VBA.VarType(value)
        Case VBA.vbBoolean
            If Not value Then Exit Function
        Case VBA.vbString
            If VBA.Len(VBA.Trim$(value)) = 0 Then Exit Function
            If VBA.IsNumeric(value) Then
                If VBA.CDbl(value) = 0 Then Exit Function
            End If
        Case VBA.vbByte, VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
            If value = 0 Then Exit Function
    End Select
    private_ExportValue = value
End Function

Private Sub private_AddMappings( _
    ByVal prefix As String, _
    ByVal fields As Variant, _
    ByVal headers As Variant _
)
    Dim index As Long

    If UBound(fields) <> UBound(headers) Then
        Err.Raise 5, , "Invalid native Movement mapping: " & prefix
    End If
    m_values.Add prefix & ".count", VBA.CStr(UBound(fields) + 1)
    For index = 0 To UBound(fields)
        m_values.Add prefix & "." & (index + 1) & ".field", fields(index)
        m_values.Add prefix & "." & (index + 1) & ".header", headers(index)
    Next index
End Sub

Private Sub private_AddEvent( _
    ByVal eventId As String, _
    ByVal operation As String, _
    ByVal profile As String, _
    ByVal sourceStatus As String, _
    ByVal targetStatus As String, _
    ByVal destinationField As String, _
    ByVal destinationValue As String, _
    ByVal sourceDestination As String, _
    ByVal required As String _
)
    Dim plan As Object
    Dim field As Variant
    Dim expected As String

    Select Case operation
        Case "Arrival"
            expected = "close"
        Case "Departure"
            expected = "append"
        Case "Transition"
            expected = "transition"
        Case Else
            Err.Raise 5, , "Invalid native Movement operation: " & eventId
    End Select
    If m_values("profiles." & profile & ".operation") <> expected Then
        Err.Raise 5, , "Invalid native Movement profile: " & eventId
    End If
    If operation <> "Departure" And VBA.Len(sourceStatus) = 0 Then
        Err.Raise 5, , "Native source status is missing: " & eventId
    End If
    If operation <> "Arrival" And VBA.Len(targetStatus) = 0 Then
        Err.Raise 5, , "Native target status is missing: " & eventId
    End If
    If operation = "Departure" And VBA.Len(sourceStatus) > 0 Then
        Err.Raise 5, , "Native departure declares a source: " & eventId
    End If
    If operation = "Arrival" And VBA.Len(targetStatus) > 0 Then
        Err.Raise 5, , "Native arrival declares a target: " & eventId
    End If
    For Each field In VBA.Split(required & "|" & destinationField, "|")
        If VBA.InStr(1, m_formFields, "|" & field & "|", VBA.vbBinaryCompare) = 0 Then
            Err.Raise 5, , "Unknown native Movement input: " & field
        End If
    Next field
    If VBA.Len(sourceStatus) > 0 Then
        If Not m_statuses.Exists(sourceStatus) Then
            Err.Raise 5, , "Unknown native source status: " & sourceStatus
        End If
    End If
    If VBA.Len(targetStatus) > 0 Then
        If Not m_statuses.Exists(targetStatus) Then
            Err.Raise 5, , "Unknown native target status: " & targetStatus
        End If
    End If
    Set plan = VBA.CreateObject("Scripting.Dictionary")
    plan.Add "operation", operation
    plan.Add "profile", profile
    plan.Add "sourceStatus", sourceStatus
    plan.Add "targetStatus", targetStatus
    plan.Add "destinationField", destinationField
    plan.Add "destinationValue", destinationValue
    plan.Add "sourceDestination", sourceDestination
    plan.Add "required", required
    m_events.Add eventId, plan
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------