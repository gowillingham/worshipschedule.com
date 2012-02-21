<%
Option Explicit

Call Main()

Sub Main
	Dim str
	
	Dim action					: action = Request.Form("act")
	Dim eventAvailabilityId		: eventAvailabilityId = Request.Form("eaid")
	Dim availability_note		: availability_note = Request.Form("availability_note")
	
	Call Wait(0.5)		
	
	Select Case action
		Case DELETE_RECORD
			Call DoDeleteAvailabilityNote(eventAvailabilityId, "")
		Case UPDATE_RECORD
			Call DoUpdateAvailabilityNote(eventAvailabilityId, availability_note, str, "")
		Case GET_AVAILABILITY_FORM
			str = FormAvailabilityNoteToString(eventAvailabilityId)    
		Case SET_MEMBER_TO_AVAILABLE
			Call DoUpdateAvailability(eventAvailabilityId, 1, "")
		Case SET_MEMBER_TO_NOT_AVAILABLE
			Call DoUpdateAvailability(eventAvailabilityId, 0, "")
	End Select

	Response.Write str
End Sub

Sub DoDeleteAvailabilityNote(id, outError)
	If Len(id) = 0 Then Exit Sub
	
	Dim eventAvailability			: Set eventAvailability = New cEventAvailability
	
	eventAvailability.EventAvailabilityId = id
	Call eventAvailability.Load()
	
	eventAvailability.MemberNote = ""
	Call eventAvailability.Save(outError)
End Sub

Sub DoUpdateAvailabilityNote(id, val, outHtml, outError)
	If Len(id) = 0 Then Exit Sub
	
	Dim eventAvailability			: Set eventAvailability = New cEventAvailability
	
	eventAvailability.EventAvailabilityId = id
	Call eventAvailability.Load()
	
	eventAvailability.MemberNote = val
	Call eventAvailability.Save(outError)
	
	Call eventAvailability.Load()
	outHtml = AvailabilityNoteToString(eventAvailability.MemberNote, eventAvailability.DateModified)
End Sub

Sub DoUpdateAvailability(id, val, outError)
	If Len(id) = 0 Then Exit Sub
	
	Dim eventAvailability			: Set eventAvailability = New cEventAvailability
	
	eventAvailability.EventAvailabilityId = id
	eventAvailability.Load()
	
	eventAvailability.IsAvailable = val
	eventAvailability.IsViewedByMember = 1
	Call eventAvailability.Save(outError)
End Sub

Function FormAvailabilityNoteToString(id)
	Dim str
	
	Dim eventAvailability			: Set eventAvailability = New cEventAvailability
	eventAvailability.EventAvailabilityId = id
	If Len(eventAvailability.EventAvailabilityId) > 0 Then Call eventAvailability.Load()
	
	str = str & "<p>You may leave a note about your availability with this event. "
	str = str & "The person who does the schedule for this program will see this note when they are assigning members to the event team. </p>"
	str = str & "<form method=""post"" action=""/_incs/script/ajax/_member_events.asp"" id=""form-availability-note"">"
	str = str & "<input type=""hidden"" name=""act"" value=""" & UPDATE_RECORD & """ />"
	str = str & "<input type=""hidden"" name=""eaid"" value=""" & id & """ />"
	str = str & "<textarea name=""availability_note"">" & Server.HTMLEncode(eventAvailability.MemberNote & "") & "</textarea>"
	str = str & "</form>"
	
	FormAvailabilityNoteToString = str
End Function	
%>

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_availability_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_Wait.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_AvailabilityNoteToString.asp"-->
