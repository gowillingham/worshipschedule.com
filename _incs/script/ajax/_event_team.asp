<%
Option Explicit

Call wait(.5)
Call Main()

Sub Main
	Dim str
	Dim rv
	
	Dim action				: action = Request.QueryString("action")							
	Dim item_type			: item_type = Request.QueryString("item_type")							
	Dim eventId				: eventId = Request.QueryString("event_id")					
	Dim scheduleId			: scheduleId = Request.QueryString("schedule_id")	
	Dim programId			: programId = Request.QueryString("program_id")				
	Dim memberId			: memberId = Request.QueryString("member_id")						
	Dim copyToEventId		: copyToEventId = Request.QueryString("to_event_id")				
	Dim copyFromEventId		: copyFromEventId = Request.QueryString("from_event_id")			
	Dim removeTeamEventId	: removeTeamEventId = Request.QueryString("remove_team_event_id")
	Dim publishTeamEventId	: publishTeamEventId = Request.QueryString("publish_team_event_id")	
	
	Select Case action
		Case PUBLISH_EVENT
			Call DoPublishEvent(memberId, eventId, rv)
			Call GetScheduleIdProgramIdForEventId(eventId, scheduleId, programId)			
			Call OutPutToJson("", ScheduleViewItemToString(eventId, programId, scheduleId, item_type), "", TeamAccordionToString(eventId))

		Case CLEAR_EVENT_TEAM_FROM_EVENT
			Call DoDeleteEventTeam(eventId, rv)
			Call GetScheduleIdProgramIdForEventId(eventId, scheduleId, programId)			
			Call OutPutToJson("", ScheduleViewItemToString(eventId, programId, scheduleId, item_type), "", TeamAccordionToString(eventId))
			
		Case COPY_EVENT_TEAM_TO_EVENT
			Call CopyEventToTeam(copyFromEventId, copyToEventId, rv)
			Call OutPutToJson("", ScheduleViewItemToString(copyToEventId, programId, scheduleId, SCHEDULE_ITEM_TYPE_COPY_TO), "", "")
			
'		Case RETURN_EVENT_TEAM
'			If Len(copyToEventId) > 0 Then
'				Call OutPutToJson("", ScheduleViewItemToString(copyToEventId, memberId), "", "")
'			Else
'				Call OutPutToJson("", ScheduleViewItemToString(copyFromEventId, memberId), "", "")
'			End If
			
		Case RETURN_SCHEDULE_ITEM
			Call OutputToJson(str, ScheduleViewItemToString(eventId, programId, scheduleId, item_type), "", "")
			
		Case Else
			Call Err.Raise(vbObjectError, "Main()", "Reached uncaught else clause. ")		
	End Select
	
End Sub

Sub GetScheduleIdProgramIdForEventId(eventId, outScheduleId, outProgramId)
	Dim evnt			: Set evnt = New cEvent
	evnt.EventId = eventId
	If Len(evnt.EventId) > 0 Then Call evnt.Load()
	
	outScheduleId = evnt.ScheduleId
	outProgramId = evnt.ProgramId
End Sub

Sub DoPublishEvent(memberId, eventId, outError)
	Dim evnt			: Set evnt = New cEvent
	
	Dim member			: Set member = New cMember
	member.MemberId = memberId
	If Len(member.MemberId) > 0 Then Call member.Load()
	
	Dim memberName		: memberName = member.NameLast & ", " & member.NameFirst
	If Len(memberName) = 2 Then memberName = "Unknown"
	
	evnt.EventID = eventId
	Call evnt.Publish(member.NameLast & ", " & member.NameFirst, outError)
End Sub

Sub DoDeleteEventTeam(eventId, outError)
	Dim scheduleBuild
	
	Set scheduleBuild = New cScheduleBuild
	scheduleBuild.EventID = eventID
	Call scheduleBuild.ClearAllByEventID(outError)
End Sub

Sub CopyEventToTeam(fromEventId, toEventId, outError)
		Dim scheduleBuild			: Set scheduleBuild = New cScheduleBuild
		
		outError = 0
		If Len(toEventID) = 0 Or Len(fromEventID) = 0 Then
			outError = -1 
			Exit Sub
		End If

		outError = 0
		scheduleBuild.EventID = toEventID
		Call scheduleBuild.CopyFromEvent(fromEventID, outError)
		
		Set scheduleBuild = Nothing
End Sub

Function OptionListToString(scheduleId)
	Dim str
	Dim schedule		: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	
	str = str & "<option value="""">" & server.HTMLEncode("< Select event >") & "</option>"	
	str = str & EventDropdownOptionsToString(schedule, "")
	
	OptionListToString = str
End Function

Sub OutputToJson(optionList, scheduleViewItem, eventTeam, teamAccordion)
	Dim str
	
	str = str & "{ "
	str = str & " ""optionList"": """ & optionList & """, "
	str = str & " ""scheduleViewItem"": """ & Replace(scheduleViewItem, Chr(34), "'") & """, "
	str = str & " ""teamAccordion"": """ & Replace(teamAccordion, Chr(34), "'") & """, "
	str = str & " ""eventTeam"": """ & Replace(eventTeam, Chr(34), "'") & """ "
	str = str & "}"
	
	Response.Write(str)
End Sub

%>
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EventTeamToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EventTeamMembersForSkillToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_ScheduleViewItemToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_ScheduleDropdownOptionsToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EventDropdownOptionsToString.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_TeamAccordionToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_AvailableSelectToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_NotAvailableSelectToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_MemberNotesToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_ScheduledSelectToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/sub_SetScheduledOptionsList.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/sub_SetUnscheduledOptionsLists.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_Wait.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_build_cls.asp"-->
