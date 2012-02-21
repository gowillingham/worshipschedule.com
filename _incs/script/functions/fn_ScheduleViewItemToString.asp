<script runat="server" type="text/vbscript" language="vbscript">

Function ScheduleViewItemToString(eventId, programId, scheduleId, itemType)
	Dim str, i
	
	Dim evnt				: Set evnt = New cEvent
	evnt.EventId = eventId
	If Len(evnt.EventId) > 0 Then Call evnt.Load()
	Dim qs
	
	Dim program				: Set program = New cProgram
	program.ProgramId = programId
	
	Dim schedule			: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	
	Dim item
	Dim itemEventDate
	Dim itemEventName
	Dim itemScheduleName
	Dim itemBackgroundColor
	Dim itemEventTeam
	Dim itemToolbar
	Dim itemScheduleDropdown
	Dim itemEventDropdown
	
	If Len(evnt.EventId) > 0 Then
		itemEventDate					= WeekdayName(Weekday(evnt.EventDate), True) & " " & Day(evnt.EventDate) & " " & MonthName(Month(evnt.EventDate), True)
		itemEventName					= Server.HTMLEncode(evnt.EventName)
		itemScheduleName				= server.HTMLEncode(evnt.ScheduleName)
		itemBackgroundColor				= evnt.HtmlBackgroundColor
		itemEventTeam					= EventTeamToString(eventId)
		
		itemToolbar = itemToolbar & "<div class=""toolbar"">"
		qs = "?eid=" & Encrypt(eventId) & "&amp;act=" & Encrypt(UPDATE_RECORD)
		itemToolbar = itemToolbar & "<a href=""/schedule/teams.asp" & qs & """ title=""Edit Team""><img src=""/_images/icons/pencil.png"" alt="""" /></a>"
		qs = "?eid=" & Encrypt(eventId) & "&amp;act=" & Encrypt(COPY_EVENT_TEAM_TO_EVENT)
		itemToolbar = itemToolbar & "<a href=""/schedule/teams.asp" & qs & """ title=""Copy Team""><img src=""/_images/icons/paste_group.png"" alt="""" /></a>"
		itemToolbar = itemToolbar & "<a href=""#"" class=""ajax-publish-" & eventId & """ title=""Publish""><img src=""/_images/icons/arrow_rotate_clockwise.png"" alt="""" /></a>"
		itemToolbar = itemToolbar & "<a href=""#"" class=""ajax-remove-team-" & eventId & """ title=""Remove team""><img src=""/_images/icons/cross.png"" alt="""" /></a>"
		itemToolbar = itemToolbar & "</div>"
		
		itemScheduleDropdown = "<select name=""schedule_id"" class=""schedule-dropdown"">" 
		itemScheduleDropdown = itemScheduleDropdown & ScheduleDropdownOptionsToString(program, scheduleId) & "</select>"
		
		itemEventDropdown =	"<select name=""event_id"" class=""event-dropdown"">"
		itemEventDropdown = itemEventDropdown & "<option value="""">" & Server.HTMLEncode("Select an event ..") & "</option>"
		itemEventDropdown = itemEventDropdown & EventDropdownOptionsToString(schedule, eventId) & "</select>"
	Else
		If itemType = SCHEDULE_ITEM_TYPE_COPY_TO Then
			itemEventDate = "Copy into .."
			itemEventTeam = "<p class=""alert"">Select an event from the dropdown to copy into ..</p>"
		ElseIf itemType = SCHEDULE_ITEM_TYPE_COPY_FROM Then
			itemEventDate = "Copy from .."
			itemEventTeam = "<p class=""alert"">Select an event from the dropdown to copy from ..</p>"
		Else
			itemEventDate = "Select an event from the dropdown .."
			itemEventTeam = "<p class=""alert"">Select an event from the dropdown ..</p>"
		End If


		itemScheduleDropdown = "<select name=""schedule_id"" class=""schedule-dropdown"">" 
		itemScheduleDropdown = itemScheduleDropdown & ScheduleDropdownOptionsToString(program, scheduleId)
		itemScheduleDropdown = itemScheduleDropdown & "</select>"

		itemEventDropdown =	"<select name=""event_id"" class=""event-dropdown"">"
		itemEventDropdown = itemEventDropdown & "<option value="""">" & Server.HTMLEncode("Select an event ..") & "</option>"
		itemEventDropdown = itemEventDropdown & EventDropdownOptionsToString(schedule, eventId)
		itemEventDropdown = itemEventDropdown & "</select>"
	End If
	
	str = str & "<div class=""event-item"" style=""background-color:" & itemBackgroundColor & ";"">"
	str = str & "<h4>" & itemEventDate & "</h4>"
	If Len(itemScheduleName) > 0 Then str = str & "<h5>" & itemScheduleName & "</h5>"
	If Len(itemEventName) > 0 Then str = str & "<h6>" & itemEventName & "</h6>"
	If Len(itemToolbar) > 0 Then str = str & itemToolbar
	str = str & itemEventTeam

	str = str & "<div class=""bottom"">"
	If (itemType = SCHEDULE_ITEM_TYPE_COPY_TO) Or (itemType = SCHEDULE_ITEM_TYPE_COPY_FROM) Then
		str = str & "<form class=""form-set-schedule-item"" method=""post"" action=""/_incs/script/ajax/_event_team.asp"">"
		str = str & "<input type=""hidden"" name=""program_id"" value=""" & programId & """ />"
		str = str & "<input type=""hidden"" name=""action"" value=""" & RETURN_SCHEDULE_ITEM & """ />"
		str = str & "<input type=""hidden"" name=""item_type"" value=""" & itemType & """ />"
		str = str & "Schedule"
		str = str & itemScheduleDropdown
		str = str & "Event"
		str = str & itemEventDropdown
		str = str & "</form>"
	ElseIf itemType = SCHEDULE_ITEM_TYPE_EDITOR Then
		str = str & "<form class=""form-goto-event-dropdown"" method=""post"" action=""#"">"
		str = str & "<input type=""hidden"" name=""form_go_to_event_dropdown_is_postback"" value=""" & IS_POSTBACK & """ />"
		str = str & "Event"
		str = str & "<select name=""event_id"" class=""event-dropdown"">"
		str = str & EventDropdownOptionsToString(schedule, eventId)
		str = str & "</select>"
		str = str & "</form>"
	Else
		str = str & "&nbsp;"
	End If
	str = str & "</div>"
	
	str = str & "</div>"
		
	ScheduleViewItemToString = str 
End Function

</script>