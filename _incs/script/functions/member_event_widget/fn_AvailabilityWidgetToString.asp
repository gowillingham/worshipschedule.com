<script runat="server" type="text/vbscript" language="vbscript">

Function AvailabilityWidgetToString(scheduleId, showPastEvents)
	Dim str, i
	
	Dim schedule				: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	If Len(schedule.ScheduleId & "") > 0 Then Call schedule.Load()
	
	Dim program					: Set program = New cProgram
	program.ProgramId = schedule.ProgramId
	
	Dim members
	If Len(program.ProgramId & "") > 0 Then members = program.MemberList()
	
	str = str & "<div id=""availability-widget"">"
	str = str & "<h4>Find events for member</h4>"
	
	str = str & "<ul class=""event-list"">"
	str = str & "<li class=""error"">Click on a member name in the listing to see a list of their <span style=""white-space:nowrap;"">events ..</span></li></ul>"
	
	str = str & "<div class=""bottom"">"
	str = str & "<form method=""post"" action=""#"" id=""form-select-member"">"
	str = str & "<input type=""hidden"" name=""act"" value=""get"" />"
	str = str & "<input type=""hidden"" name=""scid"" value=""" & scheduleId & """ id=""schedule-id"" />"
	str = str & "<input type=""hidden"" name=""show_past_events"" value=""" & showPastEvents & """ />"
	str = str & "<select name=""mid"" id=""member-dropdown"">"
	str = str & "<option value="""">" & Server.HTMLEncode(".. or select a member") & "</option>"
	str = str & "<option value=""0"" disabled=""disabled"">" & Server.HTMLEncode("-- ") & "</option>"
	str = str & OptionListForAvailabilityWidgetToString(members)
	str = str & "</select></form></div>"
	str = str & "</div>"	
	
	AvailabilityWidgetToString = str
End Function

</script>