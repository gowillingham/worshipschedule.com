<script type="text/vbscript" runat="server" language="vbscript">

Function EventManagerHeaderItemToString(schedule, outError)
	Dim str, dateTime, events, firstEvent, lastEvent, eventCount
	Set dateTime = New cFormatDate
	
	' get events for this schedule
	events = schedule.EventList("")

	' if no events then return with error ..
	If Not IsArray(events) Then
		outError = -1
		Exit Function
	End If
	
	' get some stats for schedule
	firstEvent = events(3,0)
	lastEvent = events(3, UBound(events,2))
	eventCount = UBound(events,2) + 1
	
	str = str & "<table class=""eventManagerItem"" style=""margin-bottom:25px;background-color:#" & schedule.BackgroundColor & ";border-width:3px;"">"
	str = str & "<tr><td>"
	' scheduleName header
	str = str & "<img src=""/_images/icons/calendar.png"" class=""eventManagerItemIcon"" title=""Schedule"" alt=""Schedule"" />"
	str = str & "<div><strong>" & HTML(schedule.ScheduleName) & "</strong> (" & HTML(schedule.ProgramName) & ")</div>"
	' date range and event count
	str = str & "<div class=""eventDate""> " & dateTime.Convert(firstEvent, "DDD MMM dd, YYYY") & " to " & dateTime.Convert(lastEvent, "DDD MMM dd, YYYY") & "<br />(" & eventCount & " events)</div>"
	' desc
	If Len(schedule.ScheduleDesc) > 0 Then
		str = str & "<div class=""eventNote""><strong>Description: </strong>" & HTML(schedule.ScheduleDesc) & "</div>"
	End If
	
	' schedule is hidden ..
	If CInt(schedule.IsVisible) = 0 Then
		str = str & "<br /><img src=""/_images/icons/monitor_error.png"" title=""Hidden"" alt=""Hidden"" style=""display:inline;"" />&nbsp;&nbsp;<strong>Hidden: </strong>Events for this schedule are hidden and will not appear on the member calendar."
	End If
	
	' publish status ..
	If CInt(schedule.PublishStatus) = 2 Then
		str = str & "<br /><img src=""/_images/icons/group_error.png"" title=""No Schedule"" alt=""No Schedule"" style=""display:inline;"" />&nbsp;&nbsp;<strong>No Members Scheduled: </strong>No members have yet been assigned to any of the events on this schedule."
	ElseIf CInt(schedule.PublishStatus) = 1 Then
		str = str & "<br /><img src=""/_images/icons/group_delete.png"" title=""Unpublished Changes"" alt=""Unpublished"" style=""display:inline;"" />&nbsp;&nbsp;<strong>Not Published: </strong>This schedule has changes that have not yet been published to the member calendar."
	Else
		str = str & "<br /><img src=""/_images/icons/group_add.png"" title=""Published"" alt=""Published"" style=""display:inline;"" />&nbsp;&nbsp;<strong>Published: </strong>This schedule has been published to the member calendar."
	End If
	
	str = str & "</td></tr>"
	str = str & "</table>" 
	
	EventManagerHeaderItemToString = str
	
	Set dateTime = Nothing
End Function

</script>