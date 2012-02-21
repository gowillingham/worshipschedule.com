<script runat="server" type="text/vbscript" language="vbscript">

Function EventTeamToString(eventId)
	Dim str, i
	
	Dim evnt				: Set evnt = New cEvent
	evnt.EventId = eventId
	Call evnt.Load()
	
	Dim schedule			: Set schedule = New cSchedule
	schedule.ScheduleId = evnt.ScheduleId
	
	Dim program				: Set program = New cProgram
	program.ProgramId = evnt.ProgramId
	
	Dim skills				: skills = program.SkillList("")
	Dim members				: members = schedule.ScheduleBuildList()
	
	Dim item				: item = ""
	Dim items				: items = ""
	
	Dim isSkillEnabled		: isSkillEnabled = True
	Dim isSkillGroupEnabled	: isSkillGroupEnabled = True
	
	' 0-SkillID 1-SkillName 2-SkillDesc 3-IsSkillEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-IsGroupEnabled 8-LastModified 9-DateCreated
	If IsArray(skills) Then
		For i = 0 To UBound(skills,2)
			isSkillEnabled = True		: If skills(3,i) = 0 Then isSkillEnabled = False
			isSkillGroupEnabled = True	: If skills(7,i) = 0 Then isSkillGroupEnabled = False
			
			If isSkillEnabled And isSkillGroupEnabled Then
				item = EventTeamMembersForSkillToString(members, skills(0,i), eventId)
				If Len(item) > 0 Then items = items & "<li>" & Server.HTMLEncode(skills(1,i)) & item & "</li>"
			End If
		Next
	End If
	If Len(items) = 0 Then
		items = items & "<li class=""alert"">"
		items = items & "<h5>No team members!</h5>"
		items = items & "<p>No members have been assigned to this event team. </p></li>"
	End If
	str = str & "<ul class=""event-team"">" & items & "</ul>"
	
	EventTeamToString = str
End Function

</script>