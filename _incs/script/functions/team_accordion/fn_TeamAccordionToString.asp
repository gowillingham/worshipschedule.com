<%

Function TeamAccordionToString(eventId)
	Dim  str, i
	
	Dim evnt				: Set evnt = New cEvent
	evnt.EventId = eventId
	Call evnt.Load()
	
	Dim program				: Set program = New cProgram
	program.ProgramId = evnt.ProgramId
	
	Dim scheduled			: scheduled = ""
	Dim skills				: skills = program.SkillList("")
	
	Dim scheduledOptions
	Dim availableOptions
	Dim notAvailableOptions
	
	' 0-ProgramMemberSkillID 1-SkillID 2-SkillName 3-SkillIsEnabled 4-SkillGroupID
	' 5-SkillGroupName 6-SkillGroupIsEnabled 7-NameLast 8-NameFirst 9-MemberActiveStatus
	' 10-IsAvailable 11-MemberNote 12-AvailabilityDateModified 13-PublishStatus
	' 14-ProgramMemberIsActive
	Dim buildList			: buildList = evnt.ScheduledMemberList()
	
	' 0-ProgramMemberSkillID 1-NameLast 2-NameFirst 3-MemberEnabled 4-ProgramMemberEnabled 5-SkillName
	' 6-SkillGroupName 7-SkillEnabled 8-SkillGroupEnabled 9-IsAvailable 10-AvailabilityNote 
	' 11-IsViewedByMember 12-DateAvailabilityModified 13-ProgramMemberID 14-MemberID 15-SkillID 
	' 16-SkillGroupID
	Dim memberList			: memberList = evnt.AvailableMemberList()
	
	str = str & "<ul id=""team-editor"">"
	For i = 0 To UBound(skills,2)
		str = str & "<li><h5 class=""head"">" & Server.HtmlEncode(skills(1,i)) & "</h5>"

		' this form will be replaced by ajax call ..
		str = str & "<form method=""post"" action=""/_incs/script/ajax/_update_schedule_view.asp"">"
		str = str & "<input type=""hidden"" name=""skill_id"" value=""" & skills(0,i) & """ />"
		str = str & "<input type=""hidden"" name=""event_id"" value=""" & evnt.EventId & """ />"
		str = str & "<table><tbody>"
		str = str & "<tr class=""header""><td class=""scheduled"">Scheduled</td>"
		str = str & "<td class=""buttons"">&nbsp;</td>"
		str = str & "<td class=""available"">Available</td>"
		str = str & "<td class=""not-available"">Not Available</td></tr>"
		
		Call SetScheduleOptionList(skills(0,i), buildList, scheduled, scheduledOptions)
		Call SetUnscheduledOptionLists(skills(0,i), memberList, scheduled, availableOptions, notAvailableOptions)
		
		str = str & "<tr><td>" & ScheduledSelectToString(scheduledOptions) & "</td>"
		str = str & "<td class=""button-cell""><input type=""submit"" name=""remove_members"" value=""" & Server.HtmlEncode(">>") & """ class=""button"" />"
		str = str & "<br /><input type=""submit"" name=""add_members"" value=""" & Server.HtmlEncode("<<") & """ class=""button"" />"
		str = str & "</td>"
		str = str & "<td>" & AvailableSelectToString(availableOptions) & "</td>"
		str = str & "<td>" & NotAvailableSelectToString(notAvailableOptions) & "</td></tr>"
		str = str & "<tr><td colspan=""4"">" & MemberNotesToString(skills(0,i), memberList) & "</td></tr>"
		str = str & "</tbody></table></form>"
		
		str = str & "</li>"
	Next
	str = str & "</ul>"
	
	TeamAccordionToString = str
End Function

%>