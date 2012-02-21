<script runat="server" type="text/vbscript" language="vbscript">

Function EventTeamMembersForSkillToString(members, skillId, eventId)
	Dim str, i
	
	If Not IsArray(members) Then Exit Function
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberIsActive 
	' 5-PublishStatus 6-SkillID 7-SkillName 8-SkillIsEnabled 9-SkillGroupID 
	' 10-SkillGroupName 11-SkillGroupIsEnabled 12-ProgramMemberSkillID 13-EventId
	' 14-EventName 15-EventDate 16-TimeStart 17-TimeEnd 18-IsAvailable 
	' 19-IsAvailabilityViewedByMember 20-AvailabilityDateModified 21-ScheduleID
	' 22-ScheduleName
	
	Dim isAccountEnabled		: isAccountEnabled = False
	Dim isProgramMemberEnabled	: isProgramMemberEnabled = False
	Dim isSkillEnabled			: isSkillEnabled = False
	Dim isSkillGroupEnabled		: isSkillGroupEnabled = False
	Dim displayAsPublished		: displayAsPublished = True
	
	Dim cls
	Dim isAvailable
	Dim isAvailabilityViewed
	
	For i = 0 To UBound(members,2)
	
		' test for this skillId, eventId
		If (CStr(members(6,i) & "") = CStr(skillId & "")) And (CLng(members(13,i)) = CLng(eventId)) Then
		
			' test for enabled account, programMember, skill, skillgroup
			isAccountEnabled = False			: If members(3,i) = 1 Then isAccountEnabled = True
			isProgramMemberEnabled = False		: If members(4,i) = 1 Then isProgramMemberEnabled = True
			isSkillEnabled = False				: If members(8,i) = 1 Then isSkillEnabled = True
			isSkillGroupEnabled = False			: If members(11,i) = 1 Then isSkillGroupEnabled = True
			
			' don't display unpublished ..
			displayAsPublished = True			: If members(5,i) = IS_MARKED_FOR_UNPUBLISH Then displayAsPublished = False
			
			isAvailable = True					: If members(18,i) = 0 Then isAvailable = False
			isAvailabilityViewed = True			: If members(19,i) = 0 Then isAvailabilityViewed = False
			If isAvailabilityViewed Then
				If isAvailable Then
					cls = "available"
				Else
					cls = "not-available"
				End If
			Else
				cls = "unknown-available"
			End If
			
			If displayAsPublished And isAccountEnabled And isProgramMemberEnabled And isSkillEnabled And isSkillGroupEnabled Then
				str = str & "<li class=""" & cls & """><a href=""#"" class=""mid-" & members(0,i) & """>" & server.HTMLEncode(members(1,i) & ", " & members(2,i)) & "</a></li>"
			End If
		End If
	Next
	
	If Len(str) > 0 Then str = "<ul class=""team-list-for-skill"">" & str & "</ul>"
	
	EventTeamMembersForSkillToString = str
End Function

</script>