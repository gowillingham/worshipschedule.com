<%
Option Explicit

Call Main()

Sub Main
	Dim str
	
	Dim action					: action = Request.QueryString("act")
	Dim programId				: programId = Request.QueryString("pid")
	Dim emailGroupId			: emailGroupId = Request.QueryString("emgid")
	Dim skillGroupId			: skillGroupId = Request.QueryString("skgid")
	Dim skillId					: skillId = Request.QueryString("skid")
	Dim scheduleId				: scheduleId = Request.QueryString("scid")
	Dim eventId					: eventId = Request.QueryString("eid")
	
	Call Wait(0.25)		' debug: wait half second ..
	
	Select Case action
		Case SMART_GROUP_PROGRAM
			str = ProgramMemberIdListToString(programId)
		Case SMART_GROUP_CUSTOM_GROUP
			str = EmailGroupMemberIdListToString(emailGroupId)
		Case SMART_GROUP_SKILL_UNGROUPED
			str = UngroupedSkillMemberIdListToSTring(programId)
		Case SMART_GROUP_SKILL_GROUP
			str = SkillGroupMemberIdListToString(skillGroupId)
		Case SMART_GROUP_SKILL
			str = SkillMemberIdListToString(skillId)
		Case SMART_GROUP_SCHEDULE_TEAM
			str = ScheduleTeamMemberIdListToString(scheduleId)
		Case SMART_GROUP_SCHEDULE_AVAILABILITY_MISSING
			str = ScheduleMissingAvailabilityMemberIdListToString(scheduleId)
		Case SMART_GROUP_EVENT_TEAM
			str = EventTeamMemberIdListToString(eventId)
		Case SMART_GROUP_EVENT_AVAILABLE
			str = AvailableForEventMemberIdListToString(eventId)
		Case SMART_GROUP_EVENT_NOT_AVAILABLE
			str = NotAvailableForEventMemberIdListToString(eventId)
		Case SMART_GROUP_EVENT_AVAILABILITY_MISSING
			str = MissingAvailabilityForEventMemberIdListToString(eventId)
		Case SMART_GROUP_SKILL_AVAILABILITY_MISSING
			str = MissingAvailabilityForSkillMemberIdListToString(skillId)
		Case SMART_GROUP_PROGRAM_AVAILABILITY_MISSING
			str = MissingAvailabilityForProgramMemberIdListToString(programId)
			
	End Select
	Response.Write str
End Sub

Function MissingAvailabilityForProgramMemberIdListToString(programId)
	Dim str, i
	
	If Len(programId & "") = 0 Then Exit Function
	
	Dim program				: Set program = New cProgram
	program.ProgramId = programId
	
	Dim list				: list = program.MemberList()
	If Not IsArray(list) Then Exit Function
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	Dim isMissingAvailabilityInfo
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 6-ProgramMemberIsEnabled 7-MemberActiveStatus 8-DateCreated 9-DateModified
	' 10-ProgramMemberID 11-IsApproved 12-HasMissingAvailability 13-Email

	For i = 0 To UBound(list,2)
		isMemberEnabled = True					: If list(7,i) = 0 Then isMemberEnabled = False
		isProgramMemberEnabled = True			: If list(6,i) = 0 Then isProgramMemberEnabled = False
		isMissingAvailabilityInfo = True		: If list(12,i) = 0 Then isMissingAvailabilityInfo = False
		
		If IsMemberEnabled And isProgramMemberEnabled And isMissingAvailabilityInfo Then
			str = str & list(0,i) & ","
		End If				
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	MissingAvailabilityForProgramMemberIdListToString = str
End Function

Function MissingAvailabilityForSkillMemberIdListToString(skillId)
	Dim str, i
	
	If Len(skillId & "") = 0 Then Exit Function
	
	Dim skill				: Set skill = New cSkill
	skill.SkillId = skillId
	
	Dim list				: list = skill.MemberList()
	If Not IsArray(list) Then Exit Function
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	Dim isMissingAvailabilityInfo
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberSkillID 5-ProgramMemberID
	' 6-IsApproved 7-ProgramMemberIsActive 8-Email 9-IsMissingAvailabilityInfoForSkill
	
	For i = 0 To UBound(list,2) 
		isMemberEnabled = True					: If list(3,i) = 0 Then isMemberEnabled = False
		isProgramMemberEnabled = True			: If list(7,i) = 0 Then isProgramMemberEnabled = False
		isMissingAvailabilityInfo = True		: If list(9,i) = 0 Then isMissingAvailabilityInfo = False
		
		If IsMemberEnabled And isProgramMemberEnabled And isMissingAvailabilityInfo Then
			str = str & list(0,i) & ","
		End If				
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	MissingAvailabilityForSkillMemberIdListToString = str
End Function

Function MissingAvailabilityForEventMemberIdListToString(eventId)
	Dim str, i
	
	Dim eventAvailability	: Set eventAvailability = New cEventAvailability
	eventAvailability.EventId = eventId
	
	Dim list
	If Len(eventAvailability.EventId) > 0 Then list = eventAvailability.EventAvailabilityList()
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberIsActive 
	' 5-EventAvailabilityID 6-IsAvailable 7-IsViewedByMember 8-DateAvailabilityModified
	
	Dim isViewedByMember
	Dim isMemberEnabled
	Dim isProgramMemberEnabled

	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isViewedByMember = True				: If list(7,i) = 0 Then isViewedByMember = False
		isMemberEnabled = True				: If list(3,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(4,i) = 0 Then isProgramMemberEnabled = False
		
		If Not isViewedByMember And isMemberEnabled And isProgramMemberEnabled Then
			str = str & list(0,i) & ","
		End If
	Next	
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	MissingAvailabilityForEventMemberIdListToString = str
End Function

Function NotAvailableForEventMemberIdListToString(eventId)
	Dim str, i
	
	Dim eventAvailability	: Set eventAvailability = New cEventAvailability
	eventAvailability.EventId = eventId
	
	Dim list
	If Len(eventAvailability.EventId) > 0 Then list = eventAvailability.EventAvailabilityList()
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberIsActive 
	' 5-EventAvailabilityID 6-IsAvailable 7-IsViewedByMember 8-DateAvailabilityModified
	
	Dim isAvailable
	Dim isMemberEnabled
	Dim isProgramMemberEnabled

	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isAvailable = True				: If list(6,i) = 0 Then isAvailable = False
		isMemberEnabled = True				: If list(3,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(4,i) = 0 Then isProgramMemberEnabled = False
		
		If Not isAvailable And isMemberEnabled And isProgramMemberEnabled Then
			str = str & list(0,i) & ","
		End If
	Next	
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	NotAvailableForEventMemberIdListToString = str
End Function

Function AvailableForEventMemberIdListToString(eventId)
	Dim str, i
	
	Dim eventAvailability	: Set eventAvailability = New cEventAvailability
	eventAvailability.EventId = eventId
	
	Dim list
	If Len(eventAvailability.EventId) > 0 Then list = eventAvailability.EventAvailabilityList()
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberIsActive 
	' 5-EventAvailabilityID 6-IsAvailable 7-IsViewedByMember 8-DateAvailabilityModified
	
	Dim isAvailable
	Dim isMemberEnabled
	Dim isProgramMemberEnabled

	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isAvailable = True					: If list(6,i) = 0 Then isAvailable = False
		isMemberEnabled = True				: If list(3,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(4,i) = 0 Then isProgramMemberEnabled = False
		
		If isAvailable And isMemberEnabled And isProgramMemberEnabled Then
			str = str & list(0,i) & ","
		End If
	Next	
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	AvailableForEventMemberIdListToString = str
End Function

Function EventTeamMemberIdListToString(eventId)
	Dim str, i
	
	Dim evnt				: Set evnt = New cEvent
	evnt.EventId = eventId
	
	Dim list
	If Len(evnt.EventId) > 0 Then list = evnt.EventTeamDetailsList()
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-IsMemberEnabled 4-IsProgramMemberEnabled
	' 5-SkillListingXmlFragment 6-IsAvailable 7-IsAvailabilityViewedByMember
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isMemberEnabled = True				: If list(3,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(4,i) = 0 Then isProgramMemberEnabled = False

		If isMemberEnabled And isProgramMemberEnabled Then
			str = str & list(0,i) & ","
		End If
	Next	
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	EventTeamMemberIdListToString = str
End Function

Function ScheduleMissingAvailabilityMemberIdListToString(scheduleId)
	Dim str, i
	
	Dim schedule			: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	
	Dim list
	If Len(schedule.ScheduleId) > 0 Then list = schedule.AvailabilityList()
	
	Dim isMissingAvailability
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-LastLogin 4-Email 5-IsMemberAccountEnabled
	' 6-IsMissingAvailabilityInfo 7-IsProgramMemberEnabled 8-ProgramMemberId
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isMissingAvailability = False		: If list(6,i) = 1 Then isMissingAvailability = True
		isMemberEnabled = True				: If list(5,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(7,i) = 0 Then isProgramMemberEnabled = False
		
		If isMissingAvailability And isMemberEnabled And isProgramMemberEnabled Then
			str = str & list(0,i) & ","
		End If
	Next	
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	ScheduleMissingAvailabilityMemberIdListToString = str
End Function

Function ScheduleTeamMemberIdListToString(scheduleId)
	Dim str, i
	
	Dim schedule			: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	
	Dim list
	If Len(schedule.ScheduleId) > 0 Then list = schedule.EventTeamDetailsList()
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberIsEnabled 5-ProgramMemberIsEnabled 
	' 6-EventId 7-EventName 8-EventDate 9-TimeStart 10-TimeEnd 11-EventNote 12-SkillListXmlFragment
	' 13-FileListXmlFragment 14-ProgramMemberId
			
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isMemberEnabled = True				: If list(4,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(5,i) = 0 Then isProgramMemberEnabled = False
		
		str = str & list(0,i) & ","
	Next	
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	ScheduleTeamMemberIdListToString = str
End Function

Function SkillMemberIdListToString(skillId)
	Dim str, i
	
	Dim skill				: Set skill = New cSkill
	skill.SkillId = skillId
	
	Dim list
	If Len(skill.SkillId) > 0 Then list = skill.MemberList()
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberSkillID 5-ProgramMemberID
	' 6-IsApproved 7-ProgramMemberIsActive 8-Email
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isMemberEnabled = True				: If list(3,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(7,i) = 0 Then isProgramMemberEnabled = False
		
		If isMemberEnabled And isProgramMemberEnabled Then
			str = str & list(0,i) & ","
		End If
	Next	
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	SkillMemberIdListToString = str
End Function

Function UngroupedSkillMemberIdListToSTring(programId)
	Dim str, i
	
	Dim skillGroup				: Set skillGroup = New cSkillGroup
	
	Dim list
	If Len(programId) > 0 Then list = skillGroup.UngroupedSkillMemberList(programId)
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberActiveStatus 5-ProgramMemberID
	' 6-ProgramMemberIsActive 7-SkillListXmlFragment

	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isMemberEnabled = True				: If list(4,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(6,i) = 0 Then isProgramMemberEnabled = False
		
		If isMemberEnabled And isProgramMemberEnabled Then
			str = str & list(0,i) & ","
		End If
	Next	
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	UngroupedSkillMemberIdListToString = str
End Function

Function SkillGroupMemberIdListToString(skillGroupId)
	Dim str, i
	
	Dim skillGroup			: Set skillGroup = New cSkillGroup
	skillGroup.SkillGroupId = skillGroupId
	
	Dim list
	If Len(skillgroup.SkillGroupId) > 0 Then list = skillGroup.MemberList()
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberActiveStatus 5-ProgramMemberID
	' 6-ProgramMemberIsActive 7-SkillListXmlFragment

	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isMemberEnabled = True				: If list(4,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(6,i) = 0 Then isProgramMemberEnabled = False
		
		If isMemberEnabled And isProgramMemberEnabled Then
			str = str & list(0,i) & ","
		End If
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	SkillGroupMemberIdListToString = str
End Function

Function EmailGroupMemberIdListToString(emailGroupId)
	Dim str, i
	
	Dim emailGroup			: Set emailGroup = New cEmailGroup
	emailGroup.emailGroupId = emailGroupId
	
	Dim list
	If Len(emailGroup.EmailGroupId) > 0 Then list = emailGroup.MemberList()
	
	' 0-EmailGroupMemberID 1-EmailGroupId 2-MemberID 3-Email 4-NameLast 
	' 5-NameFirst 6-DateCreated 7-MemberActiveStatus

	Dim isMemberEnabled
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isMemberEnabled = True					: If list(7,i) = 0 Then isMemberEnabled = False
		
		If isMemberEnabled Then
			str = str & list(2,i) & ","
		End If
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	EmailGroupMemberIdListToString = str
End Function

Function ProgramMemberIdListToString(programId)
	Dim str, i
	
	Dim program				: Set program = New cProgram
	program.ProgramId = programId
	
	Dim list				
	If Len(program.ProgramId) > 0 Then list = program.MemberList()
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 6-ProgramMemberIsEnabled 7-MemberActiveStatus 8-DateCreated 9-DateModified
	' 10-ProgramMemberID 11-IsApproved 12-HasMissingAvailability 13-Email

	Dim isProgramMemberEnabled
	Dim isMemberEnabled
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		isMemberEnabled = True				: If list(7,i) = 0 Then isMemberEnabled = False
		isProgramMemberenabled = True		: If list(6,i) = 0 Then isProgramMemberEnabled = False
		
		If isMemberEnabled And isProgramMemberEnabled Then
			str = str & list(0,i) & ","
		End If
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 1)
	
	ProgramMemberIdListToString = str
End Function
%>

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_group_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_group_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_availability_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_Wait.asp"-->
