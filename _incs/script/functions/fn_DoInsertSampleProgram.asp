<script runat="server" type="text/vbscript" language="vbscript">

Sub DoInsertSampleProgram(clientId, sampleProgramName, sampleMember1, sampleMember2, sampleMember3, outError)
	Dim tempLastName
	Dim count
		
	Dim programDesc
	programDesc = programDesc & "This is a program that was created for you automatically by " & Application.Value("APPLICATION_NAME") & ". "
	programDesc = programDesc & "It's members, skills, and events are made up. "
	programDesc = programDesc & "You are free to modify this program to use with your own account members or delete it at any time. "

	Dim scheduleDesc
	scheduleDesc = scheduleDesc & "This is just a schedule that belongs to the sample program (" & sampleProgramName & ") "
	scheduleDesc = scheduleDesc & "that was created for you automatically when you created your " & Application.Value("APPLICATION_NAME") & " administrative account. "
	scheduleDesc = scheduleDesc & "You can safely modify or remove this schedule (and the events that belong to it) at any time. "

	Dim eventDesc
	eventDesc = eventDesc & "This is an event that belongs to the sample program (" & sampleProgramName & ") "
	eventDesc = eventDesc & "that was created for you automatically when you created your " & Application.Value("APPLICATION_NAME") & " administrative account. "
	eventDesc = eventDesc & "You can safely modify or remove this event at any time. "

	Dim program						: Set program = New cProgram
	program.ProgramName	= sampleProgramName
	program.ClientId = clientId
	program.IsEnabled = 1
	program.ProgramDesc = programDesc
	Call program.Add(outError)
	
	If outError = DB_ERROR_DUPLICATE_PROGRAM_NAME Then
		count = 0
		Do While outError = DB_ERROR_DUPLICATE_PROGRAM_NAME
			count = count + 1
			program.ProgramName = sampleProgramName & " (" & count & ")"
			Call program.Add(outError)
		Loop	
	End If
	
	Dim schedule					: Set schedule = New cSchedule
	schedule.ScheduleName = "A Sample Schedule"
	schedule.ProgramId = program.ProgramId
	schedule.ScheduleDesc = scheduleDesc
	schedule.Add(outError)
	
	sampleMember1.ClientId = clientId
	sampleMember1.Email = sampleMember1.NameFirst & "@example.com"
	Call sampleMember1.QuickAdd(program.ProgramId, outError)
	
	If outError = DB_ERROR_DUPLICATE_MEMBER Then
		count = 0
		tempLastName = sampleMember1.NameLast
		Do While outError = DB_ERROR_DUPLICATE_MEMBER
			count = count + 1
			sampleMember1.NameLast = tempLastName & "(" & count & ")"
			Call sampleMember1.Add(outError)
		Loop
	End If
	
	sampleMember2.ClientId = clientId
	sampleMember2.Email = sampleMember2.NameFirst & "@example.com"
	Call sampleMember2.QuickAdd(program.ProgramId, outError)
	
	If outError = DB_ERROR_DUPLICATE_MEMBER Then
		count = 0
		tempLastName = sampleMember2.NameLast
		Do While outError = DB_ERROR_DUPLICATE_MEMBER
			count = count + 1
			sampleMember2.NameLast = tempLastName & "(" & count & ")"
			Call sampleMember2.Add(outError)
		Loop
	End If
	
	sampleMember3.ClientId = clientId
	sampleMember3.Email = sampleMember3.NameFirst & "@example.com"
	Call sampleMember3.QuickAdd(program.ProgramId, outError)
	
	If outError = DB_ERROR_DUPLICATE_MEMBER Then
		count = 0
		tempLastName = sampleMember3.NameLast
		Do While outError = DB_ERROR_DUPLICATE_MEMBER
			count = count + 1
			sampleMember3.NameLast = tempLastName & "(" & count & ")"
			Call sampleMember3.Add(outError)
		Loop
	End If
	
	Dim programMemberId1
	Dim programMemberId2
	Dim programMemberId3
	
	Dim programMember			: Set programMember = New cProgramMember
	programMember.ProgramId = program.ProgramId

	programMember.MemberId = sampleMember1.MemberId
	Call programMember.LoadByMemberProgram()
	programMemberId1 = programMember.ProgramMemberId

	programMember.MemberId = sampleMember2.MemberId
	Call programMember.LoadByMemberProgram()
	programMemberId2 = programMember.ProgramMemberId

	programMember.MemberId = sampleMember3.MemberId
	Call programMember.LoadByMemberProgram()
	programMemberId3 = programMember.ProgramMemberId
	
	Dim i
	Dim list					: list = Split("Electric Guitar,Acoustic Guitar,Drums,Bass,Keyboard,Piano,Vocal,Worship Leader", ",")
	Dim programMemberSkill		: Set programMemberSkill = New cProgramMemberSkill
	Dim skill					: Set skill = New cSkill
	
	Dim programMemberSkillDictionary1			: Set programMemberSkillDictionary1 = Server.CreateObject("Scripting.Dictionary")
	Dim programMemberSkillDictionary2			: Set programMemberSkillDictionary2 = Server.CreateObject("Scripting.Dictionary")
	Dim programMemberSkillDictionary3			: Set programMemberSkillDictionary3= Server.CreateObject("Scripting.Dictionary")
	
	For i = 0 To UBound(list)
		skill.ProgramId = program.ProgramId
		skill.SkillName = list(i)
		skill.IsEnabled = 1
		Call skill.Add(outError)
		
		programMemberSkill.ProgramMemberId = programMemberId1
		programMemberSkill.SkillId = skill.SkillId
		programMemberSkill.Add(outError)
		Call programMemberSkillDictionary1.Add(list(i), programMemberSkill.ProgramMemberSkillId)
		
		programMemberSkill.ProgramMemberId = programMemberId2
		programMemberSkill.SkillId = skill.SkillId
		programMemberSkill.Add(outError)
		Call programMemberSkillDictionary2.Add(list(i), programMemberSkill.ProgramMemberSkillId)
		
		programMemberSkill.ProgramMemberId = programMemberId3
		programMemberSkill.SkillId = skill.SkillId
		programMemberSkill.Add(outError)
		Call programMemberSkillDictionary3.Add(list(i), programMemberSkill.ProgramMemberSkillId)
	Next

	Dim eventDate				: eventDate = DateAdd("ww", 2, Date())
	Dim evnt					: Set evnt = New cEvent
	
	Dim scheduleBuild			: Set scheduleBuild = New cScheduleBuild
	scheduleBuild.PublishStatus = IS_PUBLISHED
	
	evnt.ScheduleId = schedule.ScheduleId
	evnt.EventName = "Worship Service"
	evnt.EventNote = eventDesc
	evnt.EventDate = eventDate
	evnt.TimeStart = "7:00 PM"
	
	' first event
	evnt.Add(outError)
	
	scheduleBuild.EventId = evnt.EventId
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Electric Guitar")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Worship Leader")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Vocal")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary2.Item("Drums")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary3.Item("Bass")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary3.Item("Vocal")
	scheduleBuild.Add(outError)
	
	' second event
	eventDate = DateAdd("ww", 1, eventDate)
	evnt.EventDate = eventDate
	evnt.Add(outError)
	
	scheduleBuild.EventId = evnt.EventId
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Acoustic Guitar")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary2.Item("Worship Leader")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Vocal")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary2.Item("Keyboard")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary3.Item("Bass")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary3.Item("Vocal")
	scheduleBuild.Add(outError)

	' third event
	eventDate = DateAdd("ww", 1, eventDate)
	evnt.EventDate = eventDate
	evnt.Add(outError)
	
	scheduleBuild.EventId = evnt.EventId
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Electric Guitar")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Worship Leader")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Vocal")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary2.Item("Drums")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary3.Item("Bass")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary3.Item("Vocal")
	scheduleBuild.Add(outError)
	
	' fourth event
	eventDate = DateAdd("ww", 1, eventDate)
	evnt.EventDate = eventDate
	evnt.Add(outError)
	
	scheduleBuild.EventId = evnt.EventId
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Acoustic Guitar")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Worship Leader")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary1.Item("Vocal")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary2.Item("Vocal")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary2.Item("Acoustic Guitar")
	scheduleBuild.Add(outError)
	scheduleBuild.ProgramMemberSkillId = programMemberSkillDictionary3.Item("Vocal")
	scheduleBuild.Add(outError)
End Sub

</script>