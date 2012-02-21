<%
Option Explicit

Call wait(.5)
Call Main()

Sub Main
	Dim str
	Dim rv
	
	
	Dim eventId					: eventId = Request.Form("event_id")					': if Len(eventId) = 0 Then eventId = 19
	Dim skillId					: skillId = Request.Form("skill_id")					': if Len(skillId) = 0 Then skillId = 108
	Dim scheduledIdList			: scheduledIdList = Request.Form("scheduled_id_list")			
	Dim unScheduledIdList		: unScheduledIdList = Request.Form("unscheduled_id_list")
	
	If Len(Request.Form("add_members")) > 0 Then 
		Call InsertScheduleBuild(unscheduledIdList, eventId, rv)
	End If
	If Len(Request.Form("remove_members")) > 0 Then
		Call RemoveScheduleBuild(scheduledIdList, eventId, rv)
	End If
		
	Call OutputToJson(skillId, eventId)
End Sub

Sub InsertScheduleBuild(unScheduledIdList, eventId, outError)
	Dim i 
	Dim cacheError				: cacheError = 0
	outError = 0
	
	If Len(unscheduledIdList) = 0 Then Exit Sub
	Dim scheduleBuild			: Set scheduleBuild = New cScheduleBuild
	Dim unscheduled				: unscheduled = Split(Replace(unscheduledIdList, " ", ""), ",")
	
	If IsArray(unscheduled) Then
		For i = 0 To UBound(unscheduled)
			If Len(unscheduled(i)) > 0 Then
			
				Set scheduleBuild = New cScheduleBuild
				scheduleBuild.EventId = eventId
				scheduleBuild.ProgramMemberSkillId = unscheduled(i)
				Call scheduleBuild.Load()
				
				' check if this row is found
				If Len(scheduleBuild.DateCreated) > 0 Then
				
					' row will be removed on next publish so set as published ..
					If scheduleBuild.PublishStatus = IS_MARKED_FOR_UNPUBLISH Then
						scheduleBuild.PublishStatus = IS_PUBLISHED
						scheduleBuild.Save(cacheError)
						outError = outError + cacheError
					End If
				Else
					' row doesn't exist so add it and set for next publish ..
					scheduleBuild.PublishStatus = IS_MARKED_FOR_PUBLISH
					scheduleBuild.Add(cacheError)
					outError = outError + cacheError				
				End If					
			End If
			
		Next
	End If
End Sub

Sub RemoveScheduleBuild(scheduledIdList, eventId, outError)
	Dim i
	Dim cacheError				: cacheError = 0
	outError = 0
	
	If Len(scheduledIdList) = 0 Then Exit Sub
	Dim scheduleBuild			: Set scheduleBuild = New cScheduleBuild
	Dim scheduled				: scheduled = Split(Replace(scheduledIdList, " ", ""), ",")
	
	If IsArray(scheduled) Then
		For i = 0 To UBound(scheduled)
			If Len(scheduled(i)) > 0 Then
			
				scheduleBuild.EventId = eventId
				scheduleBuild.ProgramMemberSkillId = scheduled(i)
				Call scheduleBuild.Load()
				
				' check if row is found ..
				If Len(scheduleBuild.DateCreated) > 0 Then
				
					' set to remove row on next publish ..
					If scheduleBuild.PublishStatus = IS_PUBLISHED Then
						scheduleBuild.PublishStatus = IS_MARKED_FOR_UNPUBLISH
						scheduleBuild.Save(cacheError)
						outError = outError + cacheError
						
					' remove row as was never published in first place ..
					ElseIf scheduleBuild.PublishStatus = IS_MARKED_FOR_PUBLISH Then
						scheduleBuild.Delete(cacheError)
						outError = outError + cacheError
					End If
				End If
			End If
		Next
	End If	
End Sub

Sub OutputToJson(skillId, eventId)
	Dim str
	
	Dim evnt					: Set evnt = New cEvent
	evnt.EventID = eventId
	Call evnt.Load()
	
	Dim program					: Set program = New cProgram
	program.ProgramID = evnt.ProgramID
	program.Load()
	
	' comma-delim list returned by SetScheduleOptionsList()
	Dim scheduledList			: scheduledList = ""
	
	Dim buildList				: buildList = evnt.ScheduledMemberList()
	Dim memberList				: memberList = evnt.AvailableMemberList()
	
	Dim scheduledOptions
	Dim availableOptions
	Dim notAvailableOptions
	Dim eventItem
	
	Dim scheduled				: scheduled = ""
	Dim available				: available = ""
	Dim notAvailable			: notAvailable = ""
	
	Call SetScheduleOptionList(skillId, buildList, scheduledList, scheduledOptions)
	Call SetUnscheduledOptionLists(skillId, memberList, scheduledList, availableOptions, notAvailableOptions)
	
	scheduled = scheduled & scheduledOptions
	available = available & availableOptions
	notAvailable = notAvailable & notAvailableOptions
	
	eventItem = ScheduleViewItemToString(evnt.eventId, evnt.ProgramId, evnt.ScheduleId, SCHEDULE_ITEM_TYPE_EDITOR)
	
	str = str & "{ "
	str = str & "scheduled: '" & scheduled & "', "
	str = str & "available: '" & available & "', "
	str = str & "notAvailable: '" & notAvailable & "', "
	str = str & "eventItem: '" & eventItem & "' "
	str = str & "}"
	
	Response.Write(str)
End Sub


%>
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/sub_SetScheduledOptionsList.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/sub_SetUnscheduledOptionsLists.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EventTeamToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EventTeamMembersForSkillToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_ScheduleViewItemToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_ScheduleDropdownOptionsToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EventDropdownOptionsToString.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_Wait.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_build_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
