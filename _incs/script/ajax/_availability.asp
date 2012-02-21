<%
Option Explicit

Call Wait(.5)
Call Main()

Sub Main
	Dim str
	
	Dim action					: action = Request.Form("act")
	Dim memberId				: memberId = Request.Form("mid")
	Dim scheduleId				: scheduleId = Request.Form("scid")
	Dim show_past_events		: show_past_events = Request.Form("show_past_events")
	
	Select Case action
		Case Else
			str = ScheduledEventListToString(memberId, scheduleId, show_past_events)
	End Select
	
	Response.Write str
End Sub

Function HasPublishedSkillForEvent(list, eventId, memberId)
	Dim i
	
	Dim isThisEvent
	Dim isThisMember
	Dim isPublished
	
	HasPublishedSkillForEvent = False
	
	' 0-EventID 1-EventName 2-EventDate 3-TimeStart 4-TimeEnd 
	' 5-MemberId 19-PublishStatus
	
	For i = 0 To UBound(list,2)
		isThisMember = False		: If CStr(memberId & "") = CStr(list(5,i) & "") Then isThisMember = True
		isThisEvent = False			: If CStr(eventId & "") = CStr(list(0,i) & "") Then isThisEvent = True
		isPublished = True		: If list(19,i) = IS_MARKED_FOR_UNPUBLISH Then isPublished = False
		
		If isThisMember And isThisEvent And isPublished Then
			HasPublishedSkillForEvent = True
			Exit For
		End If
	Next
End Function

Function AvailabilityNoteListToString(memberId, scheduleId, show_past_events)
	Dim str, i
	
	Dim schedule		: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	If Len(schedule.ScheduleId) > 0 Then Call schedule.Load()
	
	Dim member			: Set member = New cMember
	member.MemberId = memberId
	
	Dim list
		If Len(show_past_events & "") > 0 Then 
			list = member.EventList(Null, schedule.ProgramId, schedule.ScheduleId)
		Else
			list = member.EventList(Now(), schedule.ProgramId, schedule.ScheduleId)
		End If
	
	Dim hasNote
	
	If Not IsArray(list) Then Exit Function
	
	' 0-EventID 1-EventName 2-EventNote 3-EventDate 4-TimeStart 5-TimeEnd 6-ClientID
	' 7-MemberID 8-ProgramName 9-ProgramID 10-ScheduleID 11-ScheduleName 12-ScheduleDesc
	' 13-ScheduleIsVisible 14-IsScheduled 15-SkillListXmlFrag 16-FileListXmlFrag 17-IsAvailable
	' 18-EventAvailabilityID 19-ProgramIsEnabled 20-ProgramMemberIsActive 
	' 21-AvailabilityViewedByMember 22-MemberActiveStatus 23-AvailabilityNote 
	' 24-EventAvailabilityDateModified
	
	For i = 0 To UBound(list,2)
		hasNote	= False				: If Len(list(23,i) & "") > 0 Then hasNote = True
		
		If hasNote Then
			str = str & "<li class=""note"">"
			str = str & "<h5>" & Server.HTMLEncode(list(1,i)) 
			str = str & "<br /><span>" & MonthName(Month(list(3,i)), True) & " " & Day(List(3,i)) & " " & Year(list(3,i)) & "</span></h5>" 
			str = str & Server.HTMLEncode(list(23,i)) & "</li>"
		End If
	Next
	
	AvailabilityNoteListToString = str
End Function

Function ScheduledEventListToString(memberId, scheduleId, show_past_events)
	Dim str, i, j
	Dim count					: count = 0
	
	Dim noMemberListItem			: noMemberListItem			= "<li class=""error"">Click on a member name in the listing to see a list of their events ..</li>"
	Dim noEventsReturnedListItem	: noEventsReturnedListItem	= "<li class=""error"">No events were returned ..</li>"
	
	If Len(scheduleId & "") = 0 Then 
		ScheduledEventListToString = noEventsReturnedListItem
		Exit Function
	End If
	If Len(memberId & "") = 0 Then 
		ScheduledEventListToString = noMemberListItem
		Exit Function
	End If
	
	Dim schedule					: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	If Len(schedule.ScheduleId) > 0 Then Call schedule.Load()
	
	Dim notScheduledListItem
	notScheduledListItem = notScheduledListItem & "<li class=""error"">You selected a member who doesn't belong to any event teams! </li>"
	notScheduledListItem = notScheduledListItem & "<li class=""count"">"
	notScheduledListItem = notScheduledListItem & "<strong>" & Server.HTMLEncode(schedule.ScheduleName) & "</strong>"
	notScheduledListItem = notScheduledListItem & "<br />no events</li>"
	
	Dim scheduleBuild			: Set scheduleBuild = New cScheduleBuild
	Dim teams					: teams = scheduleBuild.TeamList(scheduleId)
	
	Dim eventIdList
	Dim eventIdArray
	
	Dim showPastEvents			: showPastEvents = False
	If Len(show_past_events) > 0 Then showPastEvents = True
	
	Dim isThisMember
	Dim isThisEvent
	Dim isPastEvent
	Dim hasPublishedSkill
	
	Dim dateClass

	If Not IsArray(teams) Then 
		ScheduledEventListToString = noEventsReturnedListItem
		Exit Function
	End If

	' 0-EventID 1-EventName 2-EventDate 3-TimeStart 4-TimeEnd 5-MemberId 6-NameLast 7-NameFirst 
	' 8-MemberIsEnabled 9-ProgramMemberId 10-ProgramMemberIsEnabled 11-IsAvailable 
	' 12-AvailabilityIsViewedByMember 13-SkillGroupId 14-SkillGroupName 15-SkillGroupIsEnabled
	' 16-SkillId 17-SkillName 18-SkillIsEnabled 19-PublishStatus
	
	' build event id list
	For i = 0 To UBound(teams,2)
		isThisMember = False			: If CStr(memberId & "") = CStr(teams(5,i) & "") Then isThisMember = True
		
		If isThisMember Then 
			eventIdList = eventIdList & teams(0,i) & ","
		End If
	Next	
	If Len(eventIdList) > 0 Then eventIdList = Left(eventIdList, Len(eventIdList) - 1)
	eventIdList = RemoveDupesFromStringList(eventIdList)
	
	eventIdArray = Split(eventIdList, ",")
	If Not IsArray(eventIdArray) Then 
		ScheduledEventListToString = noEventsReturnedListItem
		Exit Function
	End If
	
	' spin through teams again with the eventIdArray and generate the list items for ajax request ..
	For i = 0 To UBound(eventIdArray)
	
		For j = 0 To UBound(teams,2)
			isThisEvent = False			: If CStr(eventIdArray(i) & "") = CStr(teams(0,j) & "") Then isThisEvent = True
			isThisMember = False		: If CStr(memberId & "") = CStr(teams(5,j) & "") Then isThisMember = True
			
			isPastEvent = True			: If CDate(Month(teams(2,j)) & "-" & Day(teams(2,j)) & "-" & Year(teams(2,j))) => CDate(Month(Now()) & "-" & Day(Now()) & "-" & Year(Now())) Then isPastEvent = False
			If showPastEvents = True Then isPastEvent = False
			
			dateClass = "date"
			If teams(12,j) = 0 Then
				dateClass = dateClass & " unknown-available"
			Else
				If teams(11,j) = 0 Then
					dateClass = dateClass & " not-available"
				Else
					dateClass = dateClass & " available"
				End If
			End If
			
			If isThisEvent And isThisMember And (Not isPastEvent) Then
				
				' test for a published skill to see if this should be displayed ..
				If HasPublishedSkillForEvent(teams, eventIdArray(i), memberId) Then
					str = str & "<li><div class=""" & dateClass & """>"
					str = str & "<div class=""month"">" & UCase(MonthName(Month(teams(2,j)), True)) & "</div>"
					str = str & "<div class=""day"">" & Day(teams(2,j)) & "</div>"
					str = str & "</div>"
					str = str & "<p>" & Server.HTMLEncode(teams(1,j)) & "</p></li>"
					
					count = count + 1
					
					Exit For
				End If
			End If
		Next
	Next
	
	str = str & AvailabilityNoteListToString(memberId, scheduleId, show_past_events)

	str = str & "<li class=""count"">" 
	str = str & "<strong>" & Server.HTMLEncode(schedule.ScheduleName) & "</strong>"
	str = str & "<br />" & count & " event"
	If count <> 1 Then str = str & "s"
	str = str & "</li>"
	
	If count = 0 Then 
		str = notScheduledListItem
	End If
	
	ScheduledEventListToString = str
End Function

%>
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_build_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_Wait.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_RemoveDupesFromStringList.asp"-->
