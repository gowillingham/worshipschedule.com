<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Const NUMBER_OF_EVENT_COLUMNS = 5

Dim m_pageTabLocation	: m_pageTabLocation = "admin-schedules"
Dim m_pageHeaderText	: m_pageHeaderText = "&nbsp;"
Dim m_impersonateText	: m_impersonateText = ""
Dim m_pageTitleText		: m_pageTitleText = ""
Dim m_topBarText		: m_topBarText = "&nbsp;"
Dim m_bodyText			: m_bodyText = ""
Dim m_tabStripText		: m_tabStripText = ""
Dim m_tabLinkBarText	: m_tabLinkBarText = ""
Dim m_appMessageText	: m_appMessageText = ""
Dim m_acctExpiresText	: m_acctExpiresText = ""

Sub OnPageLoad(ByRef page)
	Dim sess			: Set sess = New cSession
	sess.SessionID = Request.Cookies("sid")
	Call CheckSession(sess, PERMIT_LEADER)
	
	page.MessageID = Request.QueryString("msgid")
	page.ReturnContext = Request.QueryString("rc")
	page.ShowPastEvents = Request.QueryString("hpe")
	
	page.Month = Request.Querystring("m")
	page.Day = Request.Querystring("d")
	page.Year = Request.Querystring("y")
	
	page.Action = Decrypt(Request.QueryString("act"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.ScheduleID = Decrypt(Request.QueryString("scid"))
	page.FilterScheduleId = Decrypt(Request.QueryString("fscid"))
	page.EventID = Decrypt(Request.QueryString("eid"))
	page.SkillID = Decrypt(Request.QueryString("skid"))
	
	' postbacks ..
	If Request.Form("form_owned_program_dropdwon_is_postback") = IS_POSTBACK Then
		If Len(Request.Form("new_program_id")) > 0 Then 
			page.ProgramId = Request.Form("new_program_id")
			
			' reset scheduleId/skillId
			page.ScheduleId = ""
			page.SkillId = ""
		End If
	End If
	
	If Request.Form("form_schedule_dropdown_is_postback") = IS_POSTBACK Then
		If Len(Request.Form("new_schedule_id")) Then page.ScheduleId = Request.Form("new_schedule_id")
	End If
	
	If Request.Form("form_goto_skill_dropdown_is_postback") = IS_POSTBACK Then
		If Len(Request.Form("new_skill_id")) > 0 Then page.SkillId = Request.Form("new_skill_id")
	End If
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()
	Set page.Schedule = New cSchedule
	page.Schedule.ScheduleID = page.ScheduleID
	If Len(page.Schedule.ScheduleID) > 0 Then page.Schedule.Load()

	' set the view tokens
	m_appMessageText = ApplicationMessageToString(page.MessageID)
	page.MessageID = ""
	Call SetTopBar(page)
	Call SetPageHeader(page)
	Call SetPageTitle(page)
	Call SetTabLinkBar(page)
	Call SetTabList(m_pageTabLocation, page)
	Call SetImpersonateText(sess)
	Call SetAccountNotifier(sess)
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/inside.css" />
		<link rel="stylesheet" type="text/css" href="availability.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<!-- todo: remove this-->
				<script language="javascript" type="text/javascript" src="/_incs/script/jquery/jquery-1.2.6.js"></script>
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="availability.js"></script>
		<title><%=m_pageTitleText%></title>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)
	
	Select Case page.Action
		Case PUBLISH_SCHEDULE
			Call DoPublishSchedule(page.ScheduleId, page.Member.NameLast & ", " & page.Member.NameFirst, rv)
			page.MessageId = 6019: page.Action = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case PUBLISH_PROGRAM
			Call DoPublishEventsForProgramId(page.Program, rv)
			page.MessageId = 6059: page.Action = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))

		Case Else
			str = str & AvailabilityGridToString(page)
			
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoPublishEventsForProgramId(program, outError)
	Call program.PublishEvents(outError)
End Sub

Sub DoPublishSchedule(scheduleId, publisher, outError)
	Dim schedule		: Set schedule = New cSchedule
	
	schedule.ScheduleId = scheduleId
	Call schedule.Publish(publisher, outError)
End Sub

Function GetFirstSkillId(program)
	Dim id, list, i
	
	Dim isSkillEnabled
	Dim isSkillGroupEnabled
	Dim skillHasMembers
	
	' 0-SkillID 1-SkillName 2-SkillDesc 3-IsSkillEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-IsGroupEnabled 8-LastModified 9-DateCreated 10-ActiveMemberCount
		
	If Len(program.ProgramId) = 0 Then Exit Function
	
	list = program.SkillList("SkillName")
	If Not IsArray(list) Then Exit Function
	
	For i = 0 To UBound(list,2)
		isSkillEnabled = True			: If list(3,i) = 0 Then isSkillEnabled = False
		isSkillGroupEnabled = True		: If list(7,i) = 0 Then isSkillGroupEnabled = False
		skillHasMembers = True			: If list(10,i) = 0 Then skillHasMembers = False
		
		If isSkillEnabled And isSkillGroupEnabled And skillHasMembers Then
			id = list(0,i)
			Exit For
		End If
	Next
	
	GetFirstSkillId = id
End Function

Function AvailabilityCellInfoToString(memberId, eventId, list)
	Dim str, i
	
	Dim isThisMemberId
	Dim isThisEventId
	Dim isThisSkillId
	
	Dim isAvailable
	Dim isViewedByMember
	Dim isScheduled
	Dim isPublished
	Dim toBePublished
	Dim toBeRemoved
	Dim checked
	
	Dim availableClass
	Dim availableIndicator
	
	If Not IsArray(list) Then
		Call Err.Raise(vbObjectError + 1, "/schedule/availability.asp: Function AvailabilityCellInfoToString();", "ASSERT: No rows returned for availabilityList [should have been array]. ")
	End If	

	' 0-MemberId 1-NameLast 2-NameFirst 3-MemberIsEnabled 4-ProgramMemberIsEnabled 5-IsAvailable 6-AvailabilityNote
	' 7-AvailabilityIsViewedByMember 8-AvailabilityDateModified 9-PublishStatus 10-EventID 11-EventName 
	' 12-EventDescription 13-EventDate 14-TimeStart 15-TimeEnd 16-ScheduleID 17-ScheduleName 
	' 18-IsVisible 19-EventAvailabilityID

	For i = 0 To UBound(list,2)
		isThisMemberId = False			: If CStr(memberId & "") = CStr(list(0,i) & "") Then isThisMemberId = True
		isThisEventId = False			: If CStr(eventId & "") = CStr(list(10,i) & "") Then isThisEventId = True
		
		If isThisMemberId And isThisEventId Then
			isAvailable = True			: If list(5,i) = 0 Then isAvailable = False
			isViewedByMember = True		: If list(7,i) = 0 Then isViewedByMember = False
				
			isScheduled = False
			isPublished = False
			toBePublished = False
			toBeRemoved = False
			checked = ""
			
			availableIndicator = "&nbsp;"
			
			If Len(list(9,i) & "") > 0 Then 
				isScheduled = True
				
				If list(9,i) = IS_PUBLISHED Then isPublished = True
				If list(9,i) = IS_MARKED_FOR_PUBLISH Then toBePublished = True 
				If list(9,i) = IS_MARKED_FOR_UNPUBLISH Then toBeRemoved = True
			End If
			
			If Not isViewedByMember Then
				availableClass = " class=""unknown"""
			Else
				If Not isAvailable Then 
					availableClass = " class=""not-available"""
				End If
			End If
			
			If isPublished Then 
				availableIndicator = "<img src=""/_images/icons/user.png"" title=""Scheduled"" alt=""Published"" />"
				checked = " checked=""checked"""
			End If
			If toBePublished Then 
				availableIndicator = "<img src=""/_images/icons/user.png"" alt=""Add"" title=""Add to team"" />"
				checked = " checked=""checked"""
			End If
			If toBeRemoved Then 
				availableIndicator = "&nbsp;"
				checked = ""
			End If
			
			str = str & "<td" & availableClass & " id=""eaid-" & list(19,i) & """>" 
			str = str & "<div>" & availableIndicator & "</div>"
			str = str & "<input type=""checkbox"" name=""event_availability_id"" class=""checkbox"" value=""" & list(19,i) &  """" & checked & " style=""display:none;"" />"
			str = str & "</td>" 
			Exit For
		End If
	Next

	AvailabilityCellInfoToString = str
End Function

Function NoProgramSelectedDialogToString(page)
	Dim dialog			: Set dialog = New cDialog
	Dim pg				: Set pg = page.Clone()
	
	dialog.HeadLine = "Sorry! Can't show you anything yet .."
	
	dialog.Text = dialog.Text & "<p>You are in the Manage schedules section of your account, trying to view your members availability. "
	dialog.Text = dialog.Text & "However, you haven't indicated which program you would like to view. </p>"
	dialog.Text = dialog.Text & "<p>To fix this, select a program from the dropdown list in the toolbar. </p>"
	
	dialog.SubText = dialog.Subtext & "<p>Once you have selected a program, you'll use this page to get a global view or your member's availability (organized by skill) for your program's events. "
	dialog.SubText = dialog.Subtext &  "You can also use this page to easily assign your members to an event team by skill. </p>"

	pg.Action = "": pg.ProgramId = "": pg.SkillId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Back to my schedules</a></li>"
	
	NoProgramSelectedDialogToString = dialog.ToString
End Function

Function NoSkillsDialogToString(page)
	Dim dialog			: Set dialog = New cDialog
	Dim pg				: Set pg = page.Clone()
	
	dialog.HeadLine = "Sorry! Can't show you anything yet .."
	
	dialog.Text = dialog.Text & "<p>You are in the Manage schedules section of your account, trying to view your member availability. "
	dialog.Text = dialog.Text & "However, the program you selected (" & html(page.Program.ProgramName) & ") does not yet have any skills set up. </p>"
	dialog.Text = dialog.Text & "<p>To get started fixing this, click the link for <strong>Create my first skill</strong>. </p>"
	
	dialog.SubText = dialog.Subtext & "<p>Once you have created some program skills and assigned members from your program to those skills, you'll use this page to get a global view of your members' availability (organized by skill) for your program's events. </p>"
	dialog.SubText = dialog.Subtext & "<p>You can also use this page to easily assign your members to an event team by skill. </p>"
	
	
	pg.Action = ADDNEW_RECORD
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>Create my first skill</a></li>"
	pg.Action = "": pg.ProgramId = "": pg.SkillId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Back to my schedules</a></li>"
	
	NoSkillsDialogToString = dialog.ToString
End Function

Function NoMemberSkillsDialogToString(page)
	Dim dialog			: Set dialog = New cDialog
	Dim pg				: Set pg = page.Clone()
	
	dialog.HeadLine = "Sorry! Can't show you anything yet .."
	
	dialog.Text = dialog.Text & "<p>You are in the manage schedule section of your account, trying to view your member availability. "
	dialog.Text = dialog.Text & "However, the program you selected (" & html(page.Program.ProgramName) & ") either doesn't have any members or doesn't have members assigned to skills yet. </p>"
	dialog.Text = dialog.Text & "<p>To get started fixing this, click <strong>Set member skills</strong> or <strong>Set program members</strong>. </p>"
	
	dialog.SubText = "<p>Once you have associated your members with the different skills that belong to this program, you'll use this page to get a global view of your members' availability (organized by skill) for your program's events. </p>"
	dialog.SubText = dialog.Subtext & "<p>You can also use this page to easily assign your members to an event team by skill. </p>"
	
	pg.Action = ASSIGN_SKILLS_TO_MEMBERS: pg.ScheduleID = "": pg.SkillId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>Set member skills</a></li>"
	pg.Action = CONFIGURE_PROGRAM_MEMBERS
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/members.asp" & pg.UrlParamsToString(True) & """>Set program members</a></li>"
	pg.Action = "": pg.ProgramId = "": pg.SkillId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Back to my schedules</a></li>"
	
	NoMemberSkillsDialogToString = dialog.ToString
End Function

Function NoEventsDialogToString(page)
	Dim dialog			: Set dialog = New cDialog
	Dim pg				: Set pg = page.Clone()
	
	dialog.HeadLine = "Whoa! Where's the availability ..?"
	
	dialog.Text = dialog.Text & "<p>You are in the manage schedule section of your account, trying to view your member availability. "
	dialog.Text = dialog.Text & "However, the program you selected (" & html(page.Program.ProgramName) & ") doesn't have a schedule with events set up yet. </p>"
	dialog.Text = dialog.Text & "<p>To get started fixing this, click <strong>Create an event</strong>. </p>"
	
	dialog.SubText = "<p>Once you have associated your members with the different skills that belong to this program, you'll use this page to get a global view of your members' availability (organized by skill) for your program's events. </p>"
	dialog.SubText = dialog.Subtext & "<p>You can also use this page to easily assign your members to an event team by skill. </p>"
	
	pg.Action = ADDNEW_RECORD: pg.ProgramId = "": pg.SkillId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/schedule/events.asp" & pg.UrlParamsToString(True) & """>Create an event</a></li>"
	pg.Action = "": pg.ProgramId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Back to my schedules</a></li>"
	
	NoEventsDialogToString = dialog.ToString
End Function

Function AvailabilityGridToString(page)
	Dim str, i, j, k

	Dim dateTime			: Set dateTime = New cFormatDate
	Dim skill				: Set skill = New cSkill
	
	Dim monthCells
	Dim dayCells
	Dim timeCells
	Dim eventNameCells
	Dim availabilityCells
	
	Dim scheduleName
	
	Dim memberRows
	
	Dim events
	Dim memberList
	Dim scheduleList
	
	Dim eventCount			: eventCount = 0
	
	If Len(page.Schedule.ScheduleId) > 0 Then
		events = page.Schedule.EventList("EventDate, TimeStart")
'	ElseIf Len(page.Program.ProgramId) > 0 Then
'		events = page.Program.EventList("EventDate, TimeStart")
	Else 
		AvailabilityGridToString = NoProgramSelectedDialogToString(page)
		Exit function
	End If			

	memberList = page.Program.MemberList()	
	
	' check for skills ..
	If Not page.Program.HasSkills() Then
		AvailabilityGridToString = NoSkillsDialogToString(page)
		Exit Function
	End If
	
	' check for memberSkills ..
	If Not page.Program.HasEnabledMemberSkills() Then
		AvailabilityGridToString = NoMemberSkillsDialogToString(page)
		Exit Function
	End If
	
	If Len(page.SkillId) = 0 Then
	End If

	' set to first skill with members ..
	If Len(page.SkillId) = 0 Then page.SkillId = GetFirstSkillId(page.Program)

	skill.SkillId = page.SkillId
	skill.Load()
	scheduleList = skill.ScheduleInfoList()
		
	' 0-EventID 1-EventName 2-EventDate 3-EventNote 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified
	' 8-ProgramID 9-ProgramName 10-ScheduleID 11-ScheduleName 12-ScheduleEventCount
	' 13-ScheduleIsVisible 14-HasUnpublishedChanges 15-EventFileCount 16-ScheduledMemberCount
	
	Dim currentMonth
	Dim currentYear
	Dim currentCell
	
	Dim colspan
	Dim lastCell
	Dim doClosePreviousCell
	
	str = str & m_appMessageText
	str = str & "<h3 class="""">Availability for: " & html(skill.SkillName) & "</h3>"
	str = str & "<h4 class=""first"">" & html(page.Schedule.ScheduleName) & "</h4>"
	str = str & "<p>Use this page to see (by skill) when your members are available to be scheduled for events. "
	str = str & "Add or remove a member from an event by clicking their name in the left most column of the grid. "
	str = str & "Change skills by selecting a new skill from the list. </p>"

	str = str & "<div class=""tip-box""><h3>Tip!</h3><p>"
	str = str & "If you make any changes to this page, be sure to click <strong>Publish all</strong> to sync those changes with your member's calendars. "
	str = str & "</p></div>"
	
	str = str & "<div class=""tip-box""><h3>Key</h3><ul>"
	str = str & "<li><span class=""available"">&nbsp;</span>Available</li>"
	str = str & "<li><span class=""not-available"" style="""">&nbsp;</span>Not available</li>"
	str = str & "<li><span class=""unknown"" style="""">&nbsp;</span>Unknown</li>"
	str = str & "</ul></div>"
	
	str = str & "<div id=""availability-grid"" class=""skid-" & page.SkillId & """>"
	
	If IsArray(events) Then
		For j = 0 To UBound(events,2) Step NUMBER_OF_EVENT_COLUMNS
			monthCells = ""
			dayCells = ""
			timeCells = ""
			eventNameCells = ""
			memberRows = ""

			i = j
			Do While i < j + NUMBER_OF_EVENT_COLUMNS
				
				If i <= UBound(events,2) Then
				
					' begin month/year header
					If i = j Then
						colspan = 0
						
						lastCell = False
						doClosePreviousCell = False
						
						currentMonth = Month(events(2,i))
						currentYear = Year(events(2,i))
						currentCell = MonthName(Month(events(2,i))) & " " & Year(events(2,i))
					End If
					
					' close month/year cell because end of row ..
					If i = j + NUMBER_OF_EVENT_COLUMNS - 1 Then
						doClosePreviousCell = True
						colspan = colspan + 1
					End If
					
					' close month/year cell because end of array ..
					If i = UBound(events,2) Then
						doClosePreviousCell = True
						colspan = colspan + 1
					End If
					
					' close month/year cell because new month/year ..
					If (Month(events(2,i)) <> currentMonth) Or (Year(events(2,i)) <> currentYear) Then
						doClosePreviousCell = True
						
					End If
					
					If doClosePreviousCell Then

						' save the current cell
						monthCells = monthCells & "<td colspan=""" & colspan & """>" & currentCell & "</td>"

						' reset flags and colspan						
						lastCell = False
						doClosePreviousCell = False
						colspan = 1

						' get current month, year ..
						currentCell = MonthName(Month(events(2,i))) & " " & Year(events(2,i))
					Else
					
						' don't close cell, increment colspan instead ..
						colspan = colspan + 1
					End If		
					
					' reset current month ..
					currentMonth = Month(events(2,i))
					currentYear = Year(events(2,i))
					
					dayCells = dayCells & "<td>" & WeekdayName(Weekday(events(2,i)), True) & " " & Day(events(2,i)) & "</td>"

					timeCells = timeCells & "<td>"
					If Len(events(4,i)) > 0 Then 
						timeCells = timeCells & dateTime.Convert(events(4,i), "hh:nn PP")
					Else
						timeCells = timeCells & " - "
					End If
					timeCells = timeCells & "</td>"
					
					eventNameCells = eventNameCells & "<td>" & html(events(1,i)) & "</td>"
					
					' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 6-ProgramMemberIsEnabled 7-MemberActiveStatus 8-DateCreated 9-DateModified
					' 10-ProgramMemberID 11-IsApproved 12-HasMissingAvailability 13-Email

					eventCount = eventCount + 1
				End If
				
				i = i + 1
			Loop

			str = str & "<table><thead>"
			str = str & "<tr class=""month""><td class=""month"" rowspan=""4"" style=""background-color:#ccc;vertical-align:middle;font-size:1.3em;padding:10px;"">" & Server.htmlEncode(skill.SkillName) & "</td>" & monthCells & "</tr>"
			str = str & "<tr class=""day"">" & dayCells & "</tr>"
			str = str & "<tr class=""time"">" & timeCells & "</tr>"
			str = str & "<tr class=""event-name"">" & eventNameCells & "</tr>"
			str = str & "</thead>"
			str = str & "<tbody>"
			
			If IsArray(memberList) Then
			
				For k = 0 To UBound(memberList,2)
					availabilityCells = ""
			
					i = j
					Do While i < j + NUMBER_OF_EVENT_COLUMNS
						
						If i <= UBound(events,2) Then
							availabilityCells = availabilityCells & AvailabilityCellInfoToString(memberList(0,k), events(0,i), scheduleList)
						End If
						
						i = i + 1
					Loop
					
					If Len(availabilityCells) > 0 Then
						str = str & "<tr class=""available""><td class=""member-name"" title=""Edit team""><img class=""icon"" src=""/_images/icons/group_edit.png"" alt=""Edit team"" />" & html(memberList(1,k) & ", " & memberList(2,k)) & "</td>"
						str = str & availabilityCells
						str = str & "</tr>"
					End If
				Next
			End If				
			str = str & "</tbody></table>"
		Next
	End If
	str = str & "</div>"

	If eventCount = 0 Then 
		AvailabilityGridToString = NoEventsDialogToString(page)
		Exit Function
	End If	

	AvailabilityGridToString = str
End Function

Function OptionGroupForGoToSkillDropdownToString(list, groupId, groupName)
	Dim str, i
	
	Dim isSkillEnabled
	Dim isGroupEnabled
	Dim optionDisabled
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)

			If (CStr(groupId & "") = CStr(list(4,i) & "")) Then
			
				isSkillEnabled = True		: If list(3,i) = 0 Then isSkillEnabled = False
				isGroupEnabled = True		: If list(7,i) = 0 Then isGroupEnabled = False
				
				' disable skill options that have active member count = 0 ..
				optionDisabled = ""			: If list(10,i) = 0 Then optionDisabled = " disabled=""disabled"""
			
				If isSkillEnabled And isGroupEnabled Then
					str = str & "<option value=""" & list(0,i) & """" & optionDisabled & ">" & html(list(1,i)) & "</option>"
				End if
			End If
		Next
	End If
	
	If Len(str) > 0 Then 
			str = "<optgroup label=""" & html(groupName) & """>" & str & "</optgroup>"
	End If

	OptionGroupForGoToSkillDropdownToString = str
End Function

Function GoToSkillDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	Dim list			: If Len(page.Program.ProgramId) > 0 Then list = page.Program.SkillList("")
	Dim groupList		: If Len(page.Program.ProgramId) > 0 Then groupList = page.Program.SkillGroupList()
	
	Dim options			: options = ""
	Dim hasSkills
	
	Dim disabled		: disabled = ""
	If Len(page.ProgramId & "") = 0 Then disabled = " disabled=""disabled"""

	' list() ..
	' 0-SkillID 1-SkillName 2-SkillDesc 3-IsSkillEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-IsGroupEnabled 8-LastModified 9-DateCreated 10-ActiveMemberCount
	
	' groupList() ..
	' 0-SkillGroupID 1-GroupName 2-GroupDesc 3-IsEnabled 4-AllowMultiple 5-DateModified 6-DateCreated
	
	If IsArray(groupList) Then 
		For i = 0 To UBound(groupList,2)
			options = options & OptionGroupForGoToSkillDropdownToString(list, groupList(0,i), groupList(1,i))
		Next
	End If
	
	' list ungrouped skills last ..
	options = options & OptionGroupForGoToSkillDropdownToString(list, Null, "No group")

	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-goto-skill-dropdown"">"
	str = str & "<input type=""hidden"" name=""form_goto_skill_dropdown_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""new_skill_id"" id=""goto-skill-dropdown""" & disabled & ">"
	str = str & "<option value="""">" & html("< Go to skill >") & "</option>"
	str = str & "<option value="""">" & html(" -- ") & "</option>"

	str = str & options
	str = str & "</select></form></li>"
	
	GoToSkillDropdownToString = str
End Function

Function OwnedProgramDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	Dim list			: list = page.Member.OwnedProgramsList()
	Dim selected		: selected = ""
	Dim disabled		: disabled = ""
	
	' 0-ProgramId 1-ProgramName 2-IsEnabled 3-ScheduleCount 4-EventCount

	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-owned-program-dropdown"">"
	str = str & "<input type=""hidden"" name=""form_owned_program_dropdwon_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""new_program_id"" id=""owned-program-dropdown"">"
	str = str & "<option value="""">" & html("< Select a program >") & "</option>"
	str = str & "<option value="""">" & html(" -- ") & "</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			selected = "":		If CStr(list(0,i) & "") = CStr(page.Program.ProgramId & "") Then selected = " selected=""selected"""
			
			disabled = ""
			If list(3,i) = 0 Then disabled = " disabled=""disabled"""
			If list(4,i) = 0 Then disabled = " disabled=""disabled"""

			str = str & "<option value=""" & list(0,i) & """" & selected & disabled & ">" & html(list(1,i)) & "</option>"
		Next
	End If
	str = str & "</select></form></li>"
	
	OwnedProgramDropdownToString = str
End Function

Function ScheduleDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	Dim list	
	Dim selected		: selected = ""		
	Dim disabled		: disabled = " disabled=""disabled"""
	If Len(page.Program.ProgramId & "") > 0 Then 
		disabled = ""
		list = page.Program.ScheduleList()
	End If
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-schedule-dropdown"">"
	str = str & "<input type=""hidden"" name=""form_schedule_dropdown_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""new_schedule_id"" id=""schedule-dropdown""" & disabled & ">"
	str = str & "<option value="""">" & html("< Select a schedule >") & "</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			selected = ""			: If CStr(list(0,i) & "") = CStr(page.Schedule.ScheduleId & "") Then selected = " selected=""selected"""
			
			str = str & "<option value=""" & list(0,i) & """" & selected & ">" & html(list(1,i)) & "</option>"
		Next
	End If	
	str = str & "</select></form></li>"

	ScheduleDropdownToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim skill		: Set skill = New cSkill
	
	pg.ScheduleId = "": pg.ProgramId = "": pg.SkillId = ""
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	str = str & "<a href=""/schedule/schedules.asp"">Schedules</a> / "
	
	If Len(page.ProgramId) = 0 Then 
		str = str & "Availability"
	Else
		str = str & "<a href=""/schedule/availability.asp" & pg.UrlParamsToString(True) & """>Availability</a> / "
		
		Set pg = page.Clone()
		pg.ScheduleId = "": pg.SkillId = ""
		str = str & "<a href=""/schedule/availability.asp" & pg.UrlParamsToString(True) & """>" & html(page.Program.ProgramName) & "</a> / "
		If Len(page.ScheduleId & "") > 0 Then
			Set pg = page.Clone()
			pg.SkillId = ""
			str = str & "<a href=""/schedule/availability.asp" & pg.UrlParamsToString(True) & """>" & html(page.Schedule.ScheduleName) & "</a> / "
		End If
		
		If Len(page.SkillId & "") = 0 Then 
			page.SkillId = GetFirstSkillId(page.Program)
		End If
		skill.SkillId = page.SkillId
		If Len(skill.SkillId) > 0 Then Call skill.Load()
		
		str = str & html(skill.SkillName)
	End If

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Dim buttonStyle				: buttonStyle = " style=""display:none;"""
	Dim buttonText				: buttonText = "Publish all"
	
	' conditionally show publish button ..
	If Len(page.ScheduleId & "") > 0 Then
		pg.Action = PUBLISH_SCHEDULE
		If page.Schedule.PublishStatus = SCHEDULE_HAS_UNPUBLISHED_CHANGES Then
			' show publish schedule button ..
			buttonStyle = ""
		End If
	Else
		pg.Action = PUBLISH_PROGRAM
		If page.Program.PublishStatus = SCHEDULE_HAS_UNPUBLISHED_CHANGES Then
			' show publish program button ..
			buttonStyle = ""
		End If
	End If
	
	Dim publishButton
	href = pg.Url & pg.UrlParamsToString(True)
	publishButton = "<li " & buttonStyle & " id=""publish-button""><a href=""" & href & """><img src=""/_images/icons/arrow_rotate_clockwise.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>" & buttonText & "</a></li>"
	
	Dim eventTeamsButton
	pg.Action = ""
	href = "/schedule/teams.asp" & pg.UrlParamsToString(True)
	eventTeamsButton = "<li><a href=""" & href & """><img src=""/_images/icons/group.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Event Teams</a></li>"
	
	Dim schedulesButton
	pg.EventID = "": pg.Action = "": pg.ScheduleID = "": pg.ReturnContext = ""
	href = "/schedule/schedules.asp" & pg.UrlParamsToString(True)
	schedulesButton = schedulesButton & "<li><a href=""" & href & """><img src=""/_images/icons/calendar.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Schedules</a></li>"
	
	str = str & publishButton
	str = str & GoToSkillDropdownToString(page)
	str = str & ScheduleDropdownToString(page)
	str = str & OwnedProgramDropdownToString(page)
	str = str & eventTeamsButton
	str = str & SchedulesButton
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->

<%
Class cPage
	' unencrypted
	Public MessageID
	Public ReturnContext
	Public ShowPastEvents
	Public Month
	Public Day
	Public Year

	' encrypted
	Public Action
	Public ProgramID
	Public ScheduleID
	Public FilterScheduleId
	Public EventID
	Public MemberId
	Public SkillID

	' objects
	Public Member
	Public Client
	Public Program	
	Public Schedule
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(ReturnContext) > 0 Then str = str & "rc=" & ReturnContext & amp
		If Len(ShowPastEvents) > 0 Then str = str & "hpe=" & ShowPastEvents & amp
		If Len(Month) > 0 Then str = str & "m=" & Month & amp
		If Len(Day) > 0 Then str = str & "d=" & Day & amp
		If Len(Year) > 0 Then str = str & "y=" & Year & amp
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(ScheduleID) > 0 Then str = str & "scid=" & Encrypt(ScheduleID) & amp
		If Len(FilterScheduleId) > 0 Then str = str & "fscid=" & Encrypt(FilterScheduleId) & amp
		If Len(EventID) > 0 Then str = str & "eid=" & Encrypt(EventID) & amp
		If Len(MemberId) > 0 Then str = str & "mid=" & Encrypt(MemberId) & amp
		If Len(SkillID) > 0 Then str = str & "skid=" & Encrypt(SkillID) & amp
		
		If Len(str) > 0 Then 
			str = Left(str, Len(str) - Len(amp))
		Else
			' qstring needs at least one param in case more params are appended ..
			str = str & "noparm=true"
		End If
		str = "?" & str
		
		UrlParamsToString = str
	End Function
	
	Public Function Clone()
		Dim c
		Set c = New cPage

		c.MessageID = MessageID
		c.ReturnContext = ReturnContext
		c.ShowPastEvents = ShowPastEvents
		c.Month = Month
		c.Day = Day
		c.Year = Year

		c.Action = Action
		c.ProgramID = ProgramID
		c.ScheduleID = ScheduleID
		c.FilterScheduleId = FilterScheduleId
		c.EventID = EventID
		c.MemberId = MemberId
		c.SkillId = SkillId
				
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Schedule = Schedule
		
		Set Clone = c
	End Function
End Class
%>

