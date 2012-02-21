<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Const NUMBER_OF_EVENT_COLUMNS = 4

' global view tokens
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
	
	page.Month = Request.Querystring("m")
	page.Day = Request.Querystring("d")
	page.Year = Request.Querystring("y")
	
	page.Action = Decrypt(Request.QueryString("act"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.ScheduleID = Decrypt(Request.QueryString("scid"))
	page.FilterScheduleId = Decrypt(Request.QueryString("fscid"))
	page.EventID = Decrypt(Request.QueryString("eid"))
	
	' reset page.EventId if postback from dropdown
	If Request.Form("form_go_to_event_dropdown_is_postback") = IS_POSTBACK Then
		page.EventId = Request.Form("event_id")
	End If

	If Request.Form("form_go_to_schedule_is_postback") = IS_POSTBACK Then
		If Len(Request.Form("new_schedule_id")) > 0 Then
			page.ScheduleId = Request.Form("new_schedule_id")
		End If
	End If
		
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	' set show past events state ..
	If Request.Form("form_show_events_is_postback") = IS_POSTBACK Then
		page.ShowPastEvents = Request.Form("show_events")
	Else
		' read from cookie ..
		page.ShowPastEvents = Request.Cookies("past-events-view")
	End If

	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()
	Set page.Schedule = New cSchedule
	page.Schedule.ScheduleID = page.ScheduleID
	If Len(page.Schedule.ScheduleID) > 0 Then Call page.Schedule.Load()
	Set page.Evnt = New cEvent
	page.Evnt.EventID = page.EventID
	If Len(page.Evnt.EventID) > 0 Then Call page.Evnt.Load()

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
		<link rel="stylesheet" type="text/css" href="/_incs/script/functions/member_event_widget/member_event_widget.css" />
		<link rel="stylesheet" type="text/css" href="teams.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/form/jquery.form.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/accordion/jquery.accordion.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/dimensions/jquery.dimensions.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/livequery/jquery.livequery.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/jtruncate/jquery.jtruncate.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/cookie/jquery.cookie.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/functions/member_event_widget/member_event_widget.js"></script>
		<script language="javascript" type="text/javascript" src="teams.js"></script>

		<script language="javascript" type="text/javascript">
		
			// convert server-side comments to jscript ..
			var COPY_EVENT_TEAM_TO_EVENT		= <%=COPY_EVENT_TEAM_TO_EVENT %>
			var RETURN_SCHEDULE_ITEM			= <%=RETURN_SCHEDULE_ITEM %>
			var RETURN_EVENT_TEAM				= <%=RETURN_EVENT_TEAM %>
			var PUBLISH_EVENT					= <%=PUBLISH_EVENT %>
			var CLEAR_EVENT_TEAM_FROM_EVENT		= <%=CLEAR_EVENT_TEAM_FROM_EVENT %>
			var SCHEDULE_ITEM_TYPE_COPY_TO		= <%=SCHEDULE_ITEM_TYPE_COPY_TO %>
			var SCHEDULE_ITEM_TYPE_COPY_FROM	= <%=SCHEDULE_ITEM_TYPE_COPY_FROM %>
			
		</script>

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
		Case PUBLISH_EVENT
			Call DoPublishEvent(page.Evnt.EventId, page.Member.NameLast & ", " & page.Member.NameFirst, rv)
			Select Case rv
				Case 0	
					page.EventId = "": page.Action = "": page.MessageId = 5022
				Case Else
					page.EventId = "": page.Action = "": page.MessageId = 5023
				
			End Select
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case CLEAR_EVENT_TEAM_FROM_EVENT
			Call DoDeleteEventTeam(page.Evnt.EventId, rv)
			Select Case rv
				Case 0
					page.EventId = "": page.Action = "": page.MessageId = 6040
				Case Else
					page.EventId = "": page.Action = "": page.MessageId = 6041
				
			End Select
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
			
		Case PUBLISH_SCHEDULE
			Call DoPublishSchedule(page.ScheduleId, page.Member.NameLast & ", " & page.Member.NameFirst, rv)
			page.Action = "": page.MessageId = 6019
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case COPY_EVENT_TEAM_TO_EVENT
			If Not ProgramHasSkills(page.Evnt.ProgramId) Then
				str = str & NoSkillsDialogToString(page)
			Else
				str = str & DuplicateEventTeamGridToString(page)
			End If

		Case UPDATE_RECORD
			If Not ProgramHasSkills(page.Evnt.ProgramId) Then
				str = str & NoSkillsDialogToString(page)
			Else
				str = str & EditEventTeamGridToString(page)
			End If
			
		Case Else
			str = str & MasterEventViewToString(page)
			
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoDeleteEventTeam(eventId, outError)
	Dim scheduleBuild
	
	Set scheduleBuild = New cScheduleBuild
	scheduleBuild.EventID = eventID
	Call scheduleBuild.ClearAllByEventID(outError)
End Sub

Sub DoPublishEvent(eventId, publishedBy, outError)
	outError = 0
	
	Dim evnt			: Set evnt = New cEvent
	evnt.EventId = eventId
	
	Call evnt.Publish(publishedBy, outError)
End Sub

Sub DoPublishSchedule(scheduleId, publishedBy, outError)
	outError = 0
	
	Dim schedule		: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	
	Call schedule.Publish(publishedBy, outError)
End Sub

Function MasterEventViewToString(page)
	Dim str, msg, i
	Dim pg				: Set pg = page.Clone()
	
	Dim program			: Set program = New cProgram
	program.ProgramId = page.Schedule.ProgramId
	Call program.Load()
	Dim skillList		: skillList = program.SkillList("")
	
	Dim startDate			: If Len(page.ShowPastEvents) = 0 Then startDate = Date()
	Dim eventList			: eventList = page.Schedule.EventListForPeriod(startDate, Null, "")
	
	' check for no skills ..
	If Not program.HasSkills() Then
		MasterEventViewToString = NoSkillsDialogToString(page)
		Exit Function
	End If
	
	' check for events ..
	If Not IsArray(eventList) Then
		MasterEventViewToString = NoEventsDialogToString(page)
		Exit Function
	End If
	
	Dim scheduleBuild	: Set scheduleBuild = New cScheduleBuild
	Dim teamList		: teamList = scheduleBuild.TeamList(page.Schedule.ScheduleId)
	
	Dim eventTotal		: eventTotal = UBound(eventList,2) + 1
	Dim eventIdx		' base 0
	
	Dim tablesNeeded	: tablesNeeded = Int(eventTotal/NUMBER_OF_EVENT_COLUMNS)
	Dim tableIdx		' base 0
	If eventTotal Mod NUMBER_OF_EVENT_COLUMNS > 0 Then tablesNeeded = tablesNeeded + 1
	
	Dim startIdx
	Dim endIdx
	
	Dim headerCells

	str = m_appMessageText
	str = str & "<h3 class="""">" & server.HTMLEncode(page.Schedule.ProgramName) & "</h3>"
	str = str & "<h4 class=""first"">" & server.HTMLEncode(page.Schedule.ScheduleName) & "</h4>"
	str = str & "<p>This is the master view for all events belonging to this schedule. "
	str = str & "Choose <strong>Edit team</strong> from the toolbar for any event to add or remove members from a team. </p>"

	str = str & AvailabilityWidgetToString(page.ScheduleId, page.ShowPastEvents)
	str = str & ViewFilterButtonsToString(page.UrlParamsToString(True), page.ShowPastEvents, True, True, True)
	str = str & "<div class=""tip-box view-filter-key""><h3>Tip!</h3>"
	str = str & "<p>The <strong>Available</strong> and <strong>Published</strong> buttons show more information about this event team schedule. </p>"
	str = str & "<ul><li class=""style-for-default"">Available</li>"
	str = str & "<li class=""style-for-not-available"">Not available</li>"
	str = str & "<li class=""style-for-unknown-available"">Unknown availability</li>"
	str = str & "<li class=""style-for-default"">Published</li>"
	str = str & "<li class=""style-for-publish"">Unpublished (add)</li>"
	str = str & "<li class=""style-for-unpublish"">Unpublished (remove)</li></ul>"
	str = str & "</div>"	

	str = str & "<div class=""tip-box filters""><h3>Filters! </h3>"
	
	' showPastEvents, hideEmptySkills	
	str = str & "<p>You can select or clear the checkboxes to show or hide skills in the team view. </p>"
	str = str & "<ul class=""filters"">"
	str = str & "<li><input type=""checkbox"" name=""hide_empty_skills"" id=""hide-empty-skills"" class=""checkbox"" />"
	str = str & "Show empty skills. </li></ul>"
	
	' skill listing checkboxes
	str = str & "<ul class=""filters"" id=""hide-skills"">"
	For i = 0 To UBound(skillList,2)
		str = str & "<li><input type=""checkbox"" name=""skill_name"" class=""checkbox"" id=""skill-checkbox-id-" & skillList(0,i) & """ />" 
		str = str & html(skillList(1,i)) & "</li>"
	Next
	str = str & "</ul>"
	str = str & "</div>"

	str = str & "<div id=""master-team-view"">"
	
	tableIdx = 0
	Do While tableIdx < tablesNeeded
		startIdx = (tableIdx * NUMBER_OF_EVENT_COLUMNS)
		endIdx = (tableIdx * NUMBER_OF_EVENT_COLUMNS) + (NUMBER_OF_EVENT_COLUMNS - 1)
		
		str = str & "<table>"
		str = str & MasterEventViewHeaderRowToString(startIdx, endIdx, eventList, page)
		str = str & MasterViewSkillRowsToString(startIdx, endIdx, skillList, eventList, teamList, page.Schedule.HtmlBackgroundColor)
		str = str & "</table>"
		
		tableIdx = tableIdx + 1
	Loop
	
	str = str & "</div>"

	MasterEventViewToString = str
End Function

Function MasterEventViewHeaderRowToString(startIdx, endIdx, eventList, page)
	Dim str, i
	Dim pg						: Set pg = page.Clone()
	Dim dateTime				: Set dateTime = New cFormatDate
	
	Dim bottom
	Dim eventDateRow
	Dim eventNameRow
	Dim eventToolbarRow
	
	Dim eventTotal				: eventTotal = UBound(eventList,2) + 1
	
	eventDateRow = eventDateRow & "<th class=""skill-header-cell"">&nbsp;</th>"
	eventNameRow = eventNameRow & "<th class=""skill-header-cell"">&nbsp;</th>"
	eventToolbarRow = eventToolbarRow & "<th class=""skill-header-cell"">&nbsp;</th>"
	
	For i = startIdx To endIdx
		If i > eventTotal - 1 Then
			eventDateRow = eventDateRow & "<th class=""event-header"">&nbsp;</th>"
			eventNameRow = eventNameRow & "<th style=""background-color:" & page.Schedule.HtmlBackgroundColor & ";"" class=""event-item top"">&nbsp;</th>"
			eventToolbarRow = eventToolbarRow & "<th style=""background-color:" & page.Schedule.HtmlBackgroundColor & ";"" class=""event-item bottom"">&nbsp;</th>"
		Else
			eventDateRow = eventDateRow & "<th class=""event-header"">" & WeekdayName(Weekday(eventList(2,i)), True) & " " & Day(eventList(2,i)) & " " & MonthName(Month(eventList(2,i)), True) & "</th>"
			
			eventNameRow = eventNameRow & "<td style=""background-color:" & page.Schedule.HtmlBackgroundColor & ";"" class=""event-item top"">"
			eventNameRow = eventNameRow & "<p style=""text-align:center;""><strong>" & html(eventList(1,i)) & "</strong></p>"
			
			eventNameRow = eventNameRow & "</td>"
	
			eventToolbarRow = eventToolbarRow & "<td style=""background-color:" & page.Schedule.HtmlBackgroundColor & ";"" class=""event-item bottom""><div class=""toolbar"">"
			' edit team
			pg.Action = UPDATE_RECORD: pg.EventId = eventList(0,i)
			eventToolbarRow = eventToolbarRow & "<a href=""/schedule/teams.asp" & pg.UrlparamsToString(True) & """ title=""Edit Team"">"
			eventToolbarRow = eventToolbarRow & "<img src=""/_images/icons/group_edit.png"" alt="""" /></a>"
			' copy team
			pg.Action = COPY_EVENT_TEAM_TO_EVENT: pg.EventId = eventList(0,i)
			eventToolbarRow = eventToolbarRow & "<a href=""/schedule/teams.asp" & pg.UrlparamsToString(True) & """ title=""Copy Team"">"
			eventToolbarRow = eventToolbarRow & "<img src=""/_images/icons/paste_group.png"" alt="""" /></a>"
			' publish event
			pg.Action = PUBLISH_EVENT: pg.EventId = eventList(0,i)
			eventToolbarRow = eventToolbarRow & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Publish Team"">"
			eventToolbarRow = eventToolbarRow & "<img src=""/_images/icons/arrow_rotate_clockwise.png"" alt="""" /></a>"
			' clear team
			pg.Action = CLEAR_EVENT_TEAM_FROM_EVENT: pg.EventId = eventList(0,i)
			eventToolbarRow = eventToolbarRow & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove Team"">"
			eventToolbarRow = eventToolbarRow & "<img src=""/_images/icons/cross.png"" alt="""" /></a>"
			eventToolbarRow = eventToolbarRow & "</div></td>"
		End If
	Next
	str = str & "<tr>" & eventDateRow & "</tr><tr>" & eventNameRow & "</tr><tr>" & eventToolbarRow & "</tr>"
	
	MasterEventViewHeaderRowToString = str
End Function

Function MasterViewMemberSkillCellsToString(startIdx, endIdx, firstInGroupClass, skillId, eventList, teamList)
	Dim str, i, j
	
	Dim memberString			: memberString = ""
	Dim cls
	
	' 0-EventID 1-EventName 2-EventDate 3-TimeStart 4-TimeEnd 5-MemberId 6-NameLast 7-NameFirst 
	' 8-MemberIsEnabled 9-ProgramMemberId 10-ProgramMemberIsEnabled 11-IsAvailable 
	' 12-AvailabilityIsViewedByMember 13-SkillGroupId 14-SkillGroupName 15-SkillGroupIsEnabled
	' 16-SkillId 17-SkillName 18-SkillIsEnabled 19-PublishStatus
	
	memberString = ""
	For i = startIdx To endIdx
		
		' only generate for non-emtpy columns
		If i <= UBound(eventList,2) Then
			
			' spin through the teamList, 
			' add names for this eventId, skillId
			If IsArray(teamList) Then
				For j = 0 To UBound(teamList,2)
					If CStr(teamList(0,j)) = CStr(eventList(0,i)) Then
						If CStr(teamList(16,j)) = CStr(skillId) Then
						
							' set class based on publishStatus()
							cls = "" 
							If teamList(19,j) = 1 Then
								cls = "publish"
							ElseIf teamList(19,j) = 2 Then
								cls = "unpublish"
							End If
							
							If teamList(12,j) = 0 Then
								If Len(cls) > 0 Then cls = cls & " "
								cls = cls & "unknown-available"
							Else
								If teamList(11,j) = 1 Then
									If Len(cls) > 0 Then cls = cls & " "
									cls = cls & "available"
								Else
									If Len(cls) > 0 Then cls = cls & " "
									cls = cls & "not-available"
								End If
							End If
							
							memberString = memberString & "<li class=""" & cls & """><a href=""#"" class=""mid-" & teamList(5,j) & """>" & html(teamList(6,j) & ", " & teamList(7,j)) & "</a></li>"
						End If
					End If
				Next
				If Len(memberString) > 0 Then memberString = "<ul class=""team-list-for-skill"">" & memberString & "</ul>"
			End If
		End If
		str = str & "<td class=""" & firstInGroupClass & """>" & memberString & "</td>"

		' reset member string ..
		memberString = ""
	Next
	
	MasterViewMemberSkillCellsToString = str
End Function

Function MasterViewSkillRowsToString(startIdx, endIdx, skillList, eventList, teamList, htmlBackgroundColor)
	Dim str, i
	
	' 0-SkillID 1-SkillName 2-SkillDesc 3-IsSkillEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-IsGroupEnabled 8-LastModified 9-DateCreated
		
	Dim skillEnabled				: skillEnabled = True
	Dim skillGroupEnabled			: skillGroupEnabled = True
	
	Dim lastGroupId						: lastGroupId = ""
	Dim firstInGroupClass				: firstInGroupClass = ""
	
	Dim emptyRowClass					: emptyRowClass = ""
	
	Dim classes
	
	' set skillgroupid for first record ..
	lastGroupId = skillList(4,0) & ""
	
	For i = 0 To UBound(skillList,2)
		skillEnabled = True				: If skillList(3,i) = 0 Then skillEnabled = False
		skillGroupEnabled = True		: If skillList(7,i) = 0 Then skillGroupEnabled = False
		
		If skillEnabled And skillGroupEnabled Then 
			firstInGroupClass = ""
			If CStr(skillList(4,i) & "") <> CStr(lastGroupId) Then
				firstInGroupClass = "first-in-group"
				lastGroupId = skillList(4,i) & ""
			End If
			
			classes = ""
			If Not SkillHasTeam(skillList(0,i), teamList) Then
				classes = "empty"
			Else
				classes = "not-empty"
			End If
			
			classes = classes & " skill-row-id-" & skillList(0,i)
			
			str = str & "<tr class=""" & classes & """><td class=""skill-label-cell " & firstInGroupClass & """ style=""background-color:" & htmlBackgroundColor & ";"">" & skillList(1,i) & "</td>"
			str = str & MasterViewMemberSkillCellsToString(startIdx, endIdx, firstInGroupClass, skillList(0,i), eventList, teamList)
			str = str & "</tr>"
		End If
	Next
	
	MasterViewSkillRowsToString = str
End Function

Function SkillHasTeam(skillId, teamList)
	Dim i
	
	' 16-SkillId 17-SkillName 18-SkillIsEnabled 19-PublishStatus
	
	SkillHasTeam = False
	
	If Not IsArray(teamList) Then
		SkillHasTeam = False
		Exit Function
	End If
	
	For i = 0 To UBound(teamList,2)
		If CStr(skillId) = CStr(teamList(16,i)) Then
			SkillHasTeam = True
			Exit For
		End If
	Next
End Function

Function ViewFilterButtonsToString(url, showPastEvents, showAvailableButton, showPublishedButton, showPastEventsButton)
	Dim str
	
	Dim newValue
	Dim linkText		: linkText = ""
	Dim cls				: cls = ""
	
	If showAvailableButton Then str = str & "<a href=""#"" id=""see-available-button"">Available</a>"
	If showPublishedButton Then str = str & "<a href=""#"" id=""see-published-button"">Published</a>"
	If Len(str) > 0 Then str = "<div id=""view-button-container"">" & str & "</div>"
	
	If Len(showPastEvents) > 0 Then
		newValue = ""
		cls = "selected"
		linkText = "Hide past events .."
		Response.Cookies("past-events-view") = "on"
	Else
		newValue = "on"
		cls = ""
		linkText = "Show past events .."
		Response.Cookies("past-events-view") = ""
	End If
	
	If showPastEventsButton Then
		str = str & "<div id=""past-event-button-container"">"
		str = str & "<a href=""#"" id=""past-events-button"" class=""" & cls & """>" & linkText & "</a>"
		
		str = str & "<form action=""" & url & """ method=""post"" id=""form-show-past-events"" style=""display:inline;"">"
		str = str & "<input type=""hidden"" name=""form_show_events_is_postback"" value=""" & IS_POSTBACK & """ />"
		str = str & "<input type=""hidden"" name=""show_events"" id=""show-past-events"" value=""" & newValue & """ />"
		str = str & "</form></div>"
	End If
	
	ViewFilterButtonsToString = str
End Function

Function NoEventsDialogToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim dialog			: Set dialog = New cDialog
	
	dialog.HeadLine = "Ok, there's nothing to see here ..!"
	
	dialog.Text = dialog.Text & "<p>You are currently in the Manage Event Teams section of your account. "
	dialog.Text = dialog.Text & "Either this schedule (" & html(page.Schedule.ScheduleName) & ") doesn't have any events set up yet, "
	dialog.Text = dialog.Text & "or all your events occurred in the past. </p>"
	dialog.Text = dialog.Text & "<p>You can get started fixing this by clicking <strong>Create your first event</strong>. </p>"
	
	dialog.SubText = dialog.SubText & "<p>Once you start creating events and organizing your event teams, "
	dialog.SubText = dialog.SubText & "this page will show you who you have scheduled for each event (and what they will be doing). </p>"

	pg.Action = ADDNEW_RECORD
	dialog.LinkList = dialog.LinkList & "<li><a href=""/schedule/events.asp" & pg.UrlParamsToString(True) & """>Create your first event</a></li>"
	
	Set pg = page.Clone()
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """ id=""show-past-events-link"">Show past events</a>"
	dialog.LinkList = dialog.LinkList & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-show-past-events"">"
	dialog.LinkList = dialog.LinkList & "<input type=""hidden"" name=""form_show_events_is_postback"" value=""" & IS_POSTBACK & """ />"
	dialog.LinkList = dialog.LinkList & "<input type=""hidden"" name=""show_events"" id=""show-past-events"" value=""on"" />"
	dialog.LinkList = dialog.LinkList & "</form></li>"

	NoEventsDialogToString = dialog.ToString()
End Function

Function NoSkillsDialogToString(page)
	Dim dialog			: Set dialog = New cDialog
	Dim pg				: Set pg = page.Clone()
	
	dialog.HeadLine = "Whoa, something's missing here ..!"
	
	dialog.Text = dialog.Text & "<p>You are in the Manage Schedules section of your account, "
	dialog.Text = dialog.Text & "and you are trying to work with the event team for your " & html(page.Evnt.EventName) & " (" & page.Evnt.EventDate & ") event. "
	dialog.Text = dialog.Text & "Either you haven't set up any skills for the program this event belongs to (" & html(page.Evnt.ProgramName) & "), "
	dialog.Text = dialog.Text & "or you have disabled all of the skills for this program. </p>"
	dialog.Text = dialog.Text & "<p>To get started on fixing this, click <strong>Create your first skill</strong>. </p>"

	dialog.SubText = dialog.SubText & "<p>Once you have some skills enabled for your program, "
	dialog.SubText = dialog.SubText & "this page is where you'll organize your members into event teams by what they will be doing at your events. "
	dialog.SubText = dialog.SubText & "When your event teams are set up, you will use this page to publish your completed schedules to your member's calendars. </p>"

	pg.Action = ADDNEW_RECORD: pg.ProgramId = pg.Evnt.ProgramId: pg.EventId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>Create your first skill</a></li>"
	pg.ProgramId = pg.Evnt.ProgramId: pg.Action = "": page.EventId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>Check for disabled skills</a></li>"
	
	NoSkillsDialogToString = dialog.ToString()
End Function

Function ProgramHasSkills(programId)
	Dim program			: Set program = New cProgram
	
	ProgramHasSkills = True
	
	program.ProgramId = programId
	Call program.Load()
	
	If Not program.HasSkills() Then ProgramHasSkills = False
End Function

Function AvailableButtonTipBoxToString()
	Dim str

	str = str & "<div class=""tip-box view-filter-key""><h3>Tip!</h3>"
	str = str & "<p>The <strong>Available</strong> button shows more information about this event team schedule. </p>"
	str = str & "<ul><li class=""style-for-default"">Available</li>"
	str = str & "<li class=""style-for-not-available"">Not available</li>"
	str = str & "<li class=""style-for-unknown-available"">Unknown availability</li></ul></div>"	

	AvailableButtonTipBoxToString = str
End Function

Function DuplicateEventTeamGridToString(page)
	Dim str
	
	Dim evnt			: Set evnt = New cEvent
	evnt.EventId = page.EventId
	Call evnt.Load()
	
	Dim program			: Set program = New cProgram
	program.ProgramId = page.Evnt.ProgramId
	
	Dim schedule		: Set schedule = New cSchedule
	schedule.ScheduleId = evnt.ScheduleId
	If Len(schedule.ScheduleId) > 0 Then Call schedule.Load()
	
	str = str & "<h3>Copy event team</h3>"
	str = str & "<p>Select an event to copy from and an even to copy to. "
	str = str & "Then use the copy button to copy the same event team into a new event. </p>"
	str = str & AvailabilityWidgetToString(evnt.ScheduleId, "on")
	str = str & ViewFilterButtonsToString("", "", True, False, False)
	str = str & AvailableButtonTipBoxToString()
	
	str = str & "<div id=""copy-event-team-grid"">"
	str = str & "<table><tr>"
	str = str & "<td id=""copy-from-item"">" & ScheduleViewItemToString(evnt.EventId, evnt.ProgramId, evnt.ScheduleId, SCHEDULE_ITEM_TYPE_COPY_FROM)
	str = str & "</td>"
	str = str & "<td id=""copy-button"">"
	str = str & "<form method=""post"" id=""form-copy-event"" action=""/_incs/script/ajax/_event_team.asp"">"
	str = str & "<input type=""hidden"" name=""program_id"" value=""" & schedule.ProgramId & """ />"
	str = str & "<input type=""hidden"" name=""action"" value= """ & COPY_EVENT_TEAM_TO_EVENT & """ />"
	str = str & "</form>"
	str = str & "<a href=""#"" title=""Copy event team"">"
	str = str & "<img src=""/_images/icons/greenarrow_lg.gif"" alt=""Copy"" />"
	str = str & "</a></td>"
	str = str & "<td id=""copy-to-item"">" & ScheduleViewItemToString("", evnt.ProgramId, evnt.ScheduleId, SCHEDULE_ITEM_TYPE_COPY_TO) & "</td>"
	
	str = str & "</tr></table></div>"

	DuplicateEventTeamGridToString = str
End Function

Function EditEventTeamGridToString(page)
	Dim str, i
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim evnt				: Set evnt = New cEvent
	evnt.EventId = page.EventID
	evnt.Load()
	
	Dim schedule			: Set schedule = New cSchedule
	schedule.ScheduleId = evnt.ScheduleId
	Dim eventList			: eventList = schedule.EventList("")
	
	str = str & "<h3>Event team for: " & Server.HtmlEncode(evnt.EventName) & "<span style=""font-weight:normal;font-size:.8em;""> on " & WeekdayName(Weekday(evnt.EventDate), True) & " " & Day(evnt.EventDate) & " " & MonthName(Month(evnt.EventDate), True) & " " & Year(evnt.EventDate) & "</span></h3>"
	str = str & "<h4 class=""first"">" & Server.htmlEncode(evnt.ScheduleName) & "</h4>"
	str = str & "<p>Click on a skill in the list to add or remove members for that skill from this event team. </p>"
	str = str & AvailabilityWidgetToString(evnt.ScheduleId, "on")
	str = str & ViewFilterButtonsToString("", "", True, False, False)
	str = str & AvailableButtonTipBoxToString()
	
	str = str & "<div id=""copy-event-team-grid"">"	
	str = str & "<table><tr>"
	str = str & "<td id=""edit-item"">" & ScheduleViewItemToString(evnt.EventID, evnt.ProgramId, evnt.ScheduleId, SCHEDULE_ITEM_TYPE_EDITOR) & "</td>"
	str = str & "<td id=""editor"">" & TeamAccordionToString(evnt.EventId) & "</td>"
	str = str & "</tr></table></div>"

	EditEventTeamGridToString = str
End Function

Function GoToScheduleDropdownToString(page, scheduleId)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	Dim program			: Set program = New cProgram
	program.ProgramId = page.Schedule.ProgramId
	
	Dim list			: list = program.ScheduleList()
	
	Dim disabled
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-go-to-schedule"">"
	str = str & "<input type=""hidden"" name=""form_go_to_schedule_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""new_schedule_id"" id=""go-to-schedule-dropdown"">"
	str = str & "<option value="""">" & html("< Go to schedule >") & "</option>"
	str = str & "<option value="""">" & html("--") & "</option>"
	
	' 0-ScheduleID 1-ScheduleName 2-ScheduleDesc 3-DateCreated 4-DateModified 5-IsVisible
	' 6-DatePublished 7-HasUnpublishedChanges 8-EventCount
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			disabled = ""		: If CStr(scheduleId & "") = CStr(list(0,i) & "") Then disabled = " disabled=""disabled"""
		
			str = str & "<option value=""" & list(0,i) & """" & disabled & ">" & Server.HTMLEncode(list(1,i)) & "</option>"
		Next
	End If
	str = str & "</select></form></li>"
	
	GoToScheduleDropdownToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href
	
	Dim evnt		: Set evnt = New cEvent
	evnt.EventId = page.EventId
	If Len(page.EventId) > 0 Then evnt.Load()
	
	Dim programLink
	If pg.ScheduleId <> pg.FilterScheduleId Then pg.FilterScheduleId = ""
	pg.ProgramId = pg.Evnt.ProgramID: pg.Action = "": pg.EventId = "": pg.ScheduleId = ""
	href = "/schedule/schedules.asp" & pg.UrlParamsToSTring(True)
	programLink = "<a href=""" & href & """>" & html(page.Evnt.ProgramName) & "</a> / "
	
	Dim scheduleLink
	pg.ProgramId = pg.Evnt.ProgramId: pg.Action = "": pg.EventId = "": pg.ScheduleId = "": pg.FilterScheduleId = pg.Evnt.ScheduleId
	scheduleLink = "<a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>" & html(page.Evnt.ScheduleName) & "</a> / "

	'reset pg ..
	Set pg = page.Clone()
		
	Dim scheduleRootLink
	If pg.ScheduleId <> pg.FilterScheduleId Then pg.FilterScheduleId = ""
	pg.ProgramId = "": pg.Action = "": pg.EventId = "": pg.ScheduleId = ""
	href = "/schedule/schedules.asp" & pg.UrlParamsToSTring(True)
	scheduleRootLink = "<a href=""" & href & """>Schedules</a> / "
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case UPDATE_RECORD
			str = str & scheduleRootLink & programLink & scheduleLink & html(evnt.EventName) & " event team"
		Case COPY_EVENT_TEAM_TO_EVENT
			str = str & scheduleRootLink & programLink & scheduleLink & "Copy event team"
		Case Else
			str = str & scheduleRootLink & server.HTMLEncode(page.Schedule.ScheduleName)
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Dim publishScheduleButton
	pg.Action = PUBLISH_SCHEDULE
	href = pg.Url & pg.UrlParamsToString(True)
	publishScheduleButton = publishScheduleButton & "<li><a href=""" & href & """><img src=""/_images/icons/arrow_rotate_clockwise.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Publish all</a></li>"
	
	Dim eventGridButton
	pg.Action = "": pg.EventId = ""
	href = "/schedule/events.asp" & pg.UrlParamsToString(True)
	eventGridButton = eventGridButton & "<li><a href=""" & href & """><img src=""/_images/icons/event_multiple_2.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Event list</a></li>"
	
	Dim availabilityViewButton
	pg.EventId = "": pg.ProgramId = pg.Schedule.ProgramId
	href = "/schedule/availability.asp" & pg.UrlParamsToString(True)
	availabilityViewButton = "<li><a href=""" & href & """><img src=""/_images/icons/clock.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Availability view</a></li>"
	
	Dim masterTeamViewButton
	pg.EventId = "": pg.ScheduleId = pg.Evnt.ScheduleId
	href = pg.Url & pg.UrlParamsToString(True)
	masterTeamViewButton = "<li><a href=""" & href & """><img src=""/_images/icons/group.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Team view</a></li>"
	
	Dim schedulesButton
	pg.EventID = "": pg.Action = "": pg.ScheduleID = "": pg.ReturnContext = ""
	href = "/schedule/schedules.asp" & pg.UrlParamsToString(True)
	schedulesButton = schedulesButton & "<li><a href=""" & href & """><img src=""/_images/icons/calendar.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Schedules</a></li>"
	
	Select Case page.Action
		Case COPY_EVENT_TEAM_TO_EVENT
			str = str & schedulesButton & masterTeamViewButton & eventGridButton
		Case UPDATE_RECORD
			str = str & schedulesButton & masterTeamViewButton & eventGridButton
		Case Else
			str = str & publishScheduleButton & GoToScheduleDropdownToString(page, page.Schedule.ScheduleId) & availabilityViewButton & eventGridButton & schedulesButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EventTeamToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EventTeamMembersForSkillToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_ScheduleViewItemToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_ScheduleDropdownOptionsToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EventDropdownOptionsToString.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_TeamAccordionToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_AvailableSelectToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_NotAvailableSelectToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_MemberNotesToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/fn_ScheduledSelectToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/sub_SetScheduledOptionsList.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/team_accordion/sub_SetUnscheduledOptionsLists.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/member_event_widget/fn_AvailabilityWidgetToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/member_event_widget/fn_OptionListForAvailabilityWidgetToString.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_build_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	Public ReturnContext
	Public Year
	Public Month
	Public Day
	Public ShowPastEvents
	
	' encrypted
	Public Action
	Public ProgramID
	Public ScheduleID
	Public FilterScheduleId
	Public EventID

	' objects
	Public Member
	Public Client
	Public Program
	Public Schedule
	Public Evnt
	
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
		If Len(Year) > 0 Then str = str & "y=" & Year & amp
		If Len(Month) > 0 Then str = str & "m=" & Month & amp
		If Len(Day) > 0 Then str = str & "d=" & Day & amp
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(ScheduleID) > 0 Then str = str & "scid=" & Encrypt(ScheduleID) & amp
		If Len(FilterScheduleId) > 0 Then str = str & "fscid=" & Encrypt(FilterScheduleId) & amp
		If Len(EventID) > 0 Then str = str & "eid=" & Encrypt(EventID) & amp
		
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
		c.Year = Year
		c.Month = Month
		c.Day = Day

		c.Action = Action
		c.ProgramID = ProgramID
		c.ScheduleID = ScheduleID
		c.FilterScheduleId = FilterScheduleId
		c.EventID = EventID
				
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Schedule = Schedule
		Set c.Evnt = Evnt
		
		Set Clone = c
	End Function
End Class
%>

