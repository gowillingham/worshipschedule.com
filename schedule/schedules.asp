<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Const NUMBER_OF_EVENT_COLUMNS = 4

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
	
	' always clear return context for this page ..
	page.ReturnContext = ""
	
	If Request.Form("form_program_dropdown_is_postback") = IS_POSTBACK Then
		page.ProgramId = Request.Form("new_program_id")
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
		<link rel="stylesheet" type="text/css" href="/_incs/script/jquery/plugins/colorpicker/syronex-colorpicker.css" />
		<link rel="stylesheet" type="text/css" href="/_incs/script/jquery/jquery-ui-1.6.custom/development-bundle/themes/start/ui.all.css" />	
		<link rel="stylesheet" type="text/css" href="schedules.css" />
		
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>

		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/colorpicker/syronex-colorpicker.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/jquery-ui-1.6.custom/js/jquery-ui-1.6.custom.min.js"></script>
		<script language="javascript" type="text/javascript" src="schedules.js"></script>

		<script language="javascript" type="text/javascript">
			// translate server side constants to jscript ..
			var PUBLISH_SCHEDULE_ENCRYPTED			= "<%=Encrypt(PUBLISH_SCHEDULE) %>"
			var UNPUBLISH_SCHEDULE_ENCRYPTED		= "<%=Encrypt(UNPUBLISH_SCHEDULE) %>"
			var SCHEDULE_HTML_BACKGROUND_COLOR		= '<%=Application.Value("SCHEDULE_HTML_BACKGROUND_COLOR") %>'
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
		Case SHOW_SCHEDULE_DETAILS
			str = str & ScheduleSummaryToString(page)
		
		Case TOGGLE_IS_VISIBLE
			Call DoToggleIsVisible(page.Schedule, rv)
			If CInt(page.Schedule.IsVisible) = 1 Then
				page.MessageId = 6056
			Else
				page.MessageId = 6055
			End If
			
			page.Action = "": page.ScheduleId = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case PUBLISH_SCHEDULE
			Call DoPublishSchedule(page.ScheduleId, page.Member.NameLast & ", " & page.Member.NameFirst, rv)
			page.MessageId = 6019: page.Action = "": page.ScheduleId = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case UNPUBLISH_SCHEDULE
			Call DoUnpublishSchedule(page.ScheduleId, rv)
			page.MessageId = 6021: page.Action = "": page.ScheduleId = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
	
		Case UPDATE_RECORD
			If Request.Form("form_schedule_is_postback") = IS_POSTBACK Then
				Call LoadScheduleFromRequest(page.Schedule)
				If ValidSchedule(page.Schedule) Then
					Call DoUpdateSchedule(page.Schedule, rv)
					Select Case rv
						Case 0
							page.MessageID = 6000
						Case Else
							page.MessageID = 6001
					End Select
					If page.FilterScheduleId <> page.ScheduleId Then page.FilterScheduleId = ""
					page.Action = "": page.ScheduleID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormScheduleToString(page)
				End If
			Else
				str = str & FormScheduleToString(page)
			End If
			
		Case ADDNEW_RECORD
			If Request.Form("form_schedule_is_postback") = IS_POSTBACK Then
				Call LoadScheduleFromRequest(page.Schedule)
				If ValidSchedule(page.Schedule) Then
					Call DoInsertSchedule(page.Schedule, rv)
					Select Case rv
						Case 0
							page.MessageID = 6002
						Case Else
							page.MessageID = 6001
					End Select
					If page.FilterScheduleId <> page.ScheduleId Then page.FilterScheduleId = ""
					page.Action = "": page.ScheduleID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormScheduleToString(page)
				End If
			Else
				str = str & FormScheduleToString(page)
			End If
			
		Case DELETE_RECORD
			If Request.Form("form_confirm_delete_schedule_is_postback") = IS_POSTBACK Then
				Call DoDeleteSchedule(page.Schedule, rv)
					Select Case rv
						Case 0
							page.MessageID = 6003
						Case Else
							page.MessageID = 6006
					End Select
					If page.FilterScheduleId <> page.ScheduleId Then page.FilterScheduleId = ""
					page.Action = "": page.ScheduleID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmDeleteScheduleToString(page)
			End If
			
		Case Else
			If Not page.Client.HasPrograms Then
				str = str & NoProgramsDialogToString(page)
				Call DoDisableTabLinkBar()
			ElseIf Not page.Client.HasSchedules Then
				str = str & NoSchedulesDialogToString(page)
			Else
				str = str & CalendarToString(page)
			End If
			
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoDisableTablinkBar()
	m_tabLinkBarText = "<li>&nbsp;</li>"
End Sub

Sub DoToggleIsVisible(schedule, outError)
	If schedule.IsVisible = 0 Then
		schedule.IsVisible = 1
	Else
		schedule.IsVisible = 0
	End If
	Call schedule.Save(outError)
End Sub

Sub DoUnpublishSchedule(scheduleId, outError)
	Dim schedule		: Set schedule = New cSchedule
	
	schedule.ScheduleId = scheduleId
	Call schedule.RemovePublish(outError)
End Sub

Sub DoPublishSchedule(scheduleId, publisher, outError)
	Dim schedule		: Set schedule = New cSchedule
	
	schedule.ScheduleId = scheduleId
	Call schedule.Publish(publisher, outError)
End Sub

Sub DoUpdateSchedule(schedule, outError)
	Call schedule.Save(outError)	
End Sub

Sub DoInsertSchedule(schedule, outError)
	Call schedule.Add(outError)
End Sub

Sub DoDeleteSchedule(schedule, outError)
	Call schedule.Delete(outError)
End Sub

Function EventGridForScheduleSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim list				: list = page.Schedule.EventList("")
	Dim count				: count = 0
	Dim alt					: alt = ""
	Dim rows				: rows = ""
	
	Dim publishedText	
	Dim href		
	
	' 0-EventID 1-EventName 2-EventDate 3-EventNote 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified
	' 8-ProgramID 9-ProgramName 10-ScheduleID 11-ScheduleName 12-ScheduleEventCount
	' 13-ScheduleIsVisible 14-HasUnpublishedChanges 15-EventFileCount 16-ScheduledMemberCount
	' 17-HtmlBackgroundColor 18-FileListXMLFragment
		
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			alt = ""				: If count Mod 2 <> 0 Then alt=" class=""alt"""
			
			publishedText = "Yes"	: If list(14,i) = 1 Then publishedText = "<span style=""color:red"">No</span>"
			
			pg.Action = SHOW_EVENT_DETAILS: pg.EventId = list(0,i):
			href = "/schedule/events.asp" & pg.UrlParamsToString(True)
			
			rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
			rows = rows & "<strong>" & html(list(11,i)) & "</strong> | "
			rows = rows & "<a href=""" & href & """ title=""Details""><strong>" & html(list(1,i)) & "</strong></a></td>"
			rows = rows & "<td>" & dateTime.Convert(list(2,i), "MM-DD-YYYY")
			If Len(list(4,i) & "") > 0 Then rows = rows & " at " & dateTime.Convert(list(4,i), "hh:00 pp")
			rows = rows & "</td>"
			rows = rows & "<td>" & publishedText & "</td>"
			rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
			rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
			
			count = count + 1
		Next
	End If
	
	If count > 0 Then
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Event</th><th>When</th><th>Published</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = str & "<p class=""alert"">This schedule does not have any events. </p>"
	End If
	
	EventGridForScheduleSummaryToString = str
End Function

Function GetUniqueMemberIdListForSchedule(list)
	Dim str, i
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberIsEnabled 5-ProgramMemberIsEnabled 
	' 6-EventId 7-EventName 8-EventDate 9-TimeStart 10-TimeEnd 11-EventNote 12-SkillListXmlFragment
	' 13-FileListXmlFragment
	
	' iterate and create comma delim string of IDs
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isMemberEnabled = True				: If list(4,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True		: If list(5,i) = 0 Then isProgramMemberEnabled = False
			
			If isMemberEnabled And isProgramMemberEnabled Then
				str = str & list(0,i) & ","
			End If
		Next
	End If
	
	' remove dupes
	If Len(str) > 0 Then
		str = Left(str, Len(str) - 1)
		str = RemoveDupesFromStringList(str)
	End If
	
	GetUniqueMemberIdListForSchedule = str
End Function

Function MemberGridForScheduleSummaryToString(page)
	Dim str, i, j
	Dim pg					: Set pg = page.Clone()
	
	Dim list				: list = page.Schedule.EventTeamDetailsList()
	
	Dim uniqueList			: uniqueList = Split(GetUniqueMemberIdListForSchedule(list), ",")
	Dim count				: count = 0
	Dim alt					: alt = ""
	Dim rows				: rows = ""
	Dim href
	
	Dim isFound
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberIsEnabled 5-ProgramMemberIsEnabled 
	' 6-EventId 7-EventName 8-EventDate 9-TimeStart 10-TimeEnd 11-EventNote 12-SkillListXmlFragment
	' 13-FileListXmlFragment 14-ProgramMemberId
	
	If IsArray(uniqueList) Then
		For i = 0 To UBound(uniqueList)
		
			isFound = False
			For j = 0 To UBound(list,2)
				If Not isFound Then
					If CStr(uniqueList(i) & "") = CStr(list(0,j) & "") Then
						alt = ""				: If count Mod 2 > 0 Then alt = " class=""alt"""
						
						pg.Action = SHOW_PROGRAM_MEMBER_DETAILS: pg.MemberId = list(0,j): pg.ProgramMemberId = list(14,j)
						href = "/admin/profile.asp" & pg.UrlParamsToString(True)
						
						rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/user.png"" alt="""" />"
						rows = rows & "<strong>" & html(page.Schedule.ScheduleName) & "</strong> | "
						rows = rows & "<a href=""" & href & """><strong>" & html(list(1,j) & ", " & list(2,j)) & "</strong></a></td>"
						rows = rows & "<td class=""toolbar""><a href=""" & href & """>"
						rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
						
						isFound = True
						count = count + 1
					End If
				End If
			Next
		Next
	End If
				
	If count > 0 Then
		str = str & "<p>This list of program members belongs to at least one event team for this schedule. </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Member</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = str & "<p class=""alert"">No members have been assigned to any event teams for this schedule. </p>"
	End If
	
	MemberGridForScheduleSummaryToString = str
End Function

Function AvailabilityGridForScheduleSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	
	Dim list				: list = page.Schedule.AvailabilityList()
	Dim alt
	Dim href
	Dim rows
	Dim count				: count = 0
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	Dim isMissingInfo
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-LastLogin 4-Email 5-IsMemberAccountEnabled
	' 6-IsMissingAvailabilityInfo 7-IsProgramMemberEnabled 8-ProgramMemberId
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isMemberEnabled = True					: If list(5,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True			: If list(7,i) = 0 Then isProgramMemberEnabled = False
			isMissingInfo = True					: If list(6,i) = 0 Then isMissingInfo = False

			If isMemberEnabled And isProgramMemberEnabled And IsMissingInfo Then
				alt = "":					If count Mod 2 > 0 Then alt = " class=""alt"""
				
				pg.Action = SHOW_PROGRAM_MEMBER_DETAILS: pg.MemberId = list(0,i): pg.ProgramMemberId = list(8,i)
				href = "/admin/profile.asp" & pg.UrlParamsToString(True)
				
				rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/user.png"" alt="""" />"
				rows = rows & "<strong>" & html(page.Schedule.ScheduleName) & "</strong> | "
				rows = rows & "<a href=""" & href & """ title=""Details""><strong>" & html(list(1,i) & ", " & list(2,i)) & "</strong></a></td>"
				rows = rows & "<td>??</td>"
				rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
				rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
				
				count = count + 1
			End If
		Next
	Else
	
	End If				
	
	If count > 0 Then
		str = str & "<p>This list of program members has not logged in with up-to-date availability info for some or all of the events on this schedule. </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Member</th><th>Available</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = "<p class=""alert"">All members have logged in with their up-to-date availability info for the events on this schedule. </p>"
	End If
	
	AvailabilityGridForScheduleSummaryToString = str
End Function

Function ScheduleSummaryToString(page)
	Dim str
	Dim dateTime				: Set dateTime = New cFormatDate
	
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & html(page.Schedule.ScheduleName) & "</h3>"
	
	str = str & "<h5 class=""description"">Description</h5>"
	If Len(page.Schedule.ScheduleDesc) > 0 Then
		str = str & "<p>" & html(page.Schedule.ScheduleDesc) & "</p>"
	Else
		str = str & "<p class=""alert"">No description available. </p>"
	End If
	
	If page.Schedule.PublishStatus = SCHEDULE_HAS_UNPUBLISHED_CHANGES Then
		str = str & "<h5 class=""not-published"">Unpublished changes</h5>"
		str = str & "<p class=""alert"">You have made changes to one or more event teams for this schedule but have not published those changes to your member calendar. </p>"
	ElseIf page.Schedule.PublishStatus = SCHEDULE_HAS_NO_SCHEDULE_INFORMATION Then
		' do nothing for now ..
	End If
	
	If page.Schedule.IsVisible = 0 Then
		str = str & "<h5 class=""not-visible"">Not visible</h5>"
		str = str & "<p class=""alert"">You have set this schedule to not visible. "
		str = str & "Events on this schedule will not show in your member calendars and availability lists. </p>"
	End If	
	
	str = str & "<h5 class=""schedule"">Event list</h5>"
	str = str & EventGridForScheduleSummaryToString(page)
	
	str = str & "<h5 class=""program-member"">Member listing</h5>"
	str = str & MemberGridForScheduleSummaryToString(page)
	
	str = str & "<h5 class=""availability"">Member availability</h5>"
	str = str & AvailabilityGridForScheduleSummaryToString(page)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul><li>Created on " & dateTime.Convert(page.Schedule.DateCreated, "DDDD MMMM dd, YYYY around hh:nn pp") & ". </li>"
	str = str & "</ul>"
	
	str = str & "</div>"
	
	ScheduleSummaryToString = str
End Function

Function NoSchedulesDialogToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim dialog			: Set dialog = New cDialog
	
	dialog.Headline = "Ok, something's missing here .."
	
	dialog.Text = dialog.Text & "<p>You are in the Manage Schedules section of your account, "
	dialog.Text = dialog.Text & "but you were expecting to see a nicely formatted calendar of your account's schedules and events. "
	dialog.Text = dialog.Text & "Well, it looks like there aren't any schedules set up for your account yet - that's easily fixed! </p>"
	dialog.Text = dialog.Text & "<p>Get started by clicking <strong>Create my first schedule</strong>. </p>"

	dialog.SubText = dialog.SubText & "<p>Once you have created your first schedule, "
	dialog.SubText = dialog.SubText & "this page will show you a list of all your schedules and a "
	dialog.subText = dialog.SubText & "master calendar of events and event teams (the members you schedule for your events). </p>"

	pg.Action = ADDNEW_RECORD: pg.ScheduleId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Create your first schedule</a></li>"

	NoSchedulesDialogToString = dialog.ToString()
End Function

Function NoProgramsDialogToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim dialog			: Set dialog = New cDialog
	
	dialog.HeadLine = "Whoa, where are my schedules ..?"

	dialog.Text = dialog.Text & "<p>You are in the Manage Schedules section of your account, "
	dialog.Text = dialog.Text & "and you're probably wondering why you're not seeing a calendar displaying your account's events. "
	dialog.Text = dialog.Text & "Either you haven't set up any programs in your account to manage, "
	dialog.Text = dialog.Text & "or you have all of your programs set to disabled. </p>"
	dialog.Text = dialog.Text & "<p>You can get started fixing this by clicking <strong>Create your first program</strong>. </p>"
	
	dialog.SubText = dialog.SubText & "<p>Once you start creating programs and adding your members to them, "
	dialog.SubText = dialog.SubText & "this page will show a calendar where you can manage your schedules and events "
	dialog.SubText = dialog.SubText & "(and then assign your members to them). "
	
	pg.Action = ADDNEW_RECORD: pg.ProgramId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/programs.asp" & pg.UrlParamsToString(True) & """>Create your first program</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/programs.asp"">Check for disabled programs</a></li>"
	
	NoProgramsDialogToString = dialog.ToString()
End Function

Function NoSkillsDialogToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim dialog			: Set dialog = New cDialog
	
	dialog.HeadLine = "Wait, where are my event teams ..?"
	
	dialog.Text = dialog.Text & "<p>You are currently in the Manage Event Teams section of your account. "
	dialog.Text = dialog.Text & "Either this schedule (" & html(page.Schedule.ScheduleName) & ") belongs to a program (" & html(page.Schedule.ProgramName) & ") that doesn't have any skills set up yet, "
	dialog.Text = dialog.Text & "or you have disabled all of the skills for this program. </p>"
	
	dialog.SubText = dialog.SubText & "<p>Once you have set up some skills and members for the " & html(page.Schedule.ProgramName) & " program, "
	dialog.SubText = dialog.SubText & "this page will show you who you have scheduled for each event (and what they will be doing). </p>"

	pg.Action = ADDNEW_RECORD: pg.ProgramId = page.Schedule.ProgramId: pg.ScheduleId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>Create your first skill</a></li>"
	
	Set pg = page.Clone()
	pg.Action = "": pg.ProgramId = page.Schedule.ProgramId
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>Check for disabled skills</a></li>"

	NoSkillsDialogToString = dialog.ToString()
End Function

Function ScheduleListToString(page)
	Dim str, i, j
	Dim pg				: Set pg = page.Clone()
	Dim schedule		: Set schedule = New cSchedule
	
	Dim programIdList
	Dim ownedProgramList
	Dim list
	
	' display all schedules in list always ..
	ownedProgramList = page.Member.OwnedProgramsList()
	If IsArray(ownedProgramList) Then
		Redim programIdList(UBound(ownedProgramList,2))
		For i = 0 To UBound(programIdList)
			programIdList(i) = ownedProgramList(0,i)
		Next
	End If
	
	str = str & "<ul id=""schedule-list"">"
	
	' here's the header for schedule list ..
	str = str & "<li class=""schedule-list-header"">"
	str = str & "Schedules"
	str = str & "</li>"
	
	' for each program in the list get a list of schedules ..
	For i = 0 To UBound(programIdList)
		
		' 0-ScheduleId 1-ScheduleName 2-ScheduleDesc 3-IsVisible 4-DateCreated 5-DateModified
		' 6-ProgramID 7-ProgramName 8-HtmlBackgroundColor 9-EventCount 

		schedule.ProgramID = programIDList(i)
		list = schedule.List()
		If IsArray(list) Then
			For j = 0 To UBound(list,2)
			
				' reset pg as some values are cleared in this loop ..
				Set pg = page.Clone()
				
				str = str & "<li style=""background-color:" & list(8,j) & """>"
				pg.Action = "": pg.FilterScheduleId = list(0,j): pg.ProgramId = list(6,j)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>"
				str = str & "<strong>" & html(list(7,j)) & ": </strong>" & html(list(1,j)) & "</a> (" & list(9,j) & ")"
				str = str & "<div class=""toolbar"">"
				
				' reset as I need to get the old page.FilterScheduleId if there is one ..
				Set pg = page.Clone()
				
				
				pg.ProgramId = ""
				' details
				pg.Action = SHOW_SCHEDULE_DETAILS: pg.ScheduleId = list(0,j)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details"">"
				str = str & "<img class=""icon"" src=""/_images/icons/magnifier.png"" alt="""" /></a>"
				' edit
				pg.Action = UPDATE_RECORD: pg.ScheduleID = list(0,j)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Edit Schedule"">"
				str = str & "<img class=""icon"" src=""/_images/icons/pencil.png"" alt="""" /></a>"
				' insert event
				pg.Action = ADDNEW_RECORD: pg.ScheduleID = list(0,j)
				str = str & "<a href=""/schedule/events.asp" & pg.UrlParamsToString(True) & """ title=""Add Event"">"
				str = str & "<img class=""icon"" src=""/_images/icons/date_add.png"" alt="""" /></a>"
				' event grid
				pg.Action = "": pg.ScheduleID = list(0,j): pg.ReturnContext = CONTEXT_EVENT_MANAGER
				str = str & "<a href=""/schedule/events.asp" & pg.UrlParamsToString(True) & """ title=""Event List"">"
				str = str & "<img class=""icon"" src=""/_images/icons/event_multiple_2.png"" alt="""" /></a>"
				' event teams
				pg.Action = "": pg.ScheduleID = list(0,j): pg.ShowPastEvents = "": pg.ReturnContext = ""
				str = str & "<a href=""/schedule/teams.asp" & pg.UrlParamsToString(True) & """ title=""Event Teams"">"
				str = str & "<img class=""icon"" src=""/_images/icons/group.png"" alt="""" /></a>"
				' availability
				pg.Action = "": pg.ScheduleId = list(0,j): pg.ReturnContext = "": pg.ProgramId = list(6,j)
				str = str & "<a href=""/schedule/availability.asp" & pg.UrlParamsToString(True) & """ title=""Team availability"">"
				str = str & "<img class=""icon"" src=""/_images/icons/clock.png"" alt="""" /></a>"
				' publish/unpublish (intercepted by ajax call that appends act=?? to href) ..
				pg.Action = "": pg.ScheduleId = list(0,j)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Publish schedule"" class=""publish-schedule-link"">"
				str = str & "<img class=""icon"" src=""/_images/icons/arrow_rotate_clockwise.png"" alt="""" /></a>"
				' email schedule 
				pg.Action = SEND_SCHEDULE_BY_EMAIL: pg.ReturnContext = ""
				str = str & "<a href=""/schedule/email.asp" & pg.UrlParamsToString(True) & """ title=""Send message"">"
				str = str & "<img class=""icon"" src=""/_images/icons/email.png"" alt="""" /></a>"
				' delete
				pg.Action = DELETE_RECORD: pg.ReturnContext = ""
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove Schedule"">"
				str = str & "<img class=""icon"" src=""/_images/icons/cross.png"" alt="""" /></a>"
				str = str & "</div>"
				str = str & "</li>"
			Next
		End If
	Next
	
	str = str & "<li class=""tip-box"">"
	str = str & "<h3>Tip!</h3>"
	str = str & "<p>Click on any schedule in the list to hide the others on your calendar. </p>"
	str = str & "</li>"
	
	str = str & "</ul>"
	
	ScheduleListToString = str
End Function

Function EventItemToString(page, list, idx)
	Dim str
	Dim dateTime			: Set dateTime = New cFormatDate
	Dim pg					: Set pg = page.Clone()
	
	str = str & "<strong>" & html(list(9,idx)) & ": </strong>" & html(list(1,idx))
	If Len(list(4,idx)) > 0 Then
		str = str & "<br />" & dateTime.Convert(list(4,idx), "hh:nnpx")
	End If
	
	str = str & "<div class=""toolbar"">"
	pg.Action = SHOW_EVENT_DETAILS: pg.ScheduleId = list(10,idx): pg.EventId = list(0,idx)
	str = str & "<a href=""/schedule/events.asp" & pg.UrlParamsToString(True) & """ title=""View Details"">"
	str = str & "<img class=""icon"" src=""/_images/icons/magnifier.png"" alt="""" /></a>"
	pg.ScheduleId = list(10,idx): pg.Action = UPDATE_RECORD: pg.EventID = list(0,idx)
	str = str & "<a href=""/schedule/events.asp" & pg.UrlParamsToString(True) & """ title=""Edit Event"">"
	str = str & "<img class=""icon"" src=""/_images/icons/pencil.png"" alt="""" /></a>"
	pg.Action = DUPLICATE_EVENT: pg.EventID = list(0,idx)
	str = str & "<a href=""/schedule/events.asp" & pg.UrlParamsToString(True) & """ title=""Copy Event"">"
	str = str & "<img class=""icon extra-room"" src=""/_images/icons/paste_date.png"" alt="""" /></a>"
	pg.Action = UPDATE_RECORD: pg.EventID = list(0,idx)
	str = str & "<a href=""/schedule/teams.asp" & pg.UrlParamsToString(True) & """ title=""Event Team"">"
	str = str & "<img class=""icon"" src=""/_images/icons/group_edit.png"" alt="""" /></a>"
	pg.Action = DELETE_RECORD: pg.EventID = list(0,idx)
	str = str & "<a href=""/schedule/events.asp" & pg.UrlParamsToString(True) & """ title=""Remove Event"">"
	str = str & "<img class=""icon"" src=""/_images/icons/cross.png"" alt="""" /></a>"
	str = str & "</div>"
	
	EventItemToString = str
End Function

Function CalendarToString(page)
	Dim str, i
	Dim pg			: Set pg = page.Clone()
	
	Dim cal			: Set cal = New gtdCalendar
	
	' 0-EventID 1-EventName 2-EventDate 3-EventNote 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified
	' 8-ProgramID 9-ProgramName 10-ScheduleID 11-ScheduleName 12-ScheduleEventCount
	' 13-ScheduleIsVisible 14-HasUnpublishedChanges 15-EventFileCount 16-ScheduledMemberCount
	' 17-HtmlBackgroundColor
	
	Dim list		: list = page.Member.AdminEventList(page.Program.ProgramID, page.FilterScheduleId)
	Dim item
	
	str = str & ScheduleListToString(page)
	str = str & m_appMessageText
	
	Call cal.SetDate(page.Year, page.Month)
	
	cal.AddNavUrlParams "pid", Encrypt(page.Program.ProgramID)
	cal.AddNavUrlParams "scid", Encrypt(page.Schedule.ScheduleID)
	cal.AddNavUrlParams "act", Encrypt(page.Action)
	
	cal.DisplayHeader = True
	cal.DisplayTopNav = True
	cal.DisplayDayNumbersAsLinks = False
	cal.DisplayWeekdayRowStyle = 3
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			item = EventItemToString(page, list, i)
			Call cal.AddItem(list(2,i), item, "background-color:" & list(17,i) & ";")
		Next	
	End If
	
	str = str & "<div>"
	str = str & cal.ToString()
	str = str & "</div>"
		
	CalendarToString = str 
End Function

Function FormConfirmDeleteScheduleToString(page)
	Dim str, msg
	Dim pg				: Set pg = page.Clone()
	
	msg = msg & "You will permanently remove the schedule <strong>" & html(page.Schedule.ScheduleName) & "</strong> from the " & html(page.Schedule.ProgramName) & " program. "
	msg = msg & "You will also be removing the events (" & page.Schedule.EventCount & "), calendar, and/or schedule information that belongs to this schedule. "
	msg = msg & "This action cannot be reversed. "
	
	str = str & CustomApplicationMessageToString("Please confirm this action!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-confirm-delete-schedule"">"
	str = str & "<input type=""hidden"" name=""form_confirm_delete_schedule_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = "": pg.ScheduleId = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "</p></form>"
	
	FormConfirmDeleteScheduleToString = str
End Function

Function ValidSchedule(schedule)
	ValidSchedule = True
	If Application.Value("FORM_VALIDATION_OFF") Then Exit Function	

	If Len(schedule.ProgramID) = 0 Then
		AddCustomFrmError("You need to select a program for your schedule to belong to.")
		ValidSchedule = False
	End If
	If Not ValidData(schedule.ScheduleName, True, 0, 100, "Schedule Name", "") Then ValidSchedule = False
	If Not ValidData(schedule.ScheduleDesc, False, 0, 1000, "Description", "") Then ValidSchedule = False
End Function

Sub LoadScheduleFromRequest(schedule)
	schedule.ScheduleName = Request.Form("schedule_name")
	schedule.ScheduleDesc = Request.Form("schedule_desc")
	schedule.ProgramID = Request.Form("program_id")
	schedule.HtmlBackgroundColor = Request.Form("html_background_color")
	schedule.IsVisible = Request.Form("is_visible")
End Sub

Function OwnedProgramDropdownToString(page, id)
	Dim str, i
	
	Dim list			: list = page.Member.OwnedProgramsList()
	Dim selected		: selected = ""
	
	' select current programId if no id is passed in
	Dim selectedId		: selectedId = id
	If Len(selectedId) = 0 Then selectedId = page.Program.ProgramID
	
	Dim disabled		: disabled = ""
	If page.Action = UPDATE_RECORD Then 
		disabled = " disabled=""disabled"" class=""disabled"""
		str = str & "<input type=""hidden"" name=""program_id"" value=""" & id & """ />"
	End If
	
	str = str & "<select name=""program_id""" & disabled & ">"
	str = str & "<option value="""">" & html("< Select a program >") & "</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			selected = ""
			If CStr(list(0,i) & "") = CStr(selectedId & "") Then selected = " selected=""selected"""
			str = str & "<option value=""" & list(0,i) & """" & selected & ">" & html(list(1,i)) & "</option>"
		Next
	End If
	str = str & "</select>"
	
	OwnedProgramDropdownToString = str
End Function

Function FormScheduleToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-schedule"">"
	str = str & "<input type=""hidden"" name=""form_schedule_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><tbody>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Program") & "</td>"
	str = str & "<td>" & OwnedProgramDropdownToString(page, page.Schedule.ProgramID) & "</td></tr>"
	If page.Action = ADDNEW_RECORD Then
		str = str & "<tr><td>&nbsp;</td><td class=""hint"">The program that this schedule will belong <br />to. You won't be able to change this after this <br />schedule is saved. </td></tr>"
	End If
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Schedule Name") & "</td>"
	str = str & "<td><input type=""text"" class=""medium"" name=""schedule_name"" value=""" & html(page.Schedule.ScheduleName) & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Description") & "</td>"
	str = str & "<td><textarea class=""medium"" name=""schedule_desc"">" & html(page.Schedule.ScheduleDesc) & "</textarea></td></tr>"

	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Visible</td>"
	str = str & "<td>" & YesNoDropdownToString(page.Schedule.IsVisible, "is_visible") & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">If this is set to no, your members will not see "
	str = str & "<br />events for this schedule on their member <br />calendar (until you change it to yes). </td></tr>"

	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Color")
	str = str & "<input type=""hidden"" value=""" & page.Schedule.HtmlBackgroundColor & """ name=""html_background_color"" id=""html-color"" /></td>"
	str = str & "<td id=""color-picker""></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">Events for this schedule will be this color on <br />your calendar. </td></tr>"
	
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a></td></tr>"
	str = str & "</tbody></table></form></div>"
	
	FormScheduleToString = str
End Function

Function ProgramDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	Dim list			: list = page.Member.OwnedProgramsList()
	Dim selected		: selected = ""
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ id=""form-program-dropdown"" method=""post"">"
	str = str & "<input type=""hidden"" name=""form_program_dropdown_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""new_program_id"" id=""program-select"">"
	If Len(page.Program.ProgramId) > 0 Then
		str = str & "<option value="""">" & html("< Show all >") & "</option>"
	Else
		str = str & "<option value="""">" & html("< Choose program >") & "</option>"
	End If
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			selected = ""
			If CStr(list(0,i) & "") = CStr(page.Program.ProgramId & "") Then selected = " selected=""selected"""
			str = str & "<option value=""" & list(0,i) & """" & selected & ">" & html(list(1,i)) & "</option>"
		Next
	End If
	str = str & "</select></form></li>"
	
	ProgramDropdownToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim schedule	: Set schedule = New cSchedule
	schedule.ScheduleId = page.FilterScheduleId
	If Len(schedule.ScheduleId) > 0 Then schedule.Load()
	
	Dim thisScheduleLink
	pg.FilterScheduleId = page.ScheduleId
	pg.ScheduleId = "": pg.ProgramId = "": pg.Action = ""
	thisScheduleLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.Schedule.ScheduleName) & "</a> / "
	
	Dim programLink
	Set pg = page.Clone()
	pg.ScheduleId = "": pg.Action = "": pg.FilterScheduleId = ""
	If Len(page.ProgramId) > 0 Then
		programLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.Program.ProgramName) & "</a> / "
	ElseIf Len(page.ScheduleId) > 0 Then
		pg.ProgramId = page.Schedule.ProgramId
		programLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.Schedule.ProgramName) & "</a> / "
	End If
	
	Dim scheduleLink
	pg.ProgramId = "": pg.Action = "": pg.FilterScheduleId = ""
	scheduleLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Schedules</a> / "
	
	Dim rootLink
	If Len(page.ProgramId) > 0 Then 
		If Len(page.FilterScheduleId) > 0 Then
			rootLink = rootLink & scheduleLink & programLink & html(schedule.ScheduleName)
		Else
			rootLink = rootLink & scheduleLink & html(page.Program.ProgramName)
		End If
	ElseIf Len(page.FilterScheduleId) > 0 Then
		rootLink = rootLink & scheduleLink & html(schedule.ScheduleName)
	Else
		rootLink = rootLink & "Schedules"
	End If
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case SHOW_SCHEDULE_DETAILS
			str = str & scheduleLink & programLink & html(page.Schedule.ScheduleName)
		Case DISPLAY_MASTER_SCHEDULE
			str = str & scheduleLink & programLink & thisScheduleLink & "Event teams"
		Case DELETE_RECORD
			str = str & scheduleLink & programLink & "Remove '" & html(page.Schedule.ScheduleName) & "'"
		Case UPDATE_RECORD
			str = str & scheduleLink & programLink & "Edit '" & html(page.Schedule.ScheduleName) & "'"
		Case ADDNEW_RECORD
			str = str & scheduleLink & programLink & "New Schedule"
		Case Else
			str = str & rootLink
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Dim scheduleHomeButton
	pg.Action = ""
	href = pg.Url & pg.UrlParamsToString(True)
	scheduleHomeButton = "<li><a href=""" & href & """><img src=""/_images/icons/calendar.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Schedules</a></li>"
	
	Dim eventListButton
	pg.Action = ""
	href = "/schedule/events.asp" & pg.UrlParamsToString(True)
	eventListButton = "<li><a href=""" & href & """><img src=""/_images/icons/event_multiple_2.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Event List</a></li>"
	
	Dim eventTeamsButton
	pg.Action = ""
	href = "/schedule/teams.asp" & pg.UrlParamsToString(True)
	eventTeamsButton = "<li><a href=""" & href & """><img src=""/_images/icons/group.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Event Teams</a></li>"
	
	Dim publishScheduleButton
	pg.Action = PUBLISH_SCHEDULE
	href = pg.Url & pg.UrlParamsToString(True)
	publishScheduleButton = "<li><a href=""" & href & """><img src=""/_images/icons/arrow_rotate_clockwise.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Publish all</a></li>"

	Dim newScheduleButton
	pg.Action = ADDNEW_RECORD: pg.ScheduleID = ""
	href= pg.Url & pg.UrlParamsToString(True)
	newScheduleButton = "<li><a href=""" & href & """><img src=""/_images/icons/calendar_add.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>New Schedule</a></li>"
	
	Select Case page.Action
		Case SHOW_SCHEDULE_DETAILS
			str = str & scheduleHomeButton & eventListButton & eventTeamsButton
		Case DISPLAY_MASTER_SCHEDULE
			str = str & GoToScheduleDropDownToString(page, "") & publishScheduleButton & scheduleHomeButton & eventListButton
		Case UPDATE_RECORD
			str = str & scheduleHomeButton
		Case ADDNEW_RECORD
			str = str & scheduleHomeButton
		Case DELETE_RECORD
			str = str & scheduleHomeButton
		Case Else
			str = str & ProgramDropdownToString(page) & newScheduleButton
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

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_build_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_ScheduleDropdownOptionsToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_YesNoDropdownToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_RemoveDupesFromStringList.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/gtdCalendar/gtdCal.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/member_event_widget/fn_AvailabilityWidgetToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/member_event_widget/fn_OptionListForAvailabilityWidgetToString.asp"-->
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
	Public ProgramMemberId

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
		If Len(ProgramMemberId) > 0 Then str = str & "pmid=" & Encrypt(ProgramMemberId) & amp
		
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
		c.ProgramMemberId = ProgramMemberId
				
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Schedule = Schedule
		
		Set Clone = c
	End Function
End Class
%>

