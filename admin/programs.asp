<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-programs"
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
	page.Action = Decrypt(Request.QueryString("act"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.ScheduleID = Decrypt(Request.QueryString("scid"))
	page.EmailID = Decrypt(Request.QueryString("emid"))
	page.MemberID = Decrypt(Request.QueryString("mid"))
	page.EventId = Decrypt(Request.QueryString("eid"))
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then page.Program.Load()
	
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
		<link rel="stylesheet" type="text/css" href="programs.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=m_pageTitleText%></title>
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="programs.js"></script>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)

	Select Case page.Action
		Case SHOW_PROGRAM_DETAILS
			str = str & ProgramSummaryToString(page)
		
		Case ADDNEW_RECORD
			If Request.Form("FormProgramIsPostback") = IS_POSTBACK Then
				Call LoadProgramFromForm(page.Program)
				If ValidProgram(page.Program) Then
					Call DoInsertProgram(page, rv)
					Select Case rv
						Case 0
							' success
							page.MessageID = 3000
						Case -2
							' dupe program
							page.MessageID = 3001
						Case Else
							' unknown error	
							page.MessageID = 3002
					End Select
					page.Action = "": page.ProgramID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormProgramToString(page)
				End If
			Else
				str = str & FormProgramToString(page)
			End If
		
		Case UPDATE_RECORD
			If Request.Form("FormProgramIsPostback") = IS_POSTBACK Then
				Call LoadProgramFromForm(page.Program)
				If ValidProgram(page.Program) Then
					Call DoUpdateProgram(page, rv)
					Select Case rv
						Case 0
							' success
							page.MessageID = 3003
						Case -2
							' dupe program
							page.MessageID = 3034
						Case Else
							' unknown error	
							page.MessageID = 3002
					End Select
					page.Action = "": page.ProgramID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormProgramToString(page)
				End If
			Else
				str = str & FormProgramToString(page)
			End If
				
		Case DELETE_RECORD
			If Request.Form("FormConfirmDeleteProgramIsPostback") = IS_POSTBACK Then
				Call DoDeleteProgram(page, rv)
				Select Case rv
					Case 0
						page.MessageID = 3004
					Case Else
						page.MessageID = 3006
				End Select
				page.Action = "": page.ProgramID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else	
				str = str & FormConfirmDeleteProgramToString(page)
			End If
		
		Case SEND_MESSAGE
			Call GenerateEmail(page, rv)
			page.Action = "": page.ProgramID = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
			
		Case Else
			str = str & ProgramGridToString(page)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoDeleteProgram(page, outError)
	Call page.Program.Delete(outError)
End Sub

Sub DoUpdateProgram(page, outError)
	Call page.Program.Save(outError)
End Sub

Sub DoInsertProgram(page, outError)
	page.program.ClientID = page.Client.ClientID
	Call page.program.Add(outError)
End Sub

Sub LoadProgramFromForm(program)
	program.ProgramName = Request.Form("ProgramName")
	program.ProgramDesc = Request.Form("ProgramDesc")
	program.IsEnabled = Request.Form("IsEnabled")
	program.MemberCanEnroll = Request.Form("MemberCanEnroll")
	program.MemberCanEditSkills = Request.Form("MemberCanEditSkills")
	program.DefaultAvailability = Request.Form("DefaultAvailability")
End Sub

Function ValidProgram(program)
	ValidProgram = True
	If Application.Value("FORM_VALIDATION_OFF") Then Exit Function
	
	
	If Not ValidData(program.ProgramName, True, 0, 100, "Program Name", "") Then ValidProgram = False
	If Not ValidData(program.ProgramDesc, False, 0, 2000, "Description", "") Then ValidProgram = False
End Function

Function FormConfirmDeleteProgramToString(page)
	Dim str, msg
	Dim pg				: Set pg = page.Clone()
	
	msg = msg & "You will permanently remove the program <strong>" & html(page.Program.ProgramName) & "</strong> from your " & Application.Value("APPLICATION_NAME") & " account. "
	msg = msg & "You will lose any associated schedule and calendar information. "
	msg = msg & "This action cannot be reversed. "
	
	str = str & CustomApplicationMessageToString("Please confirm this action!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formConfirmDeleteProgram"">"
	str = str & "<input type=""hidden"" name=""FormConfirmDeleteProgramIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = "": pg.ProgramID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "</p></form>"
	
	FormConfirmDeleteProgramToString = str
End Function

Function FormProgramToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" name=""formProgram"">"
	str = str & "<table>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Program Name") & "</td>"
	str = str & "<td><input class=""large gets-focus"" type=""text"" name=""ProgramName"" value=""" & HTML(page.Program.ProgramName) & """ title=""Program Name"" /></td></tr>"
	str = str & "<tr><td class=""label"">Description</td>"
	str = str & "<td><textarea class=""large"" name=""ProgramDesc"" title=""Description"">" & HTML(page.Program.ProgramDesc) & "</textarea></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Program is Enabled</td>"
	str = str & "<td>" & YesNoDropdownToString(page.Program.IsEnabled, "IsEnabled") & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "Disabled programs are hidden from your <br />member accounts and calendars. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Allow Members to Enroll</td>"
	str = str & "<td>" & YesNoDropdownToString(page.Program.MemberCanEnroll, "MemberCanEnroll") & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "When set to yes, members can add themselves <br />to a program. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Allow Members to Edit Skills</td>"
	str = str & "<td>" & YesNoDropdownToString(page.Program.MemberCanEditSkills, "MemberCanEditSkills") & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "When set to yes, members can add or remove <br/>skills from a their profiles. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Set Member Availability</td>"
	str = str & "<td>" & AvailabilityDropdownToString(page.Program.DefaultAvailability) & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "When set to available, " & Application.Value("APPLICATION_NAME") & " considers <br />your members to be available unless they have <br />logged in to set their own availability.</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Save"" />"
	pg.ProgramID = "": pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormProgramIsPostback"" value=""" & IS_POSTBACK & """ /></td></tr>"
	str = str & "</table></form></div>"
	
	FormProgramToString = str
End Function

Function SettingsListForProgramSummaryToString(page)
	Dim str
	
	Dim availableText					: availableText = "not available"
	If page.Program.DefaultAvailability = 1 Then availableText = "available"
	
	Dim enrollmentText					: enrollmentText = "<strong>restricted</strong> (only an account or program administrator can add this program to a member's profile)"
	If page.Program.MemberCanEnroll = 1 Then enrollmentText = "<strong>open</strong> (your members are free to add the program to their profile from their program page)"
	
	Dim editSkillText					: editSkillText = "<strong>allowed</strong> to change skills in their program profile. "
	If page.Program.MemberCanEditSkills = 1 Then editSkillText = "<strong>not allowed</strong> to change skills in their program profile (only an account or program administrator can change skills for program members). "

	str = str & "<ul>"
	str = str & "<li>Members for this program are considered <strong>" & availableText & "</strong> for program events until they login and indicate otherwise. </li>"
	str = str & "<li>Enrollment for this program is " & enrollmentText & ". </li>"
	str = str & "<li>Members are " & editSkillText & "</li>"
	str = str & "</ul>"
	
	SettingsListForProgramSummaryToString = str
End Function

Function EventListForProgramSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim list				: list = page.Program.EventList("")
	
	Dim count				: count = 0
	Dim alt					: alt = ""
	
	Dim eventDetailsHref
	pg.Action = SHOW_EVENT_DETAILS
	
	' 0-EventID 1-EventName 2-EventDate 3-EventNote 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified
	' 8-ProgramID 9-ProgramName 10-ScheduleID 11-ScheduleName 12-ScheduleEventCount
	' 13-ScheduleIsVisible 14-HasUnpublishedChanges 15-EventFileCount 16-ScheduledMemberCount

	If IsArray(list) Then	
		str = str & "<div class=""grid""><table><thead><tr>"
		str = str & "<th>Event</th><th>When</th><th>Schedule</th><th>&nbsp;</th></tr></thead><tbody>"
		For i = 0 To UBound(list,2)
			alt = "":			If count Mod 2 <> 0 Then alt = " class=""alt"""
		
			pg.EventId = list(0,i): pg.Action = SHOW_EVENT_DETAILS
			eventDetailsHref = "/schedule/events.asp" & pg.UrlParamsToString(True)
		
			str = str & "<tr" & alt & ">"
			str = str & "<td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
			str = str & "<a href=""" & eventDetailsHref & """ title=""Details""><strong>" & html(list(1,i)) & "</strong></a></td>"
			str = str & "<td>" & dateTime.Convert(list(2,i), "DDD MMM dd, YYYY")
			If Len(list(4,i) & "") > 0 Then str = str & " at " & dateTime.Convert(list(4,i), "hh:nn pm")
			str = str & "</td>"
			str = str & "<td>" & html(list(11,i)) & "</td>"
			str = str & "<td class=""toolbar"">" 
			str = str & "<a href=""" & eventDetailsHref & """ title=""Details"">"
			str = str & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
			
			count = count + 1
		Next
		str = str & "</tbody></table></div>"
	End If
	
	If count = 0 Then
		str = "<p class=""alert"">This program has no events. </p>"
	End If
	
	EventListForProgramSummaryToString = str 
End Function

Function MemberSkillListingToString(fragment, xml)
	Dim str
	
	Dim node
	xml.Async = False
	
	Dim isSkillEnabled
	Dim isSkillGroupEnabled
	
	If Len(fragment) = 0 Then Exit Function
	
	xml.LoadXml(fragment)
	For Each node In xml.DocumentElement.ChildNodes
		isSkillEnabled = True
		If node.Attributes.GetNamedItem("IsSkillEnabled").Text = "0" Then isSkillEnabled = False
		isSkillGroupEnabled = True
		If node.Attributes.GetNamedItem("IsSkillGroupEnabled").Text = "0" Then isSkillGroupEnabled = False
		
		If isSkillEnabled And isSkillGroupEnabled Then
			str = str & node.Attributes.GetNamedItem("SkillName").Text & ", "
		End If
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 2)	
	
	MemberSkillListingToString = str 
End Function

Function MemberListForProgramSummaryToString(page, count)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim xml					: Set xml = Server.CreateObject("MSXML2.DOMDocument")
	
	Dim list				: list = page.Program.MemberList()
	Dim alt					: alt = ""
	Dim href				: href = ""
	
	Dim memberIsEnabled
	Dim memberProgramIsEnabledText
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 6-ProgramMemberIsEnabled 7-MemberActiveStatus 8-DateCreated 9-DateModified
	' 10-ProgramMemberID 11-IsApproved 12-HasMissingAvailability 13-Email 14-MemberSkillDetailsXmlFragment
	
	If IsArray(list) Then	
		str = str & "<div class=""grid""><table><thead><tr>"
		str = str & "<th>Member</th><th>Skills</th><th>Enabled</th><th>&nbsp;</th></tr></thead><tbody>"
		
		count = 0
		For i = 0 To UBound(list,2)
			memberIsEnabled = True				: If list(7,i) = 0 Then memberIsEnabled = False
		
			If memberIsEnabled Then
				alt = "":			If count Mod 2 <> 0 Then alt = " class=""alt"""
				memberProgramIsEnabledText = "Yes"	: If list(6,i) = 0 Then memberProgramIsEnabledText = "<span style=""color:red;"">No</span>"
			
				pg.MemberId = list(0,i): pg.Action = ""
				href = "/admin/profile.asp" & pg.UrlParamsToString(True)
			
				str = str & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/user.png"" alt="""" />"
				str = str & "<a href=""" & href & """ title=""Profile""><strong>" & html(list(1,i) & ", " & list(2,i)) & "</strong></a></td>"
				str = str & "<td>" & MemberSkillListingToString(list(14,i), xml) & "</td>"
				str = str & "<td>" & memberProgramIsEnabledText & "</td>"
				str = str & "<td class=""toolbar""><a href=""" & href & """ title=""Profile"">"
				str = str & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td>"
				str = str & "</tr>"

				count = count + 1
			End If
		Next
		str = str & "</tbody></table></div>"
	End If
	
	If count = 0 Then
		str = "<p class=""alert"">No members belong to this program. </p>"
	End If
	
	MemberListForProgramSummaryToString = str
End Function

Function AvailabilityGridForProgramSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	
	Dim list				: list = page.Program.MemberList()	
	Dim alt
	Dim href
	Dim rows
	Dim count				: count = 0
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	Dim isMissingInfo
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 6-ProgramMemberIsEnabled 
	' 7-MemberActiveStatus 8-DateCreated 9-DateModified 10-ProgramMemberID 11-IsApproved 
	' 12-HasMissingAvailability 13-Email
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isMemberEnabled = True					: If list(7,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True			: If list(6,i) = 0 Then isProgramMemberEnabled = False
			isMissingInfo = True					: If list(12,i) = 0 Then isMissingInfo = False

			If isMemberEnabled And isProgramMemberEnabled And IsMissingInfo Then
				alt = "":					If count Mod 2 > 0 Then alt = " class=""alt"""
				
				pg.Action = SHOW_PROGRAM_MEMBER_DETAILS: pg.MemberId = list(0,i): pg.ProgramMemberId = list(10,i)
				href = "/admin/profile.asp" & pg.UrlParamsToString(True)
				
				rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/user.png"" alt="""" />"
				rows = rows & "<strong>" & html(page.Program.ProgramName) & "</strong> | "
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
		str = str & "<p>This list of program members has not logged in with up-to-date availability info for some or all of this program's events. </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Member</th><th>Available</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = "<p class=""alert"">All members have logged in with their up-to-date availability info for all of this program's events. </p>"
	End If
	AvailabilityGridForProgramSummaryToString = str
End Function

Function ProgramSummaryToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim memberCount
	Dim memberCountText
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	str = str & "<ul><li><a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>Change program skills</a></li>"
	str = str & "<li><a href=""/admin/members.asp" & pg.UrlParamsToString(True) & """>Work with program members</a></li>"
	pg.Action = UPDATE_RECORD
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Change program settings</a></li>"
	str = str & "<li><a href=""/help/topic.asp?hid=14#anchor-add-program"" target=""_blank"">Learn more about programs</a></li>"
	str = str & "</ul></div>"
	
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & html(page.Program.ProgramName) & "</h3>"
	
	If page.Program.IsEnabled = 0 Then
		str = str & "<h5 class=""disabled"">Program is disabled</h5>"
		str = str & "<p class=""alert"">This program has been set to disabled for your account. </p>"
	End If
	
	str = str & "<h5 class=""description"">Description</h5>"
	If Len(page.Program.ProgramDesc) > 0 Then 
		str = str & "<p>" & html(page.Program.ProgramDesc) & "</p>"
	Else
		str = str & "<p class=""alert"">No description available. </p>"
	End If
	
	str = str & "<h5 class=""settings"">Settings</h5>"
	str = str & SettingsListForProgramSummaryToString(page)
	
	str = str & "<h5 class=""schedule"">Events</h5>"
	str = str & EventListForProgramSummaryToString(page)
	
	str = str & "<h5 class=""program-member"">Members</h5>"
	str = str & MemberListForProgramSummaryToString(page, memberCount)
	
	str = str & "<h5 class=""availability"">Member availability</h5>"
	str = str & AvailabilityGridForProgramSummaryToString(page)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5><ul>"
	memberCountText = memberCount & " member"
	If memberCount <> 1 Then memberCountText = memberCountText & "s"
	str = str & "<li>The " & html(page.Program.ProgramName) & " program has " & memberCountText & ". </li>"
	str = str & "<li>Created on " & dateTime.Convert(page.Program.DateCreated, "DDDD MMM dd, YYYY at about hh:00 pp") & ". </li>"
	str = str & "</ul></div>"
		
	ProgramSummaryToString = str
End Function

Function NoProgramsDialogToString(page)
	Dim dialog					: Set dialog = New cDialog
	Dim pg						: Set pg = page.Clone()
	
	Dim href					: href = ""
	
	dialog.Headline = "No programs here yet ..!"
	
	dialog.Text = dialog.Text & "<p>It looks like the account you are logged into (" & html(page.Client.NameClient) & ") does not have any programs set up. "
	dialog.Text = dialog.Text & "Perhaps this account is brand new and no programs have been created yet. </p>"
	dialog.Text = dialog.Text & "<p>Before you can use this account to manage your schedules and events, you'll need to set up at least one program. "
	dialog.Text = dialog.Text & "To get an empty program that you can add members and events to, click <strong>create your first program</strong>. </p>"
	dialog.Text = dialog.Text & ""

	dialog.SubText = dialog.SubText & "<p>To have " & Application.Value("APPLICATION_NAME") & " generate a sample program for you (with sample members, events, and schedules) click <strong>create a sample program</strong>. "
	dialog.SubText = dialog.SubText & "</p>"
	dialog.SubText = dialog.SubText & "<p>" & Application.Value("APPLICATION_NAME") & " uses programs to organize your church's schedule. "
	dialog.SubText = dialog.SubText & "Programs keep track of your church's events, schedules, and members. "
	dialog.SubText = dialog.SubText & "When this is fixed, you can use this page to add or change the programs you'll use to organize your members. </p>"

	pg.Action = ADDNEW_RECORD
	href = "/admin/programs.asp" & pg.UrlParamsToString(True)
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & href & """>Create your first program</a></li>"
	
	pg.Action = INSERT_SAMPLE_PROGRAM
	href = "/client/preferences.asp" & pg.UrlParamsToString(True)
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & href & """>Create a sample program</a></li>"

	dialog.LinkList = dialog.LinkList & "<li><a href=""/help/topic.asp?hid=14#anchor-add-program"" target=""_blank"">Learn more about programs</a></li>"
	
	NoProgramsDialogToString = dialog.ToString()
End Function

Function ProgramGridToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	Dim list			: list = GetProgramList(page.Member.MemberID)
	Dim rows			: rows = ""
	Dim altClass		: altClass = ""
	Dim enabledText		: enabledText = ""
	Dim count			: count = 0
		
	str = str & "<div class=""tip-box""><h3>I want to ..</h3><ul>"
	pg.Action = ADDNEW_RECORD
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Add a new program</a></li>"
	pg.Action = ""
	str = str & "<li><a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Work with schedules</a></li>"
	str = str & "<li><a href=""/help/topic.asp?hid=14#anchor-add-program"" target=""_blank"">Learn more about programs</a></li>"
	
	str = str & "</ul></div>"
	
	' 0-ProgramID 1-ProgramName 2-Desc 3-IsEnabled 4-EnrollmentType 5-DefaultAvailability
	' 6-HasSkills 7-HasUnpublishedChanges 8-HasEvents 9-MemberCount 10-dateCreated 11-HasSchedules
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			altClass = ""					: If count Mod 2 > 0 Then altClass = " class=""alt"""
			enabledText = "Yes"				: If list(3,i) = 0 Then enabledText = "<span class=""negative"">No</span>"
			
			rows = rows & "<tr" & altClass & ">"
			rows = rows & "<td><img class=""icon"" src=""/_images/icons/script.png"" alt="""" />"
			rows = rows & "<strong>" & html(page.Client.nameClient) & " | "
			pg.ProgramID = list(0,i): pg.Action = SHOW_PROGRAM_DETAILS
			rows = rows & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(1,i)) & "</a></strong></td>"
			rows = rows & "<td>" & enabledText & "</td>"
			rows = rows & "<td class=""toolbar"">"
			pg.ProgramID = list(0,i)
			rows = rows & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details""><img src=""/_images/icons/magnifier.png"" alt=""icon"" /></a>"
			pg.Action = UPDATE_RECORD
			rows = rows & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Edit""><img src=""/_images/icons/pencil.png"" alt=""icon"" /></a>"
			pg.Action = ""
			rows = rows & "<a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """ title=""Schedules""><img src=""/_images/icons/calendar.png"" alt=""icon"" /></a>"
			pg.Action = ""
			rows = rows & "<a href=""/schedule/availability.asp" & pg.UrlParamsToString(True) & """ title=""Availability""><img src=""/_images/icons/clock.png"" alt=""icon"" /></a>"
			pg.Action = ""
			rows = rows & "<a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """ title=""Skills""><img src=""/_images/icons/plugin.png"" alt=""icon"" /></a>"
			rows = rows & "<a href=""/admin/members.asp"  & pg.UrlParamsToString(True) & """ title=""Members""><img src=""/_images/icons/user.png"" alt=""icon"" /></a>"
			rows = rows & "<a href=""/admin/leaders.asp" & pg.UrlParamsToString(True) & """ title=""Leaders""><img src=""/_images/icons/medal_gold_1.png"" alt=""icon"" /></a>"
			pg.Action = SEND_MESSAGE: pg.ProgramID = list(0,i)
			rows = rows & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email""><img src=""/_images/icons/email.png"" alt=""icon"" /></a>"
			pg.Action = DELETE_RECORD
			rows = rows & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove""><img src=""/_images/icons/cross.png"" alt=""icon"" /></a>"
			rows = rows & "</td></tr>"
			
			count = count + 1
		Next
	End If
	
	If count > 0 Then
		str = str & m_appMessageText
		str = str & "<div class=""summary""><h3 class=""first"">" & html(page.Client.NameClient) & " program list</h3>"
		str = str & "<p>" & Application.Value("APPLICATION_NAME") & " uses programs to organize your account members, events, and event teams. "
		str = str & "You have permission (administrator) to manage the programs in this list. </p>"
		str = str & "<div class=""grid""><table><thead><tr>"
		str = str & "<th scope=""col"">Program</th><th scope=""col"">Enabled</th><th scope=""col"">&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div></div>"
	Else
		str = NoProgramsDialogToString(page)
	End If
	
	ProgramGridToString = str
End Function

Function AvailabilityDropdownToString(val)
	Dim str, arr
	
	ReDim arr(1,1)
	arr(0,0) = "0"
	arr(0,1) = "1"
	arr(1,0) = "Not Available"
	arr(1,1) = "Available"
	
	str = str & "<select name=""DefaultAvailability"">"
	str = str & SelectOption(arr, val)
	str = str & "</select>"
	
	AvailabilityDropdownToString = str
End Function

Function GetProgramList(memberID)
	Dim cnn, rs
	
	' 0-ProgramID 1-ProgramName 2-Desc 3-IsEnabled 4-EnrollmentType 5-DefaultAvailability
	' 6-HasSkills 7-HasUnpublishedChanges 8-HasEvents 9-MemberCount 10-dateCreated 11-HasSchedules
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.Open Application.Value("CNN_STR")
	Set rs = Server.CreateObject("ADODB.Recordset")
	cnn.up_programGetProgramListForAdmin CLng(memberID), rs
	If Not rs.EOF Then GetProgramList = rs.GetRows()
	
	If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
	Set cnn = Nothing
End Function

Sub GenerateEmail(ByRef page, ByRef outError)
	Dim str, i
	Dim tempError		: tempError = 0
	
	Dim email			: Set email = New cEmail
	Dim programMember	: Set programMember = New cProgramMember
	programMember.ProgramID = page.ProgramID
	Dim list			: list = programMember.GetMemberList()
	
	If Not IsArray(list) Then 
		outError = -1
		Exit Sub
	End If
	
	' 6-ProgramMemberIsEnabled 7-MemberActiveStatus 11-IsApproved
	For i = 0 To UBound(list,2)
		If (list(6,i) = 1) And (list(7,i) = 1) And (list(11,i) = 2) Then
			str = str & list(0,i) & ","
		End If
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) -1)
	email.RecipientIDList = str
	email.MemberID = page.Member.MemberID
	email.ClientID = page.Client.ClientID
	Call email.Add(tempError)
	If tempError <> 0 Then outError = -2
	page.EmailID = email.EmailID

	Set email = Nothing
	Set programMember = Nothing
End Sub

Sub SetPageHeader(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim programsLink
	programsLink = "<a href=""" & pg.Url & """>Programs</a> / "
	
	Dim programLink
	pg.Action = SHOW_PROGRAM_DETAILS
	programLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.Program.ProgramName) & "</a> / "
	
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case SHOW_PROGRAM_DETAILS
			str = str & programsLink
			str = str & html(page.Program.ProgramName)
		Case ADDNEW_RECORD
			str = str & programsLink
			str = str & "New Program"
		Case DELETE_RECORD
			str = str & programsLink
			str = str & programLink
			str = str & "Remove"
		Case UPDATE_RECORD
			str = str & programsLink
			str = str & programLink
			str = str & "Edit"
		Case Else
			str = str & "Programs"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg
	Dim href
	
	Dim newProgramButton
	Set pg = page.Clone()
	pg.Action = ADDNEW_RECORD: pg.ProgramID = ""
	href = pg.Url & pg.UrlParamsToString(True)
	newProgramButton = newProgramButton & "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/script_add.png"" alt="""" /></a><a href=""" & href & """>New Program</a></li>"
	
	Dim programListButton
	Set pg = page.Clone()
	pg.Action = "": pg.ProgramID = ""
	href = pg.Url & pg.UrlParamsToString(True)
	programListButton = programListButton & "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/script.png"" /></a><a href=""" & href & """>Program List</a></li>"
	
	Select Case page.Action
		Case SHOW_PROGRAM_DETAILS
			str = str & programListButton
		Case DELETE_RECORD
			str = str & programListButton
		Case ADDNEW_RECORD
			str = str & programListButton
		Case UPDATE_RECORD
			str = str & programListButton
		Case Else
			str = str & newProgramButton
	End Select
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_select_option.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_YesNoDropdownToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_GetListFromXmlFragment.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public ProgramID
	Public ScheduleID
	Public EmailID
	Public EventId
	Public MemberID
	Public ProgramMemberId
	
	' objects
	Public Member
	Public Client
	Public Program	
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(ScheduleID) > 0 Then str = str & "scid=" & Encrypt(ScheduleID) & amp
		If Len(EmailID) > 0 Then str = str & "emid=" & Encrypt(EmailID) & amp
		If Len(MemberID) > 0 Then str = str & "mid=" & Encrypt(MemberID) & amp
		If Len(EventId) > 0 Then str = str & "eid=" & Encrypt(EventId) & amp
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
		c.Action = Action
		c.ProgramID = ProgramID
		c.ScheduleID = ScheduleID
		c.EmailID = EmailID
		c.MemberID = MemberID
		c.EventId = EventId
		c.ProgramMemberId = ProgramMemberId
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		
		Set Clone = c
	End Function
End Class
%>

