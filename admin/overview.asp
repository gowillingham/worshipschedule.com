<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-overview"
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
	page.ProgramId = Decrypt(Request.QueryString("pid"))
	page.EventId = Decrypt(Request.QueryString("eid"))

	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()

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
		<link rel="stylesheet" type="text/css" href="overview.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
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
		Case Else
			str = str & AdminSummaryToString(page)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Function EventGridForAdminSummary(page)
	Dim str, i
	Dim dateTime					: Set dateTime = New cFormatDate
	Dim pg							: Set pg = page.Clone()
	
	Dim list						: list = page.Member.AdminEventList("", "")
	Dim rows						: rows = ""
	Dim href						: href = ""
	Dim alt							: alt = ""
	Dim count						: count = 0
	
	Dim isThisWeek
	Dim hasEventThisWeek
	Dim hasCurrentEvents
	
	' 0-EventID 1-EventName 2-EventDate 3-EventNote 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified
	' 8-ProgramID 9-ProgramName 10-ScheduleID 11-ScheduleName 12-ScheduleEventCount
	' 13-ScheduleIsVisible 14-HasUnpublishedChanges 15-EventFileCount 16-ScheduledMemberCount
	' 17-HtmlBackgroundColor
	
	hasCurrentEvents = False
	
	If IsArray(list) Then
	
		hasEventThisWeek = False
		For i = 0 To UBound(list,2)
			isThisWeek = False
			If (list(2,i) => Date()) And list(2,i) <= (DateAdd("ww", 1, Date())) Then
				hasEventThisWeek = True
				hasCurrentEvents = True
				isThisWeek = True
			End If
		
			If isThisWeek Then
				alt = ""				: If count Mod 2 > 0 Then alt = " class=""alt"""
				
				pg.EventId = list(0,i): pg.Action = SHOW_EVENT_DETAILS
				href = "/schedule/events.asp" & pg.UrlParamsToString(True)
				
				rows = rows & "<tr><td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
				rows = rows & "<strong>" & html(list(9,i)) & "</strong> | "
				rows = rows & "<a href=""" & href & """ title=""Details""><strong>" & html(list(1,i)) & "</strong></a></td>"
				rows = rows & "<td>" & dateTime.Convert(list(2,i), "DDD MMM dd, YYYY") & "</td>"
				rows = rows & "<td>" & html(list(11,i)) & "</td>"
				rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
				rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
				
				count = count + 1
			End If
		Next
		
		If Not hasEventThisWeek Then
			For i = 0 To UBound(list,2)
				If (list(2,i) => Date()) Then
					hasCurrentEvents = True	
				
					pg.EventId = list(0,i): pg.Action = SHOW_EVENT_DETAILS
					href = "/schedule/events.asp" & pg.UrlParamsToString(True)
				
					rows = rows & "<tr><td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
					rows = rows & "<strong>" & html(list(9,i)) & "</strong> | "
					rows = rows & "<a href=""" & href & """ title=""Details""><strong>" & html(list(1,i)) & "</strong></a></td>"
					rows = rows & "<td>" & dateTime.Convert(list(2,i), "DDD MMM dd, YYYY") & "</td>"
					rows = rows & "<td>" & html(list(11,i)) & "</td>"
					rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
					rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
					
					Exit For
				End If
			Next
		End If
	End If

	
	If hasEventThisWeek Then
		str = str & "<p>You have these events scheduled for this week. </p>"
		str = str & "<div class=""grid""><table>"
		str = str & "<thead><tr><th>Event</th><th>When</th><th>Schedule</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	ElseIf hasCurrentEvents Then
		str = str & "<p>Here is your next upcoming event. </p>"
		str = str & "<div class=""grid""><table>"
		str = str & "<thead><tr><th>Event</th><th>When</th><th>Schedule</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = str & "<p class=""alert"">Your programs have no events scheduled (or all events occur in the past). </p>"
	End If
	
	EventGridForAdminSummary = str
End Function

Function MemberOwnsProgram(programs, programId)
	Dim i
	
	' 0-ProgramId 1-ProgramName 2-IsEnabled

	MemberOwnsProgram = False
	If Not IsArray(programs) Then Exit Function
	
	For i = 0 To UBound(programs,2)
		If CStr(programs(0,i) & "") = CStr(programId & "") Then
			MemberOwnsProgram = True
			Exit For
		End If
	Next
End Function

Function ProgramsWithoutSkillsItemsToString(page, programs)
	Dim str, i
	Dim pg							: Set pg = page.Clone()
	
	Dim list						: list = page.Client.ProgramList("")
	Dim href						: href = ""
	
	Dim ownsProgram
	Dim hasSkills
	Dim isEnabled 
	
	' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-EnrollmentType 4-IsEnabled 5-DateCreated
	' 6-DateModified 7-DefaultAvailability 8-MemberCanEnroll 9-MemberCanEditSkills
	' 10-MemberCount 11-SkillCount 12-ScheduleCount 113-EventCount
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			ownsProgram = MemberOwnsProgram(programs, list(0,i))
			hasSkills = True					: If list(11,i) = 0 Then hasSkills = False
			isEnabled = True					: If list(4,i) = 0 Then isEnabled = False		
			If ownsProgram And isEnabled And Not hasSkills Then
				pg.ProgramId = list(0,i): pg.Action = ""
				href = "/admin/skills.asp" & pg.UrlParamsToString(True)
				
				str = str & "<li class=""program-error""><strong class=""warning"">" & html(list(1,i)) & ". </strong>"
				str = str & "This program has no skills set up yet (or all skills are disabled). "
				str = str & "You can add or change skills for this program <a href=""" & href & """>here</a>. </li>"
			End If		
		Next
	End If
	
	ProgramsWithoutSkillsItemsToString = str
End Function

Function ProgramsWithoutMembersItemsToString(page, programs)
	Dim str, i
	Dim pg							: Set pg = page.Clone()
	
	Dim list						: list = page.Client.ProgramList("")
	Dim href						: href = ""
	
	Dim ownsProgram
	Dim isEnabled
	Dim hasMembers 
	
	' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-EnrollmentType 4-IsEnabled 5-DateCreated
	' 6-DateModified 7-DefaultAvailability 8-MemberCanEnroll 9-MemberCanEditSkills
	' 10-MemberCount 11-SkillCount 12-ScheduleCount 113-EventCount
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		ownsProgram = MemberOwnsProgram(programs, list(0,i))
		hasMembers = True				: If list(10,i) = 0 Then hasMembers = False
		isEnabled = True				: If list(4,i) = 0 Then isEnabled = False
		
		If ownsProgram And isEnabled And Not hasMembers Then
			pg.ProgramId = list(0,i): pg.Action = CONFIGURE_PROGRAM_MEMBERS
			href = "/admin/members.asp" & pg.UrlParamsToString(True)
			
			str = str & "<li class=""program-error""><strong class=""warning"">" & html(programs(1,i)) & "</strong>. "
			str = str & "No members belong to this program (or all members are disabled). "
			str = str & "You can add or change the members for this program <a href=""" & href & """>here</a>. </li>"
		End If
	Next
	
	ProgramsWithoutMembersItemsToString = str
End Function

Function DisabledProgramsItemsToString(page, programs)
	Dim str, i
	Dim pg						: Set pg = page.Clone()
	
	Dim href
	Dim isProgramEnabled
	
	' 0-ProgramId 1-ProgramName 2-IsEnabled
	
	If Not IsArray(programs) Then Exit Function
	For i = 0 To UBound(programs,2)
		isProgramEnabled = True						: If programs(2,i) = 0 Then isProgramEnabled = False
		
		If Not isProgramEnabled Then
			pg.ProgramId = programs(0,i): pg.Action = UPDATE_RECORD
			href = "/admin/programs.asp" & pg.UrlParamsToString(True)
			
			str = str & "<li class=""error""><strong class=""warning"">" & html(programs(1,i)) & "</strong>. "
			str = str & "This program is set to disabled. "
			str = str & "You can re-enable this program <a href=""" & href & """>here</a>. </li>"
		End If
	Next

	DisabledProgramsItemsToString = str
End Function

Function UnpublishedChangesItemsToString(page, programs)
	Dim str, i
	Dim pg							: Set pg = page.Clone()
	
	Dim list						: list = page.Client.ScheduleList()
	Dim href						: href = ""
	
	Dim isProgramEnabled
	Dim hasUnpublishedChanges
	Dim ownsProgram
	
	' 0-ScheduleId 1-ScheduleName 2-ScheduleDesc 3-IsVisible 4-HtmlBackgroundColor
	' 5-DateCreated 6-DateModified 7-ProgramId 8-ProgramName 9-IsEnabled
	' 10-HasUnpublishedChanges 11-EventCount
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		ownsProgram = MemberOwnsProgram(programs, list(7,i))
		isProgramEnabled = True				: If list(9,i) = 0 Then isProgramEnabled = False
		hasUnpublishedChanges = True		: If list(10,i) = 0 Then hasUnpublishedChanges = False
		
		If ownsProgram And isProgramEnabled And hasUnpublishedChanges Then
			href = "/schedule/schedules.asp"
		
			str = str & "<li class=""publish""><strong class=""warning"">" & html(list(1,i)) & "</strong>. "
			str = str & "This schedule has changes that have not yet been published to your member calendar. "
			str = str & "You can publish or change this schedule <a href=""" & href & """>here</a>. </li>"
		End If
	Next
	
	UnpublishedChangesItemsToString = str
End Function

Function OtherStuffForSummaryToString(page)
	Dim str
	Dim pg							: Set pg = page.Clone()
	Dim dateTime					: Set dateTime = New cFormatDate
	
	Dim clientAdmin					: Set clientAdmin = New cClientAdmin
	clientAdmin.ClientAdminId = page.Member.ClientAdminId
	
	Dim programs					: programs = page.Member.OwnedProgramsList()
	
	Dim items
	Dim href						: href = ""
		
	If page.Client.HasPrograms = 0 Then
		pg.Action = ADDNEW_RECORD
		href = "/admin/programs.asp" & pg.UrlParamsToString(True)
		
		items = items & "<li class=""error""><strong class=""warning"">No programs! </strong>"
		items = items & "Your account does not have any programs. "
		items = items & Application.Value("APPLICATION_NAME") & " uses programs to organize your members and events. "
		items = items & "You can create your first program <a href=""" & href & """ title=""Add program"">here</a>. "
		
		pg.Action = INSERT_SAMPLE_PROGRAM
		href = "/client/preferences.asp" & pg.UrlParamsToString(True)
		
		items = items & "<br /><br />Click <a href=""" & href & """>here</a> to have " & Application.Value("APPLICATION_NAME") & " create a sample program (with sample members, events, and schedules) you can use to practice with your account. </li>"
	End If
	
	items = items & UnpublishedChangesItemsToString(page, programs)
	items = items & DisabledProgramsItemsToString(page, programs)
	items = items & ProgramsWithoutSkillsItemsToString(page, programs)
	items = items & ProgramsWithoutMembersItemsToString(page, programs)
	
	If Len(items) > 0 Then
		str = str & "<p>Some items in your account need your attention. </p>"
		str = str & "<ul class=""other-stuff"">" & items & "</ul>"
	End If
	
	If Len(clientAdmin.ClientAdminId) > 0 Then 
		Call clientAdmin.Load() 
		str = str & "<ul><li>Account administrator since " & dateTime.Convert(page.Member.DateCreated, "DDD MMM dd, YYYY") & ". </li></ul>"
	Else
		str = str & "<ul><li>Administrator since " & dateTime.Convert(page.Member.DateCreated, "DDD MMM dd, YYYY") & ". </li></ul>"
	End If
	
	OtherStuffForSummaryToString = str
End Function

Function AdminSummaryToString(page)
	Dim str
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	str = str & "<ul><li><a href=""/admin/programs.asp"">Add or change a program</a></li>"
	str = str & "<li><a href=""/admin/members.asp"">Add or change one of my account members</a></li>"
	str = str & "<li><a href=""/schedule/schedules.asp"">Work with my schedules or events</a></li>"
	str = str & "<li><a href=""/email/email.asp"">Send email</a></li>"
	str = str & "<li><a href=""/help/help.asp"" target=""_blank"">Learn more about my " & Application.Value("APPLICATION_NAME") & " admin account</a></li>"
	str = str & "</ul></div>"
	
	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "<p>This page is the administrative home page for your account, "
	str = str & "where you can manage your schedules and teams."
	str = str & "<br /><br />Regular members are not able to access this part of your account. </p></div>"
	
	str = str & "<div class=""summary"">"
	str = str & "<div class=""message"">"
	str = str & "<h3 class=""alert-message"">"
	str = str & "<img class=""icon"" src=""/_images/icons/critical.png"" />"
	str = str & "Important message for Worshipschedule Administrators!"
	str = str & "</h3>"
	str = str & "<div class=""listing"">"
	str = str & "As of April 2012 the Worshipschedule web scheduling application is no longer under active development. "
	str = str & "As a courtesy to you, your account will be extended until September 15 2012, free of charge. "
	str = str & "However, all access to your account will end September 15, 2012. "
	str = str & "Please contact <a href=""mailto:support@worshipschedule.com"">support@worshipschedule.com</a> with any questions. "
	str = str & "</div>"
	str = str & "</div>"



	str = str & "<h3 class=""first"">Welcome " & html(page.Member.NameFirst & " " & page.Member.NameLast) & "</h3>"
	
	str = str & "<h5 class=""schedule"">Upcoming events</h5>"
	str = str & EventGridForAdminSummary(page)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & OtherStuffForSummaryToString(page)
	
	str = str & "</div>"
	
	AdminSummaryToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "Admin Home"
	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg
	Dim href
	
	Select Case page.Action
		Case Else
			str = str & "<li>&nbsp;</li>"
	End Select
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_admin_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public ProgramId
	Public EventId
	
	' objects
	Public Member
	Public Client
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ProgramId) > 0 Then str = str & "pid=" & Encrypt(ProgramId) & amp
		If Len(EventId) > 0 Then str = str & "eid=" & Encrypt(EventId) & amp
		
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
		c.ProgramId = ProgramId
		c.EventId = EventId
		
		Set c.Member = Member
		Set c.Client = Client
		
		Set Clone = c
	End Function
End Class
%>

