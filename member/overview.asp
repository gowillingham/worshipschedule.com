<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "overview"
Dim m_pageHeaderText	: m_pageHeaderText = "&nbsp;"
Dim m_impersonateText	: m_impersonateText = ""
Dim m_pageTitleText		: m_pageTitleText = ""
Dim m_topBarText		: m_topBarText = "&nbsp;"
Dim m_bodyText			: m_bodyText = ""
Dim m_tabStripText		: m_tabStripText = ""
Dim m_tabLinkBarText	: m_tabLinkBarText = ""
Dim m_acctExpiresText	: m_acctExpiresText = ""

Sub OnPageLoad(ByRef page)
	Dim sess			: Set sess = New cSession
	sess.SessionID = Request.Cookies("sid")
	Call CheckSession(sess, PERMIT_MEMBER)
	
	page.MessageID = Request.QueryString("msgid")
	page.Action = Decrypt(Request.QueryString("act"))
	page.MemberNotificationID = DeCrypt(Request.QueryString("mnid"))
	page.ProgramMemberID = DeCrypt(Request.QueryString("pmid"))
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	' set the view tokens
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
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<style type="text/css">
			.details	{width:622px;}
			.summary	{width:650px;}
		</style>
		<title><%=m_pageTitleText%></title>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)

	str = str & ApplicationMessageToString(page.MessageID)
	page.MessageID = ""

	Select Case page.Action
		Case Else
			str = str & MemberSummaryToString(page)
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Function MemberSummaryToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	str = str & "<ul><li><a href=""/member/schedules.asp" & pg.UrlParamsToString(True) & """>View my schedule</a></li>"
	str = str & "<li><a href=""/member/programs.asp" & pg.UrlParamsToString(True) & """>Make changes to my programs</a></li>"
	str = str & "<li><a href=""/member/events.asp" & pg.UrlParamsToString(True) & """>Make changes to when I'm available</a></li>"
	str = str & "<li><a href=""/help/help.asp" & pg.UrlParamsToString(True) & """ target=""_blank"">Learn more about " & Application.Value("APPLICATION_NAME") & "</a></li>"
	str = str & "</ul></div>"

	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">Welcome " & html(page.Member.NameFirst & " " & page.Member.NameLast) & "! </h3>"
	
	If page.Client.ProgramCount = 0 Then
		MemberSummaryToString = NoProgramsForClientDialogToString(page)
		Exit Function
	End If
	
	str = str & DisabledAccountMessageToString(page)
	str = str & UpcomingEventGridToString(page)
	str = str & MemberProgramInfoForSummaryToString(page)
	
	str = str & "</div>"
	
	MemberSummaryToString = str
End Function

Function DisabledAccountMessageToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	If page.Member.ActiveStatus = 1 Then Exit Function
	
	str = str & "<h5 class=""disabled"">Your " & html(page.Member.ClientName) & " account is disabled!</h5>"
	str = str & "<p class=""alert"">You or an account administrator has set your account to disabled. "
	str = str & "This means that your account has been removed from all " & html(page.Client.NameClient) & " programs, member lists, and email lists. "
	str = str & "You cannot be added to any schedules or event teams while your account is disabled. "
	str = str & "You may re-enable your account <a href=""/member/settings.asp"">here</a>. </p>"
	
	DisabledAccountMessageToString = str
End Function

Function SkillListXmlFragmentToString(fragment, xml)
	Dim str
	
	Dim node
	Dim child
	
	If Len(fragment) = 0 Then Exit Function
	
	xml.LoadXml(fragment)
	xml.Async = False
	
	For Each node In xml.DocumentElement.ChildNodes
		str = str & node.Text & ", "
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 2)
	
	SkillListXmlFragmentToString = str
End Function

Function UpcomingEventGridToString(page)
	Dim str, i
	Dim pg			: Set pg = page.Clone()
	Dim dateTime	: Set dateTime = New cFormatDate
	
	Dim xml			: Set xml = Server.CreateObject("MSXML2.DOMDocument")
	
	Dim list		: list = page.Member.EventList(Now(), Null, Null)
	
	Dim rows		: rows = ""
	Dim alt			: alt = ""
	Dim count		: count = 0
	
	Dim isScheduled
	Dim isVisible
	Dim isProgramMemberEnabled
	
	' 0-EventID 1-EventName 2-EventNote 3-EventDate 4-TimeStart 5-TimeEnd 6-ClientID
	' 7-MemberID 8-ProgramName 9-ProgramID 10-ScheduleID 11-ScheduleName 12-ScheduleDesc
	' 13-ScheduleIsVisible 14-IsScheduled 15-SkillListXmlFrag 16-FileListXmlFrag 17-IsAvailable
	' 18-EventAvailabilityID 19-ProgramIsEnabled 20-ProgramMemberIsActive 
	' 21-AvailabilityViewedByMember
		
	str = str & "<h5 class=""schedule"">My Upcoming events</h5>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			alt = ""								: If count Mod 2 > 0 Then alt = " class=""alt"""
			isScheduled = True						: If list(14,i) = 0 Then isScheduled = False
			isVisible = True						: If list(13,i) = 0 Then isVisible = False
			isProgramMemberEnabled = True			: If list(20,i) = 0 Then isProgramMemberEnabled = False
			
			If isScheduled And isVisible And isProgramMemberEnabled Then
				rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
				rows = rows & "<strong>" & html(list(8,i)) & "</strong> | "
				pg.Action = SHOW_EVENT_DETAILS: pg.EventID = list(0,i)
				rows = rows & "<a href=""/member/schedules.asp" & pg.UrlParamsToString(True) & """ title=""Details""><strong>" & html(list(1,i)) & "</strong></a></td>"
				rows = rows & "<td>" & dateTime.Convert(list(3,i), "DDD MMM dd, YYYY")
				If Len(list(4,i)) > 0 Then rows = rows & " at " & dateTime.Convert(list(4,i), "hh:nn pp") 
				rows = rows & "</td>"
				rows = rows & "<td>" & SkillListXmlFragmentToString(list(15,i), xml) & "</td>"
				rows = rows & "<td class=""toolbar"">"
				pg.Action = SHOW_EVENT_DETAILS: pg.EventID = list(0,i)
				rows = rows & "<a href=""/member/schedules.asp" & pg.UrlParamsToString(True) & """ title=""Details""><img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
				
				count = count + 1
			End If
		Next
	End If
	
	If count = 0 Then
		str = str & "<p class=""alert"">You have no events upcoming. "
		str = str & "Get a look at your full calendar <a href=""/member/schedules.asp"">here</a>. </p>"
	Else
		str = str & "<p>Here are the upcoming events for which you have been scheduled. </p>"
		str = str & "<div class=""grid""><table><thead>"
		str = str & "<tr><th>Event</th><th>Date</th><th>Scheduled For</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	End If
	
	UpcomingEventGridToString = str
End Function

Function MemberProgramInfoForSummaryToString(page)
	Dim str, i
	Dim pg						: Set pg = page.Clone()
	Dim dateTime				: Set dateTime = New cFormatDate
	
	Dim list					: list = page.Member.ProgramList()
	Dim items					: items = ""
	
	Dim isProgramEnabled
	Dim isProgramMemberEnabled
	Dim hasSkills
	Dim hasAvailability
	
	' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-ProgramMemberID 4-SkillListXmlFrag 5-EnrollStatusID 6-EnrollStatusText
	' 7-EnrollStatusDesc 8-IsLeader 9-IsLeaderText 10-IsActive 11-IsActiveText 12-LastScheduleView 
	' 13-DateModified 14-DateCreated 15-SkillCount 16-IsMissingAvailability 17-IsAdmin 18-ProgramIsEnabled
		
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isProgramEnabled = True					: If list(18,i) = 0 Then isProgramEnabled = False
			isProgramMemberEnabled = True			: If list(10,i) = 0 Then isProgramMemberEnabled = False
			
			If isProgramEnabled Then
			
				If isProgramMemberEnabled Then
					hasSkills = True				: If list(15,i) = 0 Then hasSkills = False
					hasAvailability = True			: If list(16,i) = MEMBER_AVAILABILITY_NEEDS_UPDATE Then hasAvailability = False
				
					If Not hasSkills Then
						items = items & "<li class=""error""><strong class=""warning"">" & html(list(1,i)) & ". </strong>"
						items = items & "You belong to the " & html(list(1,i)) & " program, but you haven't set which program skills belong to your profile. "
						pg.Action = "": pg.ProgramId = list(0,i): pg.ProgramMemberId = list(3,i)
						items = items & "You can fix that <a href=""/member/skills.asp" & pg.UrlParamsToString(True) & """ title=""Skills"">here</a>. </li>"
					End If
					
					If Not hasAvailability Then
						items = items & "<li class=""availability""><strong class=""alert"">" & html(list(1,i)) & ". </strong>"
						items = items & "There are new events for the " & html(list(1,i)) & " program. "
						items = items & "You can let your account administrator know when you are available "
						pg.Action = "": pg.ProgramId = list(0,i): pg.ProgramMemberId = ""
						items = items & "<a href=""/member/events.asp" & pg.UrlParamsToString(True) & """ title="""">here</a> .</li>"
					End If
				Else
					items = items & "<li class=""error""><strong class=""warning"">" & html(list(1,i)) & ". </strong>"
					items = items & "You or an account administrator have disabled the " & html(list(1,i)) & " program from your profile. "
					pg.Action = "": pg.ProgramId = list(0,i): pg.ProgramMemberId = ""
					If page.Url = "/member/programs.asp" Then
						items = items & "Click <strong>enable</strong> in the toolbar for " & html(list(1,i)) & " to re-enable. "
					Else
						items = items & "You can re-enable that program <a href=""/member/programs.asp" & pg.UrlParamsToString(True) & """ title=""Programs"">here</a>. "
					End If
					items = items & "</li>"
				End If
			End If
		Next
	End If
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	If Len(items) > 0 Or Not IsArray(list) Then 
		str = str & "<p>Some items in your account need attention. </p>"
		If Not IsArray(list) Then 
			str = str & "<h5 class=""program"">Programs</h5>"
			str = str & "<p class=""alert"">You do not have any programs in your profile. "
			str = str & Application.Value("APPLICATION_NAME") & " uses programs to organize events, schedules and members for the " & html(page.Member.ClientName) & " account. "
			str = str & "You can add or change the programs in your profile <a href=""/member/programs.asp"">here</a>. </p>"
		End If
		
		If Len(items) > 0 Then
			str = str & "<ul class=""other-stuff"">" & items & "</ul>"
		End If
	End If
	str = str & "<ul><li>Member since " & dateTime.Convert(page.Member.DateCreated, "DDD MMM dd, YYYY") & ". </li></ul>"

	MemberProgramInfoForSummaryToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	
	str = str & "Member Home"

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim href
	Dim pg		: Set pg = page.Clone()
	
	Select Case page.Action
		Case DISPLAY_NOTIFICATION
			pg.Action = "": pg.MemberNotificationID = ""
			href = "href=""" & pg.Url & pg.UrlParamsToString(True) & """"
			str = str & "<li><a " & href & "><img class=""icon"" src=""/_images/icons/house.png"" /></a><a " & href & ">My Home</a></li>"
			pg.Action = NOTIFICATION_DISMISS
			pg.MemberNotificationID = page.MemberNotificationID
			href = "href=""" & pg.Url & pg.UrlParamsToString(True) & """"
			str = str & "<li><a " & href & "><img class=""icon"" src=""/_images/icons/cross.png"" /></a><a " & href & ">Dismiss Message</a></li>"
		Case Else
			href = "href=""/help/help.asp"" target=""_blank"""
			str = str & "<li><a " & href & "><img class=""icon"" src=""/_images/icons/help.png"" alt="""" /></a><a " & href & ">Help</a></li>"
	End Select
	

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_NoProgramsForClientDialogToString.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public MemberNotificationID
	Public ProgramID
	Public EventID
	Public EventAvailabilityID
	Public ProgramMemberID
	
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
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(MemberNotificationID) > 0 Then str = str & "mnid=" & Encrypt(MemberNotificationID) & amp
		If Len(EventID) > 0 Then str = str & "eid=" & Encrypt(EventID) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(EventAvailabilityID) > 0 Then str = str & "eaid=" & Encrypt(EventAvailabilityID) & amp
		If Len(ProgramMemberID) > 0 Then str = str & "pmid=" & Encrypt(ProgramMemberID) & amp
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		
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
		c.MemberNotificationID = MemberNotificationID
		c.EventID = EventID
		c.ProgramID = ProgramID
		c.EventAvailabilityID = EventAvailabilityID
		c.ProgramMemberID = ProgramMemberID
		
		Set c.Member = Member
		Set c.Client = Client
		
		Set Clone = c
	End Function
End Class
%>

