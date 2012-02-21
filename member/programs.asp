<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "programs"
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
	Call CheckSession(sess, PERMIT_MEMBER)
	
	page.MessageID = Request.QueryString("msgid")
	page.Action = Decrypt(Request.QueryString("act"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.ProgramMemberID = Decrypt(Request.QueryString("pmid"))
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then page.Program.Load()
	
	Set page.ProgramMember = New cProgramMember
	page.ProgramMember.ProgramMemberID = page.ProgramMemberID
	If Len(page.ProgramMember.ProgramMemberID) > 0 Then page.ProgramMember.Load()
	
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
		Case SHOW_PROGRAM_DETAILS
			str = str & ProgramSummaryToString(page)
			
		Case SHOW_AVAILABLE_PROGRAM_DETAILS
			str = str & AvailableProgramDetailsToString(page)
			
		Case DELETE_RECORD
			If Request.Form("FormConfirmDeleteProgramMemberIsPostback") = IS_POSTBACK Then
				Call DeleteProgramMember(page, rv)
				Select Case rv
					Case 0
						page.MessageID = 1026
					Case Else
						page.MessageID = 1027
				End Select
				page.Action = "": page.ProgramMemberID = "": page.ProgramID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmDeleteProgramMemberToString(page)
			End If
			
		Case ADDNEW_RECORD
			Call InsertProgramMember(page, rv)
			Select Case rv
				Case 0
					page.MessageID = 1023
					Response.Redirect("/member/skills.asp" & page.UrlParamsToString(False))
				Case Else
					page.Action = "": page.ProgramID = "": page.MessageID = 1025
					Response.Redirect(page.Url & page.UrlParamsToString(False))
			End Select
			
		Case SHOW_AVAILABLE_PROGRAMS
			str = str & AvailableProgramGridToString(page)
			
		Case TOGGLE_PROGRAM_MEMBER_IS_ACTIVE
			Call ToggleProgramMemberIsActive(page, rv)
			Select Case rv
				Case 0
					page.MessageID = 1028
				Case Else
					page.MessageID = 1029
			End Select
			page.Action = "": page.ProgramMemberID = "":					
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case Else
			str = str & ProgramGridToString(page)
	
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub ClearTablinkBar()
	m_tabLinkBarText = "<li>&nbsp;</li>"
End Sub

Sub DeleteProgramMember(page, outError)
	page.ProgramMember.MemberID = page.Member.MemberID
	Call page.ProgramMember.Delete(outError)
End Sub

Sub InsertProgramMember(page, outError)
	page.ProgramMember.MemberID = page.Member.MemberID
	page.ProgramMember.ProgramID = page.ProgramID
	page.ProgramMember.EnrollStatusID = 3 ' accepted
	page.ProgramMember.IsActive = IS_ACTIVE
	Call page.ProgramMember.Add(outError)
	page.ProgramMemberID = page.ProgramMember.ProgramMemberID
End Sub

Sub ToggleProgramMemberIsActive(page, outError)
	If page.ProgramMember.IsActive = 0 Then
		page.ProgramMember.IsActive = 1
	Else
		page.ProgramMember.IsActive = 0
	End If
	Call page.ProgramMember.Save(outError)
End Sub

Function AvailableProgramDetailsToString(page)
	Dim str, i
	Dim description			: description = ""
	Dim skills				: skills = ""
	Dim skillList			: skillList = page.Program.SkillList("")
	
	' 0-SkillID 1-SkillName 2-SkillDesc 3-IsSkillEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-IsGroupEnabled 8-LastModified 9-DateCreated

	str = str & "<h3>Program Details</h3>"
	str = str & "<div class=""item""><img class=""icon"" src=""/_images/icons/script.png"" alt="""" />"
	str = str & "<p><strong>" & html(page.Client.NameClient) & " | " & html(page.Program.ProgramName) & "</strong> "
	description = page.Program.ProgramDesc
	If Len(description & "") = 0 Then description = "No description provided"
	str = str & "<br />" & html(description) & ". </p>"
	
	str = str & "<p><strong>Program Skills: </strong>"
	If IsArray(skillList) Then
		For i = 0 To UBound(skillList,2)
			' return enabled, skillgroup enabled skills
			If skillList(7,i) = 1 Then
				If skillList(3,i) = 1 Then
					skills = skills & skillList(1,i) & ", "
				End If
			End If
		Next
		If Len(skills) > 0 Then 
			skills = Left(skills, Len(skills) - 2)
			str = str & skills
		Else
			str = str & "None selected. "
		End If
	Else
		str = str & "No skills are configured for this program. "
	End If
	str = str & "</p></div>"
	
	AvailableProgramDetailsToString = str
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

Function MemberGridForProgramSummaryToString(page, count)
	Dim str, i
	
	Dim xml				: Set xml = Server.CreateObject("MSXML2.DOMDocument")
	
	Dim list			: list = page.Program.MemberList()
	
	Dim rows			: rows = ""
	Dim alt				: alt = ""
	
	Dim isMemberEnabled	
	Dim isProgramMemberEnabled
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 6-ProgramMemberIsEnabled 7-MemberActiveStatus 8-DateCreated 9-DateModified
	' 10-ProgramMemberID 11-IsApproved 12-HasMissingAvailability 13-Email 14-MemberSkillDetailsXmlFragment
	
	count = 0
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isMemberEnabled = True			: If list(7,i) = 0 Then isMemberEnabled = False
			
			isProgramMemberEnabled = True	: If list(6,i) = 0 Then isProgramMemberEnabled = False
			' always show logged in member even if disabled ..
			If CStr(list(0,i) & "") = CStr(page.Member.MemberId & "") Then isProgramMemberEnabled = True
			
			If isMemberEnabled And isProgramMemberEnabled Then
				alt = "":			If count Mod 2 > 0 Then alt = " class=""alt"""
				rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/user.png"" alt="""" />"
				rows = rows & "<strong>" & html(list(1,i) & ", " & list(2,i)) & "</strong></td>"
				rows = rows & "<td>" & MemberSkillListingToString(list(14,i), xml) & "</td></tr>"
				
				count = count + 1
			End If
		Next
	End If
	
	str = str & "<div class=""grid""><table>"
	str = str & "<thead><tr><th>Member Name</th><th>Skills</th></tr></thead>"
	str = str & "<tbody>" & rows & "</tbody></table></div>"
	
	If count = 0 Then
		str = "<p class=""alert"">No members are assigned to this program. </p>"
	End If
	
	MemberGridForProgramSummaryToString = str
End Function

Function EventGridForProgramSummaryToString(page)
	Dim str, i
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim list				: list = page.Member.EventList(Now(), page.Program.ProgramId, Null)
	
	Dim rows				: rows = ""
	Dim alt					: alt = ""
	Dim count				: count = 0
	Dim isEventVisible		: isEventVisible = True
	
	Dim availableText		: availableText = ""
	Dim scheduledText		: scheduledText = ""
	
	' 0-EventID 1-EventName 3-EventDate 4-TimeStart 5-TimeEnd 8-ProgramName
	' 9-ProgramID 11-ScheduleName 13-ScheduleIsVisible 14-IsScheduled 15-EventList 17-IsAvailable 
	' 18-EventAvailabilityID 19-ProgramEnabled 20-IsProgramMemberEnabled 
	' 21-IsAvailabilityViewedByMember 
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isEventVisible = True		: If list(13,i) = 0 Then isEventVisible = False

			If isEventVisible Then 
				alt = "":					: If count Mod 2 > 0 Then alt = " class=""alt"""
				scheduledText = "Yes"		: If list(14,i) = 0 Then scheduledText = "No"
				
				availableText = "Yes"		: If list(17,i) = 0 Then availableText = "No"
				If list(21,i) = 0 Then availableText = "??"
				
				rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
				rows = rows & "<strong>" & html(list(1,i)) & "</strong></td>"
				rows = rows & "<td>" & dateTime.Convert(list(3,i), "DDD MMM dd, YYYY")
				If Len(list(4,i)) > 0 Then rows = rows & " at " & dateTime.Convert(list(4,i), "hh:nn pm")
				rows = rows & "</td>"
				rows = rows & "<td>" & availableText & "</td>"
				rows = rows & "<td style=""width:1%;"">" & scheduledText & "</td>"
				rows = rows & "</tr>"
				
				count = count + 1
			End If
		Next
	End If
	
	str = str & "<div class=""grid""><table>"
	str = str & "<thead><tr><th>Event</th><th>Date/Time</th><th>Available</th><th>Scheduled</th></tr></thead>"
	str = str & "<tbody>" & rows & "</tbody></table></div>"
	
	If count = 0 Then
		str = "<p class=""alert"">This program has no events, or all events occurred in the past. </p>"
	End If
	
	EventGridForProgramSummaryToString = str
End Function

Function ProgramSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim programMember		: Set programMember = New cProgramMember
	programMember.ProgramId = page.Program.ProgramId
	programMember.MemberId = page.Member.MemberId
	programMember.LoadByMemberProgram()

	Dim isEnrolled			: isEnrolled = True
	If Len(programMember.ProgramMemberId) = 0 Then isEnrolled = False
	
	Dim memberCount
	Dim countText			
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	pg.Action = "": pg.ProgramMemberId = ""
	str = str & "<ul><li><a href=""/member/schedules.asp" & pg.UrlParamsToString(True) & """>View " & html(page.Program.ProgramName) & " events on my calendar</a></li>"
	str = str & "</ul></div>"
	
	str = str & "<h3>" & html(page.Program.ProgramName) & "</h3>"
	
	str = str & "<div class=""summary"">"
	str = str & "<h5 class=""description"">Description</h5>"
	If Len(page.Program.ProgramDesc) > 0 Then 
		str = str & "<p>" & html(page.Program.ProgramDesc) & "</p>"
	Else
		str = str & "<p class=""alert"">No description available. </p>"
	End If
	
	If isEnrolled Then
		If programMember.IsActive = 0 Then
			str = str & "<h5 class=""disabled"">Disabled</h5>"
			str = str & "<p class=""alert"">You have this program disabled in your profile. </p>"
		End If
	End If
	
	str = str & "<h5 class=""schedule"">Events</h5>"
	str = str & EventGridForProgramSummaryToString(page)
	
	str = str & "<h5 class=""program-member"">Members</h5>"
	str = str & MemberGridForProgramSummaryToString(page, memberCount)
	
	countText = memberCount & " member"
	If memberCount <> 1 Then countText = countText & "s"
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5><ul>"
		
	str = str & "<li>" & html(page.Program.ProgramName) & " has " & countText & "</li>"
	str = str & "<li>It was created " & dateTime.Convert(page.Program.DateCreated, "DDD MMMM dd, YYYY") & "</li>"
	
	If isEnrolled Then
		str = str & "<li>You joined on " & dateTime.Convert(programMember.DateCreated, "DDD MMMM dd, YYYY around hh:00 pm") & ". </li>"
		If programMember.IsActive = 0 Then
			str = str & "<li>You have this program disabled in your profile. </li>"
		End If
	End If
	str = str & "</ul>"
	
	str = str & "</div>"
	
	ProgramSummaryToString = str
End Function

Function AvailableProgramGridToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim count			: count = 0
	Dim restrictedText	: restrictedText = ""
	Dim altClass
	
	page.ProgramMember.MemberID = page.Member.MemberID
	Dim list		: list = page.ProgramMember.GetAvailableProgramList()
	
	Dim isProgramEnabled
	Dim canMemberEnroll
	
	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "<p>Click <strong>Add</strong> in the toolbar for any program to add it to your profile. "
	str = str & "<br /><br />You will need an administrator to add you to any restricted programs. </p></div>"
	
	' 0-ProgramID 1-ProgramName 2-Description 3-EnrollmentType 4-IsEnabled
	' 5-DateCreated 6-DateModified 7-MemberCount 8-SkillCount 9-MemberCanEnroll
	
	If IsArray(list) Then
		str = str & "<h3 class=""first"">Available programs for " & html(page.Client.NameClient) & " </h3>"
		str = str & "<p>This listing contains available programs for " & Server.HtmlEncode(page.Client.NameClient) & " that you could add to your profile. "
		str = str & "Click <strong>Add</strong> in the toolbar for the relavant program to add it to your profile. </p>"
		
		str = str & "<div class=""grid"">"
		str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
		str = str & "<input type=""checkbox"" name=""master"" onclick="""" disabled=""disabled"" checked=""checked"" /></th>"
		str = str & "<th scope=""col"">Program</th><th scope=""col"">Enrollment</th><th scope=""col""></th></tr>"
		For i = 0 To UBound(list,2)
			isProgramEnabled = True				: If list(4,i) = 0 Then isProgramEnabled = False
			' if this program is enabled ..
			If isProgramEnabled Then
				restrictedText = "Open"
				If list(9,i) = 0 Then restrictedText = "<span style=""color:red;"">Restricted</span>"
				altClass = ""
				If count Mod 2 > 0 Then altClass = " class=""alt"""
				
				str = str & "<tr" & altClass & "><td><input type=""checkbox"" name=""Unused"" disabled=""disabled"" checked=""checked"" /></td>"
				str = str & "<td><img class=""icon"" src=""/_images/icons/script.png"" alt=""icon"" />"
				str = str & "<strong>" & html(page.Client.NameClient) & " | "
				pg.ProgramID = list(0,i): pg.Action = SHOW_PROGRAM_DETAILS
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(1,i)) & "</a></strong></td>"
				str = str & "<td>" & restrictedText & "</td>"
				str = str & "<td class=""toolbar"">"
				pg.ProgramID = list(0,i): pg.Action = SHOW_PROGRAM_DETAILS
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """><img src=""/_images/icons/magnifier.png"" alt=""icon"" /></a>"

				canMemberEnroll = True			: If list(9,i) = 0 Then canMemberEnroll = False
				If canMemberEnroll Then
					pg.ProgramID = list(0,i): pg.Action = ADDNEW_RECORD
					str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Add"">"
					str = str & "<img src=""/_images/icons/plus.png"" alt=""Add"" /></a>"
				Else
					str = str & "<img src=""/_images/icons/delete.png"" alt=""Restricted"" />"
				End If
				str = str & "</td></tr>"
				
				count = count + 1
			End If
		Next
		str = str & "</table></div>"
	End If
	
	If count = 0 Then
		str = NoAvailableProgramsDialogToString(page)
	End If
	
	AvailableProgramGridToString = str
End Function

Function NoAvailableProgramsDialogToString(page)
	Dim dialog				: Set dialog = New cDialog
	Dim pg					: Set pg = page.Clone()
	
	dialog.Headline	= "Whoa, something's missing here ..!"

	dialog.Text = dialog.Text & "<p>Were you expecting to see a list of available programs here? "
	dialog.Text = dialog.Text & "It looks like you've already added every " & html(page.Client.NameClient) & " program to your profile. </p>"
	
	dialog.SubText = dialog.SubText & "<p>This page is where you can add programs to your " & html(page.Client.NameClient) & " profile. "
	dialog.SubText = dialog.SubText & ""
	dialog.SubText = dialog.SubText & ""
	dialog.SubText = dialog.SubText & ""
	dialog.SubText = dialog.SubText & "</p>"

	pg.Action = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Back to my program list</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/help/topic.asp?hid=6"" target=""_blank"">Learn more about programs</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""""></a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""""></a></li>"

	NoAvailableProgramsDialogToString = dialog.ToString()
End Function

Function ProgramGridToString(page)
	Dim str, i
	Dim pg			: Set pg = page.Clone()
	Dim list		: list = page.Member.ProgramList()
	
	Dim enabledIcon	: enabledIcon = ""
	Dim programIcon	: programIcon = ""
	Dim altClass
	
	If Not page.Client.HasPrograms Then
		ProgramGridToString = NoProgramsForClientDialogToString(page)
		Call ClearTablinkBar()
		Exit Function
	End If
	
	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "<p><strong>Click Add Program</strong> in the toolbar to add a program to your profile. </p></div>"

	str = str & m_appMessageText

	' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-ProgramMemberID 4-SkillListXmlFrag 5-EnrollStatusID 6-EnrollStatusText
	' 7-EnrollStatusDesc 8-IsLeader 9-IsLeaderText 10-IsActive 11-IsActiveText 12-LastScheduleView 
	' 13-DateModified 14-DateCreated 15-SkillCount 16-IsMissingAvailability 17-IsAdmin 18-ProgramIsEnabled

	If IsArray(list) Then
		
		str = str & "<h3>My " & html(page.Client.NameClient) & " programs</h3>"
		str = str & "<p>A listing of all the " & Server.HtmlEncode(page.Client.NameClient) & " programs that you belong to. "
		str = str & "You can make changes by clicking the relevant button in the toolbar for that program. </p>"

		str = str & "<div class=""grid"">"
		str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
		str = str & "<input type=""checkbox"" name=""master"" onclick="""" disabled=""disabled"" checked=""checked"" /></th>"
		str = str & "<th scope=""col"">Program</th><th scope=""col""></th></tr>"
		For i = 0 To UBound(list,2)
		
			altClass = ""
			If i Mod 2 <> 0 Then altClass = " class=""alt"""
			enabledIcon = "lightning_add.png"
			programIcon = "script.png"
			If list(10,i) = 0 Then 
				enabledIcon = "lightning_delete.png"
				programIcon = "script_delete.png"
			End If
			
			str = str & "<tr" & altClass & "><td><input type=""checkbox"" name=""Unused"" disabled=""disabled"" checked=""checked"" /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/" & programIcon & """ alt=""icon"" />"
			str = str & "<strong>" & html(page.Client.nameClient) & " | " 
			pg.MessageID = "": pg.ProgramMemberID = list(3,i): pg.ProgramID = list(0,i): pg.Action = SHOW_PROGRAM_DETAILS
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(1,i)) & "</a></strong></td>"
			
			str = str & "<td class=""toolbar"">"
			
			' only show toolbar if program is enabled ..
			If list(18,i) = 1 Then
				pg.MessageID = "": pg.ProgramMemberID = list(3,i): pg.ProgramID = list(0,i): pg.Action = SHOW_PROGRAM_DETAILS
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details"">"
				str = str & "<img src=""/_images/icons/magnifier.png"" alt=""Details"" /></a>"
				pg.MessageID = "": pg.Action = "": pg.ProgramMemberID = "": pg.ProgramID = list(0,i)
				str = str & "<a href=""/member/schedules.asp" & pg.UrlParamsToString(True) & """ title=""Calendar"">"
				str = str & "<img src=""/_images/icons/calendar.png"" alt=""Calendar"" /></a>"
				
				If list(10,i) = 1 Then
					pg.MessageID = "": pg.ProgramID = list(0,i): pg.ProgramMemberID = ""
					str = str & "<a href=""/member/events.asp" & pg.UrlParamsToString(True) & """ title=""Availability"">"
				Else
					pg.ProgramID = "": pg.ProgramMemberID = "": pg.Action = "": pg.MessageID = 3040
					str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Availability"">"
				End If
				str = str & "<img src=""/_images/icons/clock.png"" alt=""Availability"" /></a>"
				
				If list(10,i) = 1 Then
					pg.ProgramID = list(0,i): pg.ProgramMemberID = list(3,i): pg.Action = "": pg.MessageID = ""
					str = str & "<a href=""/member/skills.asp" & pg.UrlParamsToString(True) & """ title=""Skills"">"
				Else
					pg.ProgramID = "": pg.ProgramMemberID = "": pg.Action = "": pg.MessageID = 3041
					str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Skills"">"
				End If
				str = str & "<img src=""/_images/icons/plugin.png"" alt=""Skills"" /></a>"
				
				pg.MessageID = "": pg.Action = TOGGLE_PROGRAM_MEMBER_IS_ACTIVE: pg.ProgramMemberID = list(3,i)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Enable"">"
				str = str & "<img src=""/_images/icons/" & enabledIcon & """ alt=""Enable"" /></a>"
				
				pg.MessageID = "": pg.Action = "": pg.ProgramMemberID = "": page.ProgramID = list(0,i)
				str = str & "<a href=""/member/contacts.asp" & pg.UrlParamsToString(True) & """ title=""Contacts"">"
				str = str & "<img src=""/_images/icons/report_user.png"" alt=""Contacts"" /></a>"
				
				pg.MessageID = "": pg.programMemberID = list(3,i): pg.Action = DELETE_RECORD
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove"">"
				str = str & "<img src=""/_images/icons/cross.png"" alt=""Remove"" /></a>"
			Else
				str = str & "<strong style=""color:red;"">Program disabled by admin</strong>"
			End If
			str = str & "</td></tr>"
		Next
		str = str & "</table></div>"
	Else
		str = NoProgramsInProfileDialogToString(page)
	End If
	
	ProgramGridToString = str
End Function

Function FormConfirmDeleteProgramMemberToString(page)
	Dim str, msg
	Dim pg				: Set pg = page.Clone()
	
	msg = msg & "You will permanently remove the program <strong>" & html(page.ProgramMember.ProgramName) & "</strong> from your <strong>" & html(page.Client.NameClient) & "</strong> " & Application.Value("APPLICATION_NAME") & " profile. "
	msg = msg & "You will lose any associated schedule and calendar information. "
	msg = msg & "This action cannot be reversed. "
	
	str = str & CustomApplicationMessageToString("Please confirm this action!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formConfirmDeleteProgramMember"">"
	str = str & "<input type=""hidden"" name=""FormConfirmDeleteProgramMemberIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Confirm"" />"
	pg.Action = "": pg.ProgramMemberID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "</p></form>"
	FormConfirmDeleteProgramMemberToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	pg.Action = ""
	Dim programsHref		: programsHref = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Programs</a> / "
	
	str = str & "<a href=""/member/overview.asp"">Member Home</a> / "
	Select Case page.Action
		Case SHOW_AVAILABLE_PROGRAM_DETAILS
			str = str & programsHref & html(page.Program.ProgramName)
		Case SHOW_PROGRAM_DETAILS
			str = str & programsHref & html(page.Program.ProgramName)
		Case DELETE_RECORD
			str = str & programsHref & "Remove program"
		Case SHOW_AVAILABLE_PROGRAMS
			str = str & programsHref & "Add program"
		Case Else
			str = str & "Programs"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href		: href = ""
	
	Dim programListButton
	pg.Action = "": pg.ProgramID = "": pg.ProgramMemberID = ""
	href = pg.Url & pg.UrlParamsToString(True) 
	programListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/script.png"" /></a><a href=""" & href & """>Program List</a></li>"
	
	Dim addProgramButton
	pg.Action = SHOW_AVAILABLE_PROGRAMS
	href = pg.Url & pg.UrlParamsToString(True) 
	addProgramButton = addProgramButton & "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/script_add.png"" alt="""" /></a><a href=""" & href & """>Add Program</a></li>"
	
	Select Case page.Action
		Case SHOW_AVAILABLE_PROGRAM_DETAILS
			str = str & addProgramButton
		Case SHOW_PROGRAM_DETAILS
			str = str & programListButton
		Case DELETE_RECORD
			str = str & programListButton
		Case SHOW_AVAILABLE_PROGRAMS
			str = str & programListButton
		Case Else
			str = str & addProgramButton
	End Select
	
	' use this if no buttons will be displayed ..
	' str = str & "<li>&nbsp;</li>"
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_NoProgramsForClientDialogToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_NoProgramsInProfileDialogToString.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public ProgramID
	Public ProgramMemberID
	
	' objects
	Public Member
	Public Client
	Public Program	
	Public ProgramMember
	
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
		If Len(ProgramMemberID) > 0 Then str = str & "pmid=" & Encrypt(ProgramMemberID) & amp
		
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
		c.ProgramID = ProgramID
		c.ProgramMemberID = ProgramMemberID
		c.Action = Action
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.ProgramMember = ProgramMember
		
		Set Clone = c
	End Function
End Class
%>

