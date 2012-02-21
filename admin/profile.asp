<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-members"
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
	page.MemberID = Decrypt(Request.QueryString("mid"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.ProgramMemberID = Decrypt(Request.QueryString("pmid"))
	page.SkillID = Decrypt(Request.QueryString("skid"))
	page.EventID = Decrypt(Request.QueryString("eid"))
	
	page.ProgramIDList = Request.Form("ProgramIDList")
	page.SkillIDList = Request.Form("SkillIDList")
	
	If Request.Form("FormGotoMemberDropdownIsPostback") = IS_POSTBACK Then
		If Len(Request.Form("NewMemberID")) > 0 Then
			page.MemberID = Request.Form("NewMemberID")
			page.Action = SHOW_MEMBER_DETAILS
		End If
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
	Set page.ThisMember = New cMember
	page.ThisMember.MemberID = page.MemberID
	If Len(page.ThisMember.MemberID) > 0 Then Call page.ThisMember.Load()
	Set page.ProgramMember = New cProgramMember
	page.ProgramMember.ProgramMemberID = page.ProgramMemberID
	If Len(page.ProgramMember.ProgramMemberID) > 0 Then Call page.ProgramMember.Load()
	
	page.Settings = GetApplicationSetting(page.Client.ClientID, "MemberRequiredField")

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
		<link rel="stylesheet" type="text/css" href="profile.css" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="profile.js"></script>
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
		Case SHOW_PROGRAM_MEMBER_DETAILS
			str = str & MemberProgramSummaryToString(page)
			
		Case SHOW_MEMBER_DETAILS
			str = str & MemberAccountSummaryToString(page)
		
		Case TOGGLE_PROGRAM_ACTIVE_STATUS
			' check that program is enabled ..
			If page.Program.IsEnabled = 0 Then
				page.Action = "": page.MessageID = 3045: page.ProgramMemberID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
			Call DoToggleProgramMemberIsActive(page, rv)
			Select Case rv
				Case 0
					page.MessageID = 3015
				Case Else
					page.MessageID = 3016
			End Select
			page.Action = "": page.ProgramMemberID = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case TOGGLE_CLIENT_ACTIVE_STATUS
			Call DoToggleMemberIsActive(page.ThisMember, rv)
			Select Case rv
				Case 0
					page.MessageID = 2013
				Case Else
					page.MessageID = 2014
			End Select
			page.Action = "":
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case REMOVE_PROGRAM_MEMBER
			' check that program is enabled ..
			If page.Program.IsEnabled = 0 Then
				page.Action = "": page.MessageID = 3045: page.ProgramMemberID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
			' check for programMember enabled ..
			If page.ProgramMember.IsActive = 0 Then
				page.MessageID = 3046: page.Action = "": page.ProgramMemberID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
			If Request.Form("FormConfirmDeleteProgramMemberIsPostback") = IS_POSTBACK Then
				Call DoDeleteProgramMember(page, rv)
				Select Case rv
					Case 0
						page.MessageID = 3017
					Case Else
						page.MessageID = 3018
				End Select
				page.Action = "": page.ProgramMemberID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmDeleteProgramMemberToString(page)
			End If		
			
		Case REMOVE_CLIENT_MEMBER
			If Request.Form("FormConfirmDeleteMemberIsPostback") = IS_POSTBACK Then
				Call DoDeleteMember(page, rv)
				Select Case rv
					Case 0 
						page.MessageID = 1019
					Case -3
						' member not owned
						page.MessageID = 1021
					Case -4
						' self-delete
						page.MessageID = 1022
					Case -5 
						' delete last admin
						page.MessageID = 1040 
					Case Else
						page.MessageID = 1020 
				End Select
				page.Action = "": page.MemberID = ""
				Response.Redirect("/admin/members.asp" & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmDeleteMemberToString(page)
			End If
			
		Case SEND_MESSAGE
			Call GenerateEmail(page, rv)
			page.Action = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
			
		Case UPDATE_MEMBER_AVAILABILITY
			' check that program is enabled ..
			If page.Program.IsEnabled = 0 Then
				page.Action = "": page.MessageID = 3045: page.ProgramMemberID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
			' check for programMember enabled ..
			If page.ProgramMember.IsActive = 0 Then
				page.MessageID = 3046: page.Action = "": page.ProgramMemberID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
			If Request.Form("FormAvailabilityIsPostback") = IS_POSTBACK Then
				Call DoUpdateAvailability(page, rv)
				Select Case rv
					Case 0
						page.MessageID = 3037
					Case Else
						page.MessageID = 3038
				End Select
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & MemberAvailabilityGridToString(page)
			End If
			
		Case CONFIGURE_PROGRAM_MEMBER_SKILLS
			' check that program is enabled ..
			If page.Program.IsEnabled = 0 Then
				page.Action = "": page.MessageID = 3045: page.ProgramMemberID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
			' check for programMember enabled ..
			If page.ProgramMember.IsActive = 0 Then
				page.MessageID = 3046: page.Action = "": page.ProgramMemberID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
			If Request.Form("FormConfirmConfigureSkillsIsPostback") = IS_POSTBACK Then
				Call UpdateProgramMemberSkill(page, rv)
				Select Case rv
					Case 0
						page.MessageID = 1062
					Case Else
						page.MessageID = 1063
				End Select
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			ElseIf Request.Form("FormSkillIsPostback") = IS_POSTBACK Then
				str = str & FormConfirmConfigureSkillsToString(page)
			Else
				str = str & MemberSkillGridToString(page)
			End If
				
		Case SHOW_AVAILABLE_PROGRAMS
			If Request.Form("FormConfirmSelectProgramsIsPostback") = IS_POSTBACK Then
				Call DoUpdateProgramsByIDList(page, rv)
				Select Case rv
					Case 0
						page.MessageID = 3042
					Case Else
						page.MessageID = 3043
				End Select
				Response.Redirect(page.Url & page.UrlParamsToString(False))
				
			ElseIf Request.Form("FormSelectProgramsIsPostback") Then 
				str = str & FormConfirmSelectProgramsToString(page)
			Else
				str = str & FormSelectProgramsToString(page)
			End If
			
		Case IMPERSONATE_MEMBER
			If Request.Form("FormConfirmImpersonateIsPostback") = IS_POSTBACK Then
				Call ImpersonateMember(page.ThisMember.MemberID, rv)
				Response.Redirect("/member/programs.asp")
			Else
				str = str & FormConfirmImpersonateToString(page, page.ThisMember)
			End If
		
		Case SEND_MEMBER_CREDENTIALS 
			If Request.Form("FormConfirmSendCredentialsIsPostback") = IS_POSTBACK Then
				Call SendCredentials(page.ThisMember.MemberID, page.Member.Email)
				page.Action = "": page.MessageID = 1037
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmSendCredentialsToString(page)
			End If
				
		Case UPDATE_PASSWORD
			If Request.Form("FormPasswordIsPostback") = IS_POSTBACK Then
				Call LoadMemberPasswordFromForm(page.ThisMember)
				If ValidFormPassword(page.Member) Then
					Call DoUpdatePassword(page.ThisMember, rv)
					Select Case rv
						Case 0
							page.MessageID = 1055
						Case Else
							page.MessageID = 1056
					End Select
					page.Action = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormPasswordToString(page, page.ThisMember)
				End If
			Else
				str = str & FormPasswordToString(page, page.ThisMember)
			End If
			
		Case UPDATE_RECORD
			If Request.Form("FormMemberIsPostback") = IS_POSTBACK Then
				Call LoadMemberFromPost(page.ThisMember)
				If ValidFormMember(page.ThisMember, page.Settings) Then
					Call DoUpdateMember(page.ThisMember, rv)
					Select Case rv
						Case 0
							page.MessageID = 1051
						Case -2
							' dupe member
							page.MessageID = 1053
						Case -3
							' dupe login
							page.MessageID = 1052
						Case Else
							' unknown error
							page.MessageID = 1054
					End Select
					page.Action = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormMemberToString(page, page.ThisMember)
				End If
			Else
				str = str & FormMemberToString(page, page.ThisMember)
			End If
			
		Case Else
			str = str & MemberAccountSettingsToString(page)
			
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub GenerateEmail(page, outError)
	Dim member		: Set member = New cMember
	member.MemberId = page.ThisMember.MemberId
	member.Load()

	Dim email		: Set email = New cEmail
	email.MemberID = page.Member.MemberID
	email.ClientID = page.Client.ClientID
	email.RecipientAddressList = member.Email()
	
	Call email.Add(outError)
	page.EmailID = email.EmailID
	
	Set email = Nothing
End Sub

Sub DoDeleteProgramMember(page, outError)
	Call page.ProgramMember.Delete(outError)
End Sub

Sub DoToggleProgramMemberIsActive(page, outError)
	
	If page.ProgramMember.IsActive = 0 Then
		page.ProgramMember.IsActive = 1
	Else
		page.ProgramMember.IsActive = 0
	End If
	Call page.ProgramMember.Save(outError)
End Sub

Sub DoToggleMemberIsActive(member, outError)
	If member.ActiveStatus = 1 Then
		member.ActiveStatus = 0
	Else
		member.ActiveStatus = 1
	End If
	Call member.Update(outError)
End Sub

Sub DoUpdateAvailability(page, outError)
	Dim i
	Dim localError
	Dim eventAvailability		: Set eventAvailability = New cEventAvailability
	
	eventAvailability.MemberID = page.ThisMember.MemberID
	Dim list					: list = eventAvailability.AvailabilityList(page.Program.ProgramID, "", "")
	Dim isAvailable
	outError = 0
	
	' 0-EventAvailabilityID 3-IsAvailable 
	
	If Not IsArray(list) Then Exit Sub
	For i = 0 To UBound(list,2)		
		' get a value from the form
		isAvailable = 0
		If Len(Request.Form("IsAvailable" & list(0,i))) > 0 Then
			isAvailable = 1
		End If
		eventAvailability.EventAvailabilityID = list(0,i)
		eventAvailability.Load()
		eventAvailability.IsAvailable = isAvailable
		eventAvailability.Save(localError)
		If localError <> 0 Then outError = -1
	Next
End Sub

Sub UpdateProgramMemberSkill(page, outError)
	Dim str, i, j, rv, deleteThis, hasError
	
	Dim programMemberSkill		: Set programMemberSkill = New cProgramMemberSkill
	
	' list of skillIDs from the form
	Dim list					: list = Split(Replace(page.SkillIDList, " ", ""), ",")
	
	' list of existing programMemberSkills
	Dim currentList				: currentList = page.programMember.GetSkillList("")
	
	' 0-SkillID 1-ProgramMemberSkillID 2-ProgramMemberID 3-SkillName 4-SkillDesc 5-SkillGroupName
	' 6-IsProgramMemberSkill 7-DateCreated
	
	hasError = False
	outError = 0
	
	If IsArray(list) Then
	
		' if it's in the form but not isProgramMemberSkill
		For i = 0 To UBound(currentList,2)
			' iterate through the IDs from the form
			For j = 0 To UBound(list)
				' found one ..
				If CLng(currentList(0,i)) = CLng(list(j)) Then
					' if not isProgramMemberSkill then add it
					If CInt(currentList(6,i)) = 0 Then
						' add the programMemberSkill
						programMemberSkill.ProgramMemberID = page.ProgramMemberID
						programMemberSkill.SkillID = list(j)
						Call programMemberSkill.Add(rv)
						If rv <> 0 Then hasError = True
						Exit For
					End If
				End If
			Next
		Next
		
		' if it's in the db but not in the form
		For i = 0 To UBound(currentList,2)
			' isProgramMemberSkill
			If CInt(currentList(6,i)) = 1 Then
				' look for it in form
				deleteThis = True
				For j = 0 To UBound(list)
					' if it's in the form then clear the flag
					If CLng(currentList(0,i)) = CLng(list(j)) Then
						deleteThis = False
						Exit For
					End If
				Next
				If deleteThis Then
					programMemberskill.ProgramMemberSkillID = CLng(currentList(1,i))
					Call programMemberSkill.Delete(rv)
					If rv <> 0 Then hasError = True
				End If
			End If
		Next
	End If
	
	If hasError Then outError = -1
	
	Set programMemberSkill = Nothing
End Sub

Sub DoUpdateProgramsByIDList(page, outError)
	Dim i, j
	
	Dim tempError						: tempError = 0
	outError = 0
	
	Dim idList							: idList = Split(page.ProgramIDList, ",")
	Dim memberProgramList				: memberProgramList = page.ThisMember.ProgramList()
	' 0-ProgramID 1-ProgramName 
		
	Dim deleteProgram					: deleteProgram = False
	Dim addProgram						: addProgram = True
	Dim thisID
	
	' find programs to add ..
	For i = 0 To UBound(idList)
		addProgram = True
		If IsArray(memberProgramList) Then
			For j = 0 To UBound(memberProgramList,2)
				If CLng(idList(i)) = CLng(memberProgramList(0,j)) Then
					addProgram = False
				End If
			Next
		End If
		
		If addProgram Then
			page.ProgramMember.MemberID = page.ThisMember.MemberID
			page.ProgramMember.ProgramID = idList(i)
			page.ProgramMember.EnrollStatusID = PROGRAM_ENROLL_STATUS_ACCEPTED
			page.ProgramMember.IsActive = PROGRAM_MEMBER_IS_ACTIVE
			page.ProgramMember.Add(tempError)
			outError = outError + tempError
		End If
	Next
	
	' find programs to remove 
	If IsArray(memberProgramList) Then
		For i = 0 To UBound(memberProgramList,2)
			deleteProgram = True
			For j = 0 To UBound(idList)
				If CLng(memberProgramList(0,i)) = CLng(idLIst(j)) Then
					deleteProgram = False
				End If
			Next
			
			If deleteProgram Then
				page.ProgramMember.MemberID = page.ThisMember.MemberID
				page.ProgramMember.ProgramID = memberProgramList(0,i)
				page.ProgramMember.LoadByMemberProgram()
				page.ProgramMember.Delete(tempError)
				outError = outError + tempError
			End If
		Next
	End If
End Sub

Sub ImpersonateMember(memberID, outError)
	Dim sess		: Set sess = New cSession
	outError = 0

	sess.SessionID = Request.Cookies("sid")
	sess.Load()
	sess.MemberID = memberID
	sess.IsImpersonated = 1
	sess.IsAdmin = 0 
	sess.IsLeader = 0
	call sess.Save(outError)

	Set sess = Nothing
End Sub

Sub DoUpdatePassword(member, outError)
	Call member.Update(outError)
End Sub

Sub DoUpdateMember(member, outError)
	Call member.Update(outError)
End Sub

Function MemberSkillGridForSummaryToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	Dim list			: list = page.ProgramMember.GetSkillList("")
	Dim count			: count = 0
	Dim alt				: alt = ""
	Dim rows			: rows = ""
	Dim href			: href = ""
	
	Dim isSkillEnabled
	Dim isSkillGroupEnabled
	Dim hasSkill
	
	Dim groupName
	
	' 0-SkillID 1-ProgramMemberSkillID 2-ProgramMemberID 3-SkillName 4-SkillDesc 5-SkillGroupName
	' 6-IsProgramMemberSkill 7-DateCreated 8-SkillIsEnabled 9-SkillGroupIsEnabled

	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			hasSkill = True						: If list(6,i) = 0 Then hasSkill = False
			isSkillEnabled = True				: If list(8,i) = 0 Then isSkillEnabled = False
			isSkillGroupEnabled = True			: If list(8,i) = 0 Then isSkillGroupEnabled = False
			
			If hasSkill And isSkillEnabled And isSkillGroupEnabled Then
				alt = ""					: If count Mod 2 <> 0 Then alt = " class=""alt"""
				groupName = list(5,i)		: If Len(groupName) = 0 Then groupName = " - "
				
				pg.Action = SHOW_SKILL_DETAILS: pg.SkillId = list(0,i): pg.MemberId = "": pg.ProgramMemberId = ""
				href = "/admin/skills.asp" & pg.UrlParamsToString(True)
				
				rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/plugin.png"" alt="""" />"
				rows = rows & "<a href="""& href & """><strong>" & html(list(3,i)) & "</strong></a></td>"
				rows = rows & "<td>" & html(groupName) & "</td>"
				rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
				rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
			
				count = count + 1
			End If
		Next
	End If	
	
	If count > 0 Then
		str = str & "<p>These are the " & html(page.ProgramMember.ProgramName) & " skills you have assigned to this member. </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Skill</th><th>Group</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str ="<p class=""alert"">No skills have beens set for this member. </p>"
	End If	
	
	MemberSkillGridForSummaryToString = str
End Function

Function MemberAvailabilityGridForSummaryToString(page, programId)
	Dim str, i
	Dim pg						: Set pg = page.Clone()
	Dim dateTime				: Set dateTime = New cFormatDate
	
	Dim list					: list = page.ThisMember.EventList(Null, Null, Null)
	Dim count					: count = 0
	Dim alt						: alt = ""
	Dim rows					: rows = ""
	Dim href
	
	Dim isViewedByMember
	Dim isAvailable
	Dim isThisProgramId
	Dim availableText

	Dim missingAvailabilityToken	: missingAvailabilityToken = "???"
	
	' 0-EventID 1-EventName 2-EventNote 3-EventDate 4-TimeStart 5-TimeEnd 6-ClientID
	' 7-MemberID 8-ProgramName 9-ProgramID 10-ScheduleID 11-ScheduleName 12-ScheduleDesc
	' 13-ScheduleIsVisible 14-IsScheduled 15-SkillListXmlFrag 16-FileListXmlFrag 17-IsAvailable
	' 18-EventAvailabilityID 19-ProgramIsEnabled 20-ProgramMemberIsActive 
	' 21-AvailabilityViewedByMember 22-MemberActiveStatus
		
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isThisProgramId = True
			If Len(programId & "") > 0 Then
				If CStr(list(9,i)) <> CStr(programId & "") Then isThisProgramId = False
			End If
		
			If isThisProgramId Then
				isViewedByMember = True					: If list(21,i) = 0 Then isViewedByMember = False
				isAvailable = True						: If list(17,i) = 0 Then isAvailable = False
				
				availableText = "Yes"
				If Not isAvailable Then availableText = "<span style=""color:red;"">No</span>"
				If Not isViewedByMember Then availableText = missingAvailabilityToken
				
				pg.EventId = list(0,i)
				pg.Action = SHOW_EVENT_DETAILS: pg.MemberId = "": pg.ProgramMemberId = "": pg.ProgramId = ""
				href= "/schedule/events.asp" & pg.UrlParamsToString(True)
				
				rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
				rows = rows & "<a href=""" & href & """ title=""Details""><strong>" & html(list(1,i)) & "</strong></a></td>"
				rows = rows & "<td>" & dateTime.Convert(list(3,i), "MM-DD-YY")
				If Len(list(4,i)) > 0 Then rows = rows & " at " & dateTime.Convert(list(4,i), "hh:nn pp")
				rows = rows & "</td>"
				rows = rows & "<td>" & html(list(11,i)) & "</td>"
				rows = rows & "<td>" & availableText & "</td>"
				rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
				rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
				
				count = count + 1
			End If
		Next
	End If
	
	If count > 0 Then
		str = str & "<p>These are the " & html(page.ProgramMember.ProgramName) & " events this member is available for "
		str = str & "(events showing '" & missingAvailabilityToken & "' in the available column are new events that haven't been viewed or saved by this member). </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Event</th><th>When</th><th>Schedule</th><th>Available</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str ="<p class=""alert"">No event information is available (or this program has no events). </p>"
	End If	
	
	MemberAvailabilityGridForSummaryToString = str
End Function

Function MemberProgramSummaryToString(page)
	Dim str
	Dim dateTime			: Set dateTime = New cFormatDate
	
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & page.ProgramMember.ProgramName & " summary for " & page.ThisMember.NameFirst & " " & page.ThisMember.NameLast & "</h3>"
	
	If page.ProgramMember.IsActive = 0 Then
		str = str & "<h5 class=""disabled"">Program disabled</h5>"
		str = str & "<p class=""alert"">This program (" & html(page.ProgramMember.ProgramName) & ") is set to disabled for this member. </p>"
	End If
	
	str = str & "<h5 class=""event-team"">Event teams</h5>"
	str = str & MemberEventGridForSummaryToString(page, page.ProgramMember.ProgramId)
	
	str = str & "<h5 class=""availability"">Event availability</h5>"
	str = str & MemberAvailabilityGridForSummaryToString(page, page.ProgramMember.ProgramId)
	
	str = str & "<h5 class=""skills"">Member skills</h5>"
	str = str & MemberSkillGridForSummaryToString(page)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul><li>Member of this program since " & dateTime.Convert(page.ProgramMember.DateCreated, "DDDD MMMM dd, YYYY around hh:00 pp") & ". </li>"
	str = str & "</ul>"
	
	str = str & "</div>"
	
	MemberProgramSummaryToString = str
End Function

Function MemberProgramGridForSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim xml					: Set xml = Server.CreateObject("MSXML2.DOMDocument")
	
	Dim list				: list = page.ThisMember.ProgramList()
	Dim rows				: rows = ""
	Dim count				: count = 0
	Dim href
	
	Dim isThisProgram		
	
	Dim enabledText			: enabledText = ""
	
	' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-ProgramMemberID 4-SkillListXmlFrag 5-EnrollStatusID 6-EnrollStatusText
	' 7-EnrollStatusDesc 8-IsLeader 9-IsLeaderText 10-IsActive 11-IsActiveText 12-LastScheduleView 
	' 13-DateModified 14-DateCreated 15-SkillCount 16-IsMissingAvailability 17-IsAdmin 18-ProgramIsEnabled
	
	If IsArray(list) Then
		For i = 0 To UBound(list, 2) 
			enabledText = "Yes"				: If list(10,i) = 0 Then enabledText = "<span style=""color:red;"">No</span>"
			
			pg.Action = SHOW_PROGRAM_MEMBER_DETAILS: pg.ProgramMemberId = list(3,i): pg.ProgramId = ""
			href = "/admin/profile.asp" & pg.UrlParamsToString(True)
			rows = rows & "<tr><td><img class=""icon"" src=""/_images/icons/script.png"" alt="""" />"
			rows = rows & "<a href=""" & href & """><strong>" & html(list(1,i)) & "</strong></a></td>"
			rows = rows & "<td>" & XmlFragmentToList(list(4,i), ", ", xml) & "</td>"
			rows = rows & "<td>" & enabledText & "</td>"
			rows = rows & "<td class=""toolbar""><a href=""" & href & """>"
			rows = rows &" <img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
			
			count = count + 1
		Next
	End If
	
	If count > 0 Then
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Program</th><th>Skills</th><th>Enabled</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else 
		str = "<p class=""alert"">This member does not belong to any programs. </p>"
	End If
	
	MemberProgramGridForSummaryToString = str
End Function

Function ContactInfoForAccountSummaryToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	
	Dim address
	Dim phoneList
	
	pg.Action = SEND_MESSAGE
	str = str & "<p><a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email"">" 
	str = str & "<strong>" & html(page.ThisMember.Email) & "</strong></a></p>"
	address = page.ThisMember.AddressToString()
	phoneList = page.ThisMember.PhoneListToString()
	If Len(address & "") > 0 Then
		str= str & "<p>" & address & "</p>"
	Else
		str = str & "<p class=""alert"">No address available. </p>"
	End If
	If Len(phoneList & "") > 0 Then
		str = str & "<p>" & page.ThisMember.PhoneListToString()
	Else
		str = str & "<p class=""alert"">No phone numbers available. </p>"
	End If
	
	ContactInfoForAccountSummaryToString = str
End Function

Function MemberEventGridForSummaryToString(page, programId)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate
	Dim xml					: Set xml = Server.CreateObject("MSXML2.DOMDocument")
	
	Dim list				: list = page.ThisMember.EventList(Null, Null, Null)
	Dim count				: count = 0
	Dim alt					: alt =""
	Dim skillListing		: skillListing = ""
	Dim rows				: rows = ""
	Dim href
	
	Dim isMemberEnabled
	Dim isProgramEnabled 		
	Dim isProgramMemberEnabled
	Dim isScheduled
	
	Dim isThisProgramId
	
	' 0-EventID 1-EventName 2-EventNote 3-EventDate 4-TimeStart 5-TimeEnd 6-ClientID
	' 7-MemberID 8-ProgramName 9-ProgramID 10-ScheduleID 11-ScheduleName 12-ScheduleDesc
	' 13-ScheduleIsVisible 14-IsScheduled 15-SkillListXmlFrag 16-FileListXmlFrag 17-IsAvailable
	' 18-EventAvailabilityID 19-ProgramIsEnabled 20-ProgramMemberIsActive 
	' 21-AvailabilityViewedByMember 22-MemberActiveStatus
		
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isThisProgramId = True
			If Len(programId & "") > 0 Then
				If CStr(list(9,i)) <> CStr(programId & "") Then isThisProgramId = False
			End If
		
			If isThisProgramId Then
				isMemberEnabled = True				: If list(22,i) = 0 Then isMemberEnabled = False
				isProgramMemberEnabled = True		: If list(20,i) = 0 Then isProgramMemberEnabled = False
				isProgramEnabled = True				: If list(19,i) = 0 Then isProgramEnabled = False
				isScheduled = True					: If list(14,i) = 0 Then isScheduled = False
				
				If isMemberEnabled And isProgramMemberEnabled And isProgramEnabled And isScheduled Then
					alt = ""							: If count Mod 2 <> 0 Then alt = " class=""alt"""
					
					pg.Action = SHOW_EVENT_DETAILS: pg.EventId = list(0,i): pg.MemberId = "": pg.ProgramMemberId = "": pg.ProgramId = ""
					href = "/schedule/events.asp" & pg.UrlParamsToString(True)
					
					skillListing = XmlFragmentToList(list(15,i), ", ", xml)
					
					rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
					If Len(programId & "") = 0 Then
						rows = rows & "<strong>" & html(list(8,i)) & "</strong> | "
					Else
						rows = rows & "<strong>" & html(list(11,i)) & "</strong> | "
					End If
					rows = rows & "<a href=""" & href & """ title=""Details""><strong>" & html(list(1,i)) & "</strong></a></td>"
					rows = rows & "<td>" & dateTime.Convert(list(3,i), "MM-DD-YY")
					If Len(list(4,i)) > 0 Then rows = rows & " " & dateTime.Convert(list(4,i), "hh:nn pp")
					rows = rows & "</td>"
					rows = rows & "<td>" & skillListing & "</td>"
					rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
					rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
					
					count = count + 1
				End If
			End If
		Next
	End If
	
	If count > 0 Then
		str = str & "<p>This member has been assigned to the event team for these events. </p>"
		str = str & "<div class=""grid""><table><thead>"
		str = str & "<tr><th>Event</th><th>When</th><th>For</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = "<p class=""alert"">This member is not assigned to an event team for any events "
		str = str & "(or their account and/or programs have been set to disabled). </p>"
	End If
	
	MemberEventGridForSummaryToString = str
End Function

Function MemberAccountSummaryToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim loginText			: loginText = "<strong>never</strong>"
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	pg.Action = SEND_MESSAGE
	str = str & "<ul><li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Email this member</a></li>"
	pg.Action = ""
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Change settings for this member</a></li>"
	str = str & "</ul></div>"

	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">Account Summary for " & html(page.ThisMember.NameFirst & " " & page.ThisMember.NameLast) & "</h3>"
	
	If page.ThisMember.ActiveStatus = 0 Then
		str = str & "<h5 class=""disabled"">Account disabled</h5>"
		str = str & "<p class=""alert"">This account is set to disabled. </p>"
	End If
	
	str = str & "<h5 class=""contact"">Contact</h5>"
	str = str & ContactInfoForAccountSummaryToString(page)
	
	str = str & "<h5 class=""program"">Programs</h5>"
	str = str & MemberProgramGridForSummaryToString(page)
	
	str = str & "<h5 class=""event-team"">Event teams</h5>"
	str = str & MemberEventGridForSummaryToString(page, Null)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul><li>" & html(page.Client.NameClient) & " account created on " 
	str = str & dateTime.Convert(page.ThisMember.DateCreated, "DDDD MMMM dd, YYYY") & ". </li>"
	If Len(page.ThisMember.LastLogin & "") > 0 Then loginText = dateTime.Convert(page.ThisMember.LastLogin, "DDDD MMMM dd, YYYY around hh:00 pp")
	str = str & "<li>Last login to this account was " & loginText & ". </li></ul>"
	
	str = str & "</div>"
	
	MemberAccountSummaryToString = str
End Function

Function OtherStuffForMemberAccountSettingsToString(page, list)
	Dim str, i
	Dim dateTime					: Set dateTime = New cFormatDate
	
	Dim items
	Dim isEnabled
	Dim hasSkills
	
	Dim lastLoginText				: lastLoginText = "This member has never logged in to their " & Application.Value("APPLICATION_NAME") & " account"
	If Len(page.ThisMember.LastLogin & "") > 0 Then lastLoginText = "Last login was " & dateTime.Convert(page.ThisMember.LastLogin, "DDDD, MMMM dd, YYYY around hh:00 pp")
	
	' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-ProgramMemberID 4-SkillListXmlFrag 5-EnrollStatusID 6-EnrollStatusText
	' 7-EnrollStatusDesc 8-IsLeader 9-IsLeaderText 10-IsActive 11-IsActiveText 12-LastScheduleView 
	' 13-DateModified 14-DateCreated 15-SkillCount 16-IsMissingAvailability 17-IsAdmin 18-ProgramIsEnabled
		
	If page.ThisMember.ActiveStatus = 0 Then
		items = items & "<li class=""error""><strong class=""warning"">This account is disabled!</strong> "
		items = items & "The " & Application.Value("APPLICATION_NAME") & " account for this member is set to disabled. </li>"
	End If
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isEnabled = True			: If list(10,i) = 0 Then isEnabled = False
			hasSkills = True			: If list(15,i) = 0 Then hasSkills = False
			
			If Not isEnabled Then
				items = items & "<li class=""warning""><strong class=""warning"">A program is disabled!</strong> "
				items = items & "The " & html(list(1,i)) & " program for this member has been set to disabled. </li>"
			End If	
			If Not hasSkills Then
				items = items & "<li class=""warning""><strong class=""warning"">No skills!</strong> "
				items = items & "No skills have been set for the " & html(list(1,i)) & " program in this member's profile. </li>"				
			End If	
		Next
	End If
	
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul class=""other-stuff"">" & items & "</ul>"
	
	str = str & "<ul><li>" & lastLoginText & "</li></ul>"
		
	OtherStuffForMemberAccountSettingsToString = str
End Function

Function MemberAccountSettingsToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim programList		: programList = page.ThisMember.ProgramList()
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	pg.Action = SEND_MESSAGE
	str = str & "<ul><li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Email to this member</a></li>"
	pg.Action = SHOW_AVAILABLE_PROGRAMS
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Set the programs for this member</a></li>"
	str = str & "</ul></div>"
	
	str = str & m_appMessageText
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">Settings for " & html(page.ThisMember.NameFirst & " " & page.ThisMember.NameLast) & "</h3>"
	
	str = str & "<h5 class=""settings"">Account settings</h5>"
	str = str & "<p>These are the account wide settings for this member. </p>"
	str = str & MemberAccountGridToString(page)

	str = str & "<h5 class=""program"">Program settings</h5>"
	str = str & MemberProgramGridToString(page, programList)
	
	str = str & OtherStuffForMemberAccountSettingsToString(page, programList)
	
	str = str & "</div>"
	
	MemberAccountSettingsToString = str
End Function

Function MemberAvailabilityGridToString(page)
	Dim str, msg, i
	Dim pg						: Set pg = page.Clone()
	Dim dateTime				: Set dateTime = New cFormatDate
	Dim eventAvailability		: Set eventAvailability = New cEventAvailability
	eventAvailability.MemberID = page.ThisMember.MemberID
	
	Dim list					: list = eventAvailability.AvailabilityList(page.Program.ProgramID, "", "")
	
	Dim count					: count = 0
	Dim altClass				: altClass = ""
	Dim checked					: checked = ""
	Dim isViewed				: isViewed = ""
	Dim eventHref				: eventHref = ""
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>If you make any changes to this page, be sure to click <strong>save</strong>. "
	str = str & "</p></div>"
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>The <strong>viewed</strong> column in the grid tells you if your member has logged in and saved their availability info for that event. "
	str = str & "</p></div>"
	
	' 0-EventAvailabilityID 1-MemberID 2-AvailabilityNote 3-IsAvailable 4-IsViewedByMember 5-DateCreated
	' 6-DateModified 7-EventID 8-EventName 9-EventDate 10-TimeStart 11-TimeEnd 13-ScheduleID 14-ScheduleName
	' 15-ScheduleIsVisible 16-ProgramID 17-ProgramName 17-ProgramIsEnabled
	
	str = str & m_appMessageText
	If IsArray(list) Then
		str = str & "<div class=""grid"">"
		str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-availability"">"
		str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
		str = str & "<input type=""checkbox"" id=""master"" /></th>"
		str = str & "<th scope=""col"">" & html(page.Program.ProgramName) & " Availability</th>"
		str = str & "<th scope=""col"">Schedule</th><th scope=""col"">Viewed</th><th scope=""col""></th></tr>"
		For i = 0 To UBound(list,2)
			count = count + 1
			altClass = ""
			If count Mod 2 = 0 Then altClass = " class=""alt"""
			
			checked = ""
			If list(3,i) = 1 Then checked = " checked=""checked"""
			isViewed = "Yes"
			If list(4,i) = 0 Then isViewed = "<span style=""color:red;"">No</span>"
			
			str = str & "<tr" & altClass & ">"
			str = str & "<td><input name=""IsAvailable" & list(0,i) & """" & checked & " type=""checkbox"" class=""checkbox is-available-checkbox"" /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/clock.png"" alt="""" />"
			str = str & "<div class=""data"">"
			pg.Action = SHOW_EVENT_DETAILS: pg.EventID = list(7,i)
			eventHref = "/admin/events.asp" & pg.UrlParamsToString(True)
			str = str & "<strong>" & html(list(17,i)) & " | <a href=""" & eventHref & """ title=""Details"">" & html(list(8,i)) & "</a></strong>"
			str = str & "<br />" & dateTime.Convert(list(9,i), "DDD MMMM dd, YYYY")
			If Len(list(10,i)) > 0 Then
				str = str & " at " & dateTime.Convert(list(10,i), "hh:nn pp")
			End If
			str = str & "</div></td>"
			str = str & "<td>" & html(list(14,i)) & "</td>"
			str = str & "<td>" & isViewed & "</td>"
			str = str & "<td class=""toolbar"">"
			str = str & "<a href=""" & eventHref & """ title=""Details""><img src=""/_images/icons/magnifier.png"" alt=""icon"" /></a>"
			str = str & "</td></tr>"
		Next
		str = str & "</table>"
		str = str & "<p style=""text-align:right;"">"
		str = str & "<input type=""submit"" name=""Submit"" value=""Save"" />"
		str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a></p>"
		str = str & "<input type=""hidden"" name=""FormAvailabilityIsPostback"" value=""" & IS_POSTBACK & """ />"
		str = str & "</form></div>"
	End If
	
	If count = 0 Then
		msg = msg & "No enabled, non-hidden events were returned for the program <strong>" & html(page.Program.ProgramName) & "</strong>. "
		msg = msg & "You will need to enable or show at least one event before you can set the availability for any of this program's members. "
		str = CustomApplicationMessageToString("No events were returned!", msg, "Error") 
	End If
	
	MemberAvailabilityGridToString = str
End Function

Function MemberSkillGridToString(page)
	Dim str, msg, i
	Dim pg				: Set pg = page.Clone()
	Dim skill			: Set skill = New cSkill
	skill.ProgramID = page.Program.ProgramID
	
	Dim list			: list = skill.SkillList("")
	Dim memberSkillList	: memberSkillList = page.ProgramMember.GetSkillList("")
	Dim count			: count = 0
	Dim altClass		: altClass = ""
	Dim isChecked		: isChecked = ""
	Dim isEnabled		: isEnabled = ""
	Dim checkboxEnabled : checkboxEnabled = ""
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>Be sure to click <strong>Save</strong> in the toolbar if you make any changes to this list. </p></div>"
	
	' 0-SkillID 1-SkillName 2-SkillDesc 3-SkillIsEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-GroupIsEnabled 8-DateGroupModified 9-DateGroupCreated 10-DateSkillModified
	' 11-DateSkillCreated 12-MemberCount

	str = str & m_appMessageText
	If IsArray(list) Then
		str = str & "<div class=""grid"">"
		str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-skill"">"
		str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
		str = str & "<input type=""checkbox"" id=""master"" /></th>"
		str = str & "<th scope=""col"">" & html(page.Program.ProgramName) & " Skills</th><th scope=""col"">Enabled</th>"
		str = str & "<th scope=""col"">Group</th><th scope=""col""></th></tr>"
		For i = 0 To UBound(list,2)
			count = count + 1
			altClass = ""
			If i Mod 2 <> 0 Then altClass = " class=""alt"""
			isEnabled = "<span style=""color:red;"">No</span>"
			If (list(3,i) = 1) And (list(7,i) = 1) Then isEnabled = "Yes"
			isChecked = ""
			If IsMemberSkill(list(0,i), memberSkillList) Then isChecked = " checked=""checked"""
			
			str = str & "<tr" & altClass & "><td><input type=""checkbox"" name=""SkillIDList"" value=""" & list(0,i) & """" & isChecked & checkboxEnabled & " /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/plugin.png"" alt=""icon"" />"
			str = str & "<strong>" & html(page.Program.ProgramName) & " | "
			pg.SkillID = list(0,i): pg.Action = SHOW_SKILL_DETAILS 
			str = str & "<a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>" & html(list(1,i)) & "</a></strong></td>"
			str = str & "<td>" & isEnabled & "</td>"
			str = str & "<td>" & html(list(5,i)) & "</td>"
			str = str & "<td class=""toolbar"">"
			pg.ProgramID = list(0,i): pg.Action = SHOW_SKILL_DETAILS
			str = str & "<a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """><img src=""/_images/icons/magnifier.png"" alt=""icon"" /></a>"
			str = str & "</td></tr>"
		Next
		str = str & "</table>"
		str = str & "<p style=""text-align:right;"">"
		str = str & "<input type=""submit"" name=""Submit"" value=""Save"" />"
		pg.Action = "": pg.ProgramMemberID = "": pg.SkillID = ""
		str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
		str = str & "<input type=""hidden"" name=""FormSkillIsPostback"" value=""" & IS_POSTBACK & """ />"
		str = str & "</p></form></div>"
	End If
	
	If count = 0 Then
		' no enabled skills returned
		msg = msg & "No enabled skills were returned for the program <strong>" & html(page.Program.ProgramName) & "</strong>. "
		msg = msg & "You will need add or enable at least one skill for this program before you can set the skills for your program members. "
		str = CustomApplicationMessageToString("No skills were returned.", msg, "Error")
	End If
	
	MemberSkillGridToString = str
End Function

Function IsMemberSkill(skillID, memberSkillList)
	Dim i
	IsMemberSkill = False
	
	' 0-SkillID 6-IsProgramMemberSkill

	If Not IsArray(memberSkillList) Then Exit Function
	For i = 0 To UBound(memberSkillList,2)
		If CStr(skillID) = CStr(memberSkillList(0,i)) Then
			If memberSkillList(6,i) = 1 Then
				IsMemberSkill = True
				Exit For
			End If
		End If
	Next
End Function

Function MemberProgramGridToString(page, list)
	Dim str, i
	Dim pg			: Set pg = page.Clone()
	
	Dim msg
	Dim count				: count = 0
	Dim altClass			: altClass = ""
	Dim enabledText			: enabledText = ""
	Dim enabledButtonIcon	: enabledButtonIcon = ""
	Dim enabledButtonTip	: enabledButtonTip = ""
	
	Dim member				: Set member = New cMember
	member.MemberId = page.Member.MemberId
	
	Dim programList			: programList = member.OwnedProgramsList()
	
	' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-ProgramMemberID 4-SkillListXmlFrag 5-EnrollStatusID 6-EnrollStatusText
	' 7-EnrollStatusDesc 8-IsLeader 9-IsLeaderText 10-IsActive 11-IsActiveText 12-LastScheduleView 
	' 13-DateModified 14-DateCreated 15-SkillCount 16-IsMissingAvailability 17-IsAdmin 18-ProgramIsEnabled
		
	If IsArray(list) Then
		str = str & "<p>These are the settings for each of this member's programs. </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th scope=""col"">Program</th><th scope=""col"">Enabled</th><th scope=""col"">&nbsp;</th></tr></thead>"
		For i = 0 To UBound(list,2)
			count = count + 1
			altClass = ""
			If count Mod 2 = 0 Then altClass = " class=""alt"""
			
			enabledText = "Yes"
			enabledButtonIcon = "lightning_add.png"
			enabledButtonTip = "Disable"
			If list(10,i) = 0 Then 
				enabledText = "<span style=""color:red;"">No</span>"
				enabledButtonIcon = "lightning_delete.png"
				enabledButtonTip = "Enable"
			End If
			
			str = str & "<tbody><tr" & altClass & "><td><img class=""icon"" src=""/_images/icons/user.png"" alt="""" />"
			str = str & "<strong>" & html(page.ThisMember.NameLast & ", " & page.ThisMember.NameFirst) & " | "
			str = str & html(list(1,i)) & "</strong>"
			str = str & "<td>" & enabledText & "</td>"
			str = str & "<td class=""toolbar"">"
			
			' check for admin or leader permission to change program ..
			If IsOwnedProgram(list(0,i), programList) Then
				pg.Action = SHOW_PROGRAM_MEMBER_DETAILS: pg.ProgramMemberID = list(3,i): pg.ProgramID = list(0,i)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details"">"
				str = str & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a>"
				pg.Action = CONFIGURE_PROGRAM_MEMBER_SKILLS: pg.ProgramID = list(0,i): pg.ProgramMemberID = list(3,i)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Skills"">"
				str = str & "<img src=""/_images/icons/plugin.png"" alt="""" /></a>"
				pg.Action = UPDATE_MEMBER_AVAILABILITY: pg.ProgramID = list(0,i): pg.ProgramMemberID = list(3,i)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Availability"">"
				str = str & "<img src=""/_images/icons/clock.png"" alt="""" /></a>"
				pg.Action = TOGGLE_PROGRAM_ACTIVE_STATUS: pg.ProgramID = list(0,i): pg.ProgramMemberID = list(3,i)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""" & enabledButtonTip & """>"
				str = str & "<img src=""/_images/icons/" & enabledButtonIcon & """ alt="""" /></a>"
				pg.Action = REMOVE_PROGRAM_MEMBER: pg.ProgramID = list(0,i): pg.ProgramMemberID = list(3,i)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove"">"
				str = str & "<img src=""/_images/icons/cross.png"" alt="""" /></a>"
			End If
			
			str = str & "</td></tr>"
		Next
		str = str & "</tbody></table></div>"
	End If
	
	If count = 0 then 
		str = "<p class=""alert"">This member does not belong to any programs. </p>"
	End If
	
	MemberProgramGridToString = str
End Function

Function IsOwnedProgram(programId, programList)
	Dim i
	
	IsOwnedProgram = False
	If Not IsArray(programList) Then Exit Function
	
	' 0-ProgramId 1-ProgramName 2-IsEnabled 3-ScheduleCount 4-EventCount
	
	For i = 0 To UBound(programList,2)
		If CStr(programList(0,i)) = CStr(programId) Then
			IsOwnedProgram = True
			Exit For
		End If
	Next
End Function

Function MemberAccountGridToString(page)
	Dim str
	Dim pg						: Set pg = page.Clone()
	
	
	Dim enabledText				: enabledText = "Yes"
	Dim enabledButtonIcon		: enabledButtonIcon = "lightning_add.png"
	Dim enabledButtonTipText	: enabledButtonTipText = "Disable"
	 
	If page.ThisMember.ActiveStatus = 0 Then 
		enabledText = "<span style=""color:red;"">No</span>"
		enabledButtonIcon = "lightning_delete.png"
		enabledButtonTipText = "Enable"
	End If
	
	str = str & "<div class=""grid"">"
	str = str & "<table><thead><tr><th scope=""col"">Member</th><th scope=""col"">Enabled</th><th scope=""col"">&nbsp;</th></tr></thead>"
	str = str & "<tbody><tr><td><img class=""icon"" src=""/_images/icons/user_red.png"" alt="""" />"
	str = str & "<strong>" & html(page.Client.NameClient) & " | "
	pg.Action = SHOW_MEMBER_DETAILS: pg.MemberID = page.ThisMember.MemberID: pg.ProgramID = ""
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.ThisMember.NameLast & ", " & page.ThisMember.NameFirst) & "</a></strong>"
	str = str & "<td>" & enabledText & "</td>"
	str = str & "<td class=""toolbar"">"
	pg.Action = SHOW_MEMBER_DETAILS: pg.MemberID = page.ThisMember.MemberID: pg.ProgramID = ""
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details"">"
	str = str & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a>"
	pg.Action = UPDATE_RECORD: pg.ProgramID = page.Program.ProgramID
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Edit"">"
	str = str & "<img src=""/_images/icons/pencil.png"" alt="""" /></a>"
	pg.Action = UPDATE_PASSWORD
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Password"">"
	str = str & "<img src=""/_images/icons/key.png"" alt="""" /></a>"
	pg.Action = IMPERSONATE_MEMBER
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Impersonate"">"
	str = str & "<img src=""/_images/icons/status_online.png"" alt="""" /></a>"
	pg.Action = SHOW_AVAILABLE_PROGRAMS
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Add Program"">"
	str = str & "<img src=""/_images/icons/script_add.png"" alt="""" /></a>"
	pg.Action = SEND_MEMBER_CREDENTIALS
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Send Login"">"
	str = str & "<img src=""/_images/icons/email_key.png"" alt="""" /></a>"
	pg.Action = SEND_MESSAGE
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email"">"
	str = str & "<img src=""/_images/icons/email.png"" alt="""" /></a>"
	pg.Action = TOGGLE_CLIENT_ACTIVE_STATUS
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""" & enabledButtonTipText & """>"
	str = str & "<img src=""/_images/icons/" & enabledButtonIcon & """ alt="""" /></a>"
	pg.Action = REMOVE_CLIENT_MEMBER
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove"">"
	str = str & "<img src=""/_images/icons/cross.png"" alt="""" /></a>"
	str = str & "</td></tr></tbody></table></div>"
	
	MemberAccountGridToString = str
End Function

Function FormSelectProgramsToString(page)
	Dim str, i, j
	Dim pg								: Set pg = page.Clone()
	Dim clientProgramList				: clientProgramList = page.Client.ProgramList("")
	
	Dim ownedProgramList				: ownedProgramList = page.Member.OwnedProgramsList()

	Dim memberProgramList				: memberProgramList = page.ThisMember.ProgramList()
	' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-ProgramMemberID 4-SkillListXmlFrag 5-EnrollStatusID 6-EnrollStatusText
	' 7-EnrollStatusDesc 8-IsLeader 9-IsLeaderText 10-IsActive 11-IsActiveText 12-LastScheduleView 
	' 13-DateModified 14-DateCreated 15-SkillCount 16-IsMissingAvailability 17-IsAdmin 18-ProgramIsEnabled

	Dim count							: count = 0
	Dim checked							: checked = ""
	Dim disabled						: disabled = ""
	Dim altClass						: altClass = ""
	Dim msg								: msg = ""
	Dim cssToHideRow					: cssToHideRow = ""
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>Be sure to click <strong>save</strong> if you make any changes to this member's program list. </p></div>"
	str = str & m_appMessageText
	str = str & "<div class=""grid"">"
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-select-program"">"
	str = str & "<table><tr class=""header""><th scope=""col"" style=""width:1%;""><input type=""checkbox"" id=""master"" /></th>"
	str = str & "<th>" & HTML(page.Client.NameClient & " Programs") & "</th><th scope=""col"">&nbsp;</th></tr>"
	If IsArray(clientProgramList) Then
		For i = 0 To UBound(clientProgramList,2)
			count = count + 1
			altClass = ""
			If count Mod 2 = 0 Then altClass = " class=""alt"""
			
			checked = ""
			If IsArray(memberProgramList) Then
				For j = 0 To UBound(memberProgramList,2)
					If CStr(clientProgramList(0,i)) = CStr(memberProgramList(0,j)) Then
						checked = " checked=""checked"""
						Exit For
					End If
				Next
			End If
			
			' hack: hide the row so leaders can't check or clear programs they don't own ..
			
			cssToHideRow = ""
			If Not IsOwnedProgram(clientProgramList(0,i), ownedProgramList) Then 
				cssToHideRow = " style=""display:none;"""
			End If
			
			str = str & "<tr" & altClass & cssToHideRow & "><td><input type=""checkbox"" name=""ProgramIDList"" value=""" & clientProgramList(0,i) & """" & checked & " class=""checkbox"" /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/script.png"" />"
			str = str & "<strong>" & html(page.Client.NameClient) & " | "
			pg.Action = SHOW_PROGRAM_DETAILS: pg.ProgramID = clientProgramList(0,i): page.MemberID = ""
			str = str & "<a href=""/admin/programs.asp" & pg.UrlParamsToString(True) & """>" & html(clientProgramList(1,i)) & "</a></strong>"
			str = str & "<td class=""toolbar"">"
			pg.Action = SHOW_PROGRAM_DETAILS: pg.ProgramID = clientProgramList(0,i): page.MemberID = ""
			str = str & "<a href=""/admin/programs.asp" & pg.UrlParamsToString(True) & """ title=""Details"">"
			str = str & "<img src=""/_images/icons/magnifier.png"" alt=""Add Program"" /></a>"
			str = str & "</td></tr>"
		Next
	End If
	str = str & "</table>"
	str = str & "<p style=""text-align:right;"">"
	str = str & "<input type=""submit"" name=""Submit"" value=""Save"" />"
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormSelectProgramsIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</p></form></div>"	
	
	FormSelectProgramsToString = str
End Function

Function FormConfirmDeleteProgramMemberToString(page)
	Dim str, msg
	Dim pg				: Set pg = page.Clone()
	
	msg = msg & "You will remove the member <strong>" & html(page.ThisMember.NameFirst & " " & page.ThisMember.NameLast) & "</strong> "
	msg = msg & "from the program <strong>" & html(page.Program.ProgramName) & "</strong>. "
	msg = msg & "You will lose any calendar or schedule information associated with this member. "
	msg = msg & "This action cannot be reversed. "
	str = str & CustomApplicationMessageToString("Please confirm remove member!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formRemoveProgramMember"">"
	str = str & "<input type=""hidden"" name=""FormConfirmDeleteProgramMemberIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = "": pg.ProgramMemberID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a></p></form>"
	
	FormConfirmDeleteProgramMemberToString = str
End Function

Function FormConfirmSelectProgramsToString(page)
	Dim str
	Dim msg
	Dim pg				: Set pg = page.Clone()
	
	msg = msg & "You are changing the programs for the member <strong>" & html(page.ThisMember.NameFirst & " " & page.ThisMember.NameLast) & "</strong>. "
	msg = msg & "If you are removing programs from this member's profile, you will lose any program, schedule, or calendar information associated with the program or programs you are removing. "
	msg = msg & "This action cannot be reversed. "
	str = CustomApplicationMessageToString("Please confirm this action!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" name=""formConfirmSelectPrograms"">"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""ProgramIDList"" value=""" & page.ProgramIDList & """ />"
	str = str & "<input type=""hidden"" name=""FormConfirmSelectProgramsIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</p></form>"
	
	FormConfirmSelectProgramsToString = str
End Function

Function FormConfirmSendCredentialsToString(page)
	Dim str
	Dim msg
	Dim pg					: Set pg = page.Clone()
	
	msg = msg & "You are about to send account member <strong>" & html(page.ThisMember.NameFirst & " " & page.ThisMember.NameLast) & "</strong> their " & Application.Value("APPLICATION_NAME") & " username and password via the email address associated with their account. "
	str = CustomApplicationMessageToString("Please confirm send credentials!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" name=""formConfirmSendCredentials"">"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Send"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormConfirmSendCredentialsIsPostback"" value=""" & IS_POSTBACK & """ /></p></form>"
	
	FormConfirmSendCredentialsToString = str
End Function

Function FormConfirmConfigureSkillsToString(page)
	Dim str
	Dim msg
	Dim pg					: Set pg = page.Clone()
	
	msg = msg & "You are changing the <strong>" & html(page.Program.ProgramName) & "</strong> skills for your member <strong>" & html(page.ThisMember.NameFirst & " " & page.ThisMember.NameLast) & "</strong>. "
	msg = msg & "If you are removing skills from this member's profile, you will lose any schedule or calendar information associated with those skills for this member. "
	msg = msg & "This action cannot be reversed. "
	str = CustomApplicationMessageToString("Please confirm this action!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" name=""formConfirmSendCredentials"">"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Save"" />"
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""SkillIDList"" value=""" & page.SkillIDList & """ />"
	str = str & "<input type=""hidden"" name=""FormConfirmConfigureSkillsIsPostback"" value=""" & IS_POSTBACK & """ /></p></form>"
	
	FormConfirmConfigureSkillsToString = str
End Function

Function FormConfirmImpersonateToString(page, member)
	Dim str
	Dim msg
	Dim pg			: Set pg = page.Clone()
	
	msg = msg & "You are about to logout of your own account and login as the member <strong>" & HTML(member.NameFirst & " " & member.NameLast) & "</strong>. "
	msg = msg & "To return to your own account, you will be required to logout of the impersonated account and re-login again to your own account. "
	str = str & CustomApplicationMessageToString("Please confirm impersonate member!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" name=""formConfirmDeleteMember"">"
	str = str & "<p><input type=""submit"" name=""submit"" value=""Impersonate"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel<a/>"
	str = str & "<input type=""hidden"" name=""FormConfirmImpersonateIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</p></form>"
	
	FormConfirmImpersonateToString = str
End Function

Function GetProgramsOwned(memberID) 'returns array
	Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs			: Set rs = Server.CreateObject("ADODB.Recordset")
	
	cnn.Open Application.Value("CNN_STR")
	cnn.up_memberGetProgramsOwnedByMemberID CLng(memberID), rs
	If Not rs.EOF then GetProgramsOwned = rs.GetRows()

	rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim memberLink
	pg.Action = SHOW_MEMBER_DETAILS
	memberLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.ThisMember.NameFirst & " " & page.ThisMember.NameLast) & "</a> / "
	
	Dim memberSettingsLink
	pg.Action = ""
	memberSettingsLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Settings</a> / "

	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	str = str & "<a href=""/admin/members.asp"">Members</a> / "
	If Len(page.Program.ProgramID) > 0 Then
		str = str & "<a href=""/admin/members.asp" & pg.UrlParamsToString(True) & """>" & html(page.Program.ProgramName) & "</a> / "
	End If
	Select Case page.Action
		Case SHOW_PROGRAM_MEMBER_DETAILS
			str = str & memberLink
			str = str & "Program Details"
		Case SHOW_MEMBER_DETAILS
			str = str & html(page.ThisMember.NameFirst & " " & page.ThisMember.NameLast)
		Case UPDATE_MEMBER_AVAILABILITY
			str = str & memberLink & memberSettingsLink
			str = str & "Availability"
		Case CONFIGURE_PROGRAM_MEMBER_SKILLS
			str = str & memberLink & memberSettingsLink
			str = str & "Skills"
		Case REMOVE_CLIENT_MEMBER
			str = str & memberLink & memberSettingsLink
			str = str & "Remove Account"
		Case SEND_MEMBER_CREDENTIALS
			str = str & memberLink & memberSettingsLink
			str = str & "Email Login"
		Case SHOW_AVAILABLE_PROGRAMS
			str = str & memberLink & memberSettingsLink
			str = str & "Select Programs"
		Case UPDATE_PASSWORD
			str = str & memberLink & memberSettingsLink
			str = str & "Change Password"
		Case UPDATE_RECORD
			str = str & memberLink & memberSettingsLink
			str = str & "Edit Profile"
		Case Else
			str = str & memberLink & "Settings"
	End Select
	
	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href			: href = ""
	
	Dim saveAvailabilityButton
	href = "#"
	saveAvailabilityButton = "<li id=""save-availability-button""><a href=""" & href & """><img class=""icon"" src=""/_images/icons/disk.png"" /></a><a href=""" & href & """ title=""Save"">Save</a></li>"
		
	Dim saveProgramsButton
	href = "#"
	saveProgramsButton = "<li id=""save-program-button""><a href=""" & href & """><img class=""icon"" src=""/_images/icons/disk.png"" /></a><a href=""" & href & """ title=""Save"">Save</a></li>"
		
	Dim saveSkillsButton
	href = "#"
	saveSkillsButton = "<li id=""save-skills-button""><a href=""" & href & """><img class=""icon"" src=""/_images/icons/disk.png"" /></a><a href=""" & href & """ title=""Save"">Save</a></li>"
		
	Dim memberProfileButton
	pg.Action = "": pg.ProgramMemberId = ""
	href = pg.Url & pg.UrlParamsToString(True)
	memberProfileButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/pencil.png"" /></a><a href=""" & href & """>Settings</a></li>"
	
	Dim memberListButton
	Dim userIcon			: userIcon = "user_red.png"
	If Len(page.ProgramId) > 0 Then userIcon = "user.png"
	pg.Action = "": pg.MemberID = "": pg.ProgramMemberId = ""
	href = "/admin/members.asp" & pg.UrlParamsToString(True)
	memberListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/" & userIcon & """ /></a><a href=""" & href & """>Member List</a></li>"
	
	Select Case page.Action
		Case SHOW_PROGRAM_MEMBER_DETAILS
			str = str & memberProfileButton
		Case SHOW_MEMBER_DETAILS
			str = str & GotoMemberDropdownToString(page)
			str = str & memberProfileButton & memberListButton
		Case UPDATE_MEMBER_AVAILABILITY
			str = str & memberProfileButton
			str = str & saveAvailabilityButton
		Case CONFIGURE_PROGRAM_MEMBER_SKILLS
			str = str & memberProfileButton
			str = str & saveSkillsButton
		Case REMOVE_CLIENT_MEMBER
			str = str & memberProfileButton	
		Case SEND_MEMBER_CREDENTIALS
			str = str & memberProfileButton
		Case SHOW_AVAILABLE_PROGRAMS
			str = str & memberProfileButton
			str = str & saveProgramsButton
		Case UPDATE_PASSWORD
			str = str & memberProfileButton
		Case UPDATE_RECORD
			str = str & memberProfileButton
		Case Else
			str = str & GotoMemberDropdownToString(page)
			str = str & memberListButton
	End Select 
		
	m_tabLinkBarText = str
End Sub

Function GotoMemberDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = page.Client.MemberList(page.Program.ProgramID, "")
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formGotoMember"">"
	str = str & "<input type=""hidden"" name=""FormGotoMemberDropdownIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""NewMemberID"" onchange=""document.forms.formGotoMember.submit();"">"
	str = str & "<option value="""">" & html("< Go to member >") & "</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			str = str & "<option value=""" & list(0,i) & """>" & html(list(1,i) & ", " & list(2,i)) & "</option>"
		Next
	End If
	str = str & "</select></form></li>"
	
	GotoMemberDropdownToString = str
End Function
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_setting.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_select_option.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_availability_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/state_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_FormMemberToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_FormPasswordToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_DoDeleteClientMember.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_OwnsMember.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_GetListFromXmlFragment.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public MemberID
	Public ProgramID
	Public ProgramMemberID
	Public EmailID
	Public SkillID
	Public EventID
	
	' not persisted
	Public ProgramIDList
	Public SkillIDList
	
	' objects
	Public Member
	Public Client
	Public ThisMember
	Public Program
	Public ProgramMember
	Public Settings			' as array
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(MemberID) > 0 Then str = str & "mid=" & Encrypt(MemberID) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(ProgramMemberID) > 0 Then str = str & "pmid=" & Encrypt(ProgramMemberID) & amp
		If Len(EmailID) > 0 Then str = str & "emid=" & Encrypt(EmailID) & amp
		If Len(SkillID) > 0 Then str = str & "skid=" & Encrypt(SkillID) & amp
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
		c.Action = Action
		c.ProgramID = ProgramID
		c.MemberID = MemberID
		c.ProgramMemberID = ProgramMemberID
		c.EmailID = EmailID
		c.SkillID = SkillID
		c.EventID = EventID
		
		c.ProgramIDList = ProgramIDList
		c.SkillIDList = SkillIDList
		c.Settings = Settings
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.ThisMember = ThisMember
		Set c.Program = Program
		Set c.ProgramMember = ProgramMember
		
		Set Clone = c
	End Function
End Class
%>

