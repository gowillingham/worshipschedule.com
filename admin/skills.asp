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
	page.SortBy = Request.QueryString("sb")
	page.Action = Decrypt(Request.QueryString("act"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.SkillID = Decrypt(Request.QueryString("skid"))
	page.MemberID = Decrypt(Request.QueryString("mid"))
	
	If Request.Form("FormSkillDropdownIsPostback") = IS_POSTBACK Then
		If Len(Request.Form("NewSkillID")) > 0 Then
			page.SkillID = Request.Form("NewSkillID")
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
	If Len(page.Program.ProgramID) > 0 Then page.Program.Load()
	Set page.Skill = New cSkill
	page.Skill.SkillID = page.SkillID
	If Len(page.Skill.SkillID) > 0 Then page.Skill.Load()
	
	If Request.Form("FormSortByDropdownIsPostback") = IS_POSTBACK Then
		page.SortBy = Request.Form("SortBy")
	End If
	
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
		<link rel="stylesheet" type="text/css" href="skills.css" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="skills.js"></script>
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
		Case SHOW_SKILL_DETAILS
			str = str & SkillSummaryToString(page)

		Case DELETE_RECORD
			If Request.Form("FormConfirmDeleteSkillIsPostback") = IS_POSTBACK Then
				Call DoDeleteSkill(page.Skill, rv)
				Select Case rv
					Case 0
						page.MessageID = 4004
					Case Else
						page.MessageID = 4004
				End Select 
				page.Action = "": page.SkillID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmDeleteSkillToString(page)
			End If
			
		Case ADDNEW_RECORD
			If Request.Form("FormSkillIsPostback") = IS_POSTBACK Then
				Call LoadSkillFromRequest(page.Skill)
				If ValidSkill(page.Skill) Then 
					Call DoInsertSkill(page, rv)
					Select Case rv
						Case 0
							page.MessageID = 4000
						Case -2
							' dupe skill name
							page.MessageID = 4003
						Case Else
							page.MessageID = 4001
					End Select
					page.SkillID = "": page.Action = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormSkillToString(page)
				End If			
			Else
				str = str & FormSkillToString(page)
			End If
			
		Case UPDATE_RECORD
			If Request.Form("FormSkillIsPostback") = IS_POSTBACK Then
				Call LoadSkillFromRequest(page.Skill)
				If ValidSkill(page.Skill) Then 
					Call DoUpdateSkill(page.Skill, rv)
					Select Case rv
						Case 0
							page.MessageID = 4000
						Case -2
							' dupe skill name
							page.MessageID = 4003
						Case Else
							page.MessageID = 4001
					End Select
					page.SkillID = "": page.Action = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormSkillToString(page)
				End If			
			Else
				str = str & FormSkillToString(page)
			End If
			
		Case ASSIGN_SKILLS_TO_MEMBERS
			' test for no skills ..
			If page.Program.HasSkills = 0 Then
				page.MessageID = 4030: page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			' test for no members ..
			If page.Program.MemberCount = 0 Then
				page.MessageID = 4029: page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If

			If Request.Form("Submit") = ">>" Then
				str = str & FormConfirmDeleteProgramMemberSkillToString(page, Request.Form("ProgramMemberSkillID"))
			ElseIf Request.Form("Submit") = "<<" Then
				Call DoInsertMemberSkills(page, Request.Form("ProgramMemberID"), rv)
				Select Case rv
					Case 0
						page.MessageID = 4026
						Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False))
					Case - 1
						str = str & MemberSkillGridToString(page)
					Case Else
						page.MessageID = 4028
						Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False))
				End Select
			ElseIf Request.Form("FormConfirmDeleteProgramMemberSkillIsPostback") = IS_POSTBACK Then
				Call DoDeleteProgramMemberSkills(page, Request.Form("ProgramMemberSkillID"), rv)
				Select Case rv
					Case 0
						page.MessageID = 4027
						Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False))
					Case - 1
						str = str & MemberSkillGridToString(page)
					Case Else
						page.MessageID = 4028
						Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False))
				End Select
			Else
				str = str & MemberSkillGridToString(page)
			End If
			
		Case SEND_MESSAGE
			Call DoInsertMessage(page, rv)
			page.Action = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
		
		Case Else
			str = str & SkillGridToString(page)
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub ClearTablinkBar()
	m_tabLinkBarText = "<li>&nbsp;</li>"
End Sub

Sub DoInsertMessage(page, outError)
	Dim email			: Set email = New cEmail
	
	email.MemberId = page.Member.MemberId
	email.ClientId = page.Member.ClientId
	email.GroupList = "skill|" & page.SkillId
	
	Call email.Add(outError)
	
	page.EmailId = email.EmailId	
End Sub

Sub DoDeleteProgramMemberSkills(page, idList, outError)
	Dim i
	Dim tempError				: tempError = 0
	Dim totalError				: totalError = 0
	outError = 0
	
	Dim programMemberSkill		
	If Len(idList) = 0 Then
		outError = -1
		Exit Sub
	End If
	Dim list					: list = Split(Replace(idList, " ", ""), ",")
	
	For i = 0 To UBound(list)
		If Len(list(i)) > 0 Then
			Set programMemberSkill = New cProgramMemberSkill
			programMemberSkill.ProgramMemberSkillID = list(i)
			Call programMemberSkill.Delete(tempError)
			totalError = totalError + tempError
			Set programMemberSkill = Nothing
		End If
	Next
	If totalError <> 0 Then outError = -2
End Sub

Sub DoInsertMemberSkills(page, idList, outError)
	Dim i
	Dim programMemberSkill
	Dim tempError				: tempError = 0
	Dim totalError				: totalError = 0
	outError = 0
	
	If Len(idList) = 0 Then
		outError = -1
		Exit Sub
	End If
	Dim list					: list = Split(Replace(idList, " ", ""), ",")
	
	If IsArray(list) Then
		For i = 0 To UBound(list)
			If Len(list(i)) > 0 Then
				Set programMemberSkill = New cProgramMemberSkill
				programMemberSkill.SkillID = page.SkillID
				programMemberSkill.ProgramMemberID = list(i)
				Call programMemberSkill.Add(tempError)
				totalError = totalError + tempError
				Set programMemberSkill = Nothing
			End If
		Next
	End If
	If totalError <> 0 Then outError = -2
	
End Sub

Sub DoUpdateSkill(skill, outError)
	Call skill.Save(outError)
End Sub

Sub DoDeleteSkill(skill, outError)
	Call skill.Delete(outError)
End Sub

Sub DoInsertSkill(page, outError)
	page.Skill.ProgramID = page.Program.ProgramID
	Call page.Skill.Add(outError)
End Sub

Function MemberSkillGridForSkillSummaryToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	Dim list			: list = page.Skill.MemberList()
	Dim count			: count = 0
	Dim rows			: rows = ""
	Dim href			: href = ""
	Dim alt				: alt = ""
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberSkillID 5-ProgramMemberID
	' 6-IsApproved 7-ProgramMemberIsActive 8-Email
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isMemberEnabled = True				: If list(3,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True		: If list(7,i) = 0 Then isProgramMemberEnabled = False
		
			If isMemberEnabled And isProgramMemberEnabled Then
				pg.Action = SHOW_PROGRAM_MEMBER_DETAILS: pg.MemberId = list(0,i): pg.ProgramMemberId = list(5,i): pg.SkillId = ""
				href = "/admin/profile.asp" & pg.UrlParamsToString(True)
				
				rows = rows & "<tr" & alt & "><td><img src=""/_images/icons/user.png"" class=""icon"" alt="""" />"
				rows = rows & "<strong>" & html(page.Skill.ProgramName) & "</strong> | "
				rows = rows & "<a href=""" & href & """><strong>" & html(list(1,i) & ", " & list(2,i)) & "</strong></a></td>"
				rows = rows & "<td class=""toolbar""><a href=""" & href & """>"
				rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
				
				count = count + 1
			End If
		Next	
	End If
	
	If count > 0 Then
		str = str & "<p>This is a list of " & html(page.Skill.ProgramName) & " members that have this skill assigned to their profile. </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Member</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = "<p class=""alert"">No program members have this skill in their profile (or you have disabled all the members that do). </p>"
	End If
	
	MemberSkillGridForSkillSummaryToString = str
End Function

Function SkillSummaryToString(page)
	Dim str
	Dim dateTime					: Set dateTime = New cFormatDate
	
	Dim isSkillEnabled				: isSkillEnabled = True
	If page.Skill.IsEnabled = 0 Then isSkillEnabled = False 
	Dim isSkillGroupEnabled			: isSkillGroupEnabled = True
	If page.Skill.SkillGroupIsEnabled = 0 Then isSkillGroupEnabled = False
	
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & html(page.Skill.SkillName) & "</h3>"
	
	If (Not isSkillEnabled) Or (Not isSkillGroupEnabled) Then
		str = str & "<h5 class=""disabled"">Skill is disabled</h5>"
		str = str & "<p class=""alert"">This skill is disabled or it belongs to a skill group that has been disabled. </p>"
	End If
	
	str = str & "<h5 class=""description"">Description</h5>"
	If Len(page.Skill.SkillDesc) > 0 Then 
		str = str & "<p>" & html(page.Skill.SkillDesc) & "</p>"
	Else
		str = str & "<p class=""alert"">No skill description available. </p>"
	End If
	str = str & "<h5 class=""program-member"">Member listing</h5>"
	str = str & MemberSkillGridForSkillSummaryToString(page)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul><li>Created on " & dateTime.Convert(page.Skill.DateCreated, "DDDD MMMM dd, YYYY around hh:00 pp") & ". </li>"
	If Len(page.Skill.GroupName) > 0 Then
		str = str & "<li>This skill belongs to the <strong>" & html(page.Skill.GroupName) & "</strong> skill grouping. </li>"
	Else	
		str = str & "<li>This skill has not been placed in a skill grouping. </li>"
	End If
	str = str & "</ul>"
	

	str = str & "</div>"
	
	SkillSummaryToString = str
End Function

Function NoSkillsDialogToString(page)
	Dim dialog				: Set dialog = New cDialog
	Dim pg					: Set pg = page.Clone()
	
	dialog.Headline = "Where are the skills ..?"
	
	dialog.Text = dialog.Text & "<p>The program you selected (" & html(page.Program.ProgramName) & ") does not have any skills set up yet. "
	dialog.Text = dialog.Text & "Before you can use this program for scheduling, it will need to have at least one skill. </p><p>"
	dialog.Text = dialog.Text & "Click <strong>Create the first skill</strong> to get started. </p>"

	dialog.SubText = dialog.SubText & "<p>When this is fixed, you will use this page to add or change the skills that the members of this program can have. "
	dialog.SubText = dialog.SubText & "You'll also set which members have which skills. </p>"
	dialog.SubText = dialog.SubText & "<p>" & Application.Value("APPLICATION_NAME") & " uses programs and skills to organize your account members. "
	dialog.SubText = dialog.SubText & "When you assign your members to your schedule's event teams, you'll schedule your members based on what they can do (their skills). "
	dialog.SubText = dialog.SubText & "</p>"

	pg.Action = ADDNEW_RECORD
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Create the first skill</a></li>"
	
	pg.Action = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/programs.asp" & pg.UrlParamsToString(True) & """>Back to my program list</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/help/topic.asp?hid=14#anchor-add-program"" target=""_blank"">Learn more about programs and skills</a></li>"

	NoSkillsDialogToString = dialog.ToString()	
End Function

Function SkillGridtoString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	page.Skill.ProgramID = page.Program.ProgramID
	Dim list			: list = page.Skill.SkillList(LookupSortParam(page.SortBy))
	
	Dim altClass		: altClass = ""
	Dim enabledText		: enabledText = ""
	Dim hasSkills		: hasSkills = False
	Dim msg
		
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	pg.Action = ADDNEW_RECORD
	str = str & "<ul><li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Create a new skill</a></li>"
	pg.Action = ASSIGN_SKILLS_TO_MEMBERS 
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Organize members by skill</a></li>"
	pg.Action = ""
	str = str & "<li><a href=""/admin/skillgroups.asp" & page.UrlParamsToString(True) & """>Organize these skills into groups</a></li>"
	str = str & "<li><a href=""/help/faq.asp?act=2&amp;fqid=96"" target=""_blank"">Learn more about skills and skill groups</a></li></ul></div>"

	' 0-SkillID 1-SkillName 2-SkillDesc 3-IsSkillEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-IsGroupEnabled 8-LastModified 9-DateCreated
		
	str = str & m_appMessageText
	str = str & "<div class=""grid"">"
	str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
	str = str & "<input type=""checkbox"" name=""master"" onclick="""" disabled=""disabled"" checked=""checked"" /></th>"
	str = str & "<th scope=""col"">Skills</th><th scope=""col"">Group</th><th scope=""col"">Enabled</th><th scope=""col"">&nbsp;</th></tr>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			hasSkills = True
			altClass = ""
			If i Mod 2 <> 0 Then altClass = " class=""alt"""
			enabledText = "Yes"
			If list(3,i) = 0 Then enabledText = "<span style=""color:red;"">No</span>"
			
			str = str & "<tr" & altClass & "><td><input type=""checkbox"" name=""Unused"" disabled=""disabled"" checked=""checked"" /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/plugin.png"" alt="""" />"
			str = str & "<strong>" & html(page.Program.ProgramName) & " | "
			pg.SkillID = list(0,i): pg.Action = SHOW_SKILL_DETAILS
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(1,i)) & "</a></strong></td>"
			str = str & "<td>" & list(5,i) & "</td>"
			str = str & "<td>" & enabledText & "</td>"
			str = str & "<td class=""toolbar"">"
			pg.SkillID = list(0,i)
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details""><img src=""/_images/icons/magnifier.png"" alt=""icon"" /></a>"
			pg.Action = UPDATE_RECORD
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Edit""><img src=""/_images/icons/pencil.png"" alt=""icon"" /></a>"
			pg.Action = ASSIGN_SKILLS_TO_MEMBERS
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Members""><img src=""/_images/icons/plugin_user.png"" alt=""icon"" /></a>"
			pg.SkillId = list(0,i): pg.Action = SEND_MESSAGE
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email""><img src=""/_images/icons/email.png"" alt=""icon"" /></a>"
			pg.Action = DELETE_RECORD
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove""><img src=""/_images/icons/cross.png"" alt=""icon"" /></a>"
			str = str & "</td></tr>"	
		Next
	End If
	str = str & "</table></div>"
	
	If Not hasSkills Then
		Call ClearTablinkBar()
		str = NoSkillsDialogToString(page)
	End If
	
	SkillGridToString = str
End Function

Function MemberSkillGridToString(page)
	Dim str
	Dim memberCount			: memberCount = 0
	Dim memberSkillCount	: memberSkillCount = 0
	
	page.Skill.SkillID = page.SkillID
	If Len(page.SkillID) > 0 Then page.Skill.Load()
	
	Dim list		: list = GetProgramMemberList(page.SkillID)
	
	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "<p>You can change the skill that is shown by selecting a different skill from the dropdown list. </p></div>"

	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "<p>Highlight the members in the desired listbox and use the buttons to move them. "
	str = str & "<br /><br />Use [CONTROL]-Click or [SHIFT]-Click to select multiple members at once. </p></div>"

	str = str & m_appMessageText
	str = str & "<div class=""grid"">"
	str = str & "<table><tr class=""header""><th scope=""col"">Assign Skills</th>"
	str = str & "<th scope=""col"" style=""text-align:right;"">Skill: " & SkillDropdownToString(page) & "</th></tr>"
	str = str & "<tr><td class=""two-way-selector"" colspan=""2"" style="""">"
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" name=""formProgramMemberSkill"">"

	str = str & "<table><tr><td class=""list-box"">"
	str = str & DeleteProgramMemberSkillListboxToString(list, page.Skill.SkillName, memberSkillCount) & "</td>"
	
	str = str & "<td class=""arrow-buttons"">"
	str = str & "<input type=""submit"" name=""Submit"" value=""" & HTML("<<") & """ />"
	str = str & "<br /><input type=""submit"" name=""Submit"" value=""" & HTML(">>") & """ /></td>"
	
	str = str & "<td class=""list-box"">" & AddProgramMemberSkillListboxToString(list, page.Program.ProgramName, memberCount)
	str = str & "</td></tr></table></form></td></tr></table></div>"
	
	MemberSkillGridToString = str
End Function

Function DeleteProgramMemberSkillListboxToString(list, skillName, count)
	Dim str, i
	count = 0
	
	str = str & "<h4 id=""skill-label"">" 
	If Len(skillName) > 0 Then
		str = str & html(skillName)
	Else
		str = str & "<strong style=""color:red;"">No Skill Selected!</strong>"
	End If
	str = str & "</h4>"
	str = str & "<select name=""ProgramMemberSkillID"" multiple=""multiple"" style="""">"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			If list(3,i) = 1 Then
				count = count + 1
				str = str & "<option value=""" & list(4,i) & """>" & HTML(list(1,i) & ", " & list(2,i)) & "</option>"
			End If
		Next
	End If
	str = str & "</select>"
	
	DeleteProgramMemberSkillListboxToString = str
End Function

Function AddProgramMemberSkillListboxToString(list, programName, count)
	Dim str, i
	count = 0
	
	str = str & "<h4 id=""program-label"">" & html(programName) & "</h4>"
	str = str & "<select name=""ProgramMemberID"" multiple=""multiple"" style="""">"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			If list(3,i) = 0 Then
				count = count + 1
				str = str & "<option value=""" & list(0,i) & """>" & HTML(list(1,i) & ", " & list(2,i)) & "</option>"
			End If
		Next
	End If
	str = str & "</select>"
	
	AddProgramMemberSkillListboxToString = str
End Function

Function FormConfirmDeleteSkillToString(page)
	Dim str, msg
	Dim pg				: Set pg = page.Clone()
	
	Dim skill			: Set skill = New cSkill
	skill.SkillID = page.SkillID
	skill.Load()
	
	msg = msg & "You will permanently delete the skill <strong>" & HTML(skill.SkillName) & "</strong> from the " & HTML(page.Program.ProgramName) & " program. "
	msg = msg & "Currently, " & skill.MemberCount & " " & HTML(page.Program.ProgramName) & " members have this skill in their profile. "
	msg = msg & "This action cannot be reversed (to preserve all existing program and calendar data, consider disabling the skill instead). "
	str = CustomApplicationMessageToString("Please confirm this action", msg, "Confirm")
	
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" name=""formConfirmDeleteSkill"">"
	str = str & "<table><tr><td>"
	str = str & "<input type=""submit"" name=""Submit"" value=""Remove Skill"" title=""Delete Skill"" />"
	pg.Action = "": pg.SkillId = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel<a/>"
	str = str & "<input type=""hidden"" name=""FormConfirmDeleteSkillIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</td></tr></table></form>"

	FormConfirmDeleteSkillToString = str	
End Function

Function FormConfirmDeleteProgramMemberSkillToString(page, idList)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim line1			: line1 = "Please confirm this action. "
	
	Dim count			: count = UBound(Split(idList, ",")) + 1
	Dim memberCount		: memberCount =  count & " member"
	Dim memberAdverb	: memberAdverb = "this member"
	If count <> 1 Then 
		memberCount = memberCount & "s"
		memberAdverb = "these members"
	End If
	page.Skill.Load()	
	
	str = str & "You are removing the skill <strong>" & HTML(page.Skill.SkillName) & "</strong> from the " & HTML(page.Program.ProgramName) & " profile of " & memberCount & ". "
	str = str & "You will lose any existing event and schedule information associated with this skill for " & memberAdverb & ". " 
	str = str & "This action cannot be reversed. "
	str = CustomApplicationMessageToString(line1, str, "Confirm")
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" name=""formConfirmDeleteProgramMemberSkill"">"
	str = str & "<input type=""hidden"" name=""FormConfirmDeleteProgramMemberSkillIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<input type=""hidden"" name=""ProgramMemberSkillID"" value=""" & idList & """ />"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = "": pg.SkillID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"	
	str = str & "</p></form>"
	
	FormConfirmDeleteProgramMemberSkillToString = str
End Function

Function FormSkillToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	
	str = str & "<form action=""" & pg.Url & page.UrlParamsToString(True) & """ method=""post"" name=""formSkill"">"
	str = str & "<table><tr><td class=""label"">" & RequiredElementToString(True, "Skill Name") & "</td>"
	str = str & "<td><input class=""large gets-focus"" type=""text"" name=""SkillName"" value=""" & HTML(page.Skill.SkillName) & """ /></td></tr>"
	str = str & "<tr><td class=""label"">Skill Description</td>"
	str = str & "<td><textarea class=""large"" name=""SkillDesc"">" & HTML(page.Skill.SkillDesc) & "</textarea></td></tr>"
	str = str & "<tr><td class=""label"">Skill Group</td>"
	str = str & "<td>" & SkillGroupDropdownToString(page) & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Enabled</td>"
	str = str & "<td>" & YesNoDropdownToString(page.Skill.IsEnabled, "IsEnabled") & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "When set to yes, this skill is available to members in their profiles <br />and is available for administrators to be scheduled. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Save"" />"
	pg.Action = "": pg.SkillID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormSkillIsPostback"" value=""" & IS_POSTBACK & """ /></td></tr>"
	str = str & "</table></form></div>"

	FormSkillToString = str
End Function

Function ValidSkill(skill)
	ValidSkill = True
	If Application.Value("FORM_VALIDATION_OFF") Then Exit Function
	
	If Not ValidData(skill.SkillName, True, 0, 100, "Skill Name", "") Then ValidSkill = False
	If Not ValidData(skill.SkillDesc, False, 0, 2000, "Skill Name", "") Then ValidSkill = False
End Function

Sub LoadSkillFromRequest(skill)
	skill.SkillName = Request.Form("SkillName")
	skill.SkillDesc = Request.Form("SkillDesc")
	skill.SkillGroupID = Request.Form("SkillGroupID")
	skill.IsEnabled = Request.Form("IsEnabled")
End Sub

Function SkillGroupDropdownToString(page)
	Dim str, i
	
	page.Skill.ProgramID = page.Program.ProgramID
	Dim list		: list = page.Skill.SkillGroupList()
	
	str = str & "<select name=""SkillGroupID"">"
	str = str & "<option value="""">Default</option>"
	str = str & SelectOption(list, page.Skill.SkillGroupID)
	str = str & "</select>"	
	
	SkillGroupDropdownToString = str
End Function

Function SkillDropdownToString(page)
	Dim str
	Dim list			: list = page.Program.SkillList("")		' returns enabled skill listing
	
	str = str & "<form style=""display:inline;"" action=""" & Request.ServerVariables("URL") & page.UrlParamsToString(True) & """ method=""post"" name=""formSkillDropdown"">"
	str = str & "<input type=""hidden"" name=""FormSkillDropdownIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""NewSkillID"" onchange=""document.formSkillDropdown.submit();"">"
	If Len(page.Skill.SkillID) = 0 Then
		str = str & "<option value="""">" & HTML("< Choose a skill .. >") & "</option>"
	End If
	str = str & SelectOption(list, page.SkillID)
	str = str & "</select></form>"
	
	SkillDropdownToString = str
End Function

Function GetProgramMemberList(skillID)
	' returns list of members with hasSkills flag
	Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs			: Set rs = Server.CreateObject("ADODB.Recordset")
	
	' 0-ProgramMemberID 1-NameLast 2-NameFirst 3-HasSkill 4-ProgramMmeberSkillID

	cnn.Open(Application.Value("CNN_STR"))
	cnn.up_programGetProgramMembersBySkillID CLng(skillID), rs
	If Not rs.EOF Then GetProgramMemberList = rs.GetRows()
	
	If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
	Set cnn = Nothing
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim skillLink
	pg.Action = SHOW_SKILL_DETAILS
	skillLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(pg.Skill.SkillName) & "</a> / "
	
	Dim programLink
	pg.SkillID = "": pg.Action = SHOW_PROGRAM_DETAILS
	programLink = "<a href=""/admin/programs.asp" & pg.UrlParamsToString(True) & """>" & html(pg.Program.ProgramName) & "</a> / "

	Dim skillsLink
	pg.SkillID = "": pg.Action = ""
	skillsLink = "<a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>Skills</a> / "
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case SHOW_SKILL_DETAILS	
			str = str & programLink
			str = str & skillsLink
			str = str & html(page.Skill.SkillName)
					
		Case DELETE_RECORD
			str = str & programLink
			str = str & skillsLink
			str = str & "Remove Skill"
					
		Case UPDATE_RECORD
			str = str & programLink
			str = str & skillsLink
			str = str & "Edit Skill"
					
		Case ADDNEW_RECORD
			str = str & programLink
			str = str & skillsLink
			str = str & "Add Skill"
			
		Case ASSIGN_SKILLS_TO_MEMBERS
			str = str & programLink
			str = str & skillsLink
			str = str & skillLink
			str = str & "Members"
		Case Else
			str = str & programLink
			str = str & "Skills"
	End Select	

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href		: href = ""
	
	Dim memberSkillsButton
	pg.Action = ASSIGN_SKILLS_TO_MEMBERS
	href = pg.Url & pg.UrlParamsToString(True)
	memberSkillsButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/plugin_user.png"" /></a><a href=""" & href & """>Member skills</a></li>"

	Dim newSkillButton
	pg.Action = ADDNEW_RECORD
	href = pg.Url & pg.UrlParamsToString(True)
	newSkillButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/plugin_add.png"" /></a><a href=""" & href & """>Add Skill</a></li>"
	
	Dim skillGroupButton
	pg.Action = "": pg.SkillID = ""
	href = "/admin/skillgroups.asp" & pg.UrlParamsToString(True)
	skillGroupButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/link.png"" /></a><a href=""" & href & """>Skill Groups</a></li>"
	
	Dim skillListButton
	pg.Action = "": pg.SkillID = ""
	href = "/admin/skills.asp" & pg.UrlParamsToString(True)
	skillListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/plugin.png"" /></a><a href=""" & href & """>Skill List</a></li>"
	
	Dim programListButton
	pg.Action = "": pg.SkillID = "": pg.ProgramID = ""
	href = "/admin/programs.asp" & pg.UrlParamsToString(True)
	programListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/script.png"" /></a><a href=""" & href & """>Program List</a></li>"
	
	Select Case page.Action
		Case UPDATE_RECORD
			str = str & skillListButton
			str = str & memberSkillsButton
			
		Case DELETE_RECORD
			str = str & skillListButton
			str = str & memberSkillsButton

		Case ADDNEW_RECORD
			str = str & skillListButton
			str = str & memberSkillsButton
			
		CASE SHOW_SKILL_DETAILS
			str = str & programListButton & skillListButton
			
		Case ASSIGN_SKILLS_TO_MEMBERS
			str = str & skillListButton
			
		Case Else
			str = str & SortByDropdownToString(page)
			str = str & newSkillButton 
			str = str & skillGroupButton
			str = str & memberSkillsButton
			str = str & programListButton
	End Select
	
	m_tabLinkBarText = str
End Sub

Function SortByDropdownToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formSortByDropdown"">"
	str = str & "<input type=""hidden"" name=""FormSortByDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""SortBy"" onchange=""document.forms.formSortByDropdown.submit();"">"
	str = str & "<option value="""">" & html("< Sort by .. >") & "</option>"
	str = str & "<option value=""" & SORT_BY_SKILLNAME & """>Skill</option>"
	str = str & "<option value=""" & SORT_BY_SKILLGROUP & """>Group</option>"
	str = str & "<option value=""" & SORT_BY_SKILL_IS_ENABLED & """>Enabled</option>"
	str = str & "</select></form></li>"	
	
	SortByDropdownToString = str
End Function

Function LookupSortParam(val)
	Dim str
	
	Select Case val
		Case SORT_BY_SKILLNAME
			str = "SkillName"
		Case SORT_BY_SKILLGROUP
			str = "GroupName, SkillName"
		Case SORT_BY_SKILL_IS_ENABLED
			str = "SkillIsEnabled DESC, SkillName"
		Case Else
			str = ""
	End Select
	
	LookupSortParam = str
End Function
%>
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_select_option.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_YesNoDropdownToString.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_skill_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	Public SortBy
	
	' encrypted
	Public Action
	Public ProgramID
	Public SkillID
	Public MemberID
	Public EmailId
	Public ProgramMemberId
	
	' objects
	Public Member
	Public Client
	Public Program
	Public Skill	
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(SortBy) > 0 Then str = str & "sb=" & SortBy & amp

		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(SkillID) > 0 Then str = str & "skid=" & Encrypt(SkillID) & amp
		If Len(MemberID) > 0 Then str = str & "mid=" & Encrypt(MemberID) & amp
		If Len(EmailId) > 0 Then str = str & "emid=" & Encrypt(EmailId) & amp
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
		c.SortBy = SortBy
		
		c.Action = Action
		c.ProgramID = ProgramID
		c.SkillID = SkillID
		c.MemberID = MemberID
		c.EmailId = EmailId
		c.ProgramMemberId = ProgramMemberId
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Skill = Skill
		
		Set Clone = c
	End Function
End Class
%>

