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
	page.SkillGroupID = Decrypt(Request.QueryString("sgid"))
	page.SkillID = Decrypt(Request.QueryString("skid"))
	page.EmailID = Decrypt(Request.QueryString("EmailID"))
	
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
	Set page.SkillGroup = New cSkillGroup
	page.SkillGroup.SkillGroupID = page.SkillGroupID
	If Len(page.SkillGroup.SkillGroupID) > 0 Then page.SkillGroup.Load()
	
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
		<link rel="stylesheet" type="text/css" href="skillgroups.css" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="skillgroups.js"></script>
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
		Case SEND_MESSAGE
			Call DoInsertEmail(page, rv)
			page.Action = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
			
		Case SHOW_SKILL_GROUP_DETAILS
			str = str & SkillGroupSummaryToString(page)

		Case ADDNEW_RECORD
			If Request.Form("FormSkillGroupIsPostback") = IS_POSTBACK Then
				Call LoadSkillGroupFromRequest(page.SkillGroup)
				If ValidSkillGroup(page.SkillGroup) Then
					Call DoInsertSkillGroup(page, rv)
					Select Case rv
						Case 0
							page.MessageID = 4031: page.Action = "": page.SkillGroupID = ""
						Case -2
							' duplicate group name
							page.MessageID = 4015: page.Action = "": page.SkillGroupID = ""
						Case Else
							page.MessageID = 4014: page.Action = "": page.SkillGroupID = ""
					End Select
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormSkillGroupToString(page)
				End If
			Else
				str = str & FormSkillGroupToString(page)
			End If
			
		Case UPDATE_RECORD
			If Request.Form("FormSkillGroupIsPostback") = IS_POSTBACK Then
				Call LoadSkillGroupFromRequest(page.SkillGroup)
				If ValidSkillGroup(page.SkillGroup) Then
					Call DoUpdateSkillGroup(page.SkillGroup, rv)
					Select Case rv
						Case 0
							page.MessageID = 4013: page.Action = "": page.SkillGroupID = ""
						Case -2
							' duplicate group name
							page.MessageID = 4015: page.Action = "": page.SkillGroupID = ""
						Case Else
							page.MessageID = 4014: page.Action = "": page.SkillGroupID = ""
					End Select
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormSkillGroupToString(page)
				End If
			Else
				str = str & FormSkillGroupToString(page)
			End If

		Case DELETE_RECORD
			Call DoDeleteSkillGroup(page.SkillGroup, rv)
			Select Case rv
				Case 0 
					page.MessageID = 4017: page.Action = "": page.SkillGroupID = ""
				Case Else
					page.MessageID = 4018: page.Action = "": page.SkillGroupID = ""
			End Select
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case ASSIGN_SKILLS_TO_GROUPS
			' check for groups ..
			If Not page.Program.HasSkillGroups Then
				page.MessageID = 4033 : page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
			If Request.Form("FormAssignGroupsIsPostback") = IS_POSTBACK Then
				Call DoUpdateGroups(page, Request.Form("IDList"), rv)
				str = str & "<p>rv=" & rv
				Select Case rv
					Case 0
						page.MessageID = 4020
					Case Else
						page.MessageID = 4021
				End Select
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormAssignSkillsToString(page)
			End If
			
		Case Else
			' check for skills ..
			If Not page.Program.HasSkills Then
				page.MessageID = 4033
				Response.Redirect("/admin/programs.asp" & page.UrlParamsToString(False))
			End If	
			
			str = str & SkillGroupGridToString(page)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoInsertEmail(page, outError)
	Dim email			: Set email = New cEmail
	
	email.MemberId = page.Member.MemberId
	email.ClientId = page.Member.ClientId
	email.GroupList = "skillgroup|" & page.SkillGroupId
	
	Call email.Add(outError)
	
	page.EmailId = email.EmailId
End Sub

Sub DoUpdateSkillGroup(skillGroup, outError)
	Call skillGroup.Save(outError)
End Sub

Sub DoDeleteSkillGroup(skillGroup, outError)
	Call skillGroup.Delete(outError)
End Sub

Sub DoInsertSkillGroup(page, outError)
	page.SkillGroup.ProgramID = page.Program.ProgramID
	page.SkillGroup.AllowMultiple = 1
	Call page.SkillGroup.Add(outError)
End Sub

Sub DoUpdateGroups(page, idLIst, outError)
	Dim i
	Dim skillID
	Dim oldGroupID
	Dim newGroupID
	
	Dim list			: list = Split(Replace(idList, " ", ""), ",")
	Dim formValues
	Dim skill			: Set skill = New cSkill
	Dim tempError		: tempError = 0
	Dim totalError		: totalError = 0
	
	For i = 0 To UBound(list)
		formValues = Split(list(i), "|")
		skillID = formValues(0)
		oldGroupID = formValues(1)
		newGroupID = formValues(2)

		' only update if changing
		If CStr(oldGroupID) <> CStr(newGroupID) Then
			skill.SkillID = skillID
			Call skill.Load()
			skill.SkillGroupID = newGroupID & ""
			Call skill.Save(tempError)
			totalError = totalError + tempError
		End If
	Next
	outError = totalError
	
	Set skill = Nothing
End Sub

Function SkillGridForSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	
	page.Skill.ProgramId = page.SkillGroup.ProgramId
	Dim list				: list = page.Skill.SkillList("")
	
	Dim rows				: rows = ""
	Dim alt					: alt = ""
	Dim href				: href = ""
	Dim count				: count = 0
	
	Dim skillEnabledText
	
	' 0-SkillID 1-SkillName 2-SkillDesc 3-IsSkillEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-IsGroupEnabled 8-LastModified 9-DateCreated
		
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
		
			If CStr(page.SkillGroup.SkillGroupId) = CStr(list(4,i) & "") Then
				alt = ""				: If count Mod 2 > 0 Then alt = " class=""alt"""
				
				pg.Action = SHOW_SKILL_DETAILS: pg.SkillId = list(0,i)
				href = "/admin/skills.asp" & pg.UrlParamsToString(True)
				
				skillEnabledText = "Yes"
				If list(3,i) = 0 Then skillEnabledText = "<span class=""negative"">No</span>"
				
				rows = rows & "<tr><td><img src=""/_images/icons/plugin.png"" class=""icon"" alt="""" />"
				rows = rows & "<strong>" & html(page.SkillGroup.GroupName) & "</strong> | "
				rows = rows & "<a href=""" & href & """ title=""Details""><strong>" & html(list(1,i)) & "</strong></a></td>"
				rows = rows & "<td>" & skillEnabledText & "</td>"
				rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
				rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
				
				count = count + 1
			End If
		Next
	End If
	
	If count > 0 Then
		str = str & "<p>This list of skills belongs to the " & html(page.SkillGroup.GroupName) & " group. </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Skill</th><th>Enabled</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = "<p class=""alert"">No skills have been assigned to this group. </p>"	
	End If
	
	SkillGridForSummaryToString = str
End Function

Function SkillGroupSummaryToString(page)
	Dim str
	Dim dateTime				: Set dateTime = New cFormatDate
	
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & html(page.SkillGroup.GroupName) & "</h3>"
	
	str = str & "<h5 class=""description"">Description</h5>"
	If Len(page.SkillGroup.GroupDesc) > 0 Then
		str = str & "<p>" & html(page.SkillGroup.GroupDesc) & "</p>"
	Else
		str = str & "<p class=""alert"">No description available. </p>"
	End If
	
	If page.SkillGroup.IsEnabled = 0 Then
		str = str & "<h5 class=""disabled"">Group is disabled</h5>"
		str = str & "<p class=""alert"">This group is disabled. Skills belonging to this group are hidden from your member accounts and cannot be added to event teams. </p>"
	End If
	
	str = str & "<h5 class=""skills"">Skill listing</h5>"
	str = str & SkillGridForSummaryToString(page)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul><li>Created on " & dateTime.Convert(page.SkillGroup.DateCreated, "DDDD MMMM dd, YYYY around hh:00 pp") & ". </li></ul>"
	
	str = str & "</div>"
	
	SkillGroupSummaryToString = str
End Function


Function SkillGroupGridToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	
	page.Skill.ProgramID = page.Program.ProgramID
	Dim list			: list = page.Skill.SkillGroupList()
	
	Dim count			: count = 0
	Dim altClass		: altClass = ""
	Dim enabledText		: enabledText = ""
	Dim hasSkills		: hasSkills = False
	Dim msg
		
	' tip-box
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	pg.Action = ADDNEW_RECORD
	str = str & "<ul><li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Create a new skill group</a></li>"
	pg.Action = ASSIGN_SKILLS_TO_GROUPS
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Assign my skills to groups</a></li>"
	str = str & "<li><a href=""/help/faq.asp?act=2&amp;fqid=96"" target=""_blank"">Learn more about skills and skill groups</a></li></ul></div>"

	' 0-SkillGroupID 1-GroupName 2-GroupDesc 3-IsEnabled 4-AllowMultiple 5-DateModified
	' 6-DateCreated 7-SkillCount 8-SkillListXMLFragment
	
	str = str & m_appMessageText
	str = str & "<div class=""grid"">"
	str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
	str = str & "<input type=""checkbox"" name=""master"" onclick="""" disabled=""disabled"" checked=""checked"" /></th>"
	str = str & "<th scope=""col"">Skill Group</th><th scope=""col"">Skills</th><th scope=""col"">Enabled</th><th scope=""col"">&nbsp;</th></tr>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			count = count + 1
			altClass = ""
			If count Mod 2 = 0 Then altClass = " class=""alt"""
			enabledText = "Yes"
			If list(3,i) = 0 Then enabledText = "<span style=""color:red;"">No</span>"
			
			str = str & "<tr" & altClass & "><td><input type=""checkbox"" name=""Unused"" disabled=""disabled"" checked=""checked"" /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/link.png"" alt="""" />"
			str = str & "<strong>" & html(page.Program.ProgramName) & " | "
			pg.SkillGroupID = list(0,i): pg.Action = SHOW_SKILL_GROUP_DETAILS
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(1,i)) & "</a></strong></td>"
			str = str & "<td>" & list(7,i) & "</td>"
			str = str & "<td>" & enabledText & "</td>"
			str = str & "<td class=""toolbar"">"
			pg.SkillGroupID = list(0,i): pg.Action = SHOW_SKILL_GROUP_DETAILS
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details""><img src=""/_images/icons/magnifier.png"" alt=""icon"" /></a>"
			pg.SkillGroupID = list(0,i): pg.Action = UPDATE_RECORD
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Edit""><img src=""/_images/icons/pencil.png"" alt=""icon"" /></a>"
			pg.SkillGroupId = list(0,i): pg.Action = SEND_MESSAGE
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email""><img src=""/_images/icons/email.png"" alt=""icon"" /></a>"
			pg.SkillGroupID = list(0,i): pg.Action = DELETE_RECORD
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove""><img src=""/_images/icons/cross.png"" alt=""icon"" /></a>"
			
			str = str & "</td></tr>"
		Next
	End If
	str = str & "</table></div>"
	
	If count = 0 Then
		str = ""
		msg = msg & "You have not created any skill groups for the <strong>" & html(page.Program.ProgramName) & "</strong> program . "
		pg.Action = ADDNEW_RECORD: pg.SkillID = "": pg.SkillGroupID = ""
		msg = msg & "Click <a href=""" & pg.Url & pg.UrlParamsToString(True) & """>here</a> to create one. "
		str = str & CustomApplicationMessageToString("No skill groups returned!", msg, "Error")
	End If
	
	SkillGroupGridToString = str
End Function

Function FormAssignSkillsToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	
	page.Skill.ProgramID = page.Program.ProgramID
	Dim list				: list = page.Skill.SkillList("")
	
	Dim groups				: groups = page.Skill.SkillGroupList()
	Dim altClass			: altClass = ""
	
	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "If you make any changes to this page, be sure to click <strong>Save</strong> in the toolbar. </div>"
	
	' 0-SkillID 1-SkillName 2-SkillDesc 3-SkillIsEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-GroupIsEnabled 8-DateGroupModified 9-DateGroupCreated 10-DateSkillModified
	' 11-DateSkillCreated 12-MemberCount

		
	str = str & m_appMessageText
	str = str & "<div class=""grid"">"
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" id=""form-assign-groups"">"
	str = str & "<input type=""hidden"" name=""FormAssignGroupsIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
	str = str & "<input type=""checkbox"" name=""master"" onclick="""" disabled=""disabled"" checked=""checked"" /></th>"
	str = str & "<th scope=""col"">Skill Group</th><th scope=""col"">&nbsp;</th></tr>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			altClass = ""
			If i Mod 2 <> 0 Then altClass = " class=""alt"""
			
			str = str & "<tr" & altClass & "><td><input type=""checkbox"" name=""Unused"" disabled=""disabled"" checked=""checked"" /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/plugin.png"" alt="""" />"
			str = str & "<strong>" & html(page.Program.ProgramName) & " | "
			pg.SkillID = list(0,i): pg.Action = SHOW_SKILL_DETAILS
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(1,i)) & "</a></strong></td>"
			str = str & "<td class=""toolbar"">" & GroupDropdownToString(groups, list(4,i), list(0,i)) & "</td>"
		Next
	End If
	str = str & "</table>"
	str = str & "<p style=""text-align:right;margin-right:0;"">"
	str = str & "<input type=""submit"" name=""Submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "</p></form></div>"
	
	FormAssignSkillsToString = str
End Function

Function FormSkillGroupToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" name=""formSkillGroup"">"
	str = str & "<table><tr><td class=""label"">" & RequiredElementToString(True, "Skill group name") & "</td>"
	str = str & "<td><input class=""medium gets-focus"" type=""text"" name=""GroupName"" value=""" & HTML(page.SkillGroup.GroupName) & """ /></td></tr>"
	str = str & "<tr><td class=""label"">Description</td>"
	str = str & "<td><textarea class=""medium"" name=""GroupDesc"">" & HTML(page.SkillGroup.GroupDesc) & "</textarea></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Enabled</td>"
	str = str & "<td>" & YesNoDropdownToString(page.SkillGroup.IsEnabled, "IsEnabled") & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "When set to yes, skills that belong to this group are available <br />to members in their profiles and are available to administrators <br />for scheduling. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td><td>"
	str = str & "<input type=""submit"" name=""Submit"" value=""Save"" />"
	pg.Action = "": pg.SkillGroupID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormSkillGroupIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</td></tr></table></form></div>"
	
	FormSkillGroupToString = str
End Function

Function LoadSkillGroupFromRequest(skillGroup)
	skillGroup.GroupName = Request.Form("GroupName")
	skillGroup.GroupDesc = Request.Form("GroupDesc")
	skillGroup.IsEnabled = Request.Form("IsEnabled")
End Function

Function ValidSkillGroup(skillGroup)
	ValidSkillGroup = True
	
	If Not ValidData(skillGroup.GroupName, True, 0, 100, "Group Name", "") Then ValidSkillGroup = False
	If Not ValidData(skillGroup.GroupDesc, False, 0, 2000, "Description", "") Then ValidSkillGroup = False
End Function

Function GroupDropdownToString(groups, groupID, skillID)
	Dim str, i

	' value of each option is a delim list .. SkillID|oldGroupID|newGroupID
	
	str = str & "<select name=""IDList"">"
	str = str & "<option value=""" & skillID & "|" & groupID & "|" & "" & """>None</option>"
	If IsArray(groups) Then
		For i = 0 To UBound(groups,2)
			Dim selected			: selected = ""
			If CStr(skillID & "|" & groupID & "|" & groups(0,i)) = CStr(skillID & "|" & groupID & "|" & groupID) Then
				selected = " selected=""selected"""
			End If
			
			str = str & "<option" & selected & " value=""" & skillID & "|" & groupID & "|" & groups(0,i) & """>" & html(groups(1,i)) & "</option>" 
		Next
	End If
	str = str & "</select>"
	
	GroupDropdownToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim programLink
	pg.SkillID = "": pg.Action = SHOW_PROGRAM_DETAILS
	programLink = "<a href=""/admin/programs.asp" & pg.UrlParamsToString(True) & """>" & html(pg.Program.ProgramName) & "</a> / "

	Dim skillsLink
	pg.SkillID = "": pg.Action = ""
	skillsLink = "<a href=""/admin/skills.asp" & pg.UrlParamsToString(True) & """>Skills</a> / "
	
	Dim skillGroupLink
	pg.Action = SHOW_SKILL_GROUP_DETAILS
	skillGroupLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.SkillGroup.GroupName) &"</a> / "

	Dim skillGroupsLink
	pg.SkillGroupID = "": pg.Action = ""
	skillGroupsLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Groups</a> / "

	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case SHOW_SKILL_GROUP_DETAILS
			str = str & programLink
			str = str & skillsLink
			str = str & skillGroupsLink
			str = str & html(page.SkillGroup.GroupName)
			
		Case ADDNEW_RECORD
			str = str & programLink
			str = str & skillsLink
			str = str & skillGroupsLink
			str = str & "Group"
			
		Case UPDATE_RECORD
			str = str & programLink
			str = str & skillsLink
			str = str & skillGroupsLink
			str = str & skillGroupLink
			str = str & "Edit"
			
		Case DELETE_RECORD
			str = str & programLink
			str = str & skillsLink
			str = str & skillGroupsLink
			str = str & skillGroupLink
			str = str & "Remove"
			
		Case ASSIGN_SKILLS_TO_GROUPS
		Case Else	
			str = str & programLink
			str = str & skillsLink
			str = str & "Groups"
			
	End Select
	
	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href			: href = ""
	
	Dim saveButton
	href = "#"
	saveButton = saveButton & "<li id=""save-button""><a href=""" & href & """><img class=""icon"" src=""/_images/icons/disk.png"" /></a><a href=""" & href & """ title=""Save"">Save</a></li>"
		
	Dim groupListButton
	pg.Action = "": pg.SkillGroupID = ""
	href = pg.Url & pg.UrlParamsToString(True)
	groupListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/link.png"" /></a><a href=""" & href & """>Skill Group List</a></li>"

	Dim organizeSkillsButton
	pg.Action = ASSIGN_SKILLS_TO_GROUPS: pg.SkillGroupID = "" 
	href = pg.Url & pg.UrlParamsToString(True)
	organizeSkillsButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/plugin_link.png"" /></a><a href=""" & href & """>Organize</a></li>"

	Dim addGroupButton
	pg.Action = ADDNEW_RECORD: pg.SkillGroupID = ""
	href = pg.Url & pg.UrlParamsToString(True)
	addGroupButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/link_add.png"" /></a><a href=""" & href & """>Add Skill Group</a></li>"

	Dim skillListButton
	pg.Action = "": pg.SkillID = ""
	href = "/admin/skills.asp" & pg.UrlParamsToString(True)
	skillListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/plugin.png"" /></a><a href=""" & href & """>Skill List</a></li>"

	Select Case page.Action
		Case SHOW_SKILL_GROUP_DETAILS
			str = str & groupListButton & skillListButton
			
		Case ADDNEW_RECORD
			str = str & groupListButton
			
		Case UPDATE_RECORD
			str = str & groupListButton
			
		Case DELETE_RECORD
			str = str & groupListButton
			
		Case ASSIGN_SKILLS_TO_GROUPS
			str = str & groupListButton
			str = str & skillListButton
			str = str & saveButton
			
		Case Else
			str = str & addGroupButton	
			str = str & organizeSkillsButton
			str = str & skillListButton
			
	End Select	

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_select_option.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_group_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_YesNoDropdownToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_RemoveDupesFromStringList.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public ProgramID
	Public SkillID
	Public SkillGroupID
	Public EmailID
	
	' objects
	Public Member
	Public Client	
	Public Program	
	Public Skill	
	Public SkillGroup	
	
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
		If Len(SkillID) > 0 Then str = str & "skid=" & Encrypt(SkillID) & amp
		If Len(SkillGroupID) > 0 Then str = str & "sgid=" & Encrypt(SkillGroupID) & amp
		If Len(EmailID) > 0 Then str = str & "emid=" & Encrypt(EmailID) & amp
		
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
		c.SkillID = SkillID
		c.SkillGroupID = SkillGroupID
		c.EmailID = EmailID
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Skill = Skill
		Set c.SkillGroup = SkillGroup
		
		Set Clone = c
	End Function
End Class
%>

