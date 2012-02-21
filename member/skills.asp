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
	page.SkillID = Decrypt(Request.QueryString("skid"))
	page.ProgramMemberID = Decrypt(Request.QueryString("pmid"))
	page.SkillIDList = Request.Form("SkillIDList")
	page.SortBy	= Request.QueryString("sb")
	
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
	
	Set page.ProgramMember = New cProgramMember
	page.ProgramMember.ProgramMemberID = page.ProgramMemberID
	If Len(page.ProgramMember.ProgramMemberID) > 0 Then page.ProgramMember.Load()
	
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
		<link rel="stylesheet" type="text/css" href="/_incs/script/jquery/jquery-ui-1.6.custom/development-bundle/themes/start/ui.all.css" />	
		<link rel="stylesheet" type="text/css" href="skills.css" />	
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=m_pageTitleText%></title>
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/jquery-ui-1.6.custom/js/jquery-ui-1.6.custom.min.js"></script>
		<script type="text/javascript" src="skills.js"></script>
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
		Case SHOW_SKILL_DETAILS
			str = str & SkillSummaryToString(page)
			
		Case UPDATE_RECORD
			If Request.Form("FormSkillIsPostback") = IS_POSTBACK Then
				str = str & FormConfirmSkillUpdateToString(page)
			ElseIf Request.Form("FormConfirmSkillUpdateIsPostback") = IS_POSTBACK Then
				Call UpdateProgramMemberSkill(page, rv)
				Select Case rv
					Case 0
						page.Action = "": page.MessageID = 4023
					Case Else
						page.Action = "": page.MessageID = 4025
				End Select
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
		Case Else
			str = str & SkillGridToString(page)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub UpdateProgramMemberSkill(page, outError)
	Dim str, i, j, rv, deleteThis, hasError
	
	Dim programMemberSkill		: Set programMemberSkill = New cProgramMemberSkill
	
	' list of skillIDs from the form
	Dim list					: list = Split(Replace(page.SkillIDList, " ", ""), ",")
	
	' list of existing programMemberSkills
	Dim currentList				: currentList = page.programMember.GetSkillList("")
	
	' 0-SkillID 1-ProgramMemberSkillID 2-ProgramMemberID 3-SkillName 4-SkillDesc 5-SkillGroupName
	' 6-IsProgramMemberSkill 7-DateCreated 8-SkillIsEnabled 9-SkillGroupIsEnabled
	
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

Function MemberSkillGridForSkillSummaryToString(page, count)
	Dim str, i
	
	Dim list				: list = page.Skill.MemberList()
	Dim rows
	Dim alt
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberSkillID 5-ProgramMemberID
	' 6-IsApproved 7-ProgramMemberIsActive 8-Email
		
	count = 0
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isMemberEnabled = True				: If list(3,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True		: If list(7,i) = 0 Then isProgramMemberEnabled = False
			
			If isMemberEnabled And isProgramMemberEnabled Then
				alt = ""		: If count Mod 2 > 0 Then alt = " class=""alt"""
				
				rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/user.png"" alt="""" />"
				rows = rows & "<strong>" & html(list(1,i) & ", " & list(2,i)) & "</strong></td>"
				rows = rows & "</tr>"
						
				count = count + 1
			End If
		Next	
	End If
	
	str = str & "<div class=""grid"">"
	str = str & "<table><thead><tr><th>Member Name</th></tr></thead>"
	str = str & "<tbody>" & rows & "</tbody></table></div>"
	
	If count = 0 Then
		str = "<p class=""alert"">No members from this program have this skill in their profile. </p>"
	End If
	
	MemberSkillGridForSkillSummaryToString = str
End Function

Function SkillSummaryToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim dateTime		: Set dateTime = New cFormatDate
	
	Dim memberCount

	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & html(page.Skill.SkillName) & "</h3>"
	
	str = str & "<h5 class=""description"">Description</h5>"
	If Len(page.Skill.SkillDesc) > 0 Then
		str = str & "<p>" & html(page.Skill.SkillDesc) & "</p>"
	Else
		str = str & "<p class=""alert"">No description available.</p>"
	End If
	
	str = str & "<h5 class=""program-member"">Members with " & html(page.Skill.SkillName) & "</h5>"
	str = str & MemberSkillGridForSkillSummaryToString(page, memberCount)

	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul>"
	str = str & "<li>Created on " & dateTime.Convert(page.Skill.DateCreated, "DDD MMM dd, YYYY") & "</li>"
	If memberCount = 0 Then
		str = str & "<li>No members have this skill in their profile. </li>"
	ElseIf memberCount = 1 Then 
		str = str & "<li>One member has this skill in their profile. </li>"
	Else
		str = str & "<li>" & memberCount & " members have this skill in their profile. </li>"
	End If
	str = str & "</ul>"
	
	str = str & "</div>"
		
	SkillSummaryToString = str
End Function

Function FormConfirmSkillUpdateToString(page)
	Dim str, msg
	Dim pg				: Set pg = page.Clone()
	
	msg = msg & "You are changing your skill profile for the program <strong>" & html(page.ProgramMember.ProgramName) & "</strong>. "
	msg = msg & "If you are removing skills, you will lose the calendar and schedule information for any skills that are being removed. "
	msg = msg & "This action cannot be reversed. "

	str = str & CustomApplicationMessageToString("Please confirm this action! ", msg, "Confirm")	
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formConfirmSkillUpdate"">"
	str = str & "<input type=""hidden"" name=""FormConfirmSkillUpdateIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<input type=""hidden"" name=""SkillIDList"" value=""" & page.SkillIDList & """ />"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Confirm"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "</p></form>"

	FormConfirmSkillUpdateToString = str
End Function

Function NoSkillsDialogToString(page)
	Dim dialog			: Set dialog = New cDialog
	
	dialog.Headline = "Whoa, where are the skills?"
	
	dialog.Text = dialog.Text & "<p>Either this program (" & html(page.Program.ProgramName) & ") doesn't have any skills set up yet, "
	dialog.Text = dialog.Text & "or they have all been disabled (for whatever reason!) by an account administrator. "
	dialog.Text = dialog.Text & ""
	dialog.Text = dialog.Text & "</p>"
	
	dialog.SubText = dialog.SubText & "<p>When this is fixed, you will use this page to select or change the skills that will belong to your profile for this program. "
	dialog.SubText = dialog.SubText & "Then whoever schedules the " & html(page.Program.ProgramName) & " program will assign or schedule you for event teams based on your skills from this list. </p>"
	
	dialog.LinkList = dialog.LinkList & "<li><a href=""/member/contacts.asp"">Email account administrator</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/member/programs.asp"">Back to my program list</a></li>"
	
	NoSkillsDialogToString = dialog.ToString()
End Function

Function SkillGridToString(page)
	Dim str, msg, i
	Dim pg				: Set pg = page.Clone()
	
	page.programMember.ProgramMemberID = page.ProgramMemberID
	Dim list			: list = page.programMember.GetSkillList(LookupSortParam(page.SortBy))
	Dim isChecked		: isChecked = ""
	Dim isDisabled		: isDisabled = ""
	Dim mySkillText		: mySkillText = ""
	Dim count			: count = 0
	Dim altClass
	
	Dim isSkillEnabled
	Dim isSkillGroupEnabled
	Dim isProgramMemberSkill
	
	Dim instructions
	
	If page.Program.MemberCanEditSkills = 0 Then
		isDisabled = " disabled=""disabled"""
		str = str & "<div class=""tip-box""><h3>Tip</h3>"
		str = str & "<p>Editing the skill list for this program has been disabled. </p>"
		str = str & "<p style=""margin-top:10px;"">You'll need to contact an administrator for the <strong>" & html(page.Program.ProgramName) & "</strong> program if you need to edit this list. </p></div>"

		instructions = "<p>The available skills for the <strong>" & server.HtmlEncode(page.Program.ProgramName) & "</strong> program are listed below. "
		instructions = instructions & "This program has been set so that only an administrator can add or remove skills from your profile. </p>"
	Else
		str = str & "<div class=""tip-box""><h3>Tip</h3>"
		str = str & "<p>If you make any changes, don't forget to <strong>click save</strong> before leaving this page. </p></div>"

		instructions = "<p>You can set you skills for the <strong>" & server.HtmlEncode(page.Program.ProgramName) & "</strong> program in the listing below. "
		instructions = instructions & "Be sure and click <strong>Save</strong> if you make any changes to this list. </p>"
	End If

	' 0-SkillID 1-ProgramMemberSkillID 2-ProgramMemberID 3-SkillName 4-SkillDesc 5-SkillGroupName
	' 6-IsProgramMemberSkill 7-DateCreated 8-SkillIsEnabled 9-SkillGroupIsEnabled



	str = str & m_appMessageText
	str = str & "<h3>" & html(page.Program.ProgramName) & " skills</h3>"
	str = str & instructions

	If IsArray(list) Then
		str = str & "<div class=""grid"">"
		pg.Action = UPDATE_RECORD
		str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" name=""formSkill"">"
		str = str & "<input type=""hidden"" name=""FormSkillIsPostback"" value=""" & IS_POSTBACK & """ />"
		str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
		str = str & "<input type=""checkbox"" id=""master""" & isDisabled & " /></th>"
		str = str & "<th scope=""col"">" & html(page.Program.ProgramName) & " Skills</th><th scope=""col"">Group</th>"
		str = str & "<th scope=""col"">My Skill</th><th scope=""col"" style=""text-align:right;"">"
		str = str & "<input type=""submit"" name=""dummy"" value=""Save""" & isDisabled & " /></th></tr>"
		
		For i = 0 To UBound(list,2)
			isSkillEnabled = False			: If list(8,i) = 1 Then isSkillEnabled = True
			isSkillGroupEnabled = False		: If list(9,i) = 1 Then isSkillGroupEnabled = True
			isProgramMemberSkill = False	: If list(6,i) = 1 Then isProgramMemberSkill = True

			' both skill and group are enabled ..
			If isSkillEnabled And isSkillGroupEnabled Then
				
				count = count + 1
				altClass = ""
				If i Mod 2 <> 0 Then altClass = " class=""alt"""
				isChecked = ""
				mySkillText = "&nbsp;"
				If isProgramMemberSkill Then 
					isChecked = " checked=""checked"""
					mySkillText = "Yes"
				End If				
				
				str = str & "<tr" & altClass & "><td><input type=""checkbox"" name=""SkillIDList"" value=""" & list(0,i) & """" & isChecked & isDisabled & " /></td>"
				str = str & "<td><img class=""icon"" src=""/_images/icons/plugin.png"" alt=""icon"" />"
				str = str & "<strong>" & html(page.Program.ProgramName) & " | "
				pg.SkillID = list(0,i): pg.Action = SHOW_SKILL_DETAILS 
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(3,i)) & "</a></strong></td>"
				str = str & "<td>" & html(list(5,i)) & "</td>"
				str = str & "<td>" & mySkillText & "</td>"
				str = str & "<td class=""toolbar"">"
				pg.SkillID = list(0,i): pg.Action = SHOW_SKILL_DETAILS
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """><img src=""/_images/icons/magnifier.png"" alt=""icon"" /></a>"
				str = str & "</td></tr>"
			End If
		Next
		str = str & "<tr><th scope=""col"" colspan=""5"" style=""text-align:right;"">"
		str = str & "<input type=""submit"" name=""dummy"" value=""Save"" /></th></tr>"
		str = str & "</table>"
		str = str & "</form></div>"
	End If
	
	If count = 0 Then
		str = NoSkillsDialogToString(page)
	End If
	
	SkillGridToString = str
End Function

Function SortByDropdownToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formSortByDropdown"">"
	str = str & "<input type=""hidden"" name=""FormSortByDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""SortBy"" onchange=""document.forms.formSortByDropdown.submit();"">"
	str = str & "<option value="""">" & html("< Sort by .. >") & "</option>"
	str = str & "<option value=""" & SORT_BY_SKILLNAME & """>Skill</option>"
	str = str & "<option value=""" & SORT_BY_SKILLGROUP & """>Skill Group</option>"
	str = str & "<option value=""" & SORT_BY_IS_MEMBER_SKILL & """>My Skills</option>"
	str = str & "</select></form></li>"	
	
	SortByDropdownToString = str
End Function

Function LookupSortParam(val)
	Dim str
	
	Select Case val
		Case SORT_BY_SKILLNAME
			str = "SkillName"
		Case SORT_BY_SKILLGROUP 
			str = "SkillGroupName, SkillName"
		Case SORT_BY_IS_MEMBER_SKILL
			str = "IsProgramMemberSkill DESC, SkillName"
		Case Else
			str = ""
	End Select
	
	LookupSortParam = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href
	
	Dim programLink
	pg.Action = SHOW_PROGRAM_DETAILS
	href = "/member/programs.asp" & pg.UrlParamsToString(True)
	programLink = "<a href=""" & href & """>" & html(page.Program.ProgramName) & "</a> / "
	
	str = str & "<a href=""/member/overview.asp"">Member Home</a> / "
	Select Case page.Action
		Case SHOW_SKILL_DETAILS
			str = str & "<a href=""/member/programs.asp"">Programs</a> / " 
			pg.Action = ""
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.Skill.ProgramName) & " Skills</a> / "
			str = str & html(page.Skill.SkillName)
		Case Else
			str = str & "<a href=""/member/programs.asp"">Programs</a> / "
			str = str & programLink
			str = str & "Skills"
	End Select
	

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	Dim href				: href = ""
	

	Dim skillListButton
	pg.Action = ""
	href = pg.Url & pg.UrlParamsToString(True)
	skillListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/plugin.png"" /></a><a href=""" & href & """>Skill List</a></li>"
	
	Dim programListButton
	pg.Action = "": pg.ProgramID = "": pg.ProgramMemberID = ""
	href = "/member/programs.asp" & pg.UrlParamsToString(True) 
	programListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/script.png"" alt="""" /></a><a href=""" & href & """>Program List</a></li>"

	Select Case page.Action
		Case SHOW_SKILL_DETAILS
			str = str & skillListButton
		Case UPDATE_RECORD
			str = str & programListButton
		Case Else
			str = str & SortByDropdownToString(page)
			str = str & programListButton
	End Select
	
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
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	Public SortBy
	
	' encrypted
	Public Action
	Public ProgramID
	Public SkillID
	Public ProgramMemberID
	
	' not persisted
	Public SkillIDList
	
	' objects
	Public Member
	Public Client
	Public Program
	Public Skill	
	Public ProgramMember
	
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
		c.SortBy = SortBy
		c.Action = Action
		c.ProgramID = ProgramID
		c.SkillID = SkillID
		c.ProgramMemberID = ProgramMemberID
		Set c.Client = Client
		Set c.Member = Member
		Set c.Program = Program
		Set c.Skill = Skill
		Set c.ProgramMember = ProgramMember
		
		Set Clone = c
	End Function
End Class
%>

