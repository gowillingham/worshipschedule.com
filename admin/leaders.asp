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
	page.ProgramMemberID = Decrypt(Request.QueryString("pmid"))
	page.MemberID = Decrypt(Request.QueryString("mid"))
	page.EmailID = Decrypt(Request.QueryString("emid"))
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()
	Set page.ProgramMember = New cProgramMember
	page.ProgramMember.ProgramMemberID = page.ProgramMemberID
	If Len(page.ProgramMember.ProgramMemberID) > 0 Then Call page.ProgramMember.Load()
	
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
		<link rel="stylesheet" type="text/css" href="leaders.css" />
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
		Case DELETE_RECORD
			Call DoDeleteLeader(page.programMember, rv)
			Select Case rv
				Case 0
					page.MessageID = 3011
				Case Else
					page.MessageID = 3013
			End Select
			page.Action = "": page.ProgramMemberID = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
	
		Case ADDNEW_RECORD
			If Request.Form("FormLeaderIsPostback") = IS_POSTBACK Then
				page.ProgramMember.ProgramMemberID = Request.Form("ProgramMemberID")
				If ValidLeader(page.ProgramMember) Then
					Call DoInsertLeader(page.ProgramMember, rv)
					Select Case rv
						Case 0
							page.MessageID = 3010
						Case Else
							page.MessageID = 3013
					End Select
					page.Action = "": page.ProgramMemberID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormLeaderToString(page)
				End If
			Else
				str = str & FormLeaderToString(page)
			End If
			
		Case SEND_MESSAGE
			Call GenerateEmail(page, rv)
			page.Action = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
				
		Case Else
			str = str & LeaderGridToString(page)
			str = str & AdminGridToString(page)
			
	End Select 

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoDeleteLeader(programMember, outError)
	programMember.Load()
	programMember.IsLeader = 0
	Call programMember.Save(outError)
End Sub

Sub DoInsertLeader(programMember, outError)
	programMember.Load()
	programMember.IsLeader = 1 
	Call programMember.Save(outError)
End Sub

Function LeaderGridToString(page)
	Dim str, i
	Dim list			: list = page.Program.MemberList()
	Dim pg				: Set pg = page.Clone()
	
	Dim msg				: msg = ""
	Dim hasLeader		: hasLeader = False
	Dim altClass		: altClass = ""
	Dim count			: count = 0
	
	Dim leaderTipBox
	leaderTipBox = leaderTipBox & "<div class=""tip-box""><h3>Tip</h3>"
	leaderTipBox = leaderTipBox & "<p>Any members you add to the program leader list will have permission to manage the <strong>" & html(page.Program.ProgramName) & "</strong> program. </p></div>"
	Dim adminTipBox
	adminTipBox = adminTipBox & "<div class=""tip-box""><h3>Tip</h3>"
	adminTipBox = adminTipBox & "<p>Any member in your account administrator list also has permission to manage this program. </p></div>"

	' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 6-ProgramMemberIsEnabled 7-MemberActiveStatus 8-DateCreated 9-DateModified
	' 10-ProgramMemberID 11-IsApproved 12-HasMissingAvailability 13-Email
	
	str = str & leaderTipBox
	str = str & adminTipBox
	str = str & m_appMessageText
	If IsArray(list) Then
		str = str & "<div class=""grid"">"
		str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
		str = str & "<input type=""checkbox"" name=""master"" onclick="""" disabled=""disabled"" checked=""checked"" /></th>"
		str = str & "<th scope=""col"">" & html(page.Program.ProgramName) & " Program Leaders</th><th scope=""col"">&nbsp;</th></tr>"
		For i = 0 To UBound(list,2)
			If list(4,i) = 1 Then
				count = count + 1
				hasLeader = True
				altClass = ""
				If count Mod 2 = 0 Then altClass = " class=""alt"""

				str = str & "<tr" & altClass & "><td><input type=""checkbox"" name=""Unused"" disabled=""disabled"" checked=""checked"" /></td>"
				str = str & "<td><img class=""icon"" src=""/_images/icons/medal_gold_1.png"" alt="""" />"
				str = str & "<strong>" & html(page.Program.ProgramName) & " | "
				str = str & html(list(1,i) & ", " & list(2,i)) & "</strong></td>"
				str = str & "<td class=""toolbar"">"
				pg.MemberID = list(0,i): pg.Action = SEND_MESSAGE
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email""><img src=""/_images/icons/email.png"" alt="""" /></a>"
				pg.ProgramMemberID = list(10,i): pg.Action = DELETE_RECORD
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove""><img src=""/_images/icons/cross.png"" alt="""" /></a>"
				str = str & "</td></tr>"
			End If
		Next
		str = str & "</table></div>"	
	End If
	
	If Not hasLeader Then
		str = ""
		str = str & leaderTipBox
		str = str & adminTipBox
		
		pg.Action = ADDNEW_RECORD
		msg = msg & "No leaders have been set for the <strong>" & html(page.Program.ProgramName) & "</strong> program. "
		msg = msg & "Click <a href=""" & pg.Url & pg.UrlParamsToString(True) & """>here</a> to add a leader to the list. "
		str = str & CustomApplicationMessageToString("No leaders.", msg, "Error")
	End If
	
	LeaderGridToString = str
End Function

Function AdminGridToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = page.Client.MemberList("", "")
	
	Dim altClass		: altClass = ""
	Dim count			: count = 0
	
	' 0-MemberID 1-NameLast 2-NameFirst 26-IsAdmin

	str = str & "<div class=""grid"">"
	str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
	str = str & "<input type=""checkbox"" name=""master"" onclick="""" disabled=""disabled"" checked=""checked"" /></th>"
	str = str & "<th scope=""col"">" & html(page.Client.NameClient) & " Account Administrators</th><th scope=""col"">&nbsp;</th></tr>"
	For i = 0 To UBound(list,2)
		If list(26,i) = 1 Then
			count = count + 1
			altClass = ""
			If count Mod 2 = 0 Then altClass = " class=""alt"""
				
			str = str & "<tr" & altClass & "><td><input type=""checkbox"" name=""Unused"" disabled=""disabled"" checked=""checked"" /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/medal_gold_2.png"" alt="""" />"
			str = str & "<strong>" & html(page.Client.NameClient) & " | "
			str = str & html(list(1,i) & ", " & list(2,i)) & "</strong></td>"
			str = str & "<td class=""toolbar"">"
				pg.MemberID = list(0,i): pg.Action = SEND_MESSAGE
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email""><img src=""/_images/icons/email.png"" alt="""" /></a>"
			str = str & "</td></tr>"
		End If
	Next
	str = str & "</table></div>"
	
	AdminGridToString = str
End Function

Function ValidLeader(programMember)
	ValidLeader = True
	If Application.Value("FORM_VALIDATION_OFF") Then Exit Function
	
	If Not ValidData(programMember.ProgramMemberID, True, 0, 100, "A member", "") Then ValidLeader = False
End Function

Function FormLeaderToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" name=""formLeader"">"
	str = str & "<table>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Member") & "</td>"
	str = str & "<td>" & CandidateDropdownToString(page) & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormLeaderIsPostback"" value=""" & IS_POSTBACK & """ /></td></tr>"
	str = str & "</table></form></div>"

	FormLeaderToString = str
End Function

Function CandidateDropdownToString(page)
	Dim str
	Dim list		: list = GetCandidateList(page.ProgramID)
	
	str = str & "<select name=""ProgramMemberID"">"
	str = str & "<option value="""">" & HTML("< Choose a member >") & "</option>"
	str = str & SelectOption(list, "")
	str = str & "</select>"
	
	CandidateDropdownToString = str 
End Function

Function GetCandidateList(programID)
	Dim i
	Dim newList()
	Dim memberProgram	: Set memberProgram = New cProgramMember
	memberProgram.ProgramID = programID
	
	Dim list			: list = memberProgram.GetMemberList()
	Dim hasCandidates	: hasCandidates = False
	
	' 0-MemberID 1-NameLast 2-NameFirst 4-IsLeader 10-ProgramMemberID
	
	If IsArray(list) Then
		ReDim newList(1,0)
		For i = 0 To UBound(list,2)
			If list(4,i) = 0 Then
				hasCandidates = True
				newList(0, UBound(newList,2)) = list(10,i)
				newList(1, UBound(newList,2)) = list(1,i) & ", " & list(2,i)
				ReDim Preserve newList(1, UBound(newList,2) + 1)
			End If
		Next
		If hasCandidates Then
			ReDim Preserve newList(1, UBound(newList,2) - 1)
			GetCandidateList = newList
		End If
	End If
	
	Set memberProgram = Nothing
End Function

Sub GenerateEmail(page, outError)
	Dim email		: Set email = New cEmail
	email.MemberID = page.Member.MemberID
	email.ClientID = page.Client.ClientID
	email.RecipientIDList = page.MemberID
	
	Call email.Add(outError)
	page.EmailID = email.EmailID
	
	Set email = Nothing
End Sub

Sub SetPageHeader(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case ADDNEW_RECORD
			str = str & "Add " & html(page.Program.ProgramName) & " Leader"
		Case Else
			pg.Action = SHOW_PROGRAM_DETAILS
			str = str & "<a href=""/admin/programs.asp" & pg.UrlParamsToString(True) & """>" & html(page.Program.ProgramName) & "</a> / "
			str = str & " Leaders"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href			: href = ""
	
	Dim addLeaderButton
	pg.Action = ADDNEW_RECORD
	href =  pg.Url & pg.UrlParamsToString(True)
	addLeaderButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/medal_gold_1_add.png""  alt="""" /></a><a href=""" & href & """>Add Leader</a></li>"
	
	Dim leaderListButton
	pg.Action = ""
	href =  pg.Url & pg.UrlParamsToString(True)
	leaderListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/table.png""  alt="""" /></a><a href=""" & href & """>Leader List</a></li>"
	
	Dim programListButton
	Set pg = page.Clone()
	pg.Action = "": pg.ProgramID = ""
	href = "/admin/programs.asp" & pg.UrlParamsToString(True)
	programListButton = programListButton & "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/table.png"" /></a><a href=""" & href & """>Program List</a></li>"
	
	Select Case page.Action
		Case ADDNEW_RECORD
			str = str & leaderListButton
		Case DELETE_RECORD
			str = str & leaderListButton
		Case Else
			str = str & addLeaderButton
			str = str & programListButton
	End Select
	
	' use this if no buttons will be displayed ..
	' str = str & "<li>&nbsp;</li>"
	
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
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public ProgramID
	Public ProgramMemberID
	Public MemberID
	Public EmailID
	
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
		If Len(MemberID) > 0 Then str = str & "mid=" & Encrypt(MemberID) & amp
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
		c.MemberID = MemberID
		c.EmailID = EmailID
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.ProgramMember = ProgramMember
		
		Set Clone = c
	End Function
End Class
%>

