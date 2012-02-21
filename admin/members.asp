<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Const SAVE_AND_NEW = "Save and New"

' hack: getting global programId for page to embed in jscript ..
Dim m_programId

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
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.MemberID = Decrypt(Request.QueryString("mid"))
	page.ProgramMemberID = Decrypt(Request.QueryString("pmid"))
	page.EmailID = Decrypt(Request.QueryString("emid"))
	
	page.EmailList = Request.Form("EmailList")
	page.InviteNote = Request.Form("InviteNote")
	
	page.MemberIDList = Request.Form("MemberIDList")
	page.ProgramMemberIDList = Request.Form("ProgramMemberIDList")
	
	If Request.Form("FormProgramDropdownIsPostback") = IS_POSTBACK Then
		page.ProgramID = Request.Form("NewProgramID")
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
	Set page.ProgramMember = New cProgramMember
	page.ProgramMember.ProgramMemberID = page.ProgramMemberID
	If Len(page.ProgramMember.ProgramMemberID) > 0 Then page.ProgramMember.Load()
	Set page.ThisMember = New cMember
	page.ThisMember.MemberID = page.MemberID
	If Len(page.ThisMember.MemberID) > 0 Then page.ThisMember.Load()
	
	m_programId = page.Program.ProgramId
	
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
		<link type="text/css" href="/_incs/script/jquery/jquery-ui-1.6.custom/development-bundle/themes/start/ui.all.css" rel="stylesheet" />	
		<link rel="stylesheet" type="text/css" href="members.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=m_pageTitleText%></title>
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/jquery-ui-1.6.custom/js/jquery-ui-1.6.custom.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/plugins/cookie/jquery.cookie.js"></script>
		<script type="text/javascript" src="members.js"></script>
		<script language="javascript" type="text/javascript">
			// translate serverside globals to client side global ..
			var ADD_PROGRAM_MEMBER = <%=ADD_PROGRAM_MEMBER %>
			var REMOVE_PROGRAM_MEMBER = <%=REMOVE_PROGRAM_MEMBER %>
			var PROGRAM_ID_ENCRYPTED = "<%=Encrypt(m_programId) %>"
		</script>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)
	
	Select Case page.Action
		Case GOTO_MEMBER_PROFILE
			If Len(Request.Form("NewMemberID")) > 0 Then
				page.MemberID = Request.Form("NewMemberID"): page.Action = ""
				Response.Redirect("/admin/profile.asp" & page.UrlParamsToString(False))
			Else
				page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
		Case ADDNEW_RECORD
			If Request.Form("FormNewMemberIsPostback") = IS_POSTBACK Then
				Call LoadNewMemberFromRequest(page.ThisMember, page.ProgramMember)
				If ValidNewMember(page.ThisMember) Then
					Call DoInsertNewMember(page, rv)
					If rv = 0 Then
						' success
						Call SendNewMemberLogin(page.ThisMember, page.Member)
						Call SendNewMemberWelcome(page.ThisMember, page.Member)
						page.MessageID = 2027
					ElseIf rv = -2 Then
						' dupe member
						page.MessageID = 2029
					Else
						' unknown error
						page.MessageID = 2028
					End If
					If Request.Form("Submit") = SAVE_AND_NEW Then
						Response.Redirect(page.Url & page.UrlParamsToString(False))
					End If
					page.Action = "": page.MemberID = ""
					Response.Redirect("/admin/members.asp" & page.UrlParamsToString(False))					
				Else
					str = str & FormNewMemberToString(page)
				End If
			Else
				str = str & FormNewMemberToString(page)
			End If
		
		Case REMOVE_PROGRAM_MEMBER
			If Request.Form("FormConfirmDeleteProgramMemberIsPostback") = IS_POSTBACK Then
				Call DoDeleteProgramMember(page, rv)
				Select Case rv
					Case 0
						page.MessageID = 3017: page.Action = "": page.MemberID = ""
					Case Else
						page.MessageID = 3018: page.Action = "": page.MemberID = ""
				End Select
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmDeleteProgramMember(page)
			End If
			
		Case BULK_REMOVE_CLIENT_MEMBERS
			If Request.Form("FormConfirmBulkRemoveClientMembersIsPostback") = IS_POSTBACK Then
				Call DoDeleteClientMemberByList(page, rv)
				Select Case rv
					Case 0
						page.MessageID = 1065
					Case Else
						page.MessageID = 1066
				End Select
				page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmBulkRemoveClientMembersToString(page)			
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
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmDeleteMemberToString(page)
			End If
			
		Case INVITE_MEMBERS
			If Request.Form("FormInviteMembersIsPostback") = IS_POSTBACK Then
				If ValidFormInvite(page) Then
					Call DoSendInvite(page)
					page.MessageID = 2038: page.Action = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormInviteMembersToString(page)
				End If
			Else 
				str = str & FormInviteMembersToString(page)
			End If
						
		Case SEND_MESSAGE
			Call GenerateEmail(page, rv)
			page.Action = "": page.MemberID = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
			
		Case Else
			str = str & MemberGridToString(page)
			
			' todo: get this into a function ..
			str = str & "<div title=""Choose program members"" id=""program-member-widget""></div>"
			str = str & "<div id=""bulk-delete-modal-dialog""></div>"
			
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoSendInvite(page)
	Dim str, i
	Dim message				: Set message = New cEmailSender
	Dim list				: list = Split(page.EmailList, ",")
	Dim memberName			: memberName = page.Member.NameFirst & " " & page.Member.NameLast
	Dim subject				: subject = "[" & Application.Value("APPLICATION_NAME") & "] ** " & page.Client.NameClient & " Member Invitation **"
	
	If Len(page.EmailList) = 0 Then Exit Sub
	If Not IsArray(list) Then Exit Sub
	
	For i = 0 To UBound(list)
		str = ""
		str = str & "Dear " & list(i)
		str = str & vbCrLf & vbCrLf & "This is an invitation from " & memberName & " with " & page.Client.NameClient & " to join their " & Application.Value("APPLICATION_NAME") & " account. "
		str = str & "Click on this link to complete the new account process .."
		str = str & vbCrLf & "http://" & Request.ServerVariables("SERVER_NAME") & "/newmember.asp?gid=" & page.Client.Guid
		
		str = str & vbCrLf & vbCrLf & "You'll just need to supply your name and email address to complete the process. "
		If Len(page.InviteNote) > 0 then
			str = str & vbCrLf & vbCrLf & "Additional info from " & memberName & ": "
			str = str & vbCrLf & String(60, "-")
			str = str & vbCrLf & page.InviteNote
			str = str & vbCrLf & String(60, "-")
		End If
		
		str = str & vbCrLf & vbCrLf & "Questions? Having trouble creating your account? Please contact " & Application.Value("APPLICATION_NAME") & " support here .."
		str = str & vbCrLf & "http://" & Request.ServerVariables("SERVER_NAME") & "/support.asp"
		str = str & vbCrLf & vbCrLf & "Thanks for your interest in " & Application.Value("APPLICATION_NAME") & ". "
		str = str & EmailDisclaimerToString(page.Client.NameClient)
		
		Call message.SendMessage(list(i), page.Member.Email, subject, str)
	Next
	
	
	Set message = Nothing
End Sub

Sub GenerateEmail(page, outError)
	Dim email		: Set email = New cEmail
	
	email.MemberID = page.Member.MemberID
	email.ClientID = page.Client.ClientID
	email.RecipientIDList = page.MemberID
	Call email.Add(outError)
	
	page.EmailID = email.EmailID
End Sub

Sub DoDeleteClientMemberByList(page, outError)
	Dim i
	Dim list			: list = Split(page.MemberIDList, ",")
	
	Dim tempError		: tempError = 0
	outError = 0
	
	If Len(page.MemberIDList) = 0 Then Exit Sub
	
	
	For i = 0 To UBound(list)
		page.ThisMember.MemberID = list(i)
		Call page.ThisMember.Delete(tempError)
		outError = outError + tempError
	Next
End Sub

Sub DoDeleteProgramMember(page, outError)
	page.ProgramMember.MemberID = page.ThisMember.MemberID
	page.ProgramMember.ProgramID = page.Program.ProgramID
	
	Call page.ProgramMember.LoadByMemberProgram()
	Call page.ProgramMember.Delete(outError)
End Sub

Sub DoInsertNewMember(page, outError)
	page.ThisMember.ClientID = page.Client.ClientID
	Call page.ThisMember.QuickAdd(page.ProgramMember.ProgramID, outError)
End Sub

Function NoMembersForProgramDialogToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	Dim dialog				: Set dialog = New cDialog
	
	dialog.Headline = "Ok, nothing to see here!"
	
	dialog.Text = dialog.Text & "<p>"
	dialog.Text = dialog.Text & "It looks like there are no members assigned to the " & html(page.Program.ProgramName) & " program. "
	dialog.Text = dialog.Text & "To fix this, click <strong>Set up my members</strong> and start assigning members from your " & html(page.Client.NameClient) & " account to this program. "
	dialog.Text = dialog.Text & "</p>"

	dialog.SubText = dialog.SubText & "<p>Once you have some members assigned to the " & html(page.Program.ProgramName) & " program, "
	dialog.SubText = dialog.SubText & "this page will show you a filtered list of your " & html(page.Client.NameClient) & " account members who belong to this program. </p>"
	dialog.SubText = dialog.SubText & "<p>You can use that list to add, remove, or change the settings for members in the " & html(page.Program.ProgramName) & " program. </p>"

	pg.Action = CONFIGURE_PROGRAM_MEMBERS
	dialog.LinkList = dialog.LinkList & "<li class=""program-member-link""><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Set up my members</a></li>"
	pg.Action = ADDNEW_RECORD
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Add a new member</a></li>"
	pg.Action = "": pg.ProgramId = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Show all my account members</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/help/topic.asp?hid=14"" target=""_blank"">Learn about members and programs</a></li>"
	
	NoMembersForProgramDialogToString = dialog.ToString()
End Function

Function MemberGridToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim list				: list = page.Client.MemberList(page.Program.ProgramID, "")
	
	Dim count				: count = 0
	Dim altClass			: altClass = ""
	Dim enabledText			: enabledText = ""
	Dim msg
	
	Dim memberIcon			: memberIcon = "user_red.png"
	Dim entityText			: entityText = page.Client.NameClient
	Dim headerText			: headerText = "Account Members"
	If Len(page.Program.ProgramID) > 0 Then 
		memberIcon = "user.png"
		entityText = page.Program.ProgramName
		headerText = "Program Members"
	End If
	
	Dim iWantToBoxText
	iWantToBoxText = iWantToBoxText & "<div class=""tip-box""><h3>I want to .. </h3><ul>"
	pg.Action = ADDNEW_RECORD: pg.MemberId = "": pg.ProgramMemberId = ""
	If Len(page.Program.ProgramID) = 0 Then
		iWantToBoxText = iWantToBoxText & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Add a member account</a></li>"
		pg.Action = ""
		iWantToBoxText = iWantToBoxText & "<li><a href=""/admin/importmembers.asp" & pg.UrlParamsToString(True) & """>Add member accounts from a file (import)</a></li>"
		pg.Action = INVITE_MEMBERS
		iWantToBoxText = iWantToBoxText & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Add a member by email</a></li>"
	Else
		pg.Action = CONFIGURE_PROGRAM_MEMBERS
		iWantToBoxText = iWantToBoxText & "<li class=""program-member-tip-link""><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Add or remove members from this list</a></li>"
	End If
	iWantToBoxText = iWantToBoxText & "</ul></div>"
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-NameLogin 4-PWord 5-Email 
	' 17-IsProfileComplete 18-IsProfileUserCertified 19-IsApproved 20-ActiveStatus 26-IsAdmin

	str = str & "<h3>" & Server.HtmlEncode(page.Client.NameClient) & " Members</h3>"
	If Len(page.Program.ProgramId) > 0 Then
		str = str & "<h4 class=""first"">" & Server.HtmlEncode(page.Program.ProgramName) & "</h4>"
		str = str & "<p>This listing includes all of your " & Application.Value("APPLICATION_NAME") & " account members that have been assigned to the <strong>" & Server.HtmlEncode(page.Program.ProgramName) & "</strong> program. "
		str = str & "Click <strong>Change members</strong> in the toolbar to add or remove members from this list. </p>"
	Else
		str = str & "<h4 class=""first"">All programs</h4>"
		str = str & "<p>This listing includes all of your " & Application.Value("APPLICATION_NAME") & " account members. "
		str = str & "Select a program from the list in the toolbar to work with the members that belong to that program. </p>"
	End If
	str = str & iWantToBoxText
	str = str & m_appMessageText
	
	str = str & "<div class=""grid"">"
	pg.Action = BULK_REMOVE_CLIENT_MEMBERS
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-bulk-delete-members"" >"
	str = str & "<table>"
	str = str & "<tr class=""header""><th scope=""col""><input type=""checkbox"" class=""checkbox"" id=""master"" /></th>"
	str = str & "<th scope=""col"">" & headerText & "</th><th scope=""col"">Enabled</th><th scope=""col"">&nbsp;</th></tr>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			count = count + 1
			enabledText = "Yes"
			If list(20,i) = 0 Then enabledText = "<span style=""color:red;"">No</span>"
			altClass = ""
			If count Mod 2 = 0 Then altClass = " class=""alt"""
		
			str = str & "<tr" & altClass & "><td style=""width:1%;""><input type=""checkbox"" class=""checkbox"" name=""MemberIDList"" value=""" & list(0,i) & """ /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/" & memberIcon & """ alt="""" />"
			str = str & "<strong>" & html(entityText) & " | "
			pg.Action = SHOW_MEMBER_DETAILS: pg.MemberID = list(0,i)
			str = str & "<a href=""/admin/profile.asp" & pg.UrlParamsToString(True) & """>" & html(list(1,i) & ", " & list(2,i)) & "</a></strong></td>"
			str = str & "<td>" & enabledText & "</td>"
			
			str = str & "<td class=""toolbar"">"
			pg.Action = SHOW_MEMBER_DETAILS: pg.MemberID = list(0,i)
			str = str & "<a href=""/admin/profile.asp" & pg.UrlParamsToString(True) & """ title=""Details"">"
			str = str & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a>"		
			pg.Action = "": pg.MemberID = list(0,i)
			str = str & "<a href=""/admin/profile.asp" & pg.UrlParamsToString(True) & """ title=""Edit"">"
			str = str & "<img src=""/_images/icons/pencil.png"" alt="""" /></a>"		
			pg.Action = SEND_MESSAGE: pg.MemberID = list(0,i)
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email"">"
			str = str & "<img src=""/_images/icons/email.png"" alt="""" /></a>"	
			If Len(page.Program.ProgramID) > 0 Then
				pg.Action = REMOVE_PROGRAM_MEMBER: pg.MemberID = list(0,i)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove"">"
			Else 
				pg.Action = REMOVE_CLIENT_MEMBER: pg.MemberID = list(0,i)
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove"">"
			End If	
			str = str & "<img src=""/_images/icons/cross.png"" alt="""" /></a>"		
			str = str & "</td></tr>"
		Next
	End If
	str = str & "</table></form></div>"
	
	If count = 0 Then
		str = NoMembersForProgramDialogToString(page)	
	End If
	
	MemberGridToString = str
End Function

Function FormInviteMembersToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>List the email address of anyone you would like to receive a link to create their own <strong>" & html(page.Client.NameClient) & "</strong> " & Application.Value("APPLICATION_NAME") & " account. </p></div>"
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" id=""formInviteMember"">"
	str = str & "<table><tr><td class=""label"">" & RequiredElementToString(True, "Email Address(es)") & "</td>"
	str = str & "<td><input type=""text"" name=""EmailList"" value=""" & page.EmailList & """ class=""extra-large gets-focus"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">You can invite multiple people to create an account, just separate their addresses <br />with a comma.</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Note</td>"
	str = str & "<td><textarea name=""InviteNote"" class=""large"">" & page.InviteNote & "</textarea></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">This note (optional) will be included with your invitation. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Send"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormInviteMembersIsPostback"" value=""" & IS_POSTBACK & """ /></td></tr>"
	str = str & "</table></form></div>"
	
	FormInviteMembersToString = str
End Function

Function FormNewMemberToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & m_appMessageText
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" name=""formNewMember"">"
	str = str & "<input type=""hidden"" name=""FormNewMemberIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><tr><td class=""label"">" & RequiredElementToString(True, "First Name") & "</td>"
	str = str & "<td><input type=""text"" class=""gets-focus medium"" name=""FirstName"" value=""" & page.ThisMember.NameFirst & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Last Name") & "</td>"
	str = str & "<td><input type=""text"" name=""LastName"" class=""medium"" value=""" & page.ThisMember.NameLast & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Email") & "</td>"
	str = str & "<td><input type=""text"" name=""Email"" class=""medium"" value=""" & page.ThisMember.Email & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Retype Email") & "</td>"
	str = str & "<td><input type=""text"" name=""EmailRetype"" class=""medium"" value=""" & page.ThisMember.EmailRetype & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Program</td><td>" & ProgramDropdownToString(page.Member.MemberID, page.Program.ProgramID) & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">If a program is selected, your new member will also <br />be added to that program when you click save. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Save"" />"
	str = str & "&nbsp;<input type=""submit"" name=""Submit"" value=""" & SAVE_AND_NEW & """ />"
	pg.Action = "": pg.MemberID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a></td></tr>"
	str = str & "</table></form></div>"
		
	FormNewMemberToString = str
End Function

Function FormConfirmBulkRemoveClientMembersToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim msg
	
	Dim list			: list = Split(page.MemberIDList, ",")
	Dim memberCount		: memberCount = 0
	
	If IsArray(list) Then memberCount = UBound(list) + 1

	Dim memberCountText	: memberCountText = memberCount & " member"
	If memberCount <> 1 Then memberCountText = memberCountText & "s"	
	
	msg = msg & "You will permanently remove " & memberCountText & " from your " & Application.Value("APPLICATION_NAME") & " account. "
	msg = msg & "You will lose any program, schedule, and calendar information associated with these member accounts. "
	msg = msg & "This action cannot be reversed. "
	
	str = str & CustomApplicationMessageToString("Please confirm remove accounts!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""FormConfirmBulkRemoveClientMembers"">"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormConfirmBulkRemoveClientMembersIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<input type=""hidden"" name=""MemberIDList"" value=""" & page.MemberIDList & """ />"
	str = str & "</p></form>"
	
	FormConfirmBulkRemoveClientMembersToString = str
End Function

Function FormConfirmDeleteProgramMember(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "You will permanently remove the member <strong>" & html(page.ThisMember.NameLast & ", " & page.ThisMember.NameFirst) & "</strong> from the <strong>" & html(page.Program.ProgramName) & "</strong> program. "
	str = str & "You will lose any schedule or calendar information for this program that is associated with this member. "
	str = str & "This action cannot be reversed. "
	str = str & "<br /> <br /><strong>Note: </strong>You are not removing this member from your " & Application.Value("APPLICATION_NAME") & " account, only from the " & html(page.Program.ProgramName) & " program. "
	str = CustomApplicationMessageToString("Please confirm remove member from " & html(page.Program.ProgramName) & "!", str, "Confirm")
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" id=""formConfirmDeleteProgramMember"">"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = "": pg.MemberID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormConfirmDeleteProgramMemberIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</p></form>"
	
	FormConfirmDeleteProgramMember = str
End Function

Function LoadNewMemberFromRequest(member, programMember)
	member.NameLast = Request.Form("LastName")
	member.NameFirst = Request.Form("FirstName")
	member.Email = Request.Form("Email")
	member.EmailRetype = Request.Form("EmailRetype")
	
	programMember.ProgramID = Request.Form("ProgramID")
End Function

Function ValidNewMember(member)
	ValidNewMember = True
	
	If Not ValidData(member.NameFirst, True, 0, 50, "First Name", "") Then ValidNewMember = False
	If Not ValidData(member.NameLast, True, 0, 50, "Last Name", "") Then ValidNewMember = False
	If Not ValidData(member.Email, True, 0, 100, "Email Address", "email") Then ValidNewMember = False

	'check that email fields match
	If UCase(member.Email) <> UCase(member.EmailRetype) Then
		AddCustomFrmError("Email and Retype Email must match exactly.")
		ValidNewMember = False
	End If	
End Function

Function ValidFormInvite(page)
	Dim i
	Dim hasBadAddress		: hasBadAddress = False
	ValidFormInvite = True
	
	' no email addresses supplied
	If Len(page.EmailList) = 0 Then
		AddCustomFrmError("At least one email address is required. ")
		ValidFormInvite = False
	End If
	
	Dim list		: list = Split(page.EmailList, ",")
	If IsArray(list) Then
		For i = 0 To UBound(list)
			If Not IsEmail(list(i)) Then
				AddCustomFrmError("An email address (" & html(list(i)) & ") does not appear to be in the correct format. ")
				hasBadAddress = True
			End If
		Next
		If hasBadAddress Then ValidFormInvite = False
	End If
End Function

Function ProgramDropdownToString(memberID, programID)
	Dim str, i
	Dim list				: list = GetProgramList(memberID)
	Dim selected			: selected = ""
	
	Dim disabled			: disabled = ""
	Dim disabledClass		: disabledClass = ""
	If Not IsArray(list) Then 
		disabled = " disabled=""disabled"""
		disabledClass = " class=""disabled"""
	End If
	
	
	str = str & "<select name=""ProgramID""" & disabled & disabledClass & ">"
	If IsArray(list) Then
		str = str & "<option value="""">&nbsp;</option>"
		For i = 0 To UBound(list,2)
			selected = ""
			If CStr(list(0,i)) = CStr(programID) Then selected = " selected=""selected"""
			
			str = str & "<option value=""" & list(0,i) & """" & selected & ">" & html(list(1,i)) & "</option>"
		Next
	Else
		str = str & "<option value="""">" & html("< No programs available >") & "</select>"
	End If
	str = str & "</select>"
	
	ProgramDropdownToString = str
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

Function GetProgramsOwned(memberID) 'returns array
	Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs			: Set rs = Server.CreateObject("ADODB.Recordset")
	
	cnn.Open Application.Value("CNN_STR")
	cnn.up_memberGetProgramsOwnedByMemberID CLng(memberID), rs
	If Not rs.EOF then GetProgramsOwned = rs.GetRows()

	rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
End Function

Sub SendNewMemberLogin(newMember, member)
	Dim loggedInMember
	
	Call SendCredentials(newMember.MemberID, member.Email)
End Sub

Sub SendNewMemberWelcome(newMember, member)
	Dim body, subject
	Dim email			: Set email = New cEmailSender
	
	Call newMember.Load()
	
	subject = "** [" & Application.Value("APPLICATION_NAME") & "] " & newMember.ClientName & " Account Information **"

	body = "Dear " & newMember.NameFirst & " " & newMember.NameLast & ":" & vbCrLf & vbCrLf
	body = body & "Welcome to " & Application.Value("APPLICATION_NAME") & ". " 
	body = body & "Your login credentials for " & newMember.ClientName & " have been sent to you in a separate email message. "
	body = body & "You may wish to print and save this message for future reference. "
	body = body & "Upon logging in, you will be asked to complete your personal profile to activate your account. " 
	body = body & "At that time you may change your login name and password to a combination that will be easier for you to manage. " 
	body = body & "Please do not reply directly to this email. " & vbCrLf & vbCrLf

	body = body & "To login for the first time, please go to " & vbCrLf
	body = body & "http://" & Request.ServerVariables("SERVER_NAME") & "/member/login.asp" & vbCrLf & vbCrLf
	body = body & "For " & Application.Value("APPLICATION_NAME") & " help go to " & vbCrLf
	body = body & "http://" & Request.ServerVariables("SERVER_NAME") & "/help/help.asp" & vbCrLf & vbCrLf
	body = body & Application.Value("APPLICATION_NAME") & " support and assistance may be found here " & vbCrLf
	body = body & "http://" & Request.ServerVariables("SERVER_NAME") & "/support.asp" & vbCrLf & vbCrLf 
	body = body & Application.Value("APPLICATION_NAME") & " Technical Staff" & vbCrLf & "Connect. Schedule. Inspire."
	
	body = body & EmailDisclaimerToString(newMember.ClientName)
	Call email.SendMessage(newMember.Email, member.Email, subject, body)
End Sub

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim programMembersLink
	pg.Action = ""
	programMembersLink = programMembersLink & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.Program.ProgramName) & "</a> / "
	
	Dim clientMembersLink
	pg.ProgramID = "": pg.Action = ""
	clientMembersLink = clientMembersLink & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Members</a> / "
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case REMOVE_PROGRAM_MEMBER
			str = str & clientMembersLink
			str = str & programMembersLink
			str = str & "Remove program member"
		Case REMOVE_CLIENT_MEMBER
			str = str & clientMembersLink
			str = str & "Remove account"
			
		Case CONFIGURE_PROGRAM_MEMBERS
			str = str & clientMembersLink
			str = str & programMembersLink
			str = str & "Change Members"
			
		Case ADDNEW_RECORD
			str = str & clientMembersLink
			str = str & "Add member account"
			
		Case Else
			If Len(page.Program.ProgramID) > 0 Then
				str = str & clientMembersLink
				str = str & html(page.Program.ProgramName)
			Else
				str = str & "Members"
			End If
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href		: href = ""
	
	Dim bulkDeleteButton
	bulkDeleteButton = "<li><a href=""#"" class=""bulk-delete-member""><img class=""icon"" src=""/_images/icons/user_red_delete.png"" alt="""" /></a><a href=""#"" class=""bulk-delete-member"">Remove Selected</a></li>"
	
	Dim inviteMemberButton: pg.MemberId = "": pg.ProgramMemberId = ""
	pg.Action = INVITE_MEMBERS
	href = pg.Url & pg.UrlParamsToString(True)
	inviteMemberButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/user_red_email.png"" alt="""" /></a><a href=""" & href & """>Invite</a></li>"
	
	Dim newMemberButton
	pg.Action = ADDNEW_RECORD: pg.MemberId = "": pg.ProgramMemberId = ""
	href = pg.Url & pg.UrlParamsToString(True)
	newMemberButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/user_red_add.png"" alt="""" /></a><a href=""" & href & """>New</a></li>"
	
	Dim importMemberButton
	pg.Action = ""
	href = "/admin/importmembers.asp" & pg.UrlParamsToString(True)
	importMemberButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/user_red_go_down.png"" alt="""" /></a><a href=""" & href & """>Import</a></li>"
	
	Dim programMemberButton
	pg.Action = CONFIGURE_PROGRAM_MEMBERS
	href = pg.Url & pg.UrlParamsToString(True)
	programMemberButton = "<li class=""program-member-button"" id=""pid-" & page.Program.ProgramId & """><a href=""" & href & """><img class=""icon"" src=""/_images/icons/user.png"" alt="""" /></a><a href=""" & href & """>Change members</a></li>"
	
	Dim memberListButton
	Dim userIcon				: userIcon = "user_red.png"
	If Len(page.ProgramId) > 0 Then userIcon = "user.png"
	pg.Action = "": pg.MemberID = ""
	href = pg.Url & pg.UrlParamsToString(True)
	memberListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/" & userIcon & """ alt="""" /></a><a href=""" & href & """>Member List</a></li>"
	
	Select Case page.Action
		Case CONFIGURE_PROGRAM_MEMBERS
			str = str & memberListButton
			
		Case INVITE_MEMBERS
			str = str & memberListButton
			
		Case ADDNEW_RECORD
			str = str & memberListButton
			
		Case DELETE_RECORD
			str = str & memberListButton
			
		Case REMOVE_PROGRAM_MEMBER
			str = str & memberListButton
			
		Case REMOVE_CLIENT_MEMBER
			str = str & memberListButton
			
		Case BULK_REMOVE_CLIENT_MEMBERS
			str = str & memberListButton
			
		Case SEND_MESSAGE
		
		Case Else
			str = str & FormGotoMemberDropdownToString(page)
			str = str & FormProgramDropdownToString(page)
			If Len(page.Program.ProgramID) > 0 Then
				str = str & programMemberButton
			Else
				str = str & bulkDeleteButton
				str = str & inviteMemberButton
				str = str & importMemberButton
				str = str & newMemberButton
			End If
			
	End Select

	m_tabLinkBarText = str
End Sub

Function FormProgramDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = GetProgramsOwned(page.Member.MemberID)
	Dim isSelected		: isSelected = ""
	
	Dim defaultText		: defaultText = "< Select a program >"
	If Len(page.Program.ProgramID) > 0 Then defaultText = "< Show all members >"
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-go-to-program"">"
	str = str & "<input type=""hidden"" name=""FormProgramDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""NewProgramID"" id=""go-to-program-dropdown"">"
	str = str & "<option value="""">" & Html(defaultText) & "</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isSelected = ""
			If CStr(list(0,i)) = CStr(page.Program.ProgramID) Then isSelected = " selected=""selected"""
			
			str = str & "<option value=""" & list(0,i) & """" & isSelected & ">" & html(list(1,i)) & "</option>"
		Next
	End If
	str = str & "</select></form></li>"
	
	FormProgramDropdownToString = str
End Function

Function FormGotoMemberDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = page.Client.MemberList(page.Program.ProgramID, "")
	
	pg.Action = GOTO_MEMBER_PROFILE
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-go-to-member"">"
	str = str & "<input type=""hidden"" name=""FormGotoMemberIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""NewMemberID"" id=""go-to-member-dropdown"">"
	str = str & "<option value="""">" & html("< Go to member >") & "</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			str = str & "<option value=""" & list(0,i) & """>" & html(list(1,i) & ", " & list(2,i)) & "</option>"
		Next
	End If
	str = str & "</select></form></li>"
	
	FormGotoMemberDropdownToString = str
End Function

%>
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_DoDeleteClientMember.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_OwnsMember.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public ProgramID
	Public MemberID
	Public ProgramMemberID
	Public EmailID
	
	' objects
	Public Member
	Public Client
	Public Program
	Public ProgramMember
	Public ThisMember
	
	' not persisted
	Public MemberIDList
	Public ProgramMemberIDList	
	Public EmailList
	Public InviteNote
	
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
		If Len(MemberID) > 0 Then str = str & "mid=" & Encrypt(MemberID) & amp
		If Len(ProgramMemberID) > 0 Then str = str & "pmid=" & Encrypt(ProgramMemberID) & amp
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
		c.MemberID = MemberID
		c.ProgramID = ProgramID
		c.ProgramMemberID = ProgramMemberID
		c.EmailID = EmailID
		
		c.MemberIDList = MemberIDList
		c.ProgramMemberIDList = ProgramMemberIDList
		c.EmailList = EmailList
		c.InviteNote = InviteNote
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.ProgramMember = ProgramMember
		Set c.ThisMember = ThisMember
		
		Set Clone = c
	End Function
End Class
%>

