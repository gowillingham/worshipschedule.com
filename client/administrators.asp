<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-settings"
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
	Call CheckSession(sess, PERMIT_ADMIN)
	
	page.MessageID = Request.QueryString("msgid")                
	
	page.Action = Decrypt(Request.QueryString("act"))
	page.ClientAdminId = Decrypt(Request.QueryString("caid"))
	page.EmailMemberId = Decrypt(Request.QueryString("emmid"))
	
	page.admin_member_id = Request.Form("member_id")

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
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<style type="text/css">
			.message, .form {width:622px;}
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
	
	Select Case page.Action
		Case DELETE_RECORD
			Call DoDeleteAdministrator(page.ClientAdminId, page.Member.MemberId, rv)
			Select Case rv
				Case 0
					page.MessageId = 2031
				Case -3
					' can't remove self from admin role ..
					page.MessageId = 1061
				Case Else
					page.MessageId = 2032
			End Select
			page.Action = "": page.ClientAdminId = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))

		Case ADDNEW_RECORD
			If Request.Form("form_administrator_is_postback") Then
				If ValidFormAdministrator(page) Then
					Call DoInsertAdministrator(page.admin_member_id, page.Client.ClientId, rv)
					Select Case rv
						Case 0
							page.MessageId = 2011
						Case Else
							page.MessageId = 2012
					End Select
					page.Action = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormAdministratorToString(page)
				End If
			Else
				str = str & FormAdministratorToString(page)
			End If
		
		Case SEND_MESSAGE
			Call DoInsertEmail(page, rv)
			page.Action = "": page.EmailMemberId = "": page.ClientAdminId = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
						
		Case Else
			str = str & AdministratorGridToString(page)
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoInsertEmail(page, outError)
	Dim member			: Set member = New cMember
	Dim email			: Set email = New cEmail
	
	member.MemberId = page.EmailMemberId
	Call member.Load()
	
	email.MemberId = page.Member.MemberId
	email.ClientId = page.Member.ClientId
	email.RecipientAddressList = member.Email
	Call email.Add(outError)
	
	page.EmailId = email.EmailId
End Sub

Sub DoDeleteAdministrator(clientAdminId, memberId, outError)
	Dim clientAdmin			: Set clientAdmin = New cClientAdmin
	
	clientAdmin.ClientAdminId = clientAdminId
	Call clientAdmin.Load()
	
	' can't remove self from admin role ..
	If CStr(memberId) = CStr(clientAdmin.MemberId) Then
		outError = -3
		Exit Sub
	End If
	
	Call clientAdmin.Delete(outError)
End Sub

Sub DoInsertAdministrator(memberId, clientId, outError)
	Dim clientAdmin			: Set clientAdmin = New cClientAdmin
	
	clientAdmin.ClientId = clientId
	clientAdmin.MemberId = memberId
	Call clientAdmin.Add(outError)
End Sub

Function MemberDropdownOptionsToString(list)
	Dim str, i
	
	Dim isAdmin				: isAdmin = False
	Dim isEnabled			: isEnabled = False
	
	str = str & "<option value="""">" & html("< Select a member >") & "</option>"
	For i = 0 To Ubound(list,2)
		isAdmin = False: If list(26,i) = 1 Then isAdmin = True
		isEnabled = False: If list(20,i) = 1 Then isEnabled = True
		
		If (Not isAdmin) And isEnabled Then
			str = str & "<option value=""" & list(0,i) & """>" & html(list(1,i) & ", " & list(2,i)) & "</option>"
		End If
	Next
	
	MemberDropdownOptionsToString = str
End Function

Function FormAdministratorToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>Administrators have full control over your " & Application.Value("APPLICATION_NAME") & " account. </p></div>"
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-administrator"">"
	str = str & "<input type=""hidden"" name=""form_administrator_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table>"
	str = str & "<tr><td class=""label"">Account Member</td>"
	str = str & "<td><select name=""member_id"">" & MemberDropdownOptionsToString(page.Client.MemberList("", "")) & "</select></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">Select a member from your account to add to your list <br />of administrators. "
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a></td></tr>"
	str = str & "</table></form></div>"
	
	FormAdministratorToString = str
End Function

Function ValidFormAdministrator(page)
	ValidFormAdministrator = True
	
	If Len(page.admin_member_id) = 0 Then
		AddCustomFrmError("You didn't select anyone from the list. Please select a member. ")
		ValidFormAdministrator = False
	End If
End Function

Function AdministratorGridToString(page)
	Dim str, i
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>Members in this list have full control over your " & Application.Value("APPLICATION_NAME") & " account, "
	str = str & "and can add, remove, or change any of your account's programs or members. </p></div>"
	
	Dim clientAdmin	: Set clientAdmin = New cClientAdmin
	clientAdmin.ClientId = page.Client.ClientId
	Dim list		: list = clientAdmin.List()
	
	' 0-ClientID 1-MemberID 2-NameFirst 3-NameLast 4-NameClient 5-DateCreated 
	' 6-DateModified 7-ClientAdminID
		
	str = str & m_appMessageText
	str = str & "<div class=""grid"">"
	str = str & "<table><thead><tr>"
	str = str & "<th style=""width:1%;""><input type=""checkbox"" name=""master"" checked=""checked"" class=""checked"" disabled=""disabled"" /></th>"
	str = str & "<th>Member</th><th>Somthing</th><th>&nbsp;</th>"
	str = str & "</tr></thead>"
	str = str & "<tbody>"
	For i = 0 To UBound(list,2)
		str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""member_id"" value="""" checked=""checked"" disabled=""disabled"" /></td>"
		str = str & "<td><img class=""icon"" src=""/_images/icons/medal_gold_2.png"" alt="""" />" 
		str = str & "<strong>" & html(list(3,i) & ", " & list(2,i)) & "</strong></td>"
		str = str & "<td>" & list(5,i) & "</td>"
		str = str & "<td class=""toolbar"">"
		pg.EmailMemberId = list(1,i): pg.Action = SEND_MESSAGE
		str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email""><img src=""/_images/icons/email.png"" alt="""" /></a>"
		pg.Action = DELETE_RECORD: pg.ClientAdminId = list(7,i)
		str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove""><img src=""/_images/icons/cross.png"" alt="""" /></a>"
		str = str & "</td></tr>"
	Next	
	str = str & "</tbody></table>"
	str = str & "</div>"

	AdministratorGridToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim preferencesLink
	preferencesLink = "<a href=""/client/preferences.asp"">Preferences</a> / "
	
	Dim administratorsLink
	pg.Action = ""
	administratorsLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Administrators</a> / "

	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case ADDNEW_RECORD
			str = str & preferencesLink & administratorsLink & "Add Administrator"
		Case Else
			str = str & preferencesLink & "Administrators"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Dim newAdministratorButton
	pg.Action = ADDNEW_RECORD
	href = pg.Url & pg.UrlParamsToString(True)
	newAdministratorButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/medal_gold_2_add.png""  alt="""" /></a><a href=""" & href & """>Add Administrator</a></li>"
	
	Dim administratorsButton
	pg.Action = ""
	href = pg.Url & pg.UrlParamsToString(True)
	administratorsButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/medal_gold_2.png""  alt="""" /></a><a href=""" & href & """>Administrators</a></li>"

	Select Case page.Action
		Case DELETE_RECORD
			str = str & administratorsButton
		Case ADDNEW_RECORD
			str = str & administratorsButton
		Case Else
			str = str & newAdministratorButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_admin_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID

	' encrypted
	Public Action
	Public ClientAdminId
	Public EmailMemberId
	Public EmailId
	
	' form data
	Public admin_member_id

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
		If Len(ClientAdminId) > 0 Then str = str & "caid=" & Encrypt(ClientAdminId) & amp
		If Len(EmailId) > 0 Then str = str & "emid=" & Encrypt(EmailId) & amp
		If Len(EmailMemberId) > 0 Then str = str & "emmid=" & Encrypt(EmailMemberId) & amp
		
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
		c.ClientAdminId = ClientAdminId
		c.EmailId = EmailId
		c.EmailMemberId = EmailMemberId

		Set c.Member = Member
		Set c.Client = Client

		Set Clone = c
	End Function
End Class
%>

