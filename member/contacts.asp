<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "contacts"
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
	page.RecipientID = Decrypt(Request.QueryString("rid"))
	
	page.Subject = Request.Form("Subject")
	page.Message = Request.Form("Message")
	
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
		<title><%=m_pageTitleText%></title>
		<style type="text/css">
			.form, .message	{width:622px;}
		</style>
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
			If Request.Form("FormEmailIsPostback") = IS_POSTBACK Then
				If ValidFormEmail(page.Subject, page.Message) Then
					Call DoSendMessage(page)
					' success
					page.MessageID = 7009: page.Action = "": page.RecipientID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormEmailToString(page)
				End If
			Else
				str = str & FormEmailToString(page)
			End If
			
		Case Else
			str = str & ContactListsToString(page)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoSendMessage(page)
	Dim str
	Dim email		: Set email = New cEmailSender
	
	Dim recipient	: Set recipient = New cMember
	recipient.MemberID = page.RecipientID
	recipient.Load()
	
	Call email.SendMessage(recipient.Email, page.member.Email, page.Subject, page.Message & EmailDisclaimerToString(page.Client.NameClient))
	Set email = Nothing
End Sub

Function FormEmailToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim recipient		: Set recipient = New cMember
	recipient.MemberID = page.RecipientID
	recipient.Load()
	
	str = str & "<div class=""form"">"
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formEmail"">"
	str = str & "<input type=""hidden"" name=""FormEmailIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & ErrorToString()
	str = str & "<table>"
	str = str & "<tr><td>&nbsp;<td><h3>Message for " & html(recipient.Email) & "</h3></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Subject</td>"
	str = str & "<td><strong><input type=""text"" name=""Subject"" value=""" & "" & """ class=""large"" /></strong></td></tr>"
	str = str & "<tr><td class=""label"">Message</td>"
	str = str & "<td><textarea name=""Message"" class=""large"" style=""height:125px;"">" & "" & "</textarea></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Send"" />"
	pg.Action = "": pg.RecipientID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a></td></tr>"
	str = str & "</table></form></div>"
	
	FormEmailToString = str
End Function

Function ValidFormEmail(subject, message)
	ValidFormEmail = True
	
	If Not ValidData(subject, False, 0, 200, "Subject", "") Then ValidFormEmail = False
	If Not ValidData(message, False, 0, 4000, "Message", "") Then ValidFormEmail = False
	If Len(subject & message) = 0 Then
		AddCustomFrmError("Your message must have either a subject or a message. They cannot both be blank.")
		ValidFormEmail = False
	End If
End Function

Function ContactListsToString(page)
	Dim str

	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "<p>Having trouble? "
	str = str & "Click <strong>Email</strong> to contact an administrator or program leader for " & html(page.Client.NameClient) & ". </p></div>"
	
	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "<p>Problem with " & Application.Value("APPLICATION_NAME") & "? "
	str = str & "Contact <a href=""/support.asp"">support</a> to get help right away. </p></div>"
	
	str = str & m_appMessageText
	str = str & "<h3 style=""margin-top:0;padding-top:0;"">" & html(page.Client.NameClient) & " Account Administrators</h3>"
	str = str & AdminGridToString(page)
	str = str & "<h3>Hosanna! Program Leaders</h3>"
	str = str & LeaderListToString(page)
	
	ContactListsToString = str
End Function

Function AdminGridToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = GetAdminList(page.Client.ClientID)
	Dim altClass		: altClass = ""
	
	str = str & "<div class=""grid""><table><tr><th scope=""col"" style=""width:1%;"">"
	str = str & "<input type=""checkbox"" name=""master"" disabled=""disabled"" /></th>"
	str = str & "<th scope=""col"">Account Administrators</th><th scope=""col""></th><th scope=""col""></th></tr>"
	For i = 0 To UBound(list,2)
		altClass = ""
		If i Mod 2 <> 0 Then altClass = " class=""alt"""

		str = str & "<tr" & altClass & ">"
		str = str & "<td><input name=""None"" disabled=""disabled"" type=""checkbox"" /></td>"
		str = str & "<td><img class=""icon"" src=""/_images/icons/user_red.png"" alt="""" />"
		str = str & "<div class=""data"">"
		str = str & "<strong>" & html(page.Client.NameClient)	 & " | " 
		pg.Action = SEND_MESSAGE: pg.RecipientID = list(0,i)
		str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(2,i) & ", " & list(1,i))
		str = str & "</a>"
		str = str & "</strong></div></td>"
		str = str & "<td>" & html(list(3,i)) & "</td>"
		str = str & "<td class=""toolbar"">"
		str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email"">"
		str = str & "<img src=""/_images/icons/email.png"" alt="""" /></a></td></tr>"
	Next	
	str = str & "</table></div>"
	
	AdminGridToString = str 
End Function

Function LeaderListToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = GetLeaderList(page.Member.MemberID)
	Dim altClass		: altClass = ""
	
	If Not IsArray(list) Then Exit Function
	
	str = str & "<div class=""grid""><table><tr><th scope=""col"" style=""width:1%;"">"
	str = str & "<input type=""checkbox"" name=""master"" disabled=""disabled"" /></th>"
	str = str & "<th scope=""col"">Program Leaders</th><th scope=""col""></th><th scope=""col""></th></tr>"
	For i = 0 To UBound(list,2)
		altClass = ""
		If i Mod 2 <> 0 Then altClass = " class=""alt"""

		str = str & "<tr" & altClass & ">"
		str = str & "<td><input name=""None"" disabled=""disabled"" type=""checkbox"" /></td>"
		str = str & "<td><img class=""icon"" src=""/_images/icons/user.png"" alt="""" />"
		str = str & "<div class=""data"">"
		str = str & "<strong>" & html(list(4,i)) & " | " 
		pg.Action = SEND_MESSAGE: pg.RecipientID = list(0,i)
		str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(2,i) & ", " & list(1,i))
		str = str & "</a>"
		str = str & "</strong></div></td>"
		str = str & "<td>" & html(list(3,i)) & "</td>"
		str = str & "<td class=""toolbar"">"
		str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email"">"
		str = str & "<img src=""/_images/icons/email.png"" alt="""" /></a></td></tr>"
	Next	
	str = str & "</table></div>"
	
	LeaderListToString = str
End Function

Function GetAdminList(clientID)
	Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs			: Set rs = Server.CreateObject("ADODB.Recordset")
	
	cnn.Open Application.Value("CNN_STR")
	cnn.up_clientGetAdminContactList CLng(clientID), rs
	If Not rs.EOF Then GetAdminList = rs.GetRows()
	
	If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
	Set cnn = Nothing
End Function

Function GetLeaderList(memberID)
	Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs			: Set rs = Server.CreateObject("ADODB.Recordset")
	
	cnn.Open Application.Value("CNN_STR")
	cnn.up_clientGetLeaderListForMemberID CLng(memberID), rs
	If Not rs.EOF Then GetLeaderList = rs.GetRows()
	
	If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
	Set cnn = Nothing
End Function

Sub SetPageHeader(page)
	Dim str

	Dim accountContactLink
	accountContactLink = accountContactLink & "<a href=""" & page.Url & page.UrlParamsToString(True) & """>Account Contacts</a>"
	
	str = str & "<a href=""/member/overview.asp"">Member Home</a> / "
	str = str & "<a href=""/member/settings.asp"">Settings</a> / "

	Select Case page.Action
		Case SEND_MESSAGE
			str = str & accountContactLink & " / "
			str = str & "Send Message"
		Case Else
			str = str & "Account Info"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	
	str = str & "<li>&nbsp;</li>"
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public RecipientID
	
	' objects
	Public Member
	Public Client	
	
	'not for url
	Public Subject
	Public Message
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(RecipientID) > 0 Then str = str & "rid=" & Encrypt(RecipientID) & amp
		
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
		c.RecipientID = RecipientID
		
		c.Subject = Subject
		c.Message = Message
		
		Set c.Member = Member
		Set c.Client = Client
				
		Set Clone = c
	End Function
End Class
%>

