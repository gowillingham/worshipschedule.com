<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Const CONFIRM_MESSAGE = "600"
Dim m_bodyText
Dim m_appMessageText

Sub OnPageLoad(ByRef page)
	page.MessageID = Request.QueryString("msgid")
	page.Action = Decrypt(Request.QueryString("act"))
	page.ClientGuid = Request.QueryString("gid")
	
	Set page.Client = New cClient
	If Len(page.ClientGuid) > 0 Then Call page.Client.LoadByGuid(page.ClientGuid)
	
	Set page.Member = New cMember
	page.Member.NameFirst = Request.Form("NameFirst")
	page.Member.NameLast = Request.Form("NameLast")
	page.Member.Email = Request.Form("Email")
	page.Member.EmailRetype = Request.Form("EmailRetype")
	page.HasTerms = Request.Form("HasTerms")
	
	m_appMessageText = "<div style=""width:540px;"">" & ApplicationMessageToString(page.MessageID) & "</div>"
	page.MessageID = ""
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/outside.css" />
		<style type="text/css">
			.form, .message {width:535px;}
		</style>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=Application.Value("APPLICATION_NAME") & " - Simple Web Scheduling for Volunteer Teams" %></title>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><a href="/">Home</a> / New Member Invitation</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="column">
					<div class="feature-box">
						<h3>More Info ..</h3>
						<ul>
							<li><a href="/requirements.asp"><strong>System Requirements</strong></a></li>
							<li><a href="/privacy.asp"><strong>Privacy</strong></a></li>
							<li><a href="/about.asp"><strong>About Us</strong></a></li>
						</ul>
					</div>
					<img style="" src="_images/hand_at.jpg" alt="Crazy Woman" />
				</div>
				<%=m_bodyText %>
			</div>
		</div>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_footer.asp"-->
	</body>
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)

	Select Case page.Action
		Case CONFIRM_MESSAGE
			str = str & ConfirmMessageToString(page)
			
		Case ADDNEW_RECORD
			If ValidFormNewMember(page) Then
				Call DoInsertNewMember(page, rv)
				Select Case rv
					Case 0
						Call SendCredentials(page.Member.MemberID, Application.Value("NO_REPLY_EMAIL_ADDRESS"))
						Call DoSendNewMemberConfirmation(page)
						page.Action = CONFIRM_MESSAGE
					Case -2
						' dupe member
						page.MessageID = 1067: page.Action = ""
					Case Else
						page.MessageID = 1068: page.Action = ""
				End Select
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormNewMemberToString(page)
			End If
		Case Else
			If Len(page.ClientGuid) = 0 Then
				Response.Redirect("/")
			End If
			
			str = str & FormNewMemberToString(page)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoInsertNewMember(page, outError)
	page.Member.ClientID = page.Client.ClientID
	Call page.Member.Add(outError)
	
	If outError = 0 Then Call page.Member.Load()
End Sub

Sub DoSendNewMemberConfirmation(page)
	Dim str
	Dim email			: Set email = New cEmailSender
	Dim fromAddress		: fromAddress = Application.Value("NO_REPLY_EMAIL_ADDRESS")
	Dim appName			: appName = Application.Value("APPLICATION_NAME")
	Dim subject			: subject = "[" & appName & "] ** " & HTML(page.Client.NameClient) & " - New Member Welcome **"
	
	str = str & "Dear " & page.member.NameFirst & " " & page.member.NameLast & ":" & vbCrLf
	str = str & vbCrLf &  "Thank you for your " & appName & " account registration for " & page.Client.NameClient & " and welcome to " & appName & ". "
	str = str & "You should receive your temporary login credentials (login/password) in a separate email message. "
	str = str & "You should change your login name and password to something easier for you to remember - it can be done at any time by editing your " & appName & " member profile. "
	str = str & vbCrLf & vbCrLf
	str = str & "Please save or print out this email and store it in a secure place. "
	str = str & "Do not reply to this email directly - this address is not monitored. "
	
	str = str & vbCrLf & vbCrLf
	str = str & "If you experience any problems or have questions about your account, please contact " & appName & " support at "
	str = str & "mailto:" & Application.Value("SUPPORT_EMAIL_ADDRESS") & ". " 
	str = str & vbCrLf & vbCrLf
	str = str & "Additional support information may be found on the " & appName & " support page here .."
	str = str & vbCrLF
	str = str & "http://" & Request.ServerVariables("SERVER_NAME") & "/support.asp"
	str = str & vbCrLf & vbCrLf
	str = str & appName & " Help is available here .."
	str = str & vbCrLF
	str = str & "http://" & Request.ServerVariables("SERVER_NAME") & "/help/help.asp"
	str = str & vbCrLf & vbCrLf
	str = str & "To get started, login here .."
	str = str & vbCrLf
	str = str & "http://" & Request.ServerVariables("SERVER_NAME") & "/member/login.asp"
	str = str & vbCrLf & vbCrLf
	str = str & "Upon first login, you will be asked to take a few moments to review your member profile. "
	str = str & "Thanks again for your interest in " & appName & ", and we look forward to working with you!"
	str = str & vbCrLf & vbCrLf
	str = str & "The " & appName & " Team"
	str = str & vbCrLf 
	str = str & "Schedule. Connect. Inspire."
	
	str = str & EmailDisclaimerToString(page.Client.NameClient)
	
	Set email = New cEmailSender
	Call email.SendMessage(page.member.Email, fromAddress, subject, str)
End Sub

Function ValidFormNewMember(page)
	ValidFormNewMember = True
	
	If Not ValidData(page.Member.NameFirst, True, 0, 50, "First Name", "") Then ValidFormNewMember = False
	If Not ValidData(page.Member.NameLast, True, 0, 50, "Last Name", "") Then ValidFormNewMember = False
	If Not ValidData(page.Member.Email, True, 0, 100, "Email Address", "email") Then ValidFormNewMember = False

	'check that email fields match
	If UCase(page.Member.Email) <> UCase(page.Member.EmailRetype) Then
		AddCustomFrmError("The email fields must match.")
		ValidFormNewMember = False
	End If	
	
	'check that terms have been read
	If Len(page.HasTerms) = 0 Then
		AddCustomFrmError("You need to have read the Terms of Service (or at least check the box!).")
		ValidFormNewMember = False
	End If
End Function

Function FormNewMemberToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim isChecked		: isChecked = ""
	If Len(page.HasTerms) > 0 Then isChecked = " checked=""checked"""
	
	str = str & "<h1>" & html(page.Client.NameClient) & " - " & Application.Value("APPLICATION_NAME") & " New Member Invitation</h1>"
	str = str & "<p>Thanks for your interest in " & Application.Value("APPLICATION_NAME") & "! "
	str = str & "We just need the following information to set up a member account for you with " & html(page.Client.NameClient) & ". "
	str = str & "Complete the form and follow the prompts to begin realizing the benefits of managing your volunteer schedule online with " & Application.Value("APPLICATION_NAME") & ". </p>"
	
	str = str & m_appMessageText
	str = str & "<div class=""form"">"
	pg.Action = ADDNEW_RECORD
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formNewMember"">"
	str = str & ErrorToString()
	str = str & "<table>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "First name") & "</td>"
	str = str & "<td><input type=""text"" name=""NameFirst"" value=""" & page.Member.NameFirst & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Last name") & "</td>"
	str = str & "<td><input type=""text"" name=""NameLast"" value=""" & page.Member.NameLast & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Email") & "</td>"
	str = str & "<td><input type=""text"" name=""Email"" value=""" & page.Member.Email & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Retype Email") & "</td>"
	str = str & "<td><input type=""text"" name=""EmailRetype"" value=""" & page.Member.EmailRetype & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>"
	str = str & "<input class=""checkbox"" type=""checkbox"" name=""HasTerms""" & isChecked & " />"
	str = str & "I have read and agree to the <a href=""/terms.asp"" target=""_blank""><strong>Terms of Service</strong></a>.</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Save"" />"
	str = str & "</td></tr>"
	str = str & "</table></form></div>"
	
	FormNewMemberToString = str
End Function

Function ConfirmMessageToString(page)
	Dim str
	
	str = str & "<h1>Thanks - Your account has been created! </h1>"
	str = str & "<p>You should receive your login information by email in a few minutes "
	str = str & "(if you use any spam filtering software, make sure it is configured to allow email from the domain '@" & Application.Value("ROOT_EMAIL_DOMAIN") & "'). "
	str = str & "Once you receive that email, you may go <a href=""/member/login.asp"">here</a> to login. </p>"
	str = str & "<p>If you do not receive your login information in a timely manner, or you experience any other problems creating your account, "
	str = str & "please contact " & Application.Value("APPLICATION_NAME") & " <a href=""" & Application.Value("SUPPORT_URL") & """>support</a>. "
	str = str & "We'll try to resolve any problems you are having right away. </p>"
	str = str & "<p>Thanks again for using " & Application.Value("APPLICATION_NAME") & "! "
	str = str & "Please send us a <a href=""" & Application.Value("SUPPORT_URL") & """>note</a> with any of your suggestions or complaints. "
	str = str & "We're working to make " & Application.Value("APPLICATION_NAME") & " the web service to best meet your scheduling and collaboration needs. </p>"
	
	ConfirmMessageToString = str
End Function

%>
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<%
Class cPage
	Public MessageID
	Public Action
	Public ClientGuid

	' obj
	Public Client
	Public Member
	
	' not persisted
	Public HasTerms
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & Encrypt(MessageID) & amp
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ClientGuid) > 0 Then str = str & "gid=" & ClientGuid & amp
		
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
		c.ClientGuid = ClientGuid
		
		Set c.Client = Client
		Set c.Member = Member
		
		Set Clone = c
	End Function
End Class
%>

