<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%

Dim m_bodyText

Sub OnPageLoad(ByRef page)
	Dim sess			: Set sess = New cSession
	sess.SessionID = Request.Cookies("sid")
	
	Call CheckSession(sess, PERMIT_ALL)
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	If Len(page.Client.ClientID) > 0 Then page.Client.Load()

	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	If Len(page.Member.MemberID) > 0 Then page.Member.Load()

	page.Action = DeCrypt(Request.QueryString("act"))
	page.MessageID = Request.QueryString("msgid")
	
	page.Member.NameLogin = Request.Form("NameLogin")
	page.Member.PWord = Request.Form("PWord")
	
	Set sess = Nothing
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/outside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=Application.Value("APPLICATION_NAME") & " - Simple Web Scheduling for Worship Teams" %></title>
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" language="javascript">
			$(document).ready(function(){
				// focus to first element in form ..
				$(".gets-focus").focus();
			});
		</script>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><a href="/">Home</a> / Login</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="contentbox"><%=m_bodyText %></div>
			</div>
		</div>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_footer.asp"-->
	</body>
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	Dim sess		: Set sess = New cSession
	Call OnPageLoad(page)

	str = str & ApplicationMessageToString(page.MessageID)
	page.MessageID = ""
	
	Select Case page.Action
		Case LOGOFF_SESSION_TIMEOUT
			' session expired, relogin
			
			str = str & FormLoginToString(page)
			
		Case LOGOFF_SESSION_ABANDON
			' session deleted, relogin
			Call LogoutMember()
			Response.Cookies("sid") = ""
			page.Action = ""
			page.MessageID = 1014
			Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False))
			
		Case LOGOFF_USER
			Call LogoutMember()
			page.Action = ""
			page.MessageID = 1013
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case LOGON_USER
			If ValidFormLogin(page.Member) Then
			
				Call LoginMember(sess, page.Member, rv)
				If rv = 0 Then 
					page.Action = ""
					Response.Redirect(page.Member.HomePageUrl & page.UrlParamsToString(False))
				End If
				
				' handle login errors
				Select Case rv
					Case -2
						' multiple rows returned
						page.MessageID = 1009
					Case -3
						' login not found
						page.MessageID = 1007
					Case -4
						' account disabled
						page.MessageID = 1047
					Case -5
						' trial expired
						page.MessageID = -5
						Response.Redirect("/client/account.asp" & page.UrlParamsToString(False))
					Case Else
						' unknown login error
						page.MessageID = 1008
				End Select
				page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormLoginToString(page)
			End If			
			
		Case GET_LOGON_CREDENTIALS
			If Request.Form("FormSendCredentialsIsPostback") = IS_POSTBACK Then
				If ValidFormSendCredentials(page.Member) Then
					Call GetCredentialsByLogin(page.Member, rv)
					If rv = 0 Then
						Call SendCredentials(page.Member.MemberID, Application.Value("NO_REPLY_EMAIL_ADDRESS"))
						page.MessageID = 1031 : page.Action = ""
					Else
						page.MessageID = 1032
					End If
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormSendCredentialsToString(page)
				End If
			Else
				str = str & FormSendCredentialsToString(page)
			End If
			
		Case Else
			' check for logged in member
			If Len(page.Member.MemberID) > 0 Then
				page.MessageID = 1012
				Response.Redirect("/member/programs.asp" & page.UrlParamsToString(False))
			End If
			
			str = str & FormLoginToString(page)
	End Select
	
	m_bodyText = str
	Set page = Nothing
	Set sess = Nothing
End Sub

Sub GetCredentialsByLogin(member, outError)
	Dim rs
	outError = 0
	
	Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.Open Application.Value("CNN_STR")
	
	Dim cmd			: Set cmd = Server.CreateObject("ADODB.Command")
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "up_memberGetLoginCredentials"
	cmd.ActiveConnection = cnn

	cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
	cmd.Parameters.Append cmd.CreateParameter("@NameLogin", adVarChar, adParamInput, 25, CStr(member.NameLogin))

	Set rs = cmd.Execute
	
	If Not rs.EOF Then
		member.MemberID = rs("MemberID").Value
		member.Load()
	End If	
	
	If rs.State = adStateOpen Then rs.Close()
	outError = cmd.Parameters("@RETURN_VALUE").Value
	
	Set cmd = Nothing
	Set rs = Nothing
	cnn.Close: Set cnn = Nothing
End Sub



Function ValidFormLogin(member)
	ValidFormLogin = True

	If Not ValidData(member.NameLogin, True, 4, 25, "Username", "") Then ValidFormLogin = False
	If Not ValidData(member.PWord, True, 4, 14, "Password", "") Then ValidFormLogin = False
End Function

Function ValidFormSendCredentials(member)
	ValidFormSendCredentials = True
	
	If Not ValidData(member.NameLogin, True, 4, 25, "Username", "") Then ValidFormSendCredentials = False
End Function

Function FormLoginToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	pg.Action = LOGON_USER
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form method=""post"" action=""" & Request.ServerVariables("URL") & pg.UrlParamsToString(True) & """ name=""formLogin"">"
	str = str & "<table>"
	str = str & "<tr><td class=""label"">Username</td>"
	str = str & "<td><input type=""text"" class=""gets-focus"" name=""NameLogin"" value=""" & page.Member.NameLogin & """ /></td></tr>"
	str = str & "<tr><td class=""label"">Password</td>"
	str = str & "<td><input type=""password"" name=""PWord"" value=""" & page.Member.PWord & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Login"" />"
	str = str & "<input type=""hidden"" name=""FormLoginIsPostback"" value=""" & IS_POSTBACK & """ />"
	pg.Action = GET_LOGON_CREDENTIALS
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Forgot password?</a>"
	str = str & "</td></tr>"
	str = str & "</table></form></div>"
	
	FormLoginToString = str
End Function

Function FormSendCredentialsToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()

	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form method=""post"" action=""" & pg.Url & pg.UrlParamsToString(True) & """ name=""formSendCredentials"">"

	str = str & "<table>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Username</td>"
	str = str & "<td><input class=""medium gets-focus"" type=""text"" name=""NameLogin"" value=""" & page.Member.NameLogin & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "Provide your username and we'll send your login <br />information to your email address. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Send"" />"
	
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & Request.ServerVariables("URL") & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormSendCredentialsIsPostBack"" value=""" & IS_POSTBACK & """ />"
	str = str & "</td></tr></table></form></div>"
	
	FormSendCredentialsToString = str
End Function
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<%
Class cPage
	' encrypted
	Public Action

	' not encrypted
	Public MessageID
	
	' object
	Dim Client
	Dim Member
	
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		
		If Len(str) > 0 Then 
			str = Left(str, Len(str) - Len(amp))
		Else
			' qstring needs at least one param in case more params are appended ..
			str = str & "noparm=true"
		End If
		str = "?" & str
		
		UrlParamsToString = str
	End Function
	
	Public Property Get Url()
		Url = Request.ServerVariables("URL")
	End Property
		
	Public Function Clone()
		Dim c
		Set c = New cPage
		
		c.MessageID = MessageID
		c.Action = Action
		Set c.Client = Client
		Set c.Member = Member
		
		Set Clone = c
	End Function
End Class
%>

