<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%

' - empty the response cache before beginning error page
Response.Clear

' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-schedules"
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
	Call CheckSession(sess, PERMIT_ALL)
	
	page.MessageID = Request.QueryString("msgid")
	page.ErrorId = Request.QueryString("errid")
	
	page.Action = Decrypt(Request.QueryString("act"))
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	If Len(page.Client.ClientId) > 0 Then Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	If Len(page.Member.MemberId) > 0 Then Call page.Member.Load()
	
	' set the view tokens
	m_appMessageText = ApplicationMessageToString(page.MessageID)
	page.MessageID = ""
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/inside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/plugins/form/jquery.form.js"></script>
		<script type="text/javascript" language="javascript">
			$(document).ready(function(){
			
				// wire up form for ajax submit ..	
				
				$("#form-error-button").click(function(){
					$("#form-error").ajaxForm({
						beforeSubmit: function(){
							var height = $("#form-container form").height();
							var width = $("#form-container form").width();
							
							$("#form-container").css("height", height)
							$("#form-container").css("width", width)
							
							$("#form-container form").hide();
							$("#form-container").addClass("loading");

						},
						success: function(responseText){
							$("#form-container").removeClass("loading");
							$("div.confirm p").html(responseText);
							
							$("div.confirm").css("display", "block");
						}
					});
				});
			});
		</script>
		<title><%=Application.Value("APPLICATION_NAME") & " - Web Scheduling"%></title>
	</head>
	<body><%=m_bodyText%></body>
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Dim exception	: Set exception = Server.GetLastError()
	Dim errorId
	
	Call OnPageLoad(page)
	
	Select Case page.Action
		Case Else
			Call InsertError(page, exception, errorId, rv)
			page.ErrorId = errorId	
			Call SendError(page, exception)
			
			str = str & FormErrorMessageToString(page, exception)	
					
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub SendError(page, exception)
	Dim str
	Dim emailSender		: Set emailSender = New cEmailSender
	Dim toAddress		: toAddress = Application.Value("ADMIN_EMAIL_ADDRESS")
	Dim fromAddress		: fromAddress = Application.Value("APPLICATION_ERROR_EMAIL_ADDRESS")
	Dim subject			: subject = "** " & Application("APPLICATION_NAME") & " Application Error **"
	
	'Generate the email body
	str = str &  "Information for Support Personnel:" & vbCrLf
	str = str &  "****************************************************************" & vbCrLf
	str = str &  "At: " & Now() & vbCrLf
	str = str &  "CustomerRefID: ClientID='" & page.Client.ClientID & "' MemberID='" & page.Member.MemberId & "'" & vbCrLf
	str = str &  "ClientName: " & page.Client.NameClient & vbCrLf
	str = str &  "MemberName: " & page.Member.NameLast & ", " & page.Member.NameFirst & vbCrLf
	str = str &  "ErrorLogID: " & page.ErrorID & vbCrLf
	str = str &  "SessionID: " & Session.SessionID & vbCrLf
	str = str &  "RequestMethod: " & Request.ServerVariables("REQUEST_METHOD") & vbCrLf
''	str = str &  "Form Data: " & Request.Form & vbCrLf
	str = str &  "ServerPort: " & Request.ServerVariables("SERVER_PORT") & vbCrLf
	str = str &  "HTTPS: " & Request.ServerVariables("HTTPS") & vbCrLf
	str = str &  "Server Address: "  & Request.ServerVariables("LOCAL_ADDR") & vbCrLf
	str = str &  "Client Address: "  & Request.ServerVariables("REMOTE_ADDR") & vbCrLf
	str = str &  "Client Browser: " & Request.ServerVariables("HTTP_USER_AGENT") & vbCrLf & vbCrLf
	str = str &  "ASP Page: " &  Request.ServerVariables("URL") & vbCrLf
	str = str &  "Error #: " & exception.ASPCode & vbCrLf
	str = str &  "COM Error #: " & exception.Number & " (Hex " & Hex(exception.Number) & ")" & vbCrLf
	str = str &  "Source: " & exception.Source & vbCrLf
	str = str &  "Category: " & exception.Category & vbCrLf
	str = str &  "File: " & "//" & Request.ServerVariables("SERVER_NAME") & exception.File & vbCrLf
	str = str &  "Line: " & exception.Line & vbCrLf
	str = str &  "Column :" & exception.Column & vbCrLf
	str = str &  "Description: " & exception.Description & vbCrLf
	str = str &  "ASP Description: " & exception.ASPDescription  & vbCrLf
	str = str &  vbCrLf & "HTTP Headers: " & vbCrLf
	str = str &  "==============================" & vbCrLf
	str = str &  Replace(Request.ServerVariables("ALL_HTTP"),vbLf,vbCrLf)
	str = str &  "------------------------------" & vbCrLf
	str = str &  "****************************************************************" & vbCrLf & vbCrLf
	
	'turn off error-checking for email send
	On Error Resume Next
		emailSender.SendMessage toAddress, fromAddress, subject, str
	On Error GoTo 0
End Sub

Function FormErrorMessageToString(page, exception)
	Dim str, msg
	Dim pg						: Set pg = page.Clone()
	
	Dim problem					: problem = ""
	If Len(problem) = 0 Then problem = "Internal error"
	If Len(exception.Number() & "") > 0 Then problem = problem & " (" & exception.Number() & ")"

	Dim description				: description = ""
	If Len(exception.Source() & "") > 0 Then description = description & "'" & exception.Source() & "' "
	If Len(exception.Description() & "") > 0 Then description = description & exception.Description() & ". "
	If Len(exception.AspDescription() & "") > 0 Then description = description & exception.AspDescription() & ". "
	If Len(exception.Category() & "") > 0 Then description = description & "(" & exception.Category() & "). "
	
	Dim scoutDescription		: scoutDescription = ""
	scoutDescription = "[" & Application.Value("APPLICATION_NAME") & "] " & exception.Description() & " (" & exception.File() & " Line " & exception.Line() & " ErrorID:" & page.ErrorID & ")"
	
	Dim scoutDefaultMessage		: scoutDefaultMessage = ""
	scoutDefaultMessage = "We're sorry that " & Application.Value("APPLICATION_NAME") & " is giving you trouble. "
	scoutDefaultMessage = scoutDefaultMessage & "We did receive your report and will respond to the email address you provided (usually within 24 hours) to let you know if there is a workaround or when this issue will be fixed. "
	
	str = str & "<div id=""container""><div class=""app-error"">"
	str = str & "<h1>" & Application.Value("APPLICATION_NAME") & " Application Error!</h1>"

	str = str & "<p>Sorry, but there has been a problem with " & Application.Value("APPLICATION_NAME") & " and whatever you were just trying to do cannot be completed. "
	str = str & "This is more than likely an issue with " & Application.Value("APPLICATION_NAME") & " and not your fault. "
	str = str & "Below you can see a desription of the actual error that occurred. </p>"
	
	str = str & "<table class=""description""><tbody>"
	str = str & "<tr><td class=""label"">Time</td>"
	str = str & "<td>" & Now() & "</td></tr>"
	str = str & "<tr><td class=""label"">Problem</td>"
	str = str & "<td>" & html(problem) & "</td></tr>"
	str = str & "<tr><td class=""label"">Description</td>"
	str = str & "<td>" & html(description) & "</td></tr>"
	str = str & "</tbody></table>"
	
	str = str & "<p>You could try what you were doing again and see if the problem clears up, "
	str = str & "but there is a very good chance that you'll continue to see this error until we can fix it on our end. "
	str = str & "To move things along, it would be helpful if you could send our support staff a short note describing what happened and what you were doing at the time this problem occurred. </p>"
	
	str = str & "<div id=""form-container"">"
	str = str & "<form action=""" & "/_incs/script/ajax/_scout_submit_proxy.asp" & """ method=""post"" id=""form-error"">"
	
	str = str & "<input type=""hidden"" name=""ScoutUserName"" id=""scout-user-name"" value=""Administrator"" />"
	str = str & "<input type=""hidden"" name=""ScoutProject"" id=""scout-project"" value=""worshipschedule.com"" />"
	str = str & "<input type=""hidden"" name=""ScoutArea"" id=""scout-area"" value=""Bug"" />"
	str = str & "<input type=""hidden"" name=""Description"" id=""description"" value=""" & server.HTMLEncode(scoutDescription) & """ />"
	str = str & "<input type=""hidden"" name=""ForceNewBug"" id=""force-new-bug"" value=""0"" />"
	str = str & "<input type=""hidden"" name=""ScoutDefaultMessage"" id=""scout-default-message"" value=""" & scoutDefaultMessage & """ />"
	str = str & "<input type=""hidden"" name=""FriendlyResponse"" id=""friendly-response"" value=""1"" />"
	
	str = str & "<table class=""form""><tbody>"
	str = str & "<tr><td class=""label"">Your Email</td>"
	str = str & "<td><input type=""text"" name=""Email"" id=""email"" value=""" & page.Member.Email & """ class=""text"" />"
	str = str & " (optional)</td></tr>"
	str = str & "<tr><td class=""label"">And what happened?</td>"
	str = str & "<td><textarea name=""Extra"" id=""extra"" class=""textarea""></textarea></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td><input type=""submit"" name=""submit"" value=""Send"" id=""form-error-button"" /></td></tr>"
	str = str & "</tbody></table></form>"
	
	' the hidden confirmation message ..
	str = str & "<div class=""confirm"" style=""display:none;"">"
	str = str & "<h5>Thanks - Your problem report was received!</h5>"
	str = str & "<p>Thanks for the information about the problem you experienced. </p></div>"	
	
	str = str & "</div></div></div>"

	FormErrorMessageToString = str
End Function

Sub InsertError(page, exception, errorId, outError)
	If Not IsObject(exception) Then Exit Sub

	Dim parm
	Dim cmd			: Set cmd = Server.CreateObject("ADODB.Command")
	Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
	
	' turn off error-checking for db write
	On Error Resume Next
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_adminInsertAppErrorLog"
			cnn.Open Application.Value("CNN_STR")
			.ActiveConnection = cnn
		End With

		Set parm = cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ErrorTime", adDBTimeStamp, adParamInput, 0, Now())
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ClientRefID", adVarChar, adParamInput, 50, "cid:" & page.Client.ClientId & "|mid:" & page.Member.MemberID)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NameMember", adVarChar, adParamInput, 101, (page.Member.NameLast & ", " & page.Member.NameFirst & ""))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NameClient", adVarChar, adParamInput, 100, (page.Client.NameClient & ""))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@SessionID", adVarChar, adParamInput, 12, CStr(Session.SessionID))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@RequestMethod", adVarChar, adParamInput, 10, Request.ServerVariables("REQUEST_METHOD"))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@FormData", adVarChar, adParamInput, 2000, Request.Form)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ServerPort", adVarChar, adParamInput, 10, Request.ServerVariables("SERVER_PORT"))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@HTTPS_Status", adVarChar, adParamInput, 10, Request.ServerVariables("HTTPS"))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ServerIP", adVarChar, adParamInput, 15, Request.ServerVariables("LOCAL_ADDR"))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ClientIP", adVarChar, adParamInput, 15, Request.ServerVariables("REMOTE_ADDR"))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@Browser", adVarChar, adParamInput, 255, Request.ServerVariables("HTTP_USER_AGENT"))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ASP_Page", adVarChar, adParamInput, 400, Request.ServerVariables("URL"))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ASP_ErrNumber", adVarChar, adParamInput, 12, exception.ASPCode)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@COM_ErrNumber", adVarChar, adParamInput, 12, exception.Number)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@Source", adVarChar, adParamInput, 255, exception.Source)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@Category", adVarChar, adParamInput, 50, exception.Category)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ErrLine", adInteger, adParamInput, 0, exception.Line)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ErrColumn", adInteger, adParamInput, 0, exception.Column)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ErrPage", adVarChar, adParamInput, 100, "//" & Request.ServerVariables("SERVER_NAME") & exception.File)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ErrDescription", adVarChar, adParamInput, 1000, exception.Description)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ASP_Description", adVarChar, adParamInput, 1000, exception.ASPDescription)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@HTTP_Headers", adVarChar, adParamInput, 2000, Request.ServerVariables("ALL_HTTP"))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NewID", adBigInt, adParamOutput)
		cmd.Parameters.Append parm

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		errorId = cmd.Parameters("@NewID").Value
		
		Set parm = Nothing: Set cmd = Nothing
		If cnn.State = adStateOpen Then cnn.Close: Set cnn = Nothing
	On Error GoTo 0
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
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	Public ErrorId
	
	' encrypted
	Public Action

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
		If Len(ErrorId) > 0 Then str = str & "errid=" & ErrorId & amp
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		
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
		c.ErrorId = ErrorId

		c.Action = Action
				
		Set c.Member = Member
		Set c.Client = Client

		Set Clone = c
	End Function
End Class
%>
