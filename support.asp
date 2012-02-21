<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Dim m_bodyText

Sub OnPageLoad(ByRef page)
	page.MessageID = Request.QueryString("msgid")
	
	Set page.Data = New cData
	page.Data.Text = Request.Form("Text")
	page.Data.Email = Request.Form("Email")
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/outside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=Application.Value("APPLICATION_NAME") & "- Simple Web Scheduling for Worship Teams" %></title>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><a href="/">Home</a> / Support</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="column">
					<div class="feature-box">
						<h3>More Info ..</h3>
						<ul>
							<li><a href="/help/help.asp"><strong>Help</strong></a></li>
							<li><a href="/help/faq.asp"><strong>FAQ</strong></a></li>
							<li><a href="/requirements.asp"><strong>System Requirements</strong></a></li>
							<li><a href="/about.asp"><strong>About Us</strong></a></li>
						</ul>
					</div>
					<img style="" src="_images/solitary_laptop.jpg" alt="Man on Rock" />
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

	str = str & "<div style=""width:540px;"">" & ApplicationMessageToString(page.MessageID) & "</div>"
	page.MessageID = ""
	
	If Request.Form("FormSupportIsPostback") = IS_POSTBACK Then
		If ValidFormSupport(page.Data) Then
			Call SendSupportMessage(page.Data)
			page.MessageID = 8002
			Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False))
		Else
			str = str & PageContentToString()
			str = str & FormSupportToString(page)
		End If
	Else
		str = str & PageContentToString()
		str = str & FormSupportToString(page)
	End If
	
	Set page = Nothing
	
	m_bodyText = str
End Sub

Sub SendSupportMessage(data)
	Dim email		: Set email = New cEmailSender
	Dim subject
	
	Dim fromAddress	: fromAddress = "anonymous@" & Application.Value("ROOT_EMAIL_DOMAIN")
	If Len(data.Email) > 0 Then
		fromAddress = data.Email
	End If
	
	subject = "[" & Application.Value("APPLICATION_NAME") & "] ** Support Request for " & fromAddress & " **"

	Call email.SendMessage(Application.Value("SUPPORT_EMAIL_ADDRESS"), fromAddress, subject, data.Text)
	Set email = Nothing
End Sub

Function ValidFormSupport(data)
	ValidFormSupport = True
	
	If Not ValidData(data.Email, False, 0, 100, "Email", "email") Then ValidFormSupport = False
	If Not ValidData(data.Text, True, 0, 8000, "Message", "") Then ValidFormSupport = False
End Function

Function FormSupportToString(page)
	Dim str

	str = str & "<div class=""form"" style=""width:535px;"">"
	str = str & ErrorToString()
	str = str & "<form method=""post"" action=""" & Request.ServerVariables("URL") & page.UrlParamsToString(True) & """ name=""formSupport"">"
	str = str & "<table>"
	str = str & "<tr><td class=""label"">Your Email:</td>"
	str = str & "<td><input type=""text"" name=""Email"" value=""" & page.Data.Email & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td class=""label"">&nbsp;</td>"
	str = str & "<td class=""hint"">(so that we can get back to you!) </td></tr>"
	str = str & "<tr><td class=""label"">&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Your Message:</td>"
	str = str & "<td><textarea name=""Text"" style=""width:325px;height:150px;"">" & page.Data.Text & "</textarea></td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Send"" />"
	str = str & "<input type=""hidden"" name=""FormSupportIsPostback"" value=""" & IS_POSTBACK & """ /></td></tr>"
	str = str & "</table></form></div>"
	
	FormSupportToString = str
End Function

Function PageContentToString()
	Dim str

	str = str & "<h1>" & Application.Value("APPLICATION_NAME") & " Support</h1>"
	str = str & "<p>We'd love to hear from you! "
	str = str & "If you are looking for a quick answer to a question, "
	str = str & "you may be interested in checking our online <a href=""/help/help.asp"">help</a> or <a href=""/help/faq.asp"">FAQ</a> first. </p>"
	str = str & "<p>You can use the form below to contact Worshipschedule with requests, suggestions, or problems you may be experiencing. "
	str = str & "Sending us a message here will create a numbered case in our support system. "
	str = str & "We'll try to investigate and respond to your question within one day. </p>"
	str = str & "<p>You may also contact support staff at any time at <a href=""mailto:" & Application.Value("SUPPORT_EMAIL_ADDRESS") & """>" & Application.Value("SUPPORT_EMAIL_ADDRESS") & "</a> directly. </p>"
	
	PageContentToString = str
End Function

%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->

<%
Class cPage
	Public MessageID
	
	' object
	Public Data
	
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
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
	
	Public Function Clone()
		Dim c
		Set c = New cPage
		
		c.MessageID = MessageID
		Set c.Data = Data
		
		Set Clone = c
	End Function
End Class

Class cData
	Public Text
	Public Email
End Class
%>

