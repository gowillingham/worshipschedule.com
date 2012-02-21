<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Sub OnPageLoad(ByRef page)
	page.MessageID = Request.QueryString("msgid")
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/outside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=Application.Value("APPLICATION_NAME") & " - Simple Web Scheduling for Worship Teams" %></title>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><a href="/">Home</a> / Policies Overview</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="column">
					<div class="feature-box">
						<h3>More Info ..</h3>
						<ul>
							<li><a href="/terms.asp"><strong>Terms of Use</strong></a></li>
							<li><a href="/privacy.asp"><strong>Privacy</strong></a></li>
						</ul>
					</div>
				</div>
				<h1>Website Policies Overview</h1>
					<h3>Email</h3>
					<p><%=Application.Value("APPLICATION_NAME")%> collects e-mail addresses from you when you try out or buy our service. 
						These addresses are used for sending login credentials and account information only. 
						Registered users can change or remove their e-mail address by
						<a href="mailto:<%=Application.Value("SUPPORT_EMAIL_ADDRESS")%>" title="Contact Us">contacting us</a>.
					</p>
					<p>If you sign up for a <%=Application.Value("APPLICATION_NAME")%> email newsletter or notification, your email 
						address will be kept confidential, and it will only be used to send you our newsletter.
					</p>
					<h3>Personal Information</h3>
					<p>We will <strong>never</strong> share your email address or personal information with third party companies.
						We also will not access the content of your account unless you specifically request it 
						(for example, if you are having technical difficulties accessing your account) or if required by law, 
						or to maintain our system, or to protect <%=Application.Value("APPLICATION_NAME")%> or the public.
					</p>
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

	page.MessageID = ""
	
	Set page = Nothing
End Sub

%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<%
Class cPage
	Public MessageID
	
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
		
		Set Clone = c
	End Function
End Class
%>

