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
		<title><%=Application.Value("APPLICATION_NAME") & " - Simple Web Scheduling for Volunteer Teams" %></title>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><a href="/">Home</a> / <a href="/overview.asp">Overview</a> / Pricing</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="column">
					<div class="feature-box">
						<h3>More Info ..</h3>
						<ul>
							<li><a href="/howitworks.asp"><strong>How It Works!</strong></a></li>
							<li><a href="/requirements.asp"><strong>System Requirements</strong></a></li>
							<li><a href="/about.asp"><strong>About Us</strong></a></li>
						</ul>
					</div>
					<img style="" src="_images/crazy_businesswoman2.jpg" alt="Crazy Woman" />
				</div>
				<h1><%=Application.Value("APPLICATION_NAME") %> Pricing</h1>
				<p>
					<strong>$25.00/Month Unlimited! </strong>
					Add as many of your members as you need. Set up as many teams as you wish. Send email as often as you want. 
					Your <%=Application.Value("APPLICATION_NAME") %> account has no restrictions on members or teams. 
					Each subscription also includes up to one gigabite of file storage space within your <%=Application.Value("APPLICATION_NAME") %> account for file uploads.
				</p>
				<h1>Ready to Go Instantly</h1>
				<p>
					<%=Application.Value("APPLICATION_NAME") %> runs in a standard web browser over the internet. 
					You and your team members already have the necessary software installed to use it right now.
					You can be up and running with <%=Application.Value("APPLICATION_NAME") %> within minutes. 
					All you need is a valid email address to <a href="/tryit.asp"><strong>try it </strong></a> (free) right now.
				</p>
				<h1>No Contracts</h1>
				<p>
					You may cancel your <%=Application.Value("APPLICATION_NAME") %> account any time for any reason.
					We'll refund the remainder of your unused subscription, no questions asked. 
					We think you'll find <%=Application.Value("APPLICATION_NAME") %> to a an invaluable aid to you in managing your teams.
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

