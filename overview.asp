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
				<h1><a href="/">Home</a> / Overview</h1>
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
							<li><a href="/pricing.asp"><strong>Pricing</strong></a></li>
							<li><a href="/about.asp"><strong>About Us</strong></a></li>
						</ul>
					</div>
					<img style="" src="_images/crazy_businesswoman.JPG" alt="Crazy Woman" />
				</div>
				<h1><%=Application.Value("APPLICATION_NAME") %> Overview</h1>
				<p>
					<%=Application.Value("APPLICATION_NAME") %> helps you to easily manage any number of worship teams or members for your church. 
					Your team members let you know when they are available for your events from any web browser. 
					When you are ready, <%=Application.Value("APPLICATION_NAME") %> helps you quickly build your schedule from their availability information. 
					See <a href="/howitworks.asp"><strong>how it works</strong></a>.
				</p>
				<h1>Set Up and Configure Your Account</h1>
				<ul>
					<li>Go to <a href="/"><strong>http://<%=Request.ServerVariables("SERVER_NAME")%></strong></a> from any web browser and login to access your secure account</li>
					<li>From here, you can do any of the following:
						<ul>
							<li><strong>Manage. </strong>Add team members to your account. All you need is their email address.</li>
							<li>
								<strong>Schedule. </strong>Setup schedules and events for your team or teams. 
								Once you have completed a schedule, your members may login to view, print, or download the schedule to their computers. 
							</li>
							<li>
								<strong>Communicate. </strong>
								Contact your team or teams by email from Worshipschedule. 
								Smart team email groups are already built-in.
								Send your schedules to your entire team with one click!
							</li>
							<li><strong>Share. </strong>Upload charts and rehearsal tracks (pdf, mp3, etc) for your team to access from their own <%=Application.Value("APPLICATION_NAME") %> accounts.</li>
						</ul>						
					</li>
				</ul>
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

