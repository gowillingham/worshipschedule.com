<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Sub OnPageLoad(ByRef page)
	page.HelpID = Request.QueryString("hid")
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
				<h1><%=Application.Value("APPLICATION_NAME")%> Home</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div id="homegraphic" style="margin-right:25px;float:left;">
					<img style="" src="_images/perspective_calendar.png" alt="<%=Application.Value("APPLICATION_NAME")%>" />
				</div>
				<h1>A Better Way to Schedule</h1>
				<p style="font-size:1.2em;">
					Manage and schedule your worship team from anywhere. 
					Worshipschedule automatically keeps track of your team members, what they can do, and when they are available, 
					saving you valuable time to focus on your events instead of who will staff them.
				</p>
				<h1>Why <%=Application.Value("APPLICATION_NAME")%>?</h1>
				<p style="font-size:1.2em;">
					Put an end to phone tag, paper shuffling, correction and revision. 
					Now your calendar can be ready and available to your members in minutes through a safe and secure web page.
					You can even upload your charts and rehearsal tracks to Worshipschedule so your team can access them from any web browser!
				</p>
					<a href="/tryit.asp"><img style="float:right;margin:10px 60px 25px 25px;" src="/_images/btn_try_it_free.png" alt="Try It Free" /></a>
					<a href="/overview.asp"><img style="float:right;margin-top:10px;margin-bottom:25px;"  src="/_images/btn_learn_more.png" alt="Learn More" /></a>
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
	Public HelpID
	Public MessageID
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(HelpID) > 0 Then str = str & "hid=" & HelpID & amp
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
		
		c.HelpID = HelpID
		c.MessageID = MessageID
		
		Set Clone = c
	End Function
End Class
%>

