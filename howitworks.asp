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
		<title><%=Application.Value("APPLICATION_NAME") & "Simple Web Scheduling for Worship Teams" %></title>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><a href="/">Home</a> / <a href="/overview.asp">Overview</a> / How It Works</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<h1>How To Schedule With <%=Application.Value("APPLICATION_NAME") %></h1>
				<div id="howitworks"><img src="_images/howto.gif" alt="How It Works" /></div>
				
				<h1>It's That Simple!</h1>
				<p>
					Of course there is much much more to it than one picture can show. 
					A more detailed description of how Worshipschedule helps you organize your teams and your schedules can be found in the FAQ topic 
					<a href="/help/topic.asp?hid=12"><strong>before you begin with <%=Application.Value("APPLICATION_NAME")%></strong></a>.				
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

