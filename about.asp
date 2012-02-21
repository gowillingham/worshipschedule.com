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
				<h1><a href="/">Home</a> / About Us</h1>
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
						</ul>
					</div>
					<img style="" src="_images/man_sits_on_rock.jpg" alt="Man on Rock" />
				</div>
				<h1>About <%=Application.Value("APPLICATION_NAME") %></h1>
				<p>
					We started <%=Application.Value("APPLICATION_NAME") %> in 2004 to solve a simple problem - 
					reliably figuring out when members of numerous volunteer teams at our large church (<a href="http://www.hosannalc.org/">Hosanna!</a>) were available for events. 
					The solution had to be simple, requiring no more than a few clicks to view and construct a schedule or to indicate schedule preferences. 
					At the same time, it had to reduce time spent on the phone or shuffling paper schedule requests.
				</p>
				<p>
					Any church, large or small that depends on volunteer teams could benefit by moving their scheduling to the web with <%=Application.Value("APPLICATION_NAME") %>.
				</p>
				<h1>Tell Us What You Think!</h1>
				<p>
					As part of the GTD Solutions family, we want you to give us your feedback, questions, ideas or complaints. 
					We promise to keep improving our service to meet your needs.
				</p>
				<p>Please direct all support related inquiries to <a href="mailto:<%=Application.Value("SUPPORT_EMAIL_ADDRESS") %>"><%=Application.Value("SUPPORT_EMAIL_ADDRESS") %></a>.</p>
				<p>
					<strong>GTD Solutions, LLC</strong>
					<br />16827 Interlachen Boulevard
					<br />Lakeville, Minnesota 55044
					<br /><a href="mailto:<%=Application.Value("INFO_EMAIL_ADDRESS") %>"><strong><%=Application.Value("INFO_EMAIL_ADDRESS") %></strong></a>
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

