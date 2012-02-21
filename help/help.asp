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
		<title><%=Application.Value("APPLICATION_NAME") & " Help" %></title>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1>Help</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="feature-box">
					<h3>New to <%=Application.Value("APPLICATION_NAME") %>?</h3>
					<p>
						You'll want to look at our help topic <a href="/help/topic.asp?hid=12"><strong>Before You Begin</strong></a>. 
						Then, the <a href="/help/topic.asp?hid=14"><strong>Getting Started</strong></a> guide has all you need to know.
					</p>
				</div>
				
				
				<div class="feature-box linklist">
					<h3>General Use: I Want To ..</h3>
					<ul>
						<li><a href="/help/topic.asp?hid=6"><strong>Work with programs</strong></a></li>
						<li><a href="/help/topic.asp?hid=17"><strong>Work with my calendar</strong></a></li>
						<li><a href="/help/topic.asp?hid=6#availability-anchor"><strong>Show when I'm available</strong></a></li>
					</ul>
					<h3 style="margin-top:25px;">Administration: I Want To ..</h3>
					<ul>
						<li><a href="/help/topic.asp?hid=1"><strong>Add members to my account</strong></a></li>
						<li><a href="/help/topic.asp?hid=14#anchor-add-program"><strong>Add a program to my account</strong></a></li>
						<li><a href="/help/topic.asp?hid=14#anchor-add-schedule"><strong>Add a schedule to a program</strong></a></li>
						<li><a href="/help/topic.asp?hid=14#anchor-add-team"><strong>Schedule an event team</strong></a></li>
						<li><a href="/help/topic.asp?hid=14#anchor-email-team"><strong>Email my team members</strong></a></li>
						<li><a href="/help/topic.asp?hid=16"><strong>Upgrade to a paid subscription</strong></a></li>
					</ul>
				</div>
				<h1>Looking for Help?</h1>
				<p style="font-size:1.2em;">
					Look no further. The best place to start is the <%=Application.Value("APPLICATION_NAME") %> <a href="/help/faq.asp"><strong>FAQ</strong></a>. 
					If you can't find your answer there, you might try the <a href="<%=Application.Value("SUPPORT_FORUM_URL")%>"><strong>forums</strong></a>!
				</p>
				<h1>Still Having Problems?</h1>
				<p style="font-size:1.2em;">Email us at <a href="/support.asp">Customer Support</a> and we will get back to you with help for any issue within one business day.</p>
				<h1>Help Us Improve!</h1>
				<p style="font-size:1.2em;">We are always looking to improve <%=Application.Value("APPLICATION_NAME") %>, and we appreciate your feedback.</p>
				<ul>
					<li><a href="/support.asp">Give us feedback</a></li>
					<li><a href="/support.asp">Submit a bug report</a></li>
					<li><a href="/support.asp">Get help from customer support</a></li>
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
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
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

