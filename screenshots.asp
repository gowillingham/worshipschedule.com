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
		<link rel="stylesheet" type="text/css" href="/_incs/script/jquery/plugins/prettyphoto/css/prettyphoto.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=Application.Value("APPLICATION_NAME") & " - Simple Web Scheduling for Worship Teams" %></title>
		
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/plugins/prettyphoto/js/jquery.prettyphoto.js"></script>
		<script language="javascript" type="text/javascript">
			$(document).ready(function(){
				$("#screenshot-list a").prettyPhoto({
					allowResize: true,
					padding: 40
				});			
			});
		</script>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><a href="/">Home</a> / Screen Shots</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="column">
					<div class="feature-box">
						<h3>More Info ..</h3>
						<ul>
							<li><a href="/requirements.asp"><strong>System Requirements</strong></a></li>
							<li><a href="/pricing.asp"><strong>Pricing</strong></a></li>
						</ul>
					</div>
					<img style="" src="_images/man_stands_on_rock.jpg" alt="Happy on Rock" />
				</div>
				<div id="screenshots">
					<h1><%=Application.Value("APPLICATION_NAME")%> Screen Shots</h1>
					<p>Take a screen shot tour of <%= Application.Value("APPLICATION_NAME") %>. Find out what our web application can do for you and your scheduling needs. </p>
					
					<table id="screenshot-list">
						<tbody>
							<tr>
								<td><a rel="prettyPhoto[gallery]" href="/_images/screenshots/member_overview.gif">
									<img src="/_images/screenshots/member_overview_thumb.gif" alt="Member home" /></a>
								</td>
								<td class="copy"><h3>Member home</h3>
									<p>A quick overview of upcoming events for your team. 
									Easily let your team know when there are new events on the calendar. </p>
								</td>
							</tr>
							<tr>
								<td><a rel="prettyPhoto[gallery]" href="/_images/screenshots/member_availability.gif">
									<img src="/_images/screenshots/member_availability_thumb.gif" alt="Member availability page" /></a>
								</td>
								<td class="copy"><h3>Availability</h3>
									<p>Find out when your team is available before you schedule. </p>
								</td>
							</tr>
							<tr>
								<td><a rel="prettyPhoto[gallery]" href="/_images/screenshots/member_calendar.gif">
									<img src="/_images/screenshots/member_calendar_thumb.gif" alt="Member calendar" /></a>
								</td>
								<td class="copy"><h3>Member calendar</h3>
									<p>Your members login to see your team's events on their calendar when you are ready. </p>
								</td>
							</tr>
							<tr>
								<td><a rel="prettyPhoto[gallery]" href="/_images/screenshots/member_event.gif">
									<img src="/_images/screenshots/member_event_thumb.gif" alt="Event page" /></a>
								</td>
								<td class="copy"><h3>Event summary</h3>
									<p><%=Application.Value("APPLICATION_NAME")%> provides an event summary page for each event. 
									</p>
								</td>
							</tr>
							<tr>
								<td><a rel="prettyPhoto[gallery]" href="/_images/screenshots/admin_home.gif">
									<img src="/_images/screenshots/admin_home_thumb.gif" alt="Administration home" /></a>
								</td>
								<td class="copy"><h3>Administration home</h3>
									<p>See your account at a glance from your account administration dashboard. </p>
								</td>
							</tr>
							<tr>
								<td><a rel="prettyPhoto[gallery]" href="/_images/screenshots/admin_calendar.gif">
									<img src="/_images/screenshots/admin_calendar_thumb.gif" alt="Master event calendar" /></a>
								</td>
								<td class="copy"><h3>Master event calendar</h3>
									<p>Organize your events into as many schedules as you need. </p>
								</td>
							</tr>
							<tr>
								<td><a rel="prettyPhoto[gallery]" href="/_images/screenshots/build_schedule.gif">
									<img src="/_images/screenshots/build_schedule_thumb.gif" alt="Event teams" /></a>
								</td>
								<td class="copy"><h3>Event teams</h3>
									<p>Each event has a team. You assign your members to the team based on what they'll do at the event. </p>
								</td>
							</tr>
							<tr>
								<td><a rel="prettyPhoto[gallery]" href="/_images/screenshots/email_page.gif">
									<img src="/_images/screenshots/email_page_thumb.gif" alt="Email" /></a>
								</td>
								<td class="copy"><h3>Email</h3>
									<p>Communicate with your team members by email. <%= Application.Value("APPLICATION_NAME")%> makes it simple to send email to individual members or groups. </p>
								</td>
							</tr>
						</tbody>
					</table>
				</div>
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

