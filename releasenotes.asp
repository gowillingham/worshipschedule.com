<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Dim m_bodyText

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
				<h1><a href="/">Home</a> / Release Notes</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
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

	page.MessageID = ""
	
	str = ReleaseNotesToString()
	
	m_bodyText = str	
	Set page = Nothing
End Sub

Function ReleaseNotesToString()
	Dim str
	
	str = str & "<div class=""release-notes"">"
	str = str & "<h1>What's new in " & Application.Value("APPLICATION_NAME") & " Version 1.9</h1> "
	str = str & "<h3>Fixed: Saving event with time information in date field causes error.</h3> <p>Fixed issue where time information could be included in the date field when editing or creating a new event, causing an error when the event was saved. Fixed an issue where Worshipschedule was mistakenly including time information in events generated for trial accounts.</p> "
	str = str & "<h3>Fixed: File numbers in member filestore grid are incorrect</h3> <p>Fixed issue where the sequential number of a file in the member file listing was incorrect.</p> "
	str = str & "<h3>Fixed: Member files store does not display 'No File' message</h3> <p>Fixed issue where member file grid did not display a message when there were no public files available for a member.</p> "
	str = str & "<h3>Fixed: Previously uploaded files with spaces in file name downloaded with corrupted name</h3> <p>Fixed issue where files on server prior to version 1.8.5 update whose file names contained spaces were renamed on download with corrupted names, preventing them from being downloaded.</p> "
	str = str & "<h3>Fixed: Selecting an email group and 'No Selection' option at same time causes error.</h3> <p>Fixed problem where selecting an email group and the 'No Selection' option at he same time on Email Groups page caused an error.</p> "
	str = str & "<h3>Fixed: Unable to delete member from account</h3> <p>Fixed problem where in certain cases deleting a member from your Worshipschedule account would cause the website to crash.</p> "
	str = str & "<h3>Fixed: View master schedule causes error</h3> <p>Fixed intermittent problem with the member calendar where sometimes selecting the schedule overview for a program from the program dropdown list would cause worshipschedule to crash.</p> "
	str = str & "<h3>Fixed: Admin calendar event item text link give error.</h3> <p>Fixed issue where clicking event text link in admin calendar returned page not found error.</p> "
	str = str & "<h3>Fixed: Deleting event in event item in admin calendar view causes error</h3> <p>Fixed problem where deleting an event from an event item in the Admin calendar caused worshipschedule to crash. The event would be deleted, but an error was caused when the application attempted to display the calendar.</p> "
	str = str & "<h3>Fixed: No state dropdown provided in form to upgrade account from trial</h3> <p>Fixed problem where the dropdown listing for US States was missing from some website edit pages.</p> "
	str = str & "<h3>Fixed: Files for events are listed with incorrect icons on member calendar</h3> <p>Fixed problem where file lists for events on the member calendar were displaying multiple files with the same (incorrect) file icon for the file extension.</p> "
	str = str & "<h3>Fixed: Member disappears from available/unavailable lists after being removed from schedule</h2> <p>Fixed problem where an available member could be removed from an event team and then would not be displayed in the available member dropdown list until the schedule was published.</p>	"
	str = str & "</div>"
	
	ReleaseNotesToString = str
End Function

%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<%
Class cPage
	Public MessageID
	Public Url
	
	
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
	
	Private Sub Class_Initialize()
		Url = Request.ServerVariables("URL")
	End Sub
End Class
%>

