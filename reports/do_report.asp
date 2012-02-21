<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Dim m_bodyText

Sub OnPageLoad(ByRef page)
	Dim sess			: Set sess = New cSession
	sess.SessionID = Request.Cookies("sid")
	Call CheckSession(sess, PERMIT_LEADER)
	
	page.MessageID = Request.QueryString("msgid")
	page.Action = Decrypt(Request.QueryString("act"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	
	If Len(page.ProgramID) > 0 Then
		page.DefaultColumnList = "[Last name],[First name],[Email],[Home phone],[Mobile phone],[Alternate phone],[Address],[Address(2)],[City],[State],[Zip],[Gender],[Birthdate],[Start date],[Last login],[Enabled]"
	Else
		page.DefaultColumnList = "[Last name],[First name],[Email],[Home phone],[Mobile phone],[Alternate phone],[Address],[Address(2)],[City],[State],[Zip],[Gender],[Birthdate],[Signup date],[Last login],[Enabled]"
	End If

	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/inside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<script src="http://www.google.com/jsapi" type="text/javascript" language="javascript"></script>
		<script language="javascript" type="text/javascript">
			google.load("jquery", "1.2.6");
		</script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/tablesorter/jquery.tablesorter.min.js"></script>

		<script language="javascript" type="text/javascript">
			$(document).ready(function(){ 
			
				// set up tablesorter plugin
				$.tablesorter.defaults.widgets = ['zebra']; 
				$("#report-table").tablesorter(); 
			}); 
		</script>

		<link href="../_incs/script/jquery/plugins/tablesorter/themes/blue/style.css" rel="stylesheet" type="text/css" />		
		<title>Report</title>
	</head>
	<body><%=m_bodyText%></body>
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)
	
	Select Case page.Action
		Case DISPLAY_PRINT_VIEW
			str = PrintViewToString(page, Request.Form("column_list"))
			
		Case STREAM_FILE_TO_BROWSER
			Call StreamFileAsExcel(page, Request.Form("column_list"))
		
		Case Else
			Call Err.Raise(vbObjectError + 100, "Main()", "ASSERT: Unexpectedly reached else clause in switch statement. ")
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub StreamFileAsExcel(page, selectedColumns)
	Dim str
	
	Dim report			: Set report = New cReport
	
	report.ClientId	= page.Client.ClientID
	report.ProgramID = page.Program.ProgramID
	report.ColumnList = page.defaultColumnList
	
	str = report.ToString(selectedColumns)

	Response.Clear
	Response.ContentType = "application/vnd.ms-excel"
	
	Response.Write str
	Response.Flush 	
	Response.End	
End Sub

Function PrintViewToString(page, selectedColumns)
	Dim str
	
	Dim report			: Set report = New cReport
	report.ClientId	= page.Client.ClientID
	report.ProgramID = page.Program.ProgramID
	report.ColumnList = page.defaultColumnList
	
	str = report.ToString(selectedColumns)

	PrintViewToString = str
End Function
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/report_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID

	' encrypted
	Public Action
	Public ProgramID
	
	' not persisted
	Public DefaultColumnList

	' objects
	Public Member
	Public Client
	Public Program	
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		
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

		c.Action = Action
		c.ProgramID = ProgramID
		
		c.DefaultColumnList = DefaultColumnList
				
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		
		Set Clone = c
	End Function
End Class
%>

