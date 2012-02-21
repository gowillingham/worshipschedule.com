<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%

' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-reports"
Dim m_pageHeaderText	: m_pageHeaderText = "&nbsp;"
Dim m_impersonateText	: m_impersonateText = ""
Dim m_pageTitleText		: m_pageTitleText = ""
Dim m_topBarText		: m_topBarText = "&nbsp;"
Dim m_bodyText			: m_bodyText = ""
Dim m_tabStripText		: m_tabStripText = ""
Dim m_tabLinkBarText	: m_tabLinkBarText = ""
Dim m_appMessageText	: m_appMessageText = ""
Dim m_acctExpiresText	: m_acctExpiresText = ""

Sub OnPageLoad(ByRef page)
	Dim sess			: Set sess = New cSession
	sess.SessionID = Request.Cookies("sid")
	Call CheckSession(sess, PERMIT_LEADER)
	
	page.MessageID = Request.QueryString("msgid")
	page.Action = Decrypt(Request.QueryString("act"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))

	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	' check for program dropdown postback
	If Request.Form("form_program_dropdown_is_postback") = IS_POSTBACK Then
		page.ProgramID = Request.Form("new_program_id")
		
		' hack: I need to redirect this page when a new program is selectd 
		' or the javascript that builds the print/download buttons break.
		Response.Redirect(page.Url & page.UrlParamsToString(False))
	End If	
	
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()

	' set the view tokens
	m_appMessageText = ApplicationMessageToString(page.MessageID)
	page.MessageID = ""
	Call SetTopBar(page)
	Call SetPageHeader(page)
	Call SetPageTitle(page)
	Call SetTabLinkBar(page)
	Call SetTabList(m_pageTabLocation, page)
	Call SetImpersonateText(sess)
	Call SetAccountNotifier(sess)
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/inside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<!--#INCLUDE VIRTUAL="/_incs/script/javascript/javascript_server_variable_wrapper.asp"-->
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
				
				// highlight clicked row
				$("#report-table tr").click(function(){
					$(this).toggleClass("active");
				});
				
				// initial state is all selected ..
				$("#column-list :checkbox").each(function(){
					$(this).attr("checked", true);
				});
				$("#column-list :checkbox").each(function(idx){
					$(this).click(function(){
						// hide or show column
						if (this.checked) {
							$("td:nth-child(" + (idx + 1) + "), th:nth-child(" + (idx + 1) + ")").show()
						}
						else {
							$("td:nth-child(" + (idx + 1) + "), th:nth-child(" + (idx + 1) + ")").hide()
						};
					});
				});
				
				// attach click event print/download button
				var formAction = "/reports/do_report.asp"
				$("#print-button a, #excel-button a").each(function(){
					var button = this
					
					$(button).click(function(){
					
						// figure out which button was clicked
						var parm
						if ($(this).parent().attr("id") == "print-button") {
							parm = <%= "'" & Encrypt(DISPLAY_PRINT_VIEW) & "'" %>;		// HACK: inject server side constant in javascript code ..
						}
						else {
							parm = <%= "'" & Encrypt(STREAM_FILE_TO_BROWSER) & "'" %>;	// HACK: inject server side constant in javascript code ..
						};
						
						if (serverVars.query_string.length > 0) {
							$("#form-column-list").attr("action", formAction + "?" + serverVars.query_string + "&act=" + parm); 
						}
						else {
							$("#form-column-list").attr("action", formAction + "?act=" + parm); 
						} ;
						$("#form-column-list").submit();
						return false;
					});
				});
				
				// attach onchange to program dropdown ..
				$("#program-dropdown").change(function(){
					$("#form-program-dropdown").submit();
				});
			}); 
		</script>

		<link href="/_incs/script/jquery/plugins/tablesorter/themes/blue/style.css" rel="stylesheet" type="text/css" />		
		<style type="text/css">
			.message, .details, .report-container {width:622px;}
		</style>
		<title><%=m_pageTitleText%></title>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)
	
	Select Case page.Action
		Case STREAM_FILE_TO_BROWSER
			Response.Redirect("/client/reports_print.asp" & page.UrlParamsToString(True))
			
		Case DISPLAY_PRINT_VIEW
			Response.Redirect("/client/reports_print.asp" & page.UrlParamsToString(True))

		Case Else
			str = str & ReportListingToString(page)
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Function ColumnListSelectorToString(columnList)
	Dim str, i
	Dim colList			: colList = Split(Replace(Replace(columnList, "[", ""), "]", ""), ",")
	
	For i = 0 To UBound(colList)
		str = str & "<li><input type=""checkbox"" name=""column_list"" value=""" & i & """ checked=""checked"" />"
		str = str & html(colList(i)) & "</li>"
	Next
	
	ColumnListSelectorToString = str
End Function

Function ReportListingToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim header			: header = ""
	
	Dim report			: Set report = New cReport
	report.ClientId	= page.Client.ClientID
	report.ProgramId = page.Program.ProgramID
	
	If Len(page.Program.ProgramID) > 0 Then
		header = "<h3 style=""padding-top:0;margin-top:0;"">" & html(page.Program.ProgramName) & " Member Report</h3>"
		report.ColumnList = "[Last name],[First name],[Email],[Home phone],[Mobile phone],[Alternate phone],[Address],[Address(2)],[City],[State],[Zip],[Gender],[Birthdate],[Start date],[Last login],[Enabled]"
	Else
		header = "<h3 style=""padding-top:0;margin-top:0;"">" & html(page.Client.NameClient) & " Member Report</h3>"
		report.ColumnList = "[Last name],[First name],[Email],[Home phone],[Mobile phone],[Alternate phone],[Address],[Address(2)],[City],[State],[Zip],[Gender],[Birthdate],[Signup date],[Last login],[Enabled]"
	End If
	
	str = str & "<div class=""tip-box""><h3>Info</h3>"
	str = str & "<p>Select the columns to be included. </p>"
	str = str & "<form method=""post"" action=""" & pg.Url & pg.UrlParamsToString(True) & """ id=""form-column-list"">"
	str = str & "<ul id=""column-list"">"
	str = str & ColumnListSelectorToString(report.ColumnList)
	str = str & "</ul></form></div>"

	str = str & header
	str = str & report.ToString("")
	
	ReportListingToString = str
End Function

Function ProgramDropdownToString(page)
	Dim str, i
	
	Dim pg						: Set pg = page.Clone()
	Dim defaultText				: defaultText = "< Select a program >"
	Dim selected				: selected = ""
	
	Dim list					: list = page.Member.OwnedProgramsList()
	If Not IsArray(list) Then Exit Function
	
	If Len(page.Program.ProgramID) > 0 Then defaultText = "< Show all >"
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-program-dropdown"">"
	str = str & "<input type=""hidden"" name=""form_program_dropdown_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""new_program_id"" id=""program-dropdown"">"
	str = str & "<option value="""">" & html(defaultText) & "</option>"
	For i = 0 To UBound(list,2)
		selected = ""
		If CStr(list(0,i)) = CStr(page.Program.ProgramID) Then selected = " selected=""selected"""
		str = str & "<option value=""" & list(0,i) & """" & selected & ">" & html(list(1,i)) & "</option>"
	Next
	str = str & "</select></form></li>"
	
	ProgramDropdownToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	If Len(page.Program.ProgramId) > 0 Then
		str = str & "<a href=""/reports/default.asp"">Reports</a> / "
		str = str & html(page.Program.ProgramName)
	Else 
		str = str & "Reports"
	End If
	
	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Dim printButton
	href = pg.Url & pg.UrlParamsToString(True)
	printButton = "<li id=""print-button""><a href=""" & href & """><img class=""icon"" src=""/_images/icons/printer.png"" alt="""" /></a><a href=""" & href & """>Print</a>"
	
	Dim excelButton
	href = pg.Url & pg.UrlParamsToString(True)
	excelButton = "<li id=""excel-button""><a href=""" & href & """><img class=""icon"" src=""/_images/icons/page_white_excel.png"" alt="""" /></a><a href=""" & href & """>Download</a>"
	
	Select Case page.Action
		Case Else
			str = str & ProgramDropdownToString(page)
			str = str & printButton
			str = str & excelButton
	End Select

	m_tabLinkBarText = str
End Sub
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
				
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		
		Set Clone = c
	End Function
End Class
%>

