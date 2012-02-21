<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Const	FILE_NAME = "members.txt"
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-members"
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
	page.MemberIDList = Decrypt(Request.QueryString("midl"))
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
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
		<link rel="stylesheet" type="text/css" href="importmembers.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=m_pageTitleText%></title>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	Dim uploader	: Set uploader = Server.CreateObject("ASPSmartUpload.SmartUpload")	
	Dim memberIDList
	Dim errors
	
	Call OnPageLoad(page)
	
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		
		Call uploader.Upload()
		
		' reset/reload page.Program ..
		If CStr(uploader.Form.Item("ProgramID")) <> CStr(page.ProgramID) Then
			page.ProgramID = uploader.Form.Item("ProgramID")
			page.Program.ProgramID = page.ProgramID
			If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()
		End If
		
		If uploader.Form.Item("Submit") = "Import" Then
			If ValidFormImport(uploader) Then
				Call TestImport(page, uploader, errors, rv)
				Select Case rv
					Case 0
						Call DoImport(page, uploader, memberIDList, rv)
						page.MemberIDList = memberIDLIst
						Call SendImportReportByEmail(page)
						page.MessageID = "": page.Action = SHOW_IMPORT_REPORT
						Response.Redirect(page.Url & page.UrlParamsToString(False))
					Case -1
						' file has no lines
						page.MessageID = 1059: page.Action = "": page.IsFormPost = ""
						Response.Redirect(page.Url & page.UrlParamsToString(False))
					Case -2
						' problem with at least one line ..
						str = str & ErrorReportToString(errors) 
					Case -3
						' problem with the first line in import file
						page.MessageID = 1060: page.Action = ""
						Response.Redirect(page.Url & page.UrlParamsToString(False))
					Case Else
						Call err.Raise(vbObjectError + 1, "Main()", "ASSERT! Else condition reached in TestImport() Select Case")
				End Select				
			End If
		End If
	End If
	
	Select Case page.Action
		Case SHOW_IMPORT_REPORT
			str = str & ImportReportToString(page, rv)
			
		Case Else
			str = str & FormImportToString(page, uploader)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
	Set uploader = Nothing
End Sub

Function ErrorReportToString(errors)
	Dim str
	Dim msg
	
	msg = msg & "There is a problem with the file that you provided. "
	msg = msg & "Please check the following lines in the file you are trying to import for errors. "
	msg = msg & "<br />" & errors
	
	str = str & CustomApplicationMessageToString("Sorry, there seems to be a problem with your import file! ", msg, "Error")

	ErrorReportToString = str
End Function

Function ImportReportToString(p, outError)
	Dim str, i
	Dim client			: Set client = New cClient
	Dim msg				: msg = ""
	
	Dim list			: list = GetMembersFromList(p.MemberIdList)
	If Not IsArray(list) Then
		outError = -1
		Exit Function
	End If
	Dim count			: count = UBound(list,2) + 1
	outError = 0
	
	msg = msg & "You added " & count & " members " 
	If Len(p.ProgramID) > 0 Then 
		msg = msg & " to the " & p.Program.ProgramName & " program "
	End If
	msg = msg & "for your " & client.NameClient & " " & Application.Value("APPLICATION_NAME") & " account. "
	msg = msg & "Find below a listing of the members imported and their temporary login credentials for " & Application.Value("APPLICATION_NAME") & ". "
	msg = msg & "You may wish to save this page or print a paper copy for reference. "
	
	str = str & CustomApplicationMessageToString("Your import was successful!", msg, "Confirm")
	
	' 0-NameLast 1-NameFirst 2-Email 3-NameLogin 4-PWord

	str = str & "<div class=""grid""><table>"
	str = str & "<tr class=""header""><th scope=""col"" style=""width:1%;"">&nbsp;</th><th scope=""col"">Member</th><th scope=""col"">Email</th><th scope=""col"">Username</th></tr>"
	For i = 0 To UBound(list,2)
		str = str & "<tr><td>" & i+1 & "</td>"
		str = str & "<td><strong>" & HTML(list(1,i) & ", " & list(0,i)) & "</strong></td>"
		str = str & "<td>" & HTML(list(2,i)) & "</td>"
		str = str & "<td>" & HTML(list(3,i)) & "</td></tr>"
	Next		
	str = str & "</table></div>"

	ImportReportToString = str
	Set client = Nothing
End Function

Sub SendImportReportByEmail(page)
	Dim i
	Dim subject			: subject = ""
	Dim body			: body = ""
	Dim list			: list = GetMembersFromList(page.MemberIDList)
	
	Dim count			: count = 0
	
	If IsArray(list) Then 
		count = UBound(list,2) + 1
	Else
		' no members imported so exit
		Exit Sub
	End If
	
	Dim email			: Set email = New cEmailSender
	
	subject = "[" & Application.Value("APPLICATION_NAME") & "] " & page.Member.ClientName & " Import Members Confirmation"
	
	body = body & "Hello " & page.Member.NameFirst & " " & page.Member.NameLast
	body = body & vbCrLf & vbCrLf & "This email is confirmation of the successful import of " & count & " members into your " & page.Member.ClientName & " " & Application.Value("APPLICATION_NAME") & " account. "
	body = body & Application.Value("APPLICATION_NAME") & " created a login (user name and password) for each of your new members, and automatically sent them their login information by email. "
	body = body & "A list of your newly imported members follows this message. You may wish to save it for your records. "
	body = body & vbCrLf & vbCrLf & "If you did not add your new members to one of your programs when you imported them, you'll need to add them to a program and assign program skills to their profile. "
	body = body & "You can manage your member programs and skills here: "
	body = body & vbCrLf & "http://" & Request.ServerVariables("SERVER_NAME") & "/admin/programs.asp"
	body = body & vbCrLf & vbCrLf & "These Members Were Imported .."
	body = body & vbCrLf & String(60, "-")
	For i = 0 To UBound(list,2)
		body = body & vbCrLf & i+1 & ". " & list(1,i) & ", " & list(0,i) & " [email:" & list(2,i) & ", login:" & list(3,i) & "]"
	Next
	body = body & vbCrLf & EmailDisclaimerToString(page.Member.ClientName)
	
	Call email.SendMessage(page.Member.Email, page.Member.Email, subject, body)
	
	Set email = Nothing
End Sub

Function GetMembersFromList(sIDList)
	Dim cnn, rs
	If Len(sIDList) = 0 Then Exit Function
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	cnn.Open Application.Value("CNN_STR")
	cnn.up_memberGetMemberDetailsByIDList CStr(sIDList), rs
	If Not rs.EOF Then GetMembersFromList = rs.GetRows()
	
	rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
End Function

Sub SendMemberWelcomeByEmail(m, fromAddress, creatorName)
	Dim subject
	Dim body
	Dim email			: Set email = New cEmailSender
	
	subject = "[" & Application.Value("APPLICATION_NAME") & "] " & m.ClientName & " - Welcome to " & Application.Value("APPLICATION_NAME")

	body = body & "Hello " & m.NameFirst & " " & m.NameLast
	body = body & vbCrLf & vbCrLf & "Welcome to " & Application.Value("APPLICATION_NAME") & " Web Scheduling. "
	body = body & "A new " & Application.Value("APPLICATION_NAME") & " account for " & m.ClientName & " has been created for you by " & creatorName & ". "
	body = body & "You should receive the login credentials (username and password) for your new account in a separate email message. "
	body = body & "You may wish to print and save this message for future reference. "
	
	body = body & vbCrLf & vbCrLf & "You will find the " & Application.Value("APPLICATION_NAME") & " login page here "
	body = body & vbCrLf & "http://" & Request.ServerVariables("SERVER_NAME") & "/member/login.asp"
	
	body = body & vbCrLf & vbCrLf & "Enter your credentials exactly as provided (your password is case-sensitive). "
	body = body & "After logging in for the first time, you may change your login name and/or password to something easier for you to remember. "
	body = body & "You may contact " & Application.Value("APPLICATION_NAME") & " techical support at mailto:" & Application.Value("SUPPORT_EMAIL_ADDRESS") & " with any questions or problems. "
	
	body = body & vbCrLf & vbCrLf & "For " & Application.Value("APPLICATION_NAME") & " help go to "
	body = body & vbCrLf & "http://" & Request.ServerVariables("SERVER_NAME") & "/help/help.asp"
	
	body = body & vbCrLf & vbCrLf & Application.Value("APPLICATION_NAME") & " technical support and assistance may be found here "
	body = body & vbCrLf & "http://" & Request.ServerVariables("SERVER_NAME") & "/support.asp"
	
	body = body & vbCrLf & vbCrLf & "Thanks for joining " & Application.Value("APPLICATION_NAME") & "!!"
	body = body & vbCrLf & EmailDisclaimerToString(m.ClientName)
	
	Call email.SendMessage(m.Email, fromAddress, subject, body)
	Set email = Nothing
End Sub

Sub DoImport(p, uploader, idList, outError)
	Dim str, rv, i
	Dim fso				: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim path			: path = Application.Value("IMPORT_MEMBER_FILE_DIRECTORY") & p.Member.MemberID & FILE_NAME
	Dim cnn				: Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs				: Set rs = Server.CreateObject("ADODB.Recordset")
	Dim member			
	Dim memberList
	
	Dim connection		: connection = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & Application.Value("IMPORT_MEMBER_FILE_DIRECTORY") & ";Extensions=asc,csv,tab,txt;Extended Properties=""text;HDR=NO"""
	Dim sql				: sql = "SELECT FirstName, LastName, Email FROM " & p.Member.MemberID & FILE_NAME
	outError = 0

	' save the file
	uploader.Files(1).SaveAs(path)
	
	' open import file into rs
	cnn.Open(connection)
	Set rs = cnn.Execute(sql)
	If Not rs.EOF Then memberList = rs.GetRows()
	
	For i = 0 To UBound(memberList,2)
		Set member = New cMember
		member.NameFirst = memberList(0,i)
		member.NameLast = memberList(1,i)
		member.Email = memberList(2,i)
		member.ClientID = p.Client.ClientID
		Call member.QuickAdd(p.ProgramID, rv)
		If rv = 0 Then
			idList = idList & member.MemberID & ","
			member.Load()
			Call SendCredentials(member.MemberID, p.Member.Email)
			Call SendMemberWelcomeByEmail(member, p.Member.Email, p.Member.NameFirst & " " & p.Member.NameLast) 
		End If
		Set member = Nothing
	Next
	If Len(idList) > 0 Then idList = Left(idList, Len(idList)-1)
	
	' clean up so I can delete file
	If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
	
	' delete the file
	If fso.FileExists(path) Then fso.DeleteFile path, True
	
	Set fso = Nothing
	Set member = Nothing
End Sub

Sub TestImport(p, uploader, errorText, outError)
	Dim str, rows
	Dim fso				: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim path			: path = Application.Value("IMPORT_MEMBER_FILE_DIRECTORY") & p.Member.MemberID & FILE_NAME
	Dim cnn				: Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs				: Set rs = Server.CreateObject("ADODB.Recordset")

	Dim connection		: connection = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & Application.Value("IMPORT_MEMBER_FILE_DIRECTORY") & ";Extensions=asc,csv,tab,txt;Extended Properties=""text;HDR=NO"""
	Dim sql				: sql = "SELECT FirstName, LastName, Email FROM " & p.Member.MemberID & FILE_NAME
	outError = 0
	
	' save the file
	uploader.Files(1).SaveAs(path)
	
	' open import file into rs
	cnn.Open(connection)
	
	' err.Number='-2147217904'
	' ------------------------------
	' this error number is thrown when the first row ..
	'	- does not have three columns
	'	- columns are mis-named
	
	On Error Resume Next
	Set rs = cnn.Execute(sql)
	If err.number = -2147217904 Then
		' problem with column name row in file
		outError = -3
		On Error GoTo 0
		Exit Sub
	End If 
	On Error GoTo 0
	
	' check for errors
	If rs.EOF Then
		' no rows in the file
		outError = -1
		Exit Sub
	End If
	
	' build string of error rows ..
	Dim newMemberList			: newMemberList = rs.GetRows()
	Call IsValidData(newMemberList, rows)
	
	If Len(rows) > 0 Then
		' at least one row doesn't have valid data
		outError = -2
		str = str & rows 
		errorText = str
	End If
	
	' clean up so I can delete file
	If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
	
	' delete the file
	If fso.FileExists(path) Then fso.DeleteFile path, True
	
	Set fso = Nothing
End Sub

Function IsValidData(ByVal memberList, ByRef report)
	Dim i, j, str
	
	'0-NameFirst 1-NameLast 2-Email
	If IsArray(memberList) Then
		For i = 0 To UBound(memberList,2)
			str = ""
			
			If Len(memberList(0,i) & "") > 50 Then 
				str = str & "First Name (" & HTML(memberList(0,i)) & ") is too long. " 
			End If
			If Len(memberList(0,i) & "") = 0 Then
				str = str & "First Name is missing. "
			End If

			If Len(memberList(1,i) & "") > 50 Then 
				str = str & "Last Name (" & HTML(memberList(1,i)) & ") is too long. " 
			End If
			If Len(memberList(1,i) & "") = 0 Then
				str = str & "Last Name is missing. "
			End If

			If Len(memberList(2,i) & "") > 100 Then 
				str = str & "Email address (" & HTML(memberList(2,i)) & ") is too long. "
			End If
			If Len(memberList(2,i) & "") = 0 Then
				str = str & "Email address is missing. "
			End If
			If Not IsEmail(memberList(2,i) & "") Then
				str = str & "Email address (" & HTML(memberList(2,i)) & ") does not seem to be in an accepted format. "
			End If
			
			If Len(str) > 0 Then
				' line number reported as i+1+1 to account for 0-based array index 
				' and first row of file is column headers
				report = report & "<br />Line " & i+1+1 & ": " & str
			End If
		Next
	End If
End Function

Function ValidFormImport(uploader)
	ValidFormImport = True
	
	'check for file provided
	If uploader.Files(1).IsMissing Then
		ValidFormImport = False
		AddCustomFrmError("No file was provided for the import. A file is required.")
	End If
End Function

Function FormImportToString(page, uploader)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>Find step-by-step instructions for importing from a file into " & Application.Value("APPLICATION_NAME") & " <a href=""/help/topic.asp?hid=1#anchor-import"" target=""_blank"">here</a>. </p></div>"
	
	str = str & "<div class=""tip-box""><h3>Still having trouble?</h3>"
	str = str & "<p>We'd be happy to import your file for you. " 
	str = str & "Contact <a href=""/support.asp"">support</a> for more help. </p></div>"
	
	str = str & m_appMessageText
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form enctype=""multipart/form-data"" method=""post"" action=""" & pg.Url & pg.UrlParamsToString(True) & """ id=""formImport"">"
	str = str & "<input type=""hidden"" name=""FormImportIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><tr><td class=""label"">" & RequiredElementToString(True, "Import File") & "</td>"
	str = str & "<td><input class=""file"" type=""file"" name=""FileName"" size=""40"" style=""width:300px;"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "The file (on your computer) your members will be imported from."
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Program</td>"
	str = str & "<td>" & ProgramDropdownToString(page) & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "If a program is selected, your members will be placed into that program <br />automatically as they are imported. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Import"" />"
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a></td></tr>"
	
	str = str & "</table></form></div>"

	FormImportToString = str
End Function

Function ProgramDropdownToString(page)
	Dim str
	Dim list			: list = GetProgramList(page.Member.MemberID)
	
	str = str & "<select name=""ProgramID"">"
	str = str & "<option value="""">&nbsp;</option>"
	str = str & SelectOption(list, page.ProgramID)
	str = str & "</select>"
	
	ProgramDropdownToString = str
End Function

Function GetProgramList(memberID)
	Dim cnn					: Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs					: Set rs = Server.CreateObject("ADODB.Recordset")
	
	cnn.Open Application.Value("CNN_STR")
	cnn.up_memberGetProgramListByMemberID CLng(memberID), rs
	If Not rs.EOF Then GetProgramList = rs.GetRows()
	
	If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
End Function

Sub SetPageHeader(page)
	Dim str
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	str = str & "<a href=""/admin/members.asp"">Members</a> / "
	str = str & "Import"

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href		: href = ""
	
	Dim memberListButton
	pg.Action = ""
	href = "/admin/members.asp" & pg.UrlParamsToString(True)
	memberListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/table.png"" /></a><a href=""" & href & """>Member List</a></li>"
	
	str = str & memberListButton
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_select_option.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public ProgramID
	
	' not persisted
	Public MemberIDList
		
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
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(programID) & amp
		If Len(MemberIDList) > 0 Then str = str & "midl=" & Encrypt(MemberIDList) & amp
		
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
		c.MemberIDLIst = MemberIDList
		c.Action = Action
		c.ProgramID = ProgramID
		
		Set c.Member = Member
		Set c.Client = Client
		
		Set Clone = c
	End Function
End Class
%>

