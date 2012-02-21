<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "files"
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
	Call CheckSession(sess, PERMIT_MEMBER)
	
	page.MessageID = Request.QueryString("msgid")
	page.SortBy = Request.QueryString("sb")
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.Action = Decrypt(Request.QueryString("act"))
	page.FileID = Decrypt(Request.QueryString("fid"))
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	If Request.Form("FormSortByDropdownIsPostback") = IS_POSTBACK Then
		page.SortBy = Request.Form("SortBy")
	End If

	If Request.Form("FormProgramDropdownIsPostback") = IS_POSTBACK Then
		page.ProgramID = Request.Form("ProgramID")
	End If
	
	Set page.Program =  New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then page.Program.Load()
	
	Set page.File = New cFile
	page.File.FileID = page.FileID
	If Len(page.File.FileID) > 0 then Call page.File.Load()
	
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
		<link rel="stylesheet" type="text/css" href="files.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<style type="text/css">
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

	str = str & ApplicationMessageToString(page.MessageID)
	page.MessageID = ""

	Select Case page.Action
		Case STREAM_FILE_TO_BROWSER
			Call page.File.StreamFile(page.Member.MemberID, rv)
			Response.End
		
		Case SHOW_DETAILS
			str = str & FileDetailsToString(page)
			
		Case Else
			str = str & FileGridToString(page)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Function FileDetailsToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim dateTime		: Set dateTime = New cFormatDate
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	pg.Action = STREAM_FILE_TO_BROWSER
	str = str & "<ul><li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>"
	str = str & "Download " & html(page.File.FriendlyName & "." & page.File.FileExtension) & "</a></li>"
	str = str & "</ul></div>"
	
	str = str & "<h3>" & html(page.File.FriendlyName & "." & page.File.FileExtension) & "</h3>"
	
	str = str & "<div class=""summary"">"
	str = str & "<h5 class=""description"">Description</h5>"
	If Len(page.File.Description) > 0 Then
		str = str & "<p>" & html(page.File.Description) & "</p>"
	Else
		str = str & "<p class=""alert"">No description is available. </p>"
	End If
	
	str = str & "<h5 class=""file-download"">Download this file</h5>"
	pg.Action = STREAM_FILE_TO_BROWSER
	str = str & "<ul><li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """><strong>" & html(page.File.FriendlyName & "." & page.File.FileExtension) & "</strong></a></li></ul>"

	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul>"
	str = str & "<li>" & FileSizeToString(page.File.FileSize) & "</li>"
	str = str & "<li>Uploaded on " & dateTime.Convert(page.File.DateCreated, "DDD MMM dd, YYYY around hh:00 pp") & "</li>"
	str = str & "<li>"
	If page.File.DownloadCount = 0 Then
		str = str & "Never downloaded. "
	ElseIf page.File.DownloadCount = 1 Then
		str = str & "Downloaded once. "
	Else
		str = str & "Downloaded about " & page.File.DownloadCount & " times. "
	End If
	str = str & "</li>"
	str = str & "</ul>"
	
	str = str & "</div>"

	FileDetailsToString = str
End Function

Function NoFilesDialogToString(page)
	Dim dialog				: Set dialog = New cDialog
	Dim pg					: Set pg = page.Clone()
	
	dialog.Headline = "Looking for a " & html(page.Program.ProgramName) & " file?"
	
	dialog.Text = dialog.Text & "<p>This page shows a list of all public files for the " & html(page.Client.NameClient) & " " & Application.Value("APPLICATION_NAME") & " account that have been made available for you to download. "
	If Len(page.Program.ProgramId) = 0 Then 
		dialog.Text = dialog.Text & "It looks like either no files have been uploaded for your account, or they have all been set to private by an account administrator. "
	Else
		dialog.Text = dialog.Text & "It looks like either no files have been uploaded for the " & html(page.Program.ProgramName) & " program, or they have all been set to private by an account administrator. </p>"
	End If

	dialog.SubText = dialog.SubText & "<p>You might try checking your calendar for the file you are looking for. "
	dialog.SubText = dialog.SubText & "You can download files that have been set to private if they have been linked to an event and you are scheduled (assigned to the event team) for that event. </p>"
	dialog.SubText = dialog.SubText & "<p>Go to your calendar and check the events you've been scheduled for to see if you can download the file you need that way. "
	dialog.SubText = dialog.SubText & "</p>"

	If Len(page.Program.ProgramId) > 0 Then
		pg.ProgramId = ""
		dialog.LinkList = dialog.LinkList & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Show files for all programs</a></li>"
	End If
	dialog.LinkList = dialog.LinkList & "<li><a href=""/member/schedules.asp"">Go to my calendar</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/member/contacts.asp"">Email an administrator</a></li>"

	NoFilesDialogToString = dialog.ToString()
End Function

Function FileGridToString(page)
	Dim str, i
	Dim pg			: Set pg = page.Clone()
	Dim display		: set display = New cFileDisplay
	Dim iconPath	: iconPath = ""
	Dim programName	: programName = ""
	Dim entityText	: entityText = ""
	
	Dim alt			: alt = ""
	Dim count		: count = 0
	
	Dim isOwned		: isOwned = True
	Dim isPublic	: isPublic = True
	
	page.File.ClientID = page.Client.ClientID
	page.File.ProgramID = page.Program.ProgramID
	Dim list		: list = page.File.List(LookupSortParam(page.SortBy))
	Dim programList		: programList = page.Member.ProgramList()
	
	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "<p>To download a file to your computer, <strong>click download</strong> in the toolbar for a file. </p></div>"
	
	' 0-FileID 1-FileName 2-FriendlyName 3-Description 4-ClientID 5-ProgramID 6-FileOwnerID
	' 7-DateCreated 8-DateModified 9-FileExtension 10-FileSize 11-MIMEType 12-MIMESubType 13-EventFileCount
	' 14-IsPublic 15-DownloadCount 16-ProgramName

	str = str & m_appMessageText
	If IsArray(list) Then
		str = str & "<h3>" & html(page.Client.NameClient) & " files</h3>"
		If Len(page.Program.ProgramId) > 0 Then 
			str = str & "<h4 class=""first"">" & Server.HtmlEncode(page.Program.ProgramName) & "</h4>"
			str = str & "<p>This list contains all the files available from your " & Application.Value("APPLICATION_NAME") & " account for the <strong>" & Server.HtmlEncode(page.Program.ProgramName) & "</strong> program. "
			str = str & "Click <strong>Download</strong> in the toolbar for any file to save it to your computer. </p>"
		Else
			str = str & "<h4 class=""first"">All files</h4>"
			str = str & "<p>This list contains all the files available from your " & Application.Value("APPLICATION_NAME") & " account. "
			str = str & "Click <strong>Download</strong> in the toolbar for any file to save it to your computer. </p>"
		End If
		
		str = str & "<div class=""grid"">"
		str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
		str = str & "<input type=""checkbox"" name=""master"" onclick="""" disabled=""disabled"" checked=""checked"" /></th>"
		str = str & "<th scope=""col"">File</th><th scope=""col"">Program</th><th>Size</th><th scope=""col""></th></tr>"
		For i = 0 To UBound(list,2)
			isOwned = False
			If IsMemberProgram(list(5,i), programList) Then isOwned = True
			isPublic = False
			If list(14,i) = 1 Then isPublic = True
		
			If isOwned And isPublic Then
				alt = ""			: If count Mod 2 > 0 Then alt = " class=""alt"""
				count = count + 1
				iconPath = display.GetIconPath(list(9,i))
				str = str & "<tr" & alt & "><td><input type=""checkbox"" name=""Unused"" disabled=""disabled"" checked=""checked"" /></td>"
				str = str & "<td><img class=""icon"" src=""" & iconPath & """ alt=""icon"" />"
				str = str & "<strong>" & html(list(2,i) & "." & list(9,i)) & "</strong></td>"
				programName = "&nbsp;"
				If Len(list(16,i)) > 0 Then programName = html(list(16,i))
				str = str & "<td>" & programName & "</td>"
				str = str & "<td>" & FileSizeToString(list(10,i)) & "</td>"
				str = str & "<td class=""toolbar"">"
				pg.FileID = list(0,i): pg.Action = SHOW_DETAILS
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details""><img src=""/_images/icons/magnifier.png"" alt="""" /></a>"
				pg.FileID = list(0,i): pg.Action = STREAM_FILE_TO_BROWSER
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Download""><img src=""/_images/icons/page_swoosh_down.png"" alt="""" /></a>"
				str = str & "</td></tr>"
			End If
		Next
		str = str & "</table></div>"
	End If
	
	' show message if no files are returned
	If count = 0 Then
		str = NoFilesDialogToString(page)
	End If
		
	FileGridToString = str
End Function

Function IsMemberProgram(programID, programList)
	Dim i
	IsMemberProgram = False
	
	If Not IsArray(programList) Then Exit Function
	If Len(programID & "") = 0 Then 
		IsMemberProgram = True
		Exit Function
	End If
	
	For i = 0 To UBound(programList,2)
		If CStr(programID & "") = CStr(programList(0,i) & "") Then
			IsMemberProgram = True
			Exit For
		End If
	Next
End Function

Function FileSizeToString(ByVal val)
	Dim str
	val = CLng(val)
		
	If val < 1000 Then
		str = "0 KB"
	ElseIf (val > 999) And (val < 1000000) Then
		str = val / 1000
		str = FormatNumber(str, 2, , , False) & " KB"
	Else
		str = val/1000000
		str = FormatNumber(str, 2, , , False) & " MB"
	End If
	
	FileSizeToString = str
End Function

Function LookupSortParam(val)
	Dim str
	
	Select Case val
		Case SORT_BY_FILE_NAME
			str = "FriendlyName"
		Case SORT_BY_FILE_EXTENSION 
			str = "FileExtension"
		Case SORT_BY_FILE_SIZE
			str = "FileSize"
		Case SORT_BY_PROGRAM_NAME
			str = "ProgramName"
		Case Else
			str = ""
	End Select
	
	LookupSortParam = str
End Function

Function SortByDropdownToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formSortByDropdown"">"
	str = str & "<input type=""hidden"" name=""FormSortByDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""SortBy"" onchange=""document.forms.formSortByDropdown.submit();"">"
	str = str & "<option value="""">" & html("< Sort by .. >") & "</option>"
	str = str & "<option value=""" & SORT_BY_FILE_NAME & """>File Name</option>"
	str = str & "<option value=""" & SORT_BY_FILE_EXTENSION & """>Extension</option>"
	str = str & "<option value=""" & SORT_BY_FILE_SIZE & """>Size</option>"
	str = str & "<option value=""" & SORT_BY_PROGRAM_NAME & """>Program</option>"
	str = str & "</select></form></li>"	
	
	SortByDropdownToString = str
End Function

Function ProgramDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = page.Member.ProgramList()
	Dim isSelected		: isSelected = ""
	
	Dim defaultText		: defaultText = "< Select a program >"
	If Len(page.Program.ProgramID) > 0 Then defaultText = "< Show all programs >"
	
	' 0-ProgramID 1-ProgramName 5-EnrollStatusID 10-IsActive 18-ProgramIsEnabled
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formProgramDropdown"">"
	str = str & "<input type=""hidden"" name=""FormProgramDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""ProgramID"" onchange=""document.forms.formProgramDropdown.submit();"">"
	str = str & "<option value="""">" & Html(defaultText) & "</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			' programmember active, enrollstatus approved, program enabled
			If (list(10,i) = 1) And (list(5,i) = 3) And (list(18,i) = 1) Then
				isSelected = ""
				If CStr(list(0,i)) = CStr(page.Program.ProgramID) Then isSelected = " selected=""selected"""
				str = str & "<option value=""" & list(0,i) & """" & isSelected & ">" & html(list(1,i)) & "</option>"
			End If
		Next
	End If
	str = str & "</select></form></li>"
	
	ProgramDropdownToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<a href=""/member/overview.asp"">Member Home</a> / "
	Select Case page.Action
		Case SHOW_DETAILS
			pg.FileID = "": pg.Action = ""
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Files</a> / " & html(page.File.FriendlyName & "." & page.File.FileExtension)
		Case Else
			str = str & "Files"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href		: href = ""
	
	Select Case page.Action
		Case SHOW_DETAILS
			pg.FileID = "": pg.Action = ""
			href = pg.Url & pg.UrlParamsToString(True)
			str = str & "<li><a href=""" & href & """>"
			str = str & "<img class=""icon"" src=""/_images/icons/page.png"" /></a><a href=""" & href & """>File List</a></li>"
			
		Case Else
			str = str & SortByDropdownToString(page)
			str = str & ProgramDropdownToString(page)
	End Select
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_displayer_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	Public SortBy
	
	' encrypted
	Public Action
	Public ProgramID
	Public FileID
	
	' objects
	Public Member
	Public Client	
	Public Program
	Public File
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(SortBy) > 0 Then str = str & "sb=" & SortBy & amp
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(FileID) > 0 Then str = str & "fid=" & Encrypt(FileID) & amp
		
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
		Dim c     : Set c = New cPage

		c.MessageID = MessageID
		c.SortBy = SortBy
		c.Action = Action
		c.ProgramID = ProgramID
		c.FileID = FileID
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.File = File
		
		Set Clone = c
	End Function
End Class
%>

