<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%Server.ScriptTimeout = 600 %>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-files"
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
	page.SortBy = Request.QueryString("sb")
	
	page.Action = Decrypt(Request.QueryString("act"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.FileID = Decrypt(Request.QueryString("fid"))
	page.EventID = Decrypt(Request.QueryString("eid"))
	page.ScheduleID = Decrypt(Request.QueryString("scid"))
	
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		If page.Uploader.Form("FormProgramDropdownIsPostback") = IS_POSTBACK Then
			page.ProgramID = page.Uploader.Form("ProgramID")
			page.ScheduleID = ""
			page.EventID = ""
		End If
		If page.Uploader.Form("FormSortByDropdownIsPostback") = IS_POSTBACK Then
			page.SortBy = page.Uploader.Form("SortBy")
		End If
		If page.Uploader.Form("FormScheduleDropdownIsPostback") = IS_POSTBACK Then
			page.ScheduleID = page.Uploader.Form("NewScheduleID")
			page.EventID = ""
		End If
		If page.Uploader.Form("FormEventDropdownIsPostback") = IS_POSTBACK Then
			page.EventID = page.Uploader.Form("NewEventID")
		End If
		
		page.FileIDList = page.Uploader.Form("FileIDList")
		page.InsertFileIDList = page.Uploader.Form("InsertFileIDList")
		page.DeleteEventFileIDList = page.Uploader.Form("DeleteEventFileIDList")
	End If

	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then page.Program.Load()
	Set page.File = New cFile
	page.File.FileID = page.FileID
	If Len(page.File.FileID) > 0 Then page.File.Load()
	Set page.Schedule = New cSchedule
	page.Schedule.ScheduleID = page.ScheduleID
	If Len(page.Schedule.ScheduleID) > 0 Then page.Schedule.Load()
	Set page.Evnt = New cEvent
	page.Evnt.EventID = page.EventID
	If Len(page.Evnt.EventID) > 0 Then page.Evnt.Load()
	
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
		<link type="text/css" href="/_incs/script/jquery/jquery-ui-1.6.custom/development-bundle/themes/start/ui.all.css" rel="stylesheet" />	
		<link rel="stylesheet" type="text/css" href="files.css" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/jquery-ui-1.6.custom/js/jquery-ui-1.6.custom.min.js"></script>
		<script type="text/javascript" src="files.js"></script>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=m_pageTitleText%></title>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Set page.Uploader = Server.CreateObject("ASPSmartUpload.SmartUpload")
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		page.Uploader.Upload()
	End If
	
	Call OnPageLoad(page)
	
	Select Case page.Action
		Case SHOW_FILE_DETAILS
			str = str & FileSummaryToString(page)

		Case ADDNEW_RECORD
			If page.Uploader.Form("FormFileIsPostback") = IS_POSTBACK Then
				Call LoadFileFromRequest(page)
				If ValidFile(page) Then
					Call DoInsertFile(page, rv)
					Select Case rv
						Case 0
							page.MessageID = 9000
						Case Else
							page.MessageID = 9001
					End Select			
					page.Action = "": page.FileID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormFileToString(page)
				End If
			Else
				str = str & FormFileToString(page)
			End If
		
		Case UPDATE_RECORD
			If page.Uploader.Form("FormFileIsPostback") = IS_POSTBACK Then
				Call LoadFileFromRequest(page)
				If ValidFile(page) Then
					Call DoUpdateFile(page, rv)
					Select Case rv
						Case 0
							page.MessageID = 9000
						Case Else
							page.MessageID = 9003
					End Select			
					page.Action = "": page.FileID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormFileToString(page)
				End If
			Else
				str = str & FormFileToString(page)
			End If
			
		Case DELETE_RECORD
			If page.Uploader.Form("FormConfirmDeleteFileIsPostback") = IS_POSTBACK Then
				Call DoDeleteFile(page, rv)
				Select Case rv
					Case 0 
						page.MessageID = 9004
					Case Else
						page.MessageID = 9010
				End Select
				page.Action = "": page.FileID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmDeleteFileToString(page)
			End If
			
		Case DELETE_FILES_BULK
			' check for files selected
			If Len(page.FileIDList) = 0 Then
				page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If
			
			If page.Uploader.Form("FormConfirmBulkDeleteFilesIsPostback") = IS_POSTBACK Then
				Call DoBulkDeleteFiles(page, rv)
				Select Case rv
					Case 0
						page.MessageID = 9009
					Case Else
						page.MessageID = 9010
				End Select
				page.Action = "": page.FileID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmBulkDeleteFilesToString(page)
			End If
			
		Case LINK_FILES_TO_EVENTS
			' check for events for program selected
			If Not page.Client.HasPrograms Then
				page.MessageID = 2039: page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			End If	

			If page.Uploader.Form("FormEventFilesIsPostback") = IS_POSTBACK Then
				If page.Uploader.Form("Submit") = "<<" Then
					Call DoInsertEventFilesByList(page, rv)
					Select Case rv
						Case 0
						Case Else
					End Select
					Response.Redirect(page.Url & page.UrlParamsToString(False))
					
				ElseIf page.Uploader.Form("Submit") = ">>" Then
					Call DoDeleteEventFilesByList(page, rv)
					Select Case rv
						Case 0
						Case Else
					End Select
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				End If
			Else
				str = str & FormEventFilesToString(page)
			End If
		
		Case STREAM_FILE_TO_BROWSER
			Call page.File.StreamFile(page.Member.MemberID, rv)
			Response.End
			
		Case Else
			str = str & FileGridToString(page)
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoClearTabLinkBar()
	m_tablinkBarText = "<li>&nbsp;</li>"
End Sub

Sub DoDeleteEventFilesByList(page, outError)
	Dim i

	If Len(page.DeleteEventFileIDList) = 0 Then Exit Sub
	
	Dim eventFile			: Set eventFile = New cEventFile
	Dim idList				: idList = Split(page.DeleteEventFileIDList, ",")
	
	Dim tempError			: tempError = 0
	outError = 0
	
	For i = 0 To UBound(idList)
		eventFile.EventFileID = idList(i)
		Call eventFile.Delete(tempError)
		outError = outError + tempError
	Next
End Sub

Sub DoInsertEventFilesByList(page, outError)
	If Len(page.InsertFileIDList) = 0 Then Exit Sub
	
	Dim i, j
	Dim eventFiles		: Set eventFiles = New cEventFile
	eventFiles.EventID = page.Evnt.EventID
	
	Dim fileList		: fileList = eventFiles.List()
	Dim idList			: idList = Split(page.InsertFileIDList, ",")
	Dim exists			: exists = False
	
	Dim tempError		: tempError = 0
	outError = 0
	
	' 0-EventFileID 1-FileName 2-FriendlyName 3-FileExtension 8-FileID 

	For i = 0 To UBound(idList)
		exists = False
		
		' check for files that are already there
		If IsArray(fileList) Then
			For j = 0 To UBound(fileList,2)
				If CStr(idList(i)) = CStr(fileList(8,j)) Then
					exists = True
				End If
			Next
		End If
	
		If Not exists Then 
			eventFiles.FileID = idList(i)
			Call eventFiles.Add(tempError)
			outError = outError + tempError
		End If
	Next
End Sub

Sub DoBulkDeleteFiles(page, outError)
	Dim i
	
	Dim file				: Set file = New cFile
	Dim fso					: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim path				: path = Application.Value("FILE_MANAGER_FILESTORE") & page.Client.ClientID
	
	Dim list				: list = Split(page.FileIDList, ",")
	Dim tempError			: tempError = 0
	outError = 0
	
	If Not IsArray(list) Then Exit Sub
	
	For i = 0 To UBound(list)
		file.FileID = list(i)
		Call file.Load()
	
		' delete the file from disk first ..
		If Len(file.ProgramID) > 0 Then
			path = path & "\" & file.ProgramID
		End If
		If fso.FileExists(path & "\" & file.FileName) Then
			fso.DeleteFile(path & "\" & file.FileName)
		End If
		
		' delete file from db ..
		Call file.Delete(tempError)
		If tempError <> 0 Then outError = -1
	Next
End Sub

Sub DoDeleteFile(page, outError)
	Dim fso			: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim path		: path = Application.Value("FILE_MANAGER_FILESTORE") & page.Client.ClientID
	If Len(page.File.ProgramID) > 0 Then
		path = path & "\" & page.File.ProgramID
	End If
	
	' delete the file from filestore
	If fso.FileExists(path & "\" & page.File.FileName) Then
		fso.DeleteFile(path & "\" & page.File.FileName)
	End If
	
	' delete the file from db ..
	Call page.File.Delete(outError)
End Sub

Sub DoUpdatefile(page, outError)
	Dim fso					: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim path				: path = Application.Value("FILE_MANAGER_FILESTORE") & page.Client.ClientID
	Dim folder 
	
	Dim oldFile				: Set oldFile = New cFile
	oldFile.FileID = page.File.FileID
	oldFile.Load()
	Dim oldFileName			: oldFileName = oldFile.FileName & "." & oldFile.FileExtension
	
	' check if moving file to a new folder
	If CStr(oldFile.ProgramID & "") <> CStr(page.File.ProgramID & "") Then
	
		' make sure the new folder exists
		If Len(page.File.ProgramID) > 0 Then
			If Not fso.FolderExists(path & "\" & page.File.ProgramID) Then Call fso.CreateFolder(path & "\" & page.File.ProgramID)
		End If
		
		' move the file
		Call fso.CopyFile(oldFile.Path, page.File.Path)
		
		' delete the old file
		Call fso.DeleteFile(oldFile.Path)
	End If 

	Call page.File.Save(outError)
End Sub

Sub DoInsertFile(page, outError)
	Dim fso			: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim path		: path = Application.Value("FILE_MANAGER_FILESTORE") & page.Client.ClientID
	Dim extension	: extension = page.Uploader.Files(1).FileExt
	Dim rootName	: rootName = Replace(page.Uploader.Files(1).FileName, "." & extension, "")
	Dim fileName	: fileName = ""
	Dim suffix		: suffix = ""
	Dim counter		: counter = 1
	
	' check for enough space
	page.File.ClientID = page.Client.ClientID
	Dim fileStore	: fileStore = page.File.GetFileStoreInfo()
	
	' create client/program dir if necessary
	If Not fso.FolderExists(path) Then fso.CreateFolder(path)
	If Len(page.File.ProgramID) > 0 Then
		path = path & "\" & page.File.ProgramID
	End If
	If Not fso.FolderExists(path) Then fso.CreateFolder(path)
	
	' clean the file name of illegal characters and spaces
	rootName = CleanFileName(rootName, "_")
	
	' rename the file if a file with that name already exists ..
	fileName = rootName & "." & extension
	Do While fso.FileExists(path & "\" & fileName)
		fileName = rootName & "(" & counter & ")" & "." & extension
		counter = counter + 1
	Loop
	
	' save the file to filestore
	Call page.Uploader.Files(1).SaveAs(path & "\" & fileName)
	
	' save the file to db
	page.File.ClientID = page.Client.ClientID
	page.File.FileName = fileName
	page.File.FriendlyName = Replace(fileName, "." & extension, "")
	page.File.FileOwnerID = page.Member.MemberID
	page.File.FileExtension = extension
	page.File.FileSize = page.Uploader.Files(1).Size
	page.File.MimeType = page.Uploader.Files(1).TypeMIME
	page.File.MimeSubType = page.Uploader.Files(1).SubTypeMIME
	
	Call page.File.Add(outError)
		
	Set fso = Nothing
End Sub

Function EventGridForFileSummaryToString(page)
	Dim str, i
	Dim pg							: Set pg = page.Clone()
	Dim dateTime					: Set dateTime = New cFormatDate
	
	Dim list						: list = page.File.EventList()
	Dim rows						: rows = ""
	Dim href						: href = ""
	Dim alt							: alt = ""
	Dim count						: count = 0
	
	' 0-EventFileID 1-EventID 2-EventName 3-EventNote 4-EventDate 5-TimeStart 6-TimeEnd
	' 7-ScheduleID 8-ScheduleName 9-ProgramID 10-ProgramName

	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			alt = ""				: If count Mod 2 > 0 Then alt = " class=""alt"""
			
			pg.Action = SHOW_EVENT_DETAILS: pg.EventId = list(1,i): pg.FileId = ""
			href = "/schedule/events.asp" & pg.UrlParamsToString(True)
		
			rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
			rows = rows & "<strong>" & html(list(10,i)) & "</strong> | "
			rows = rows & "<a href=""" & href & """ title=""Details""><strong>" & html(list(2,i)) & "</strong></a></td>"
			rows = rows & "<td>" & dateTime.Convert(list(4,i), "MM-DD-YYYY")
			If Len(list(5,i)) > 0 Then rows = rows & " at " & dateTime.Convert(list(5,i), "hh:nnpx")
			rows = rows & "</td>"
			rows = rows & "<td>" & html(list(8,i)) & "</td>"
			rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
			rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td>"
			rows = rows & "</tr>"
			
			count = count + 1
		Next
	End If
	
	If count > 0 Then
		str = str & "<p>Event team members can download this file from these events. </p>"
		str = str & "<div class=""grid""><table>"
		str = str & "<thead><tr><th>Event</th><th>When</th><th>Schedule</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = "<p class=""alert"">This file is not linked to any events. </p>"
	End If
	
	EventGridForFileSummaryToString = str
End Function

Function FileDownloadGridForFileSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim fileDownload		: Set fileDownload = New cFileDownload
	fileDownload.FileId = page.File.FileId
	
	Dim list				: list = FileDownload.List()
	Dim rows				: rows = ""
	Dim alt					: alt = ""
	Dim href				: href = ""
	Dim count				: count = 0
	
	' 0-FileDownloadID 1-FileName 2-FriendlyName 3-FileExtension 4-FileSize 5-Description 6-IsPublic
	' 7-DateFileCreated 8-FileOwnerID 9-ProgramID 10-ProgramName 11-MemberID 12-NameLast
	' 13-NameFirst 14-DownloadDate

	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			alt = ""					: If count Mod 2 > 0 Then alt = " class=""alt"""
			
			pg.Action = SHOW_MEMBER_DETAILS: pg.MemberId = list(11,i)
			href = "/admin/profile.asp" & pg.UrlParamsToString(True)
			
			rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/user_red.png"" alt="""" />"
			rows = rows & "<a href=""" & href & """ title=""Details"">"
			rows = rows & "<strong>" & html(list(12,i) & ", " & list(13,i)) & "</strong></a></td>"
			rows = rows & "<td>" & dateTime.Convert(list(14,i), "DDDD MMMM dd, YYYY at hh:nn pp") & "</td>"
			rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
			rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
			
			count = count + 1
		Next
	End If
				
	If count > 0 Then
		str = str & "<p>The list of members who have downloaded this file. </p>"
		str = str & "<div class=""grid""><table>"
		str = str & "<thead><tr><th>Member</th><th>When</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = "<p class=""alert"">This file has never been downloaded. </p>"
	End If
	
	FileDownloadGridForFileSummaryToString = str
End Function

Function FileSummaryToString(page)
	Dim str
	Dim pg						: Set pg = page.Clone()
	
	Dim dateTime				: Set dateTime = New cFormatDate
	Dim fileDisplay				: Set fileDisplay = New cFileDisplay
	
	' file owner ..
	Dim member					: Set member = New cMember
	member.MemberId = page.File.FileOwnerId
	Call member.Load()
	
	Dim fileOwnerText
	fileOwnerText = member.NameFirst & " " & member.NameLast
	'If CLng(page.File.FileOwnerId) = CLng(page.Member.MemberId) Then fileOwnerText = "you"
	
	Dim fileSizeText		
	fileSizeText = "This file is using <strong>" & FileSizeToString(page.File.FileSize) & "</strong> of your account file storage space. "
	
	Dim downloadCountText
	If page.File.DownloadCount = 0 Then 
		downloadCountText = "This file has never been downloaded"
	ElseIf page.File.DownloadCount = 1 Then
		downloadCountText = "This file has been downloaded one time"
	Else
		downloadCountText = "This file has been downloaded " & page.File.DownloadCount & " times"
	End If
	
	Dim eventLinkText
	If page.File.EventCount = 0 Then
		eventLinkText = "This file is not linked to any events"
	ElseIf page.File.EventCount = 1 Then
		eventLinkText = "This file is linked to one event"
	Else
		eventLinkText = "This file is linked to " & page.File.EventCount & " events"
	End If
	
	Dim privateFileText
	If page.File.IsPublic = 1 Then
		privateFileText = privateFileText & "This file is set to <strong>public</strong>. "
		privateFileText = privateFileText & "Your account members can download this file from their member files page. "
	Else
		privateFileText = privateFileText & "This file is set to <strong>private</strong>. "
		privateFileText = privateFileText & "Your account members will not see this file on their member files page "
		privateFileText = privateFileText & "(They will only have access if you link this file to one of your events and they belong to the event team for that event). "
	End If
	
	Dim programFileText
	If Len(page.File.ProgramId & "") = 0 Then
		programFileText = programFileText & "This file has not been set for a specific program. "
		programFileText = programFileText & "Any of your account members can download this file from their member files page. "
	Else
		programFileText = programFileText & "This file has been assigned to the <strong>" & html(page.File.ProgramName) & "</strong> program. "
		programFileText = programFileText & "Members of " & html(page.File.ProgramName) & " will be able to download this file from their member files page. "
	End If
	
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & html(page.File.FriendlyName & "." & page.File.FileExtension) & "</h3>"

	str = str & "<h5 class=""description"">Description</h5>"
	If Len(page.File.Description) > 0 Then 
		str = str & "<p>" & html(page.File.Description) & "</p>"
	Else
		str = str & "<p class=""alert"">No description is available. </p>"
	End If
	
	str = str & "<h5 class=""settings"">Settings</h5>"
	str = str & "<ul><li>" & privateFileText & "</li>"
	str = str & "<li>" & programFileText & "</li></ul>"
	
	str = str & "<h5 class=""event"">Event links</h5>"
	str = str & EventGridForFileSummaryToString(page)
	
	str = str & "<h5 class=""file-download"">Download history</h5>"
	str = str & FileDownloadGridForFileSummaryToString(page)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul><li>Uploaded by <strong>" & html(fileOwnerText) & "</strong> on " & dateTime.Convert(page.File.DateCreated, "DDDD MMMM dd, YYYY around hh:00 pp") & ". </li>"
	str = str & "<li>" & downloadCountText & ". </li>"
	str = str & "<li>" & eventLinkText & ". </li>"
	str = str & "<li>" & fileSizeText & ". </li></ul>"
	
	str = str & "</div>"
	
	FileSummaryToString = str
End Function

Function NoFilesDialogToString(page)
	Dim dialog				: Set dialog = New cDialog
	Dim pg					: Set pg = page.Clone()
	
	If Len(page.Program.ProgramId) > 0 Then
		dialog.Headline = "No files here yet .. "
	Else
		dialog.Headline = "Where are the files ..?"
	End If
	
	If Len(page.Program.ProgramId) > 0 Then
		dialog.Text = dialog.Text & "<p>It looks like you have not uploaded any files for the program you selected (" & html(page.Program.ProgramName) & "). "
		dialog.Text = dialog.Text & "That's ok, " & Application.Value("APPLICATION_NAME") & " will still work fine without any file uploads "
		dialog.Text = dialog.Text & "(you have room for " & FormatNumber(Application.Value("INITIAL_FILE_STORAGE")/1000000, 1,,,False) & " MB of files in your account). </p>"
		dialog.Text = dialog.Text & "<p>Click <strong>Upload a file</strong> to upload your first file. </p>"
	Else
		dialog.Text = dialog.Text & "<p>It looks like you have not uploaded any files to your " & html(page.Client.NameClient) & " " & Application.Value("APPLICATION_NAME") & " account. "
		dialog.Text = dialog.Text & "That's ok, " & Application.Value("APPLICATION_NAME") & " will still work fine without any file uploads "
		dialog.Text = dialog.Text & "(you have room for " & FormatNumber(Application.Value("INITIAL_FILE_STORAGE")/1000000, 1,,,False) & " MB of files in your account). </p>"
		dialog.Text = dialog.Text & "<p>Click <strong>Upload a file</strong> to upload your first file. </p>"
	End If

	dialog.SubText = dialog.SubText & "<p>Once you have uploaded a file, this page will show you a list of all of the files in your account. "
	dialog.SubText = dialog.SubText & "You will use this page to link your files (like MP3s or charts) to your calendar events. "
	dialog.SubText = dialog.SubText & "That way, your event team members have easy access to any files they need right from the event on their calendar. </p>"

	pg.Action = ADDNEW_RECORD		
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Upload a file</a></li>"
	If Len(page.Program.ProgramId) > 0 Then
		pg.Action = "": pg.ProgramId = ""
		dialog.LinkList = dialog.LinkList & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Show all of my files</a></li>"
	End If
	dialog.LinkList = dialog.LinkList & "<li><a href=""/help/faq.asp?act=2&amp;fqid=169"" target=""_blank"">Learn more about files</a></li>"
	
	NoFilesDialogToString = dialog.ToString()
End Function

Function FileGridToString(page)
	Dim str, msg, i
	Dim pg					: Set pg = page.Clone
	Dim dateTime			: Set dateTime = New cFormatDate
	Dim displayer			: Set displayer = New cFileDisplay
	Dim count				: count = 0
	Dim altClass			: altClass = ""
	Dim isPublicText		: isPublicText = ""
	Dim fileIconPath		: fileIconPath = ""
	
	page.File.ClientID = page.Client.ClientID
	page.File.ProgramID = page.Program.ProgramID
	Dim list				: list = page.File.List(LookupSortParam(page.SortBy))
	
	
	str = str & "<div class=""tip-box""><h3>I want to .. </h3><ul>"
	pg.Action = ADDNEW_RECORD
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Upload a file to my account</a></li>"
	pg.Action = LINK_FILES_TO_EVENTS
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Link a file to one of my events</a></li>"
	str = str & "<li><a href=""/help/faq.asp?act=2&amp;fqid=169"" target=""_blank"">Learn more about files</a></li>"
	str = str & "</ul></div>"
	
	str = str & FileReportTipBoxToString(page)
	
	' 0-FileID 1-FileName 2-FriendlyName 3-Description 4-ClientID 5-ProgramID 6-FileOwnerID
	' 7-DateCreated 8-DateModified 9-FileExtension 10-FileSize 11-MIMEType 12-MIMESubType 13-EventFileCount
	' 14-IsPublic 15-DownloadCount 16-ProgramName

	str = str & m_appMessageText
	If IsArray(list) Then
		str = str & "<div class=""grid"">"
		pg.Action = DELETE_FILES_BULK
		str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" id=""form-file-grid"">"
		str = str & "<table><tr><th scope=""col"" style=""width:1%;"">"
		str = str & "<input type=""checkbox"" id=""master"" /></th>"
		str = str & "<th scope=""col"">Files</th><th scope=""col"">Program</th>"
		str = str & "<th scope=""col"">Public</th><th scope=""col"">&nbsp;</th></tr>"
		For i = 0 To UBound(list,2)
			count = count + 1
			altClass = ""
			If count Mod 2 = 0 Then altClass = " class=""alt"""
			
			isPublicText = "Yes"
			If list(14,i) = 0 Then 
				isPublicText = "<span style=""color:red;"">No</span>"
			End If
			fileIconPath = displayer.GetIconPath(list(9,i))
			
			str = str & "<tr" & altClass & "><td><input class=""file-checkbox"" type=""checkbox"" name=""FileIDList"" value=""" & list(0,i) & """ /></td>"
			str = str & "<td><img class=""icon"" src=""" & fileIconPath & """ alt="""" />"
			pg.Action = SHOW_FILE_DETAILS: pg.FileID = list(0,i)
			str = str & "<strong><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(list(2,i) & "." & list(9,i)) & "</a></strong> <span style=""font-size:0.9em;color:gray;"">(" & FileSizeToString(list(10,i)) & ")</span></td>"
			str = str & "<td>" & html(list(16,i)) & "</td>"
			str = str & "<td>" & isPublicText & "</td>"
			str = str & "<td class=""toolbar"">"
			pg.Action = SHOW_FILE_DETAILS: pg.FileID = list(0,i)
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details"">"
			str = str & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a>"
			pg.Action = UPDATE_RECORD: pg.FileID = list(0,i)
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Edit"">"
			str = str & "<img src=""/_images/icons/pencil.png"" alt="""" /></a>"
			pg.Action = STREAM_FILE_TO_BROWSER: pg.FileID = list(0,i)
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Download"">"
			str = str & "<img src=""/_images/icons/page_swoosh_down.png"" alt="""" /></a>"
			pg.Action = DELETE_RECORD: pg.FileID = list(0,i)
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove"">"
			str = str & "<img src=""/_images/icons/cross.png"" alt="""" /></a>"
			str = str & "</td></tr>"
		Next
		str = str & "</table></form></div>"
	End If
	
	If count = 0 Then
		Call DoClearTabLinkBar()
		str = NoFilesDialogToString(page)
	End If
	
	FileGridToString = str
End Function

Function FileReportTipBoxToString(page)
	Dim str
	
	' 0-NameClient 1-FileCount 2-EventCount 3-Used 4-Available
	Dim fileStore		: fileStore = page.File.GetFileStoreInfo()
	
	Dim fileCount		: fileCount = fileStore(1,0) & " file"
	If fileStore(1,0) <> 1 Then fileCount = fileCount & "s"
	
	Dim eventCount		: eventCount = fileStore(2,0) & " event"
	If fileStore(2,0) <> 1 Then eventCount = eventCount & "s"

	str = str & "<div class=""tip-box""><h3>File Space</h3><p>"
	str = str & "Your account is storing " & fileCount & ", using around " & FileSizeToString(fileStore(3,0)) & " of server space. "
	str = str & "<br /><br />Your account includes " & FileSizeToString(fileStore(4,0)) & " of server space. "
	str = str & "</p></div>"
	
	FileReportTipBoxToString = str
End Function

Function FormEventFilesToString(page)
	Dim str, msg
	
	Dim eventFile				: Set eventFile = New cEventFile
	eventFile.EventID = page.Evnt.EventID
	
	Dim eventFilesList			
	If Len(eventFile.EventID) > 0 Then eventFilesList =  eventFile.List()
	
	page.File.ClientID = page.Client.ClientiD
	Dim ownedFiles			: ownedFiles = page.File.List("")
	
	str = str & "<div class=""tip-box""><h3>Tip</h3><p>"
	If Len(page.Program.ProgramID) > 0 Then
		If Len(page.Schedule.ScheduleID) > 0 then
			str = str & "Move files to the event to make them downloadable from the event on the member calendar. "
		Else
			str = str & "Select a schedule from the dropdown list to get started. "
		End If
	Else
		str = str & "Select a program from the dropdown list in the toolbar to get started. "
	End If
	str = str & "</p></div>"
	str = str & "<div class=""tip-box""><h3>Tip</h3><p>"
	str = str & "Highlight a file or files in either list, and then click on a button to move them. "
	str = str & "<br /><br />Use [CONTROL]-click or [SHIFT]-click to select multiple files. </p></div>"
	
	If Len(page.Program.ProgramID) > 0 Then
		If Not (page.Program.HasEvents Or page.Schedule.HasEvents) Then
			msg = "You have selected a program or schedule that does not have any events. "
			msg = msg & "Before you can use this page to link your files to events, you'll need to set up a schedule with at least one event for this program. "
			str = str & CustomApplicationMessageToString("No events were returned! ", msg, "Error")
		End If
	Else
		msg = "No program is selected in the toolbar. "
		msg = msg & "Select the program for the event that you would like to work with. "
		str = str & CustomApplicationMessageToString("Please select a program! ", msg, "Error")
	End If
	
	str = str & "<div class=""grid"">"
	str = str & "<table><tr class=""header""><th scope=""col"">Event Files</th>"
	str = str & "<th scope=""col"" style=""text-align:right;"">" & ScheduleDropdownToString(page) & "&nbsp;"
	str = str & EventDropdownToString(page) & "</th></tr>"
	str = str & "<tr><td class=""two-way-selector"" colspan=""2"" style="""">"
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" name=""formEventFiles"">"
	str = str & "<input type=""hidden"" name=""FormEventFilesIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><tr><td class=""list-box"">"
	str = str & DeleteEventFileDropdownToString(page, eventFilesList) & "</td>"
	
	str = str & "<td class=""arrow-buttons"">"
	str = str & "<input type=""submit"" name=""Submit"" value=""" & HTML("<<") & """ />"
	str = str & "<br /><input type=""submit"" name=""Submit"" value=""" & HTML(">>") & """ /></td>"
	
	str = str & "<td class=""list-box"">" & InsertEventFileDropdownToString(page, eventFilesList, ownedFiles)
	str = str & "</td></tr></table></form></td></tr></table></div>"
	
	FormEventFilesToString = str
End Function

Function DeleteEventFileDropdownToString(page, eventFiles)
	Dim str, i
	Dim dateTime				: Set dateTime = New cFormatDate
	
	Dim listBoxHeader			: listBoxHeader = "&nbsp;"
	Dim defaultOptionText		
	If Len(page.Evnt.EventID) = 0 Then
		defaultOptionText = "No event is selected .."
		listBoxHeader = html(page.Program.ProgramName)
	Else
		defaultOptionText = "&nbsp;"
		listBoxHeader = html(page.Evnt.EventName) & " <span style=""font-weight:normal"">(" & dateTime.Convert(page.Evnt.EventDate, "mm/dd/YYYY") & ")</span>"
	End If
	
	If Len(page.Schedule.ScheduleID) = 0 Then
		defaultOptionText = "No schedule is selected .."
		listBoxHeader = html(page.Program.ProgramName)
	End If
	If Len(page.Program.ProgramID) = 0 Then 
		defaultOptionText = "No program is selected .."
		listBoxHeader = "&nbsp;"
	End If
	
	' 0-EventFileID 1-FileName 2-FriendlyName 3-FileExtension 8-FileID
	
	str = str & "<h4 id=""event-label"">" & listBoxHeader & "</h4>"
	str = str & "<select name=""DeleteEventFileIDList"" multiple=""multiple"">"
	If IsArray(eventFiles) Then
		For i = 0 To UBound(eventFiles,2)
			str = str & "<option value=""" & eventFiles(0,i) & """>" & html(eventFiles(2,i) & "." & eventFiles(3,i)) & "</option"
		Next
	Else
		str = str & "<option style=""font-style:italic;color:gray;"" value="""">" & html(defaultOptionText) & "</option>"
	End If
	str = str & "</select>"
	
	DeleteEventFileDropdownToString = str
End Function

Function InsertEventFileDropdownToString(page, eventFiles, ownedFiles)
	Dim str, i, j

	Dim optionList			: optionList = ""
	Dim showThisFile		: showThisFile = True
	Dim isProgramFile		: isProgramFile = True
	Dim isGlobalFile		: isGlobalFile = True
	Dim count				: count = 0
	
	' 0-FileID 1-FileName 2-FriendlyName 5-ProgramID 9-FileExtension

	If IsArray(ownedFiles) Then
		For i = 0 To UBound(ownedFiles,2)
			showThisFile = True
			isProgramFile = True
			isGlobalFile = True
			
			If Len(ownedFiles(5,i) & "") > 0 Then
				isGlobalFile = False
			End If
			
			If CStr(ownedFiles(5,i) & "") <> CStr(page.Program.ProgramID) Then 
				isProgramFile = False
			End If
				
			If isGlobalFile Or isProgramFile Then
				' if not already linked ..
				If IsArray(eventFiles) Then
					For j = 0 To UBound(eventFiles,2)
						If CStr(ownedFiles(0,i)) = CStr(eventFiles(8,j)) Then
							showThisFile = False
						End If
					Next
				End If
			Else
				showThisFile = False
			End If
				
			If showThisFile Then
				count = count + 1
				optionList = optionList & "<option value=""" & ownedFiles(0,i) & """>" & html(ownedFiles(2,i) & "." & ownedFiles(9,i)) & "</option>"
			End If
		Next
	End If
	
	If count = 0 Then
		optionList = "<option value="""">&nbsp;</option>"
	End If
	
	' don't show any files until filters are set ..
	If (Len(page.Schedule.ScheduleID) = 0) Or (Len(page.Evnt.EventID) = 0) Then
		optionList = "<option value="""">&nbsp;</option>"
	End If
	
	str = str & "<h4 id=""file-label"">" & html("Available Files") & "</h4>"
	str = str & "<select name=""InsertFileIDList"" multiple=""multiple"">"
	str = str & optionList & "</select>"

	InsertEventFileDropdownToString = str
End Function

Function ScheduleDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim schedules		: If Len(page.ProgramID) > 0 Then schedules = page.Program.ScheduleList()
	Dim selected		: selected = ""
	
	Dim disabled		: disabled = ""
	If Len(page.Program.ProgramID) = 0 Then disabled = " disabled=""disabled"""
	
	str = str & "<form style=""display:inline;"" action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" name=""formScheduleDropdown"">"
	str = str & "<input type=""hidden"" name=""FormScheduleDropdownIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""NewScheduleID"" onchange=""document.forms.formScheduleDropdown.submit();""" & disabled & ">"
	If Len(page.ScheduleID) = 0 Then
		str = str & "<option value="""">" & html("< Select Schedule >") & "</option>"
		str = str & "<option value="""">" & html("--") & "</option>"
	End If
	If IsArray(schedules) Then
		For i = 0 To UBound(schedules,2)
			selected = ""
			If CStr(schedules(0,i)) = CStr(page.ScheduleID) Then selected = " selected=""selected"""
		
			str = str & "<option value=""" & schedules(0,i) & """" & selected & ">"
			str = str & html(schedules(1,i)) & "</option>"
		Next
	End If
	str = str & "</select></form>"
	
	ScheduleDropdownToString = str
End Function

Function EventDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim dateTime		: Set dateTime = New cFormatDate
	Dim events			: If Len(page.ScheduleID) > 0 Then events = page.Schedule.EventList("")
	Dim selected
	
	Dim disabled		: disabled = ""
	If Len(page.Program.ProgramID) = 0 Then disabled = " disabled=""disabled"""
	If Len(page.Schedule.ScheduleID) = 0 Then  disabled = " disabled=""disabled"""
	
	
	' 0-EventID 1-EventName 2-EventDate
	
	str = str & "<form style=""display:inline;"" action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" name=""formEventDropdown"">"
	str = str & "<input type=""hidden"" name=""FormEventDropdownIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""NewEventID"" onchange=""document.forms.formEventDropdown.submit();""" & disabled & ">"
	If Len(page.EventID) = 0 Then
		str = str & "<option value="""">" & html("< Select Event >") & "</option>"
		str = str & "<option value="""">" & html("--") & "</option>"
	End If
	If IsArray(events) Then
		For i = 0 To UBound(events,2)
			selected = ""
			If CStr(events(0,i)) = CStr(page.EventID) Then selected = " selected=""selected"""
		
			str = str & "<option value=""" & events(0,i) & """" & selected & ">"
			str = str & dateTime.Convert(events(2,i), "mm/dd/YYYY") & " | " & html(events(1,i)) & "</option>"
		Next
	End If
	str = str & "</select></form>"
	
	EventDropdownToString = str
End Function

Function FormFileToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	' set program if this is insert
	If page.Action = ADDNEW_RECORD Then
		If Len(page.File.ProgramID) = 0 Then
			page.File.ProgramID = page.Program.ProgramID
		End If
	End If
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>To add a file to your " & Application.Value("APPLICATION_NAME") & " account, click <strong>browse</strong> and navigate to a file on your computer. </p></div>"
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" name=""formFile"">"
	str = str & "<table>"
	If page.Action = UPDATE_RECORD Then
		str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "File Name") & "</td>"
		str = str & "<td><input type=""text"" class=""medium"" name=""FriendlyName"" value=""" & html(page.File.FriendlyName) & """ /> <strong style=""font-style:italic;color:gray;"">." & html(page.File.FileExtension) & "</strong></td></tr>"
	Else
		str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "File") & "</td>"
		str = str & "<td><input type=""file"" class=""file"" name=""File"" size=""46"" /></td></tr>"
	End If
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Description </td>"
	str = str & "<td><textarea name=""Description"" class=""large"">" & html(page.File.Description) & "</textarea></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Program</td>"
	str = str & "<td>" & ProgramForFileDropdownToString(page) & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "If a program is selected, only members from this program <br />will have access to this file. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Public</td>"
	str = str & "<td>" & YesNoDropdownToString(page.File.IsPublic, "IsPublic") & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">"
	str = str & "Members are only able to access private files when they are linked <br />to an event for which they are scheduled. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>"
	str = str & "<input type=""submit"" name=""submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormFileIsPostback"" value=""" & IS_POSTBACK & """ /></td></tr>"
	str = str & "</table></form></div>"
	
	FormFileToString = str
End Function

Function FormConfirmDeleteFileToString(page)
	Dim str, msg
	Dim pg			: Set pg = page.Clone()
	
	msg = msg & "You will permanently remove the file <strong>" & html(page.File.FriendlyName & "." & page.File.FileExtension) & "</strong> from your " & Application.Value("APPLICATION_NAME") & " account."
	msg = msg & "Your account members will no longer be able to access this file from their accounts. "
	msg = msg & "This action cannot be reversed. "
	str = str & CustomApplicationMessageToString("Please confirm remove file!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" id=""formConfirmDeleteFile"">"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormConfirmDeleteFileIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</p></form>"
	
	FormConfirmDeleteFileToString = str
End Function

Function FormConfirmBulkDeleteFilesToString(page)
	Dim str, msg
	Dim pg					: Set pg = page.Clone()
	Dim list				: list = Split(page.FileIDList, ",")
	
	Dim fileCount			: fileCount = UBound(list) + 1
	Dim fileCountText		: fileCountText = fileCount & " file"
	Dim thisFile			: thisFile = "this file"
	Dim errorHeader			: errorHeader = "Please confirm remove file"
	If fileCount <> 1 Then 
		fileCountText = fileCountText & "s"
		thisFile = "these files"
		errorHeader = errorHeader & "s"
	End If
	errorHeader = errorHeader & "!"
	
	msg = msg & "You will permanently remove <strong>" & fileCountText & "</strong> from your " & Application.Value("APPLICATION_NAME") & " account. "
	msg = msg & "If you continue, your members will no longer be able to access " & thisFile & " from their accounts. "
	msg = msg & "This action cannot be reversed. "
	str = str & CustomApplicationMessageToString(errorHeader, msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" id=""formConfirmDeleteFile"">"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormConfirmBulkDeleteFilesIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<input type=""hidden"" name=""FileIDList"" value=""" & page.FileIDList & """ />"
	str = str & "</p></form>"	
	
	FormConfirmBulkDeleteFilesToString = str
End Function

Function ValidFile(page)
	Dim noSpaceMessage
	ValidFile = True
	
	If Not ValidData(page.File.FriendlyName, False, 0, 256, "File Name", "") Then ValidFile = False
	If Not ValidData(page.File.Description, False, 0, 1000, "Description", "") Then ValidFile = False
	If page.Action = ADDNEW_RECORD Then
		
		' missing file
		If page.Uploader.Files(1).IsMissing Then
			AddCustomFrmError("No file was provided.")
			ValidFile = False
		End If
		
		' file too big ..
		page.File.ClientID = page.Client.ClientID
		If Not IsSpaceAvailable(page.Uploader.Files(1).Size, page.File, noSpaceMessage) Then
			AddCustomFrmError("Your account is out of disc space. You will need to remove some files to make room. ")
			ValidFile = False
		End If
	End If
	If page.Action = UPDATE_RECORD Then
		If Len(page.File.FriendlyName) = 0 Then
			AddCustomFrmError("A File Name is required.")
			ValidFile = False
		End If
	End If
End Function

Function IsSpaceAvailable(fileSize, file, outErrorMessage)
	Dim str
	IsSpaceAvailable = True
	Dim fileStore		: fileStore = file.GetFileStoreInfo()
	
	' 0-NameClient 1-FileCount 2-EventCount 3-Used 4-Available
	If (CLng(fileStore(3,0)) + CLng(fileSize)) > CLng(fileStore(4,0)) Then 
		str = str & "Your file is too large. "
		str = str & "Uploading this file (" & FileSizeToString(fileSize) & ") will cause you to exceed your total file storage "
		str = str & "(" & FileSizeToString(CLng(fileStore(4,0)) - CLng(fileStore(3,0))) & " of " & FileSizeToString(fileStore(4,0)) & " remaining). "
		str = str & "You will need to remove one or more files to make space before you will be able to upload this file. "
		outErrorMessage = str
		
		IsSpaceAvailable = False
	End If
End Function

Sub LoadFileFromRequest(page)
	page.File.FriendlyName = page.Uploader.Form.Item("FriendlyName")
	page.File.Description = page.Uploader.Form.Item("Description")
	page.File.ProgramID = page.Uploader.Form.Item("ProgramID")
	page.File.IsPublic = page.Uploader.Form.Item("IsPublic")
End Sub

Function ProgramForFileDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = page.Member.OwnedProgramsList()
	Dim selected		: selected = ""
	
	str = str & "<select name=""ProgramID"">"
	str = str & "<option value="""">&nbsp;</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			selected = ""
			If CStr(list(0,i) & "") = CStr(page.File.ProgramID & "") Then selected = " selected=""selected"""
			
			str = str & "<option value=""" & list(0,i) & """" & selected & ">" & html(list(1,i)) & "</option>"
		Next
	End If
	str = str & "</select>"
	
	ProgramForFileDropdownToString = str
End Function

Function FileSizeToString(ByVal val)
	Dim str
	val = CLng(val)
		
	If val < 1000 Then
		str = "0Kb"
	ElseIf (val > 999) And (val < 1000000) Then
		str = val / 1000
		str = FormatNumber(str, 2, , , False) & "Kb"
	Else
		str = val/1000000
		str = FormatNumber(str, 2, , , False) & "Mb"
	End If
	
	FileSizeToString = str
End Function

Function SortByDropdownToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" id=""formSortByDropdown"">"
	str = str & "<input type=""hidden"" name=""FormSortByDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""SortBy"" onchange=""document.forms.formSortByDropdown.submit();"">"
	str = str & "<option value="""">" & html("< Sort by .. >") & "</option>"
	str = str & "<option value=""" & SORT_BY_FILE_NAME & """>File Name</option>"
	str = str & "<option value=""" & SORT_BY_FILE_EXTENSION & """>Extension</option>"
	str = str & "<option value=""" & SORT_BY_FILE_SIZE & """>Size</option>"
	str = str & "<option value=""" & SORT_BY_PROGRAM_NAME & """>Program</option>"
	str = str & "<option value=""" & SORT_BY_IS_PUBLIC & """>Public Files</option>"
	str = str & "</select></form></li>"	
	
	SortByDropdownToString = str
End Function

Function ProgramDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = page.Member.OwnedProgramsList()
	Dim isSelected		: isSelected = ""
	
	Dim defaultText		: defaultText = "< Select a program >"
	If page.Action = LINK_FILES_TO_EVENTS Then
		If Len(page.Program.ProgramID) > 0 Then defaultText = "< Show all programs >"
	Else
		If Len(page.Program.ProgramID) > 0 Then defaultText = "< Show all files >"
	End If
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" id=""formProgramDropdown"">"
	str = str & "<input type=""hidden"" name=""FormProgramDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""ProgramID"" onchange=""document.forms.formProgramDropdown.submit();"">"
	str = str & "<option value="""">" & Html(defaultText) & "</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isSelected = ""
			If CStr(list(0,i)) = CStr(page.Program.ProgramID) Then isSelected = " selected=""selected"""
			str = str & "<option value=""" & list(0,i) & """" & isSelected & ">" & html(list(1,i)) & "</option>"
		Next
	End If
	str = str & "</select></form></li>"
	
	ProgramDropdownToString = str
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
		Case SORT_BY_IS_PUBLIC
			str = "IsPublic"
		Case Else
			str = ""
	End Select
	
	LookupSortParam = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim dateTime	: Set dateTime = New cFormatDate
	Dim pg			: Set pg = page.Clone()
	
		
	Dim fileDetailsLink
	pg.Action = SHOW_FILE_DETAILS
	fileDetailsLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.File.FriendlyName & "." & page.File.FileExtension) & "</a> / "

	Dim fileListLink
	pg.Action = "": pg.ProgramID = ""
	fileListLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Files</a> / "
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case LINK_FILES_TO_EVENTS
			str = str & fileListLink
			If Len(page.Evnt.EventID) > 0 Then
				str = str & "Files for " & html(page.Evnt.EventName) & " (" & dateTime.Convert(page.Evnt.EventDate, "mm/dd/YYYY") & ") Event"
			Else
				str = str & "Files for events"
			End If
		Case DELETE_RECORD
			str = str & fileListLink
			str = str & fileDetailsLink
			str = str & "Remove File"
		Case UPDATE_RECORD
			str = str & fileListLink
			str = str & fileDetailsLink
			str = str & "Edit File"
		Case ADDNEW_RECORD
			str = str & fileListLink
			str = str & "Add File"
		Case SHOW_FILE_DETAILS
			str = str & fileListLink
			str = str & html(page.File.FileName)
		Case Else
			str = str & "Files"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href		: href = ""
	
	Dim bulkDeleteButton
	href = "#"
	bulkDeleteButton = "<li><a href=""" & href & """ class=""bulk-delete-link""><img class=""icon"" src=""/_images/icons/page_copy_delete.png"" alt="""" /></a><a href=""" & href & """ class=""bulk-delete-link"">Remove Selected</a></li>"
	
	Dim uploadFileButton
	pg.Action = ADDNEW_RECORD
	href = pg.Url & pg.UrlParamsToString(True)
	uploadFileButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/page_add.png"" alt="""" /></a><a href=""" & href & """>Add File</a></li>"

	Dim linkEventFilesButton
	pg.Action = LINK_FILES_TO_EVENTS
	href = pg.Url & pg.UrlParamsToString(True)
	linkEventFilesButton =  "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/date_link.png""  alt="""" /></a><a href=""" & href & """>Event Links</a></li>"

	Dim fileListButton
	pg.Action = "": pg.EventID = "": pg.ScheduleID = ""
	href = pg.Url & pg.UrlParamsToString(True)
	fileListButton =  "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/page.png"" alt="""" /></a><a href=""" & href & """>File List</a></li>"
	
	Select Case page.Action
		Case SHOW_FILE_DETAILS
			str = str & fileListButton
		Case LINK_FILES_TO_EVENTS
			str = str & ProgramDropdownToString(page)
			str = str & fileListButton
		Case DELETE_RECORD
			str = str & fileListButton
		Case UPDATE_RECORD
			str = str & fileListButton
		Case ADDNEW_RECORD
			str = str & fileListButton
		Case Else
			str = str & SortByDropdownToString(page)
			str = str & ProgramDropdownToString(page)
			str = str & linkEventFilesButton
			str = str & BulkDeleteButton
			str = str & uploadFileButton
	End Select
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_file_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_displayer_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_download_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_YesNoDropdownToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_CleanFileName.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	Public SortBy
	
	' encrypted
	Public Action
	Public ProgramID
	Public FileID
	Public EventID
	Public ScheduleID
	Public MemberId
	
	' objects
	Public Member
	Public Client
	Public Program	
	Public File
	Public Schedule
	Public Evnt
	Public Uploader	
	
	' don't persist
	Public FileIDLIst
	Public DeleteEventFileIDList
	Public InsertFileIDList
	
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
		If Len(EventID) > 0 Then str = str & "eid=" & Encrypt(EventID) & amp
		If Len(ScheduleID) > 0 Then str = str & "scid=" & Encrypt(ScheduleID) & amp
		If Len(MemberId) > 0 Then str = str & "mid=" & Encrypt(MemberId) & amp
		
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
		c.SortBy = SortBy

		c.Action = Action
		c.ProgramID = ProgramID
		c.FileID = FileID
		c.EventID = EventID
		c.ScheduleID = ScheduleID
		c.MemberId = MemberId
				
		c.FileIDLIst = FileIDList
		c.InsertFileIDList = InsertFileIDList
		c.DeleteEventFileIDList = DeleteEventFileIDList
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.File = File
		Set c.Schedule = Schedule
		Set c.Evnt = Evnt
		Set c.Uploader = Uploader
		
		Set Clone = c
	End Function
End Class
%>

