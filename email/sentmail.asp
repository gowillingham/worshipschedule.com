<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-email"
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
	page.EmailID = Decrypt(Request.QueryString("emid"))
	
	page.EmailIDList = Request.Form("EmailIDList")

	' paging parameters
	page.PageNumber = Request.QueryString("pn")
	If Len(page.PageNumber) = 0 Then page.PageNumber = 1
	page.PageSize = Request.QueryString("ps")
	If Request.Form("FormPageSizeIsPostback") = IS_POSTBACK Then
		page.PageSize = Request.Form("PageSize")
		page.PageNumber = 1
	End If
	If Len(page.PageSize) = 0 Then page.PageSize = 25

	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()
	
	If Request.Form("FormPageSizeIsPostback") = IS_POSTBACK Then
		page.PageSize = Request.Form("PageSize")
	End If
	
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
		<link rel="stylesheet" type="text/css" href="/_incs/script/jquery/jquery-ui-1.6.custom/development-bundle/themes/start/ui.all.css" />	
		<link rel="stylesheet" type="text/css" href="sentmail.css" />
		
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/jquery-ui-1.6.custom/js/jquery-ui-1.6.custom.min.js"></script>
		<script type="text/javascript" language="javascript" src="sentmail.js"></script>
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
		Case DELETE_RECORD
			Call DoDeleteEmail(page.EmailID, rv)
			Select Case rv
				Case 0
					page.MessageID = 7010
				Case Else
					page.MessageID = 7011
			End Select
			page.Action = "": page.EmailID = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case DELETE_EMAIL_BULK
			Call DoDeleteEmailBulk(page.EmailIDList, rv)
			Select Case rv
				Case 0 
					page.MessageID = 7013
				Case -2
					' no messages selected
					page.MessageID = 7012
				Case Else
					' unknown error
					page.MessageID = 7014
			End Select
			page.Action = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case RESEND_EMAIL_MESSAGE
			page.EmailID = DoResendEmailMessage(page.EmailID, rv)
			page.Action = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
			
		Case VIEW_EMAIL_MESSAGE
			str = str & SentMessageToString(page)
			
		Case Else
			str = str & SentMailGridToString(page)
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub ClearTablinkBar()
	m_tabLinkBarText = "<li>&nbsp;</li>"
End Sub

Sub DoDeleteEmail(emailID, outError)
	Dim email		: Set email = New cEmail
	
	email.EmailID = emailID
	Call email.Delete(outError)
End Sub

Sub DoDeleteEmailBulk(idList, outError)
	Dim i
	Dim thisError			: thisError = 0
	outError = 0
		
	If Len(idList) = 0 Then 
		' no files selected
		outError = -2
		Exit Sub
	End If
	
	Dim list			: list = Split(idList, ",")
	Dim email			: Set email = New cEmail
	For i = 0 To UBound(list)
		email.EmailID = list(i)
		Call email.Delete(thisError)
		If thisError <> 0 Then
			outError = -1
		End If
	Next
End Sub

Function DoResendEmailMessage(ByRef id, outError)
	Dim email			: Set email = New cEmail
	
	' load email message from id passed in ..
	email.EmailID = id
	Call email.Load()
	
	' check to see if attachments need
	' to be copied to this message ..
	Dim fso				: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim hasFiles		: hasFiles = False
	Dim oldPath			: oldPath = Application.Value("EMAIL_ATTACHMENTS_DIRECTORY") & email.EmailID & "\"
	If fso.FolderExists(oldPath) Then
		hasFiles = True
	End If
	
	' clear relevant fields before saving this message 
	' as a new message ..
	email.EmailID = ""
	email.IsSent = ""
	email.IsMarkedForDelete = 0
	email.DateSent = ""
	Call email.Add(outError)
	If outError <> 0 Then Exit Function
	
	' copy attachments folder if necessary ..
	Dim newPath			: newPath = Application.Value("EMAIL_ATTACHMENTS_DIRECTORY") & email.EmailID & "\"
	Dim folder, files, file
	If hasFiles Then
		fso.CreateFolder(newPath)
		Set folder = fso.GetFolder(oldPath)
		Set files = folder.Files
		For Each file In files
			fso.CopyFile oldPath & file.Name, newPath & file.Name
		Next
	End If

	' hack: 
	' ---------
	' this fn was orignally a sub but the value was not being passed
	' back byref as expected when setup like this ..
	' Sub DoResendEmailMessage(ByRef id, outError)
	' if I passed id = page.EmailID and then changed val of id in sub
	' it was not reflected in value of page.EmailID

	' pass new emailID back to caller ..
	DoResendEmailMessage = email.EmailID
End Function

Function PagerDropdownToString(page)
	Dim str, i
	Dim list			: ReDim list(1,4)
	Dim arr				: arr = Split("5,10,25,50,100", ",")
	Dim selected		: selected = ""
	
	For i = 0 To UBound(arr)
		list(0,i) = arr(i)
		list(1,i) = arr(i) & " Messages"
	Next
	
	str = str & "<form class=""form"" method=""post"" action=""" & Request.ServerVariables("URL") & page.UrlParamsToString(True) & """ name=""formPageSize"">"
	str = str & "<input type=""hidden"" name=""FormPageSizeIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<select name=""PageSize"" onchange=""document.formPageSize.submit();"">"
	For i = 0 To UBound(list,2)
		selected = ""
		If CStr(page.PageSize) = CStr(list(0,i)) Then selected = " selected=""selected"""
		str = str & "<option value=""" & list(0,i) & """" & selected & ">" & list(1,i) & "</option>"
	Next
	str = str & "</select></form>"
	
	PagerDropdownToString = str
End Function

Function PagerToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	
	' set max number of page links to display 
	Dim NUM_LINKS			: NUM_LINKS = 5
	
	' get the total number of messages/rows in db
	Dim email				: Set email = New cEmail
	email.MemberID = page.Member.MemberID
	Dim sentMessageCount	: sentMessageCount = email.MessageCount(RETURN_SENT_MESSAGES)
	
	
	Dim thisPage			: thisPage = page.PageNumber
	Dim totalPages			: totalPages = Int(sentMessageCount/page.PageSize)
	If sentMessageCount Mod page.PageSize > 0 Then totalPages = totalPages + 1	
	
	
	' get first page number for pager ..
	Dim firstPage			: firstPage = thisPage - Int(NUM_LINKS/2)
	If thisPage - NUM_LINKS < 0 Then firstPage = 1
	
	' get last page number for pager ..
	Dim lastPage			: lastPage = firstPage + NUM_LINKS - 1
	If lastPage > totalPages Then lastPage = totalPages
	
	str = str & "<div class=""pager"">Page "
	pg.PageNumber = 1
	str = str & "<a href=""" & Request.ServerVariables("URL") & pg.UrlParamsToString(True) & """ title=""Most Recent Email"">"
	str = str & Html("<<") & "</a>"
	pg.PageNumber = page.PageNumber - 1
	If CInt(thisPage) <> 1 Then
		str = str & "&nbsp;&nbsp;" & "<a href=""" & Request.ServerVariables("URL") & pg.UrlParamsToString(True) & """ title=""More Recent"">"
		str = str & Html("<") & "</a>"
	End If
	
	For i = firstPage To lastPage
		If CInt(i) = CInt(thisPage) Then
			str = str & "&nbsp;<span style=""font-weight:bold;"">" & i & "</span>"
		Else 
			pg.PageNumber = i
			str = str & "&nbsp;" & "<a href=""" & Request.ServerVariables("URL") & pg.UrlParamsToString(True) & """ title=""New Page"">" & i & "</a>"
		End If
	Next
	
	pg.PageNumber = thisPage + 1
	If CInt(thisPage) <> CInt(totalPages) Then
		str = str & "&nbsp;&nbsp;" & "<a href=""" & Request.ServerVariables("URL") & pg.UrlParamsToString(True) & """ title=""Older"">" & Html(">") & "</a>"
	End If
	
	pg.PageNumber = totalPages
	str = str & "&nbsp;&nbsp;" & "<a href=""" & Request.ServerVariables("URL") & pg.UrlParamsToString(True) & """ title=""Oldest Email"">" & Html(">>") & "</a>"
	str = str & " of " & totalPages & "&nbsp;&nbsp;(" & sentMessageCount & " rows)"
	
	str = str & "&nbsp;|&nbsp;"& PagerDropdownToString(page) & "</div>"
	
	PagerToString = str 
End Function

Function SentMailGridToString(page)
	Dim str, msg, i
	Dim pg				: Set pg = page.Clone()
	Dim dateTime		: Set dateTime = New cFormatDate
	Dim count			: count = 0
	Dim altClass		: altClass = ""
	Dim emailIcon		: emailIcon = ""
	Dim tipBox
	
	Dim email			: Set email = New cEmail
	email.MemberID = page.Member.MemberID
	Dim list			: list = email.SentMessageList(page.PageNumber, page.PageSize)
	
	tipBox = tipBox & "<div class=""tip-box""><h3>Tip!</h3><p>"
	tipBox = tipBox & "This list is an archive of the email messages you have sent through your " & Application.Value("APPLICATION_NAME") & " account. </p></div>"
	
	' 0-EmailID 1-ClientID 2-Subject 3-Text 4-IsMarkedForDelete 5-IsSent 6-RecipientIDList
	' 7-RecipientAddressList 8-BccAddressList 9-CcAddressList 10-DateCreated 11-DateModified 12-DateSent 
	' 13-GroupList 14-AttachmentList 15-RowID

	str = str & tipBox
	If IsArray(list) Then
		str = str & m_appMessageText
		str = str & "<div class=""grid"">"
		str = str & PagerToString(page)
		pg.Action = DELETE_EMAIL_BULK
		str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-bulk-delete"" >"
		str = str & "<table>"
		str = str & "<tr><th scope=""col"" style=""width:1%;""><input type=""checkbox"" name=""master"" id=""master"" /></th>"
		str = str & "<th scope=""col"">Sent Mail</th><th scope=""col"">Date</th><th scope=""col"" style=""width:1%;"">&nbsp;</th></tr>"
		For i = 0 To UBound(list,2)
			count = count + 1
			altClass = ""
			If count Mod 2 = 0 Then altClass = " class=""alt"""
			emailIcon = "email.png"
			If Len(list(14,i)) > 0 Then emailIcon = "email_attach.png"
		
			str = str & "<tr" & altClass & "><td><input type=""checkbox"" value=""" & list(0,i) & """ class=""checkbox"" name=""EmailIdList"" /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/" & emailIcon & """ alt="""" />"
			pg.EmailID = list(0,i)
			pg.Action = VIEW_EMAIL_MESSAGE
			str = str & "<strong><a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details"">" & html(list(2,i)) & "</a></strong></td>"
			str = str & "<td>" & dateTime.Casual(list(12,i)) & "</td>"
			str = str & "<td class=""toolbar"">"
			
			pg.Action = VIEW_EMAIL_MESSAGE
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details"">"
			str = str & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a>"
			pg.Action = RESEND_EMAIL_MESSAGE
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Resend"">"
			str = str & "<img src=""/_images/icons/email_go.png"" alt="""" /></a>"
			pg.Action = DELETE_RECORD
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Delete"">"
			str = str & "<img src=""/_images/icons/cross.png"" alt="""" /></a>"
			str = str & "</td></tr>"
		Next
		str = str & "</table></form></div>"
	End If
	
	If count = 0 Then
		str = NoSentMailDialogToString(page)
		Call ClearTablinkBar()
	End If

	SentMailGridToString = str
End Function

Function NoSentMailDialogToString(page)
	Dim dialog				: Set dialog = New cDialog
	Dim pg					: Set pg = page.Clone()
	
	dialog.Headline = "Nothing to see here .. yet!"

	dialog.Text = dialog.Text & "<p>There are no messages in your sent mail archive to show you. "	
	dialog.Text = dialog.Text & "Either you haven't sent any messages yet, or you recently deleted all of your sent mail. "	
	dialog.Text = dialog.Text & ""	
	dialog.Text = dialog.Text & "</p>"
	
	dialog.SubText = dialog.Subtext & "<p>" & Application.Value("APPLICATION_NAME") & " saves every email message that you send through your account. "
	dialog.SubText = dialog.Subtext & "You can use this page to resend or review any messages you have sent to your members. "
	dialog.SubText = dialog.Subtext & ""
	dialog.SubText = dialog.Subtext & ""
	dialog.SubText = dialog.Subtext & "</p>"
	
	pg.Action = "": pg.PageNumber = "": pg.PageSize = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/email/email.asp" & pg.UrlParamsToString(True) & """>Send a message</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/help/topic.asp?hid=14#anchor-email-team"" target=""_blank"">Learn more about email</a></li>"

	NoSentMailDialogToString = dialog.ToString()
End Function

Function GroupListToString(stringList)
	Dim i
	
	If Len(stringList) = 0 Then Exit Function
	
	Dim outList				: outList = ""
	Dim groupName			: groupName = ""
	Dim list				: list =  Split(stringList, ",")
	If Not IsArray(list) Then Exit Function
	
	Dim smartGroup			: Set smartGroup = New cSmartGroup
	For i = 0 To UBound(list)
		groupName = ""
		smartGroup.SmartGroupID = list(i)
		groupName = smartGroup.Name()
		If Len(groupName) > 0 Then
			outList = outList & groupName & ","
		End If
	Next
	If Len(outList) > 0 Then outList = Left(outList, Len(outList) - 1)
	
	GroupListToString = outList
End Function

Function SentMessageToString(page)
	Dim str
	Dim pg							: Set pg = page.Clone()
	Dim dateTime					: Set dateTime = New cFormatDate
	
	Dim email						: Set email = New cEmail
	email.EmailID = page.EmailID		: If Len(email.EmailID) > 0 Then email.Load()
	
	Dim subject				: subject = email.Subject
	If Len(subject & "") = 0 Then subject = "<No Subject>"
	Dim text				: text = email.Text
	If Len(text & "") = 0 Then text = "<Message Contains no Text>"
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3><ul>"
	pg.Action = DELETE_RECORD
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Delete this message</a></li>"
	pg.Action = RESEND_EMAIL_MESSAGE
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Resend this message</a></li></ul></div>"
	
	str = str & "<div class=""email-message-detail"">"
	
	str = str & "<h3>" & html(email.Subject) & "</h3>"
	str = str & "<table id=""email-message-collapsed"">"
	str = str & "<tr><td class=""header""><img class=""icon"" src=""/_images/icons/bullet_green_8x8.png"" alt="""" />"
	str = str & "<strong>" & html(page.Member.NameFirst & " " & page.Member.NameLast) & "</strong> to " & html(email.FirstRecipient) & "</td>"
	str = str & "<td class=""links""><a href=""#"" id=""expand-details"">«show details</a> <span class=""softer"">" & dateTime.Casual(email.DateSent) & "</span></td></tr>"
	str = str & "</table>"
	
	str = str & "<table id=""email-message-expanded"">"
	str = str & "<tr><td class=""label softer"">from</td>"
	str = str & "<td><strong>" & html(page.Member.NameFirst & " " & page.Member.NameLast) & "</strong> <span class=""softer"">" & html("<" & page.Member.Email & ">") & "</span></td>"
	str = str & "<td class=""links""><a href=""#"" id=""collapse-details"">»hide details</a> <span class=""softer"">" & dateTime.Casual(email.DateSent) & "</span></td></tr>"
	If Len(email.RecipientAddressList) > 0 Then
		str = str & "<tr><td class=""label softer"">to</td>"
		str = str & "<td colspan=""2"">"& html(Replace(email.RecipientAddressList, ",", ", ")) & "</td></tr>"
	End If
	If Len(email.CcAddressList) > 0 Then
		str = str & "<tr><td class=""label softer"">cc</td>"
		str = str & "<td colspan=""2"">"& html(Replace(email.CcAddressList, ",", ", ")) & "</td></tr>"
	End If
	If Len(email.BccAddressList) > 0 Then
		str = str & "<tr><td class=""label softer"">bcc</td>"
		str = str & "<td colspan=""2"">"& html(Replace(email.BccAddressList, ",", ", ")) & "</td></tr>"
	End If
	If Len(email.GroupList) > 0 Then
		str = str & "<tr><td class=""label softer"">groups</td>"
		str = str & "<td colspan=""2"">"& html(Replace(GroupListToString(email.GroupList), ",", ", ")) & "</td></tr>"
	End If
	If Len(email.AttachmentList) > 0 Then
		Dim attachmentList
		attachmentList = Replace(email.AttachmentList, ",", "], [")
		attachmentList = "[" & attachmentList & "]"
		str = str & "<tr><td class=""label softer"">attach</td>"
		str = str & "<td colspan=""2"">"& html(attachmentList) & "</td></tr>"
	End If
	
	str = str & "<tr><td class=""label softer"">date</td>"
	str = str & "<td colspan=""2"">"& dateTime.Convert(email.DateSent, "DDD MMM dd, YYYY at hh:nn pp") & "</td></tr>"
	str = str & "<tr><td class=""label softer"">subject</td>"
	str = str & "<td colspan=""2"">" & html(email.Subject) & "</td></tr>"
	str = str & "<tr><td class=""label softer"">mailed-by</td>"
	str = str & "<td colspan=""2"">" & Application.Value("APPLICATION_NAME") & "</td></tr>"
	str = str & "</table>"
	str = str & "<p class=""text"">" & Replace(html(email.Text), vbCrLf, "<br />") & "</p>"
	str = str & "</div>"
	
	SentMessageToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim email
	
	
	Dim emailLink
	pg.EmailID = "": pg.Action = ""
	emailLink = "<a href=""/email/email.asp" & pg.UrlParamsToString(True) & """>Email</a> / "
	
	Dim sentEmailLink
	pg.Action = ""
	sentEmailLink = "<a href=""/email/sentmail.asp" & pg.UrlParamsToString(True) & """>Sent Email</a> / "
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case VIEW_EMAIL_MESSAGE
			Set email = New cEmail
			email.EmailID = page.EmailID
			email.Load()
			str = str & emailLink & sentEmailLink & html(email.Subject)
		Case Else
			str = str & emailLink & "Sent Email"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Dim bulkDeleteButton
	href = "#"
	bulkDeleteButton = "<li id=""bulk-delete-button""><a href=""" & href & """><img class=""icon"" src=""/_images/icons/email_multiple.png"" /></a><a href=""" & href & """>Delete Checked</a></li>"
	
	Dim resendEmailButton
	pg.Action = RESEND_EMAIL_MESSAGE
	href = pg.Url & pg.UrlParamsToString(True)
	resendEmailButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/email_go.png"" /></a><a href=""" & href & """>Resend</a></li>"
	
	Dim deleteEmailButton
	pg.Action = DELETE_RECORD
	href = pg.Url & pg.UrlParamsToString(True)
	deleteEmailButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/cross.png"" /></a><a href=""" & href & """>Delete</a></li>"
	
	Dim sentEmailListButton
	pg.Action = ""
	href = "/email/sentmail.asp" & pg.UrlParamsToString(True)
	sentEmailListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/email_date.png"" /></a><a href=""" & href & """>Sent Email</a></li>"
	
	Dim composeEmailButton
	' clear this to keep emailID for message being viewed from being sent back to compose page
	If page.Action = VIEW_EMAIL_MESSAGE Then pg.EmailID = ""
	pg.Action = ""
	href = "/email/email.asp" & pg.UrlParamsToString(True)
	composeEmailButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/email_edit.png"" /></a><a href=""" & href & """>Compose</a></li>"
	
	Select Case page.Action
		Case VIEW_EMAIL_MESSAGE
			str = str & deleteEmailButton
			str = str & resendEmailButton
			str = str & sentEmailListButton
			str = str & composeEmailButton
		Case Else
			str = str & bulkDeleteButton
			str = str & composeEmailButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/smart_group_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/email_group_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_availability_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_group_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	Public PageNumber
	Public PageSize
	Public EmailIDList
	
	' encrypted
	Public Action
	Public ProgramID
	Public EmailID
	
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
		If Len(EmailID) > 0 Then str = str & "emid=" & Encrypt(EmailID) & amp
		If Len(PageNumber) > 0 Then str = str & "pn=" & PageNumber & amp
		If Len(PageSize) > 0 Then str = str & "ps=" & PageSize & amp
		
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
		c.PageNumber = PageNumber
		c.PageSize = PageSize
		
		c.EmailIDList = EmailIDList
		
		c.Action = Action
		c.ProgramID = ProgramID
		c.EmailID = EmailID
				
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		
		Set Clone = c
	End Function
End Class
%>

