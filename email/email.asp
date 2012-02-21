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
	page.GroupID = Decrypt(Request.QueryString("emgid"))

	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()

	' if no emailID then create new email
	Set page.Email = New cEmail
	page.Email.EmailID = page.EmailID
	If Len(page.Email.EmailID) = 0 Then
		page.Email.MemberID = page.Member.MemberID
		page.Email.ClientID = page.Client.ClientID
		Call page.Email.Add("")
		page.EmailID = page.Email.EmailID
	End If
	Call page.Email.Load()
	
	Set page.Uploader = Server.CreateObject("ASPSmartUpload.SmartUpload")
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		page.Uploader.Upload()
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
		<link rel="stylesheet" type="text/css" href="/_incs/script/jquery/plugins/autocomplete/jquery.autocomplete.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/fileupload/jquery.MultiFile.pack.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/autocomplete/jquery.autocomplete.pack.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/cookie/jquery.cookie.js"></script>
		<script type="text/javascript" language="javascript">
		
			$(document).ready(function(){
				// wire up auto-complete for recipient form elements ..
				var ajax = $.ajax({
					url: "/_incs/script/ajax/_address_book.asp?sid=" + $.cookie("sid"),
					async: false
				});
				
				var data = ajax.responseText
				data = data.toString().split(",")
				
				$("#to,#cc,#bcc").autocomplete(data, {
					dataType: "json",
					multiple: true,
					matchContains: true
				});
				
				// wire up buttons that must submit email form before redirecting ..
				$("a.redirect-with-submit").click(function(){
					var url = $(this)[0].href;
					$("#form-message").attr("action", url).submit();
					return false;
				});
				
				// wire up bcc, cc text boxes ..
				$("#cc-switch").click(function(){
					$(this).hide();
					$("#row-cc").show();
					$("#cc").focus();
					
				});
				$("#bcc-switch").click(function(){
					$(this).hide();
					$("#row-bcc").show();
					$("#bcc").focus();
				});
				
				// make sure bcc, cc is visible if non-empty ..
				if ($("#cc").val().length > 0) {
					$("#row-cc").show();
					$("#cc-switch").hide();
				};
				if ($("#bcc").val().length > 0) {
					$("#row-bcc").show();
					$("#bcc-switch").hide();
				};

				// wire up file attachment ..
				$("#file-attachment").hide();
				$("#file-attachment-trigger a").click(function(){
					$("#file-attachment").show();
					$("#file-attachment-trigger").hide();
				});		
				
				// this gives the file inputs in multi-file plugin unique names ..
				$('#form-message').submit(function(){
					var files = $('#form-message input:file');
					var count=0;
					files.attr('name',function(){return this.name+''+(count++);});
				});

				// focus cursor in to field ..
				$("#to").focus();
				
			});
		</script>
		<style type="text/css">
			.form, .message {width:622px;}
			.control {margin:8px 0 0 0;}
			.attachment-list .message {}
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
		Case ADD_RECIPIENTS
			Call LoadMessageFromRequest(page)
			Call RemoveAttachments(page)
			Call AddAttachments(page)
			Call DoUpdateMessage(page, rv)
			
			page.Action = ""
			Response.Redirect("/email/contacts.asp" & page.UrlParamsToString(False))
	
		Case SEND_MESSAGE
			Call LoadMessageFromRequest(page)
			Call RemoveAttachments(page)
			Call AddAttachments(page)
			
			If ValidMessage(page, True) Then
				Call SendMessage(page, rv)
				page.MessageID = 7009: page.Action = "": page.EmailID = "": page.ProgramID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormMessageToString(page)
			End If
				
		Case Else
			str = str & FormMessageToString(page)
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoUpdateMessage(page, outError)
	Call page.Email.Save(outError)
End Sub

Sub RemoveAttachments(page)
	Dim i
	Dim fso				: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim path			: path = Application.Value("EMAIL_ATTACHMENTS_DIRECTORY") & page.Email.EmailID
	Dim files
	Dim file
	Dim removeFile
				
	Dim filesToSave
	If Len(page.AttachmentList) > 0 Then
		filesToSave = Split(page.AttachmentList, ",")
	End If
	
	If Not fso.FolderExists(path) Then
		Exit Sub
	End If
	
	Set files = fso.GetFolder(path).Files
	For Each file In files
		removeFile = True
		If IsArray(filesToSave) Then
			For i = 0 To UBound(filesToSave) 
				If CStr(filesToSave(i)) = CStr(file.Name) Then
					removeFile = False
				End If
			Next
		End If
		If removeFile Then
			Call file.Delete()
		End If
	Next
End Sub

Sub AddAttachments(page)
	Dim fso				: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim path			: path = Application.Value("EMAIL_ATTACHMENTS_DIRECTORY") & page.Email.EmailID
	Dim fileName
	
	If Not fso.FolderExists(path) Then
		fso.CreateFolder(path)
	End If
	
	Dim file
	For Each file In page.Uploader.Files
		If Not file.IsMissing Then
			fileName = file.FileName
			fileName = Replace(fileName, ",", "")
			file.SaveAs path & "\" & fileName
		End If
	Next
End Sub

Sub AddGroupsToRecipients(ByRef email)
	Dim i
	Dim smartGroup			: Set smartGroup = New cSmartGroup
	Dim recipients			: recipients = ""
	Dim outList				: outList = ""
	
	If Len(email.GroupList) = 0 Then Exit Sub
	
	Dim list			: list = Split(email.GroupList, ",")
	
	If IsArray(list) Then
		For i = 0 To UBound(list)
		
			smartGroup.SmartGroupID = list(i)
			recipients = smartGroup.GetAddressListAsString()
			
			If Len(recipients) > 0 Then
				outList = outList & recipients & ","
			End If
		Next
		If Len(outList) > 0 Then outList = Left(outList, Len(outList) - 1)
		
		If Len(email.RecipientAddressList) > 0 Then 
			email.RecipientAddressList = email.RecipientAddressList & "," & outList
		End If
	End If
End Sub

Sub SendMessage(page, outError)
	Dim str, i
	Dim message				: Set message = New cEmailSender
	Dim fso			: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim files, file
	Dim path		: path = Application.Value("EMAIL_ATTACHMENTS_DIRECTORY") & page.EmailID
	Dim list
	
	' attach files ..
	If fso.FolderExists(path) Then
	
		' clear attachment list in case this is a resend ..
		page.Email.AttachmentList = ""
		
		' generate a list of files to attach ..
		Set files = fso.GetFolder(path).Files
		For Each file In files
			message.AddAttachment(path & "\" & file.Name)
			page.Email.AttachmentList = page.Email.AttachmentList & file.Name & ","
		Next
		If Len(page.Email.AttachmentList) > 0 Then page.Email.AttachmentList = Left(page.Email.AttachmentList, Len(page.Email.AttachmentList) - 1)
	End If
	
	Call AddGroupsToRecipients(page.Email)
	
	page.Email.RecipientAddressList = RemoveDupesFromStringList(page.Email.RecipientAddressList)
	page.Email.BccAddressList = RemoveDupesFromStringList(page.Email.BccAddressList)
	page.Email.CcAddressList = RemoveDupesFromStringList(page.Email.CcAddressList)
	
	If Len(page.Email.BccAddressList & page.Email.CcAddressList) = 0 Then
		' send one message to each address ..
		list = SplitListToArray(page.Email.RecipientAddressList)
		If IsArray(list) Then
			For i = 0 To UBound(list)
				Call message.SendMessage(list(i), page.Member.Email, page.Email.Subject, page.Email.Text & EmailDisclaimerToString(page.Client.NameClient))
			Next
		End If
	Else
		' send one message with multiple recipients ..
		If Len(page.Email.RecipientAddressList) > 0 Then
			message.ToAddress = Join(SplitListToArray(page.Email.RecipientAddressList), ";")
		End If
		If Len(page.Email.CcAddressList) > 0 Then
			message.CcAddress = Join(SplitListToArray(page.Email.CcAddressList), ";")
		End If
		If Len(page.Email.BccAddressList) > 0 Then
			message.BccAddress = Join(SplitListToArray(page.Email.BccAddressList), ";")
		End If

		message.From = page.Member.Email
		message.Text = page.Email.Text & EmailDisclaimerToString(page.Client.NameClient)
		message.Subject = page.Email.Subject
		
		Call message.Send()
	End If
	
	' save this sent message to db ..
	page.Email.IsSent = IS_SENT
	page.Email.DateSent = Now()
	Call page.Email.Save(outError)
	
End Sub

Function FormMessageToString(page)
	Dim str
	Dim pg						: Set pg = page.Clone()
	
	Dim sendCancelButtons
	
	Dim showCcRowStyle				: showCcRowStyle = ""
	Dim showCcSwitchStyle			: showCcSwitchStyle = ""
	If Len(page.Email.CcAddressList) > 0 Then
		showCcRowStyle = ""
		showCcSwitchStyle = " style=""display:none;"""
	Else
		showCcRowStyle = " style=""display:none;"""
	End If	
	Dim showBccRowStyle			: showBccRowStyle= ""
	Dim ShowBccSwitchStyle		: ShowBccSwitchStyle = ""
	If Len(page.Email.BccAddressList) > 0 Then
		ShowBccRowStyle = ""
		showBccSwitchStyle = " style=""display:none;"""
	Else
		ShowBccRowStyle = " style=""display:none;"""
	End If
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3><p>"
	str = str & "Click <strong>address book</strong> to select members from your " & Application.Value("APPLICATION_NAME") & " account for the To, CC, or BCC list for this message. </p></div>"
	str = str & m_appMessageText
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	
	pg.Action = SEND_MESSAGE
	
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" enctype=""multipart/form-data"" id=""form-message"">"
	str = str & "<input type=""hidden"" name=""FormMessageIsPostback"" value=""" & IS_POSTBACK & """ />"
	
	pg.Action = "": pg.EmailID = ""
	sendCancelButtons = sendCancelButtons & "<input type=""submit"" name=""Submit"" value=""Send"" />"
	sendCancelButtons = sendCancelButtons & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"

	str = str & "<table style=""width:575px;"">"
	str = str & "<tr><td>&nbsp;</td><td style=""text-align:right;"">" & sendCancelButtons & "</td></tr>"
	str = str & "<tr id=""always-displays""><td class=""label"" style=""width:1%;"">" & RequiredElementToString(True, "To:") & "</td>"
	str = str & "<td><textarea name=""RecipientAddressList"" id=""to"">" & html(FixSpaces(page.Email.RecipientAddressList)) & "</textarea></td></tr>"
	' cc row
	str = str & "<tr id=""row-cc""" & showCcRowStyle & "><td class=""label"" style=""width:1%;"">" & RequiredElementToString(False, "CC:") & "</td>"
	str = str & "<td><textarea name=""CcAddressList"" id=""cc"">" & html(FixSpaces(page.Email.CcAddressList)) & "</textarea></td></tr>"
	' bcc row
	str = str & "<tr id=""row-bcc""" & showBccRowStyle & "><td class=""label"" style=""width:1%;"">" & RequiredElementToString(False, "BCC:") & "</td>"
	str = str & "<td><textarea name=""BccAddressList"" id=""bcc"" class=""email-recipients"">" & html(FixSpaces(page.Email.BccAddressList)) & "</textarea></td></tr>"
	' cc/bcc link row
	str = str & "<tr id=""row-switch""><td>&nbsp;</td><td>"
	str = str & "<span id=""cc-switch""" & showCcSwitchStyle & ">"
	str = str & "<a href=""#"" id=""cc-switch-href"">Add CC</a> | </span>"
	str = str & "<span id=""bcc-switch""" & showBccSwitchStyle & ">"
	str = str & "<a href=""#"" id=""bcc-switch-href"">Add BCC</a> | </span>"
	
	' refresh pg object to get repopulate emailID
	Set pg = page.Clone()
	pg.Action = ADD_RECIPIENTS
	str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ class=""redirect-with-submit"">Add Group</a> "
	str = str & "<span id=""file-attachment-trigger"">| <a href=""#"">Attach file</a></span>"
	
	' this is loaded with file inputs via javascript
	str = str & "<div id=""file-input-wrapper"" style=""margin-top:10px;"">"
	str = str & "<input type=""file"" id=""file-attachment"" name=""file_attachment_"" class=""multi"" />"
	str = str & "</div>"
	
	str = str & AttachmentListToString(page)
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Subject:</td>"
	str = str & "<td><input id=""email-subject"" type=""text"" name=""Subject"" value=""" & HTML(page.Email.Subject) & """ /></td></tr>"
	str = str & "<tr><td class=""label"">Message:</td>"
	str = str & "<td><textarea id=""email-text"" name=""text"">" & HTML(page.Email.Text) & "</textarea></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td style=""text-align:right;"">" & sendCancelButtons & "</td></tr>"
	str = str & "</table></form></div>"
	FormMessageToString = str
End Function

Function AttachmentListToString(page)
	' returns comma delim list of files from attachment folder for this messageID
	Dim str, msg
	Dim pg				: Set pg = page.Clone()
	Dim icon			: Set icon = New cFileDisplay
	Dim fso				: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim files, file
	Dim path			: path = Application.Value("EMAIL_ATTACHMENTS_DIRECTORY") & page.Email.EmailID
	Dim href			: href = ""
	Dim hasFileList		: hasFileList = False
	Dim hasFiles		: hasFiles = False
	
	' check for saved email has attachments but no attachment folder is there (exceeded attachmentTimeToLive)	
	If Len(page.Email.AttachmentList) > 0 Then hasFileList = True
	If Not fso.FolderExists(path) Then 
		If hasFileList Then
			msg = msg & "The original files attached to this message (" & html(page.Email.AttachmentList) & ") are no longer available. "
			msg = msg & Application.Value("APPLICATION_NAME") & " will only store attachments for sent email messages for " & Application.Value("EMAIL_ATTACHMENT_MONTHS_TO_LIVE") & " months before they are removed. "
			msg = msg & "You will need to re-attach your files manually. "
			AttachmentListToString = "<div class=""attachment-list"">" & CustomApplicationMessageToString("Original attachments not available! ", msg, "Error") & "</div>"
		End If
		Exit Function
	End If
	
	Set files = fso.GetFolder(path).Files
	If files.Count > 0 Then
		str = str & "<ul class=""attachment-list"" style="""">"
		For Each file In files
			str = str & "<li style=""""><input type=""checkbox"" class=""checkbox"" name=""AttachmentList"" value=""" & file.Name & """ checked=""checked"" />"
			pg.Action = STREAM_FILE_TO_BROWSER
			href = pg.Url & pg.UrlParamsToString(True)
			str = str & "<a href=""" & href & """><img class=""icon"" src=""" & icon.GetIconPath(Split(file.Name, ".")(UBound(Split(file.Name, ".")))) & """ alt="""" /></a>"
			str = str &"<a href=""" & href & """>" & html(file.Name) & "</a></li>"
		Next
		str = str & "</ul>"
	End If

	AttachmentListToString = str
End Function

Function FixSpaces(str)
	If Len(str) > 0 Then 
		' remove all spaces then add space after each comma
		str = Replace(str, " ", "")
		str = Replace(str, ",", ", ")
	End If
	
	FixSpaces = str
End Function

Function ValidMessage(page, isSend)
	Dim str, i
	Dim list
	
	ValidMessage = True
	If Not ValidData(page.Email.Subject, False, 0, 200, "Subject", "") Then ValidMessage = False
	If Not ValidData(page.Email.Text, False, 0, 4000, "Message", "") Then ValidMessage = False
	If isSend Then
		' check for both subject/text
		If Len(page.Email.Subject & page.Email.Text) = 0 Then
			AddCustomFrmError("Your message must have either a subject or a message. They cannot both be blank.")
			ValidMessage = False
		End If
		
		' check for at least one recipient
		If Len(page.Email.RecipientAddressList & page.Email.CcAddressList & page.Email.BccAddressList & page.Email.GroupList) = 0 Then
			AddCustomFrmError("Your message needs at least one recipient or group before it can be sent. ")
			ValidMessage = False
		End If
		
		' check recipient lists for invalid email addresses ..
		list = SplitListToArray(page.Email.RecipientAddressList)
		If IsArray(list) Then
			For i = 0 To UBound(list)
				If Len(Trim(list(i))) > 0 Then
					If Not IsEmail(Trim(list(i))) Then
						AddCustomFrmError("To address '" & list(i) & "' does not appear to be a valid email address. ")
						ValidMessage = False
					End If
				End If
			Next
		End If
		
		' check cc list
		list = SplitListToArray(page.Email.CcAddressList)
		If IsArray(list) Then
			For i = 0 To UBound(list)
				If Len(Trim(list(i))) > 0 Then
					If Not IsEmail(Trim(list(i))) Then
						AddCustomFrmError("CC address '" & list(i) & "' does not appear to be a valid email address. ")
						ValidMessage = False
					End If
				End If
			Next
		End If
			
		' check bcc list
		list = SplitListToArray(page.Email.BccAddressList)
		If IsArray(list) Then
			For i = 0 To UBound(list)
				If Len(Trim(list(i))) > 0 Then
					If Not IsEmail(Trim(list(i))) Then
						AddCustomFrmError("BCC address '" & list(i) & "' does not appear to be a valid email address. ")
						ValidMessage = False
					End If
				End If
			Next
		End If
	End If
	
End Function

Function LoadMessageFromRequest(page)
	page.AttachmentList = page.Uploader.Form.Item("AttachmentList")
	page.Email.GroupList = page.Uploader.Form.Item("GroupList")
	
	page.email.Subject = page.Uploader.Form.Item("Subject")
	page.email.Text = page.Uploader.Form.Item("Text")
	page.email.RecipientAddressList = page.Uploader.Form.Item("RecipientAddressList")
	page.email.CcAddressList = page.Uploader.Form.Item("CcAddressList")
	page.email.BccAddressList = page.Uploader.Form.Item("BccAddressList")
End Function

Function SplitListToArray(str)
	' this takes an address list and cleans it of spaces, semi-colons 
	' by converting to comma separated list. Then it splits to 
	' array on comma

	' replace all semi-colons with comma ..
	str = Replace(str, ";", ",")
	
	' replace all spaces with comma ..
	str = Replace(str, " ", ",")
	
	' replace doubled comma's with single ..
	Dim rv				: rv = 0
	Dim hasDoubles		: hasDoubles = True
	Do While hasDoubles
		str = Replace(str, ",,", ",")
		rv = InStr(str, ",,")
		If rv = 0 Then hasDoubles = False
	Loop
	
	' split to array ..
	SplitListToArray = Split(str, ",")
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case Else
			str = str & "Email"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Dim contactsButton
	pg.Action = ADD_RECIPIENTS
	href = pg.Url & pg.UrlParamsToString(True)
	contactsButton = "<li><a href=""" & href & """ class=""redirect-with-submit""><img class=""icon"" src=""/_images/icons/book_addresses.png"" /></a><a href=""" & href & """ class=""redirect-with-submit"">Contacts / Groups</a></li>"

	Dim sentMailButton
	pg.Action = ""
	href = "/email/sentmail.asp" & pg.UrlParamsToString(True)
	sentMailButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/email_date.png"" /></a><a href=""" & href & """>Sent Mail</a></li>"

	Select Case page.Action
		Case Else
			str = str & sentMailButton
			str = str & contactsButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_RemoveDupesFromStringList.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_group_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/smart_group_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_availability_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_group_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_displayer_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	Public AttachmentList
	
	' encrypted
	Public Action
	Public ProgramID
	Public EmailID
	Public GroupID
	
	' objects
	Public Member
	Public Client	
	Public Program	
	Public Email	
	Public Uploader
	
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
		If Len(GroupID) > 0 Then str = str & "emgid=" & Encrypt(GroupID) & amp
		
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
		c.AttachmentList = AttachmentList
		
		c.Action = Action
		c.ProgramID = ProgramID
		c.EmailID = EmailID
		c.GroupID = GroupID
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Email = Email
		Set c.Uploader = Uploader
		
		Set Clone = c
	End Function
End Class
%>

