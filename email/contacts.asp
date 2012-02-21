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
	page.EmailID = Decrypt(Request.QueryString("emid"))
	
	page.member_email_list = Request.Form("member_email_list")
	page.recipient_type = Request.Form("recipient_type")

	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()

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
		<link rel="stylesheet" type="text/css" href="/_incs/script/jquery/plugins/treeview/jquery.treeview.css" />
		<link type="text/css" href="/_incs/script/jquery/jquery-ui-1.6.custom/development-bundle/themes/start/ui.all.css" rel="stylesheet" />	
		<link rel="stylesheet" type="text/css" href="contacts.css" />

		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/jquery-ui-1.6.custom/js/jquery-ui-1.6.custom.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/plugins/treeview/jquery.treeview.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/plugins/quicksearch/jquery.quicksearch.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/plugins/form/jquery.form.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/plugins/cookie/jquery.cookie.js"></script>
		<script type="text/javascript" src="contacts.js"></script>

		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<script type="text/javascript" language="javascript">
		
			// use server to tranlate constants to javascript here ..
			var SMART_GROUP_CUSTOM_GROUP					= <%=SMART_GROUP_CUSTOM_GROUP %>
			var SMART_GROUP_PROGRAM							= <%=SMART_GROUP_PROGRAM %>
			var SMART_GROUP_SKILL_UNGROUPED					= <%=SMART_GROUP_SKILL_UNGROUPED %>
			var SMART_GROUP_SKILL_GROUP						= <%=SMART_GROUP_SKILL_GROUP %>
			var SMART_GROUP_SKILL							= <%=SMART_GROUP_SKILL %>
			var SMART_GROUP_SCHEDULE_TEAM					= <%=SMART_GROUP_SCHEDULE_TEAM %>
			var SMART_GROUP_EVENT_TEAM						= <%=SMART_GROUP_EVENT_TEAM %>
			var SMART_GROUP_EVENT_AVAILABLE					= <%=SMART_GROUP_EVENT_AVAILABLE %>
			var SMART_GROUP_EVENT_NOT_AVAILABLE				= <%=SMART_GROUP_EVENT_NOT_AVAILABLE %>
			var SMART_GROUP_EVENT_AVAILABILITY_MISSING		= <%=SMART_GROUP_EVENT_AVAILABILITY_MISSING %>
			var SMART_GROUP_SCHEDULE_AVAILABILITY_MISSING	= <%=SMART_GROUP_SCHEDULE_AVAILABILITY_MISSING %>
			var SMART_GROUP_PROGRAM_AVAILABILITY_MISSING	= <%=SMART_GROUP_PROGRAM_AVAILABILITY_MISSING %>
			var SMART_GROUP_SKILL_AVAILABILITY_MISSING		= <%=SMART_GROUP_SKILL_AVAILABILITY_MISSING %>
			
			var TYPE_EMAIL_RECIPIENTS						= <%=TYPE_EMAIL_RECIPIENTS %>
			var TYPE_CC_RECIPIENTS							= <%=TYPE_CC_RECIPIENTS %>
			var TYPE_BCC_RECIPIENTS							= <%=TYPE_BCC_RECIPIENTS %>
			
			var DELETE_RECORD								= <%=DELETE_RECORD %>
			
			var INSERT_EMAIL_GROUP_MEMBERS					= <%=INSERT_EMAIL_GROUP_MEMBERS %>
			var DELETE_EMAIL_GROUP_MEMBERS					= <%=DELETE_EMAIL_GROUP_MEMBERS %>
			
		</script>
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
			Call DoUpdateRecipients(page.Email, page.member_email_list, page.recipient_type, rv)
			page.Action = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
			
		Case Else
			str = str & ContactGridToString(page)
			
			' todo: get this into function or create via jquery ..
			str = str & "<div title=""Add a group"" id=""add-group-dialog""></div>"
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoUpdateRecipients(email, idString, recipientType, outError)
	Dim str
	outError = 0
	
	If Len(idString) = 0 Then Exit Sub
	
	str = idString
	
	' add the new recipients to the end of the list and remove duplicates ..
	Select Case recipientType
		Case TYPE_EMAIL_RECIPIENTS
			If Len(email.RecipientAddressList) > 0 Then
				str = email.RecipientAddressList & "," & str
			End If
			str = RemoveDupesFromStringList(str)
			email.RecipientAddressList = str

		Case TYPE_CC_RECIPIENTS
			If Len(email.CcAddressList) > 0 Then
				str = email.CcAddressList & "," & str
			End If
			str = RemoveDupesFromStringList(str)
			email.CcAddressList = str

		Case TYPE_BCC_RECIPIENTS
			If Len(email.BccAddressList) > 0 Then
				str = email.BccAddressList & "," & str
			End If
			str = RemoveDupesFromStringList(str)
			email.BccAddressList = str

		Case Else
			Call Err.Raise(vbObjectError + 1, "DoUpdateRecipients()", "Else clause reached when checking submit value from form")
	End Select
	
	Call email.Save(outError)
End Sub

Function ContactGridToString(page)
	Dim str
	
	str = str & "<div id=""contact-grid"">"
	str = str & "<table>"
	str = str & "<tbody>"
	str = str & "<tr><td class=""toolbar"" colspan=""3"">" & ToolBarToString() & "</td></tr>"
	str = str & "<tr><td>" & AccountTreeToString(page) & "</td>"
	str = str & "<td>" & ContactPaneToString(page) & "</td>"
	str = str & "<td>" & ActionPaneToString(page) & "</td></tr>"
	str = str & "</tbody></table></div>"
	
	ContactGridToString = str
End Function

Function ToolbarToString()
	Dim str
	
	str = str & "<input type=""button"" name=""add_group_button"" value=""Add group"" id=""add-group-button"" class=""button"" />&nbsp;"
	
	ToolbarToString = str
End Function

Function ContactPaneToString(page)
	Dim str, i

	Dim members			: members = page.Client.MemberList("", "")
	
	Dim tree			: tree = ""
	
	str = str & "<div id=""contact-pane"">"
	str = str & "<div class=""header"">Select: <a href=""#"" id=""check-all-button"">All</a>, <a href=""#"" id=""clear-all-button"">None</a></div>"

	' 0-MemberID 1-NameLast 2-NameFirst 3-NameLogin 4-PWord 5-Email 6-DOB 7-Gender
	' 8-AddressLine1 9-AddressLine2 10-City 11-StateID 12-StateCode 13-PostalCode
	' 14-PhoneHome 15-PhoneMobile 16-PhoneAlternate 17-IsProfileComplete 
	' 18-IsProfileUserCertified 19-IsApproved 20-ActiveStatus 21-LastLogin 22-SecretQuestion
	' 23-SecretAnswer 24-DateCreated 25-DateModified 26-IsAdmin 27-MemberProgramListXML
	
	Dim isEnabled
	
	page.Action = ADD_RECIPIENTS
	str = str & "<form method=""post"" action=""" & page.Url & page.UrlParamsToString(True) & """ id=""form-add-recipients"">"
	str = str & "<input type=""hidden"" id=""recipient-type"" name=""recipient_type"" value="""" />"
	str = str & "<input type=""hidden"" name=""form_add_recipients_is_postback"" value=""" & IS_POSTBACK & """ />"
	For i = 0 To UBound(members,2)
		isEnabled = True				: If members(20,i) = 0 Then isEnabled = False
		
		If isEnabled Then
			tree = tree & "<li id=""mid-" & members(0,i) & """ class=""member-item"">"
			tree = tree & "<input type=""checkbox"" id=""chx-mid-" & members(0,i) & """ name=""member_email_list"" value=""" & members(5,i) & """ />"
			tree = tree & "<label for=""chx-mid-" & members(0,i) & """>" & html(members(1,i) & ", " & members(2,i)) & "</label></li>"
		End If
	Next
	If Len(tree) > 0 Then tree = "<ul class=""listing"">" & tree & "</ul>"
	
	str = str & tree & "</form></div>"
	
	ContactPaneToString = str
End Function

Function ActionPaneToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim helpText
	
	helpText = helpText & "<p><strong>" & html(page.Client.NameClient) & " members </strong>"
	helpText = helpText & "is where you can find the email addresses for any of your account members. "
	helpText = helpText & Application.Value("APPLICATION_NAME") & " organizes your members into smart groups to make it easy for your to email just the members you need to contact. </p>"
	helpText = helpText & "<p>You can also create your own groups of members to easily email many people at once. </p>"
	helpText = helpText & "<p><strong>Missing someone? </strong>"
	helpText = helpText & "<br />Members for disabled programs and skills are not displayed on this page "
	helpText = helpText & "(also any disabled member and program accounts are hidden). </p>"
	
	str = str & "<div id=""action-pane"">"
	str = str & "<div class=""header"">"
	str = str & "<input type=""button"" id=""delete-group-button"" value=""Delete group"" style=""float:right;"" class=""button"" />"
	str = str & "<input type=""button"" id=""edit-group-button"" value=""Edit"" class=""button"" />"
	str = str & "&nbsp;" & EmailGroupMemberDropdownToString(page.Member.MemberId) 
	str = str & "</div>"

	' div to report members filtered, selected ..
	str = str & "<div id=""notifier"">"
	str = str & "<h4 class=""header""></h4>"
	str = str & "<p class=""details""></p>"
	str = str & "<p id=""recipient-buttons""><a href=""#"" id=""email-recipient-button"">Email</a>, <a href=""#"" id=""cc-recipient-button"">CC</a>, <a href=""#"" id=""bcc-recipient-button"">BCC</a></p>"
	str = str & "<div class=""help"">" & helpText & "</div>"
	
	str = str & "</div></div>"
	
	ActionPaneToString = str
End Function

Function EmailGroupMemberDropdownToString(memberId)
	Dim str
	
	str = str & "<select id=""group-member-dropdown"">"
	str = str & EmailGroupMemberDropdownOptionsToString(memberId)
	str = str & "</select>"
	
	EmailGroupMemberDropdownToString = str
End Function

Function AccountTreeToString(page)
	Dim str, i
	
	Dim programs		: programs = page.Member.OwnedProgramsList()
	Dim program			: Set program = New cProgram
	
	Dim skills
	
	Dim skillGroups
	Dim skillGroup		: Set skillGroup = New cSkillGroup
	
	Dim isEnabled

	str = str & "<div id=""group-pane"">" 
	str = str & "<div class=""header"" id=""tree-control"">Members: <a href=""#"">Collapse</a>, <a href=""#"">Expand</a></div>"

	' root level ..
	str = str & "<div id=""tree-view"">"
	str = str & "<ul class=""filetree"" id=""root"">"
	
	str = str & "<li class=""root-node"" id=""root-node"" title=""" & html(page.Client.NameClient) & " members"">"
	str = str & "<span><a href=""#"">" & html(page.Client.NameClient) & " members</a></span></li>"
	
	' custom groups ..
	str = str & CustomEmailGroupItemsToString(page.Member.MemberId, "")
	
	' 0-ProgramId 1-ProgramName 2-IsEnabled 3-ScheduleCount 4-EventCount

	If IsArray(programs) Then
		For i = 0 To UBound(programs,2)
			isEnabled = True		: If programs(2,i) = 0 Then isEnabled = False
			
			If isEnabled Then
				str = str & "<li title=""" & html(programs(1,i)) & """ class=""program-node"">" 
				str = str & "<span><a href=""#"" id=""pid-" & programs(0,i) & """ class=""program-node""><strong>" & Server.HTMLEncode(programs(1,i)) & "</strong></a></span>"
				str = str & ProgramSubtreeToString(programs(0,i))
				str = str & "</li>"
			End If
		Next		
	End If	
	str = str & "</ul></div></div>"

	AccountTreeToString = str
End Function

Function ProgramSubtreeToString(programId)
	Dim str, i
	
	Dim program			: Set program = New cProgram
	program.ProgramId = programId
	If Len(program.ProgramId) > 0 Then Call program.Load()
	Dim schedule		: Set schedule = New cSchedule
	
	Dim skillGroups		: skillGroups = program.SkillGroupList()
	Dim skills			: skills = program.SkillList("")
	Dim schedules		: schedules = program.ScheduleList()
	
	Dim ungroupedSkillNode		: ungroupedSkillNode = ""
	
	Dim isGroupEnabled
	
	' available for program node for tree
	str = str & "<li class=""program-missing-availability-node"" title=""Missing event availability information for " & html(program.ProgramName) & """>"
	str = str & "<span><a href=""#"" id=""missingavailabilityinfoforprogrampid-" & programId & """>Missing availability</a></span></li>" 

	' skill items for tree ..
	ungroupedSkillNode = SkillListItemsToString(skills, "")
	If Len(ungroupedSkillNode) > 0 Then
		str = str & "<li class=""skill-group-node"" title=""Ungrouped skills for " & html(program.ProgramName) & """>"
		str = str & "<span><a href=""#"" id=""ungroupedskillpid-" & programId & """>(ungrouped skills)</a></span>" 
		str = str & ungroupedSkillNode & "</li>"
	End If

	' 0-SkillGroupID 1-GroupName 2-GroupDesc 3-IsEnabled 4-AllowMultiple 5-DateModified 6-DateCreated

	If IsArray(skillGroups) Then
		For i = 0 To UBound(skillGroups,2)
			isGroupEnabled = True		: If skillGroups(3,i) = 0 Then isGroupEnabled = False
			
			If isGroupEnabled Then
				str = str & "<li class=""skill-group-node"" title=""" & html(skillGroups(1,i)) & """>" 
				str = str & "<span><a href=""#"" id=""skgid-" & skillGroups(0,i) & """>" & Server.HTMLEncode(skillGroups(1,i)) & "</a></span>"
				str = str & SkillListItemsToString(skills, skillGroups(0,i))
				str = str & "</li>"
			End If
		Next
	End If
	
	' schedules items for tree ..
	If IsArray(schedules) Then
		For i = 0 To UBound(schedules,2)
			str = str & "<li class=""schedule-node"" title=""" & html(schedules(1,i)) & " event team"">" 
			str = str & "<span><a href=""#"" id=""scid-" & schedules(0,i) & """>" & Server.HTMLEncode(schedules(1,i)) & "</a></span>"
			schedule.ScheduleId = schedules(0,i)
			str = str & EventListItemsToString(schedule.EventList(""), schedules(0,i))
			str = str & "</li>"
		Next
	End If
	If Len(str) > 0 Then str = "<ul>" & str & "</ul>"
	
	ProgramSubtreeToString = str
End Function

Function EventListItemsToString(events, scheduleId)
	Dim str, i
	
	Dim schedule			: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	If Len(schedule.ScheduleId) > 0 Then Call schedule.Load()
	
	Dim count				: count = 0
	
	' this are children of schedule node ..
	str = str & "<li class=""schedule-missing-availability-node"" title=""Members missing availability info for " & html(schedule.ScheduleName) & """>" 
	str = str & "<span><a href=""#"" id=""schedulemissingavailabilityscid-" & scheduleId & """>Missing availability</a></span></li>"

	If IsArray(events) Then
		For i = 0 To UBound(events,2)
		
			' hack: check for non-null eventId as this sproc returns a null event row if there are no events returned for the schedule ..
			If Len(events(0,i) & "") > 0 Then
				str = str & "<li class=""event-node"" title=""Event team for " & html(events(1,i)) & """>" 
				str = str & "<span><a href=""#"" id=""eid-" & events(0,i) & """>" & Server.HTMLEncode(events(1,i) & "") & "</a><span class=""on-event-date"">on " & Month(events(2,i)) & "-" & Day(events(2,i)) & "-" & Year(events(2,i)) & "</span></span>"
				
				' these are children of event node ..
				str = str & "<ul><li class=""available-for-event-node"" title=""Members available for " & html(events(1,i)) & """>" 
				str = str & "<span><a href=""#"" id=""availableforeventeid-" & events(0,i) & """>Available</a></span></li>"
				str = str & "<li class=""not-available-for-event-node"" title=""Members not available for " & html(events(1,i)) & """>" 
				str = str & "<span><a href=""#"" id=""notavailableforeventeid-" & events(0,i) & """>Not available</a></span></li>"
				str = str & "<li class=""missing-availability-for-event-node"" title=""Members missing availability info for " & html(events(1,i)) & """>" 
				str = str & "<span><a href=""#"" id=""missingavailabilityinfoforeventeid-" & events(0,i) & """>Missing availability</a></span></li></ul>"
				
				str = str & "</li>"
			End If
		Next
	End If
	
	If Len(str) > 0 Then str = "<ul>" & str & "</ul>"
	
	EventListItemsToString = str
End Function

Function SkillListItemsToString(skills, skillGroupId)
	Dim str, i, j
	
	If Not IsArray(skills) Then Exit Function

	' 0-SkillID 1-SkillName 2-SkillDesc 3-SkillIsEnabled 4-SkillGroupID 5-GroupName
	' 6-GroupDesc 7-GroupIsEnabled 
	
	Dim isSkillEnabled
	Dim isGroupEnabled
	Dim belongsToGroup

	For i = 0 To UBound(skills,2)
		isSkillEnabled = True			: If skills(3,i) = 0 Then isSkillEnabled = False
		isGroupEnabled = True			: If skills(7,i) = 0 Then isGroupEnabled = False
		belongsToGroup = False			: If CStr(skillGroupId & "") = CStr(skills(4,i) & "") Then belongsToGroup = True
		
		If belongsToGroup And isSkillEnabled And isGroupEnabled Then
			str = str & "<li class=""skill-node"" title=""" & html(skills(1,i)) & """>" 
			str = str & "<span><a href=""#"" id=""skid-" & skills(0,i) & """>" & Server.HTMLEncode(skills(1,i)) & "</a></span>"

			' this is child of skill node ..
			str = str & "<ul><li class=""missing-availability-for-skill-node"" title=""Members with " & html(skills(1,i)) & " missing availability"">" 
			str = str & "<span><a href=""#"" id=""missingavailabilityinfoforskillskid-" & skills(0,i) & """>Missing availability</a></span></li>"
			str = str & "</ul>"
			
			str = str & "</li>"
		End If
	Next
	If Len(str) > 0 Then str = "<ul>" & str & "</ul>"
	
	SkillListItemsToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()

	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case Else
			str = str & html(page.Client.NameClient) & " Contacts"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str, href
	Dim pg				: Set pg = page.Clone()
	
	Dim composeEmailButton
	pg.Action = ""
	href = "/email/email.asp" & pg.UrlParamsToString(True)
	composeEmailButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/email_edit.png"" alt="""" /></a><a href=""" & href & """>Compose</a></li>"

	Select Case page.Action
		Case Else
			str = str & composeEmailButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_RemoveDupesFromStringList.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_CustomEmailGroupItemsToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EmailGroupMemberDropdownOptionsToString.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_group_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_group_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID

	' encrypted
	Public Action
	Public EmailId
	
	' form post data
	Public member_email_list
	Public recipient_type
	
	' objects
	Public Member
	Public Client
	Public Email	
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(EmailId) > 0 Then str = str & "emid=" & Encrypt(EmailId) & amp

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
		c.EmailId = EmailId
		
		c.member_email_list = member_email_list
		c.recipient_type = recipient_type

		Set c.Member = Member
		Set c.Client = Client
		Set c.Email = Email

		Set Clone = c
	End Function
End Class
%>

