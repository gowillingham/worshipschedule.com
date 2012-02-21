<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-schedules"
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
	
	page.Month = Request.Querystring("m")
	page.Day = Request.Querystring("d")
	page.Year = Request.Querystring("y")
	
	page.Action = Decrypt(Request.QueryString("act"))
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.ScheduleID = Decrypt(Request.QueryString("scid"))
	page.FilterScheduleId = Decrypt(Request.QueryString("fscid"))
	
	page.include_text = Request.Form("include_text")
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()
	Set page.Schedule = New cSchedule
	page.Schedule.ScheduleID = page.ScheduleID
	If Len(page.Schedule.ScheduleID) > 0 Then Call page.Schedule.Load()

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
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>

		<script language="javascript" type="text/javascript">
			$(document).ready(function(){
				// wire up member list box ..
				$("#member-list", ".form").hide();
				$("#member-list-trigger", ".form").click(function(){
					$("#member-list", ".form").show();
					return false
				});
			});
		</script>

		<style type="text/css">
			.message, .form {width:622px;}
		</style>
		<title><%=m_pageTitleText%></title>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Dim count
	
	Call OnPageLoad(page)
	
	Select Case page.Action
		Case SEND_SCHEDULE_BY_EMAIL
			If Request.Form("form_send_schedule_by_email_is_postback") = IS_POSTBACK Then
				If ValidFormSendSchedule(page) Then
					Call DoSendScheduleToTeam(page, rv)
					page.MessageId = 6028: page.ScheduleId = "": page.Action = ""
					Response.Redirect("/schedule/schedules.asp" & page.UrlParamsToString(False))
				Else
					str = str & FormEmailScheduleToString(page, "")
				End If
			Else
				If Not page.Schedule.HasEvents Then
					page.MessageId = 6058: page.Action = "": page.ScheduleId = ""
					Response.Redirect("/schedule/schedules.asp" & page.UrlParamsToString(False))
				End If
			
				' show message or redirect if no published members returned ..
				str = str & FormEmailScheduleToString(page, count)
				If count = 0 Then
					page.MessageId = 6057: page.Action = "": page.ScheduleId = ""
					Response.Redirect("/schedule/schedules.asp" & page.UrlParamsToString(False))
				End If
			End If
			
		Case SEND_MESSAGE
			Call DoInsertEmail(page, rv)
			page.Action = "": page.ProgramId = "": page.ScheduleId = "": page.FilterScheduleId = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
			
		Case Else
			Call err.Raise(vbObjectError + 100, "Main()", "ASSERT: Unexpectedly reached else clause in switch statement. ")		
	
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoInsertEmail(page, outError)
	Dim email			: Set email = New cEMail
	
	email.MemberId = page.Member.MemberId
	email.ClientId = page.Member.ClientId
	email.GroupList = "schedule|" & page.ScheduleId
	
	Call email.Add(outError)
	
	page.EmailId = email.EmailId
End Sub

Function EventFileListToString(fragment, xml)
	Dim str
	
	Dim node
	
	If Len(fragment) = 0 Then Exit Function
	
	xml.Async = False
	xml.LoadXml(fragment)
	
	For Each node in xml.DocumentElement.ChildNodes
		str = str & node.Attributes.GetNamedItem("FileName").Text & ", "
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 2)
	
	EventFileListToString = str
End Function

Function MemberSkillListToString(fragment, xml)
	Dim str
	
	Dim node
	
	Dim isSkillEnabled
	Dim isSkillGroupEnabled
	Dim isPublished
	
	xml.Async = False
	xml.LoadXml(fragment)
	
	For Each node In xml.DocumentElement.ChildNodes

		' check to see if this skill should be included ..
		isSkillEnabled = False
		If node.Attributes.GetNamedItem("IsSkillEnabled").Text = "1" Then isSkillEnabled = True
		isSkillGroupEnabled = False
		If node.Attributes.GetNamedItem("IsSkillGroupEnabled").Text = "1" Then isSkillGroupEnabled = True
		isPublished = True
		If node.Attributes.GetNamedItem("PublishStatus").Text = CStr(IS_MARKED_FOR_PUBLISH) Then isPublished = False
		
		' add it to skill listing ..
		If isSkillEnabled And isSkillGroupEnabled And isPublished Then
			str = str & node.Attributes.GetNamedItem("SkillName").Text & ", "
		End If
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 2)
	
	MemberSkillListToString = str
End Function

Sub SetMemberList(scheduleId, htmlFragment, count)
	Dim str, i, j
	
	htmlFragment = ""
	count = 0
	
	Dim node
	Dim xml					: Set xml = Server.CreateObject("MSXML2.DOMDocument")
	xml.Async = False
	
	Dim schedule			: Set schedule = New cSchedule
	schedule.ScheduleId = scheduleId
	
	Dim isSkillEnabled
	Dim isSkillGroupEnabled
	Dim isPublished
	
	Dim list				: list = schedule.EventTeamDetailsList()
	Dim idList
	Dim idArray
	
	If Not IsArray(list) Then Exit Sub
	
	For i = 0 To UBound(list,2)
	
		' check member, programMember enabled ..
		If (list(4,i) = 1) And (list(5,i) = 1) Then
		
			xml.LoadXml(list(12,i)) 
			
			For Each node In xml.DocumentElement.ChildNodes
				isSkillEnabled = False
				If node.Attributes.GetNamedItem("IsSkillEnabled").Text = "1" Then isSkillEnabled = True
				isSkillGroupEnabled = False
				If node.Attributes.GetNamedItem("IsSkillGroupEnabled").Text = "1" Then isSkillGroupEnabled = True
				isPublished = False
				If node.Attributes.GetNamedItem("PublishStatus").Text = "0" Then isPublished = True
				
				If isSkillEnabled And isSkillGroupEnabled And isPublished Then
					idList = idList & list(0,i) & ","
				End If
			Next
		End If
	Next
	If Len(idList) > 0 Then 
		idList = Left(idList, Len(idList) - 1)
		idList = RemoveDupesFromStringList(idList)
		idArray = Split(idList, ",")
	End If
	
	If IsArray(idArray) Then
		i = 0
		Do While i <= UBound(idArray)
			For j = 0 To UBound(list,2)
				If CStr(idArray(i)) = CStr(list(0,j)) Then
					str = str & html(list(2,j)) & "&nbsp;" & html(list(1,j)) & "&nbsp;" & html("<" & list(3,j) & ">") & "<br />"
					Exit For
				End If
			Next
			i = i + 1
		Loop
		If Len(str) > 0 Then str = Left(str, Len(str) - 6)
	End If
	
	htmlFragment = str
	If IsArray(idArray) Then count = UBound(idArray) + 1
End Sub

Sub DoSendScheduleToTeam(page, outError)
	Dim i
	
	Dim emailSender		: Set emailSender = New cEmailSender
	Dim email			: Set email = New cEmail
	Dim dateTime		: Set dateTime = New cFormatDate
	Dim xml				: Set xml = Server.CreateObject("MSXML2.DOMDocument")
	
	Dim subject			: subject = ""
	Dim text			: text = ""
	Dim toAddress		: toAddress = ""
	Dim fromAddress		: fromAddress = page.Member.Email
	
	Dim eventText		: eventText = ""
	Dim skillListText	: skillListText = ""
	
	Dim list			: list = page.Schedule.EventTeamDetailsList()
	
	Dim currentMemberId
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	Dim sendNow			: sendNow = False

	' do not try to send if no events are returned ..
	If Not IsArray(list) Then Exit Sub
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberIsEnabled 5-ProgramMemberIsEnabled 
	' 6-EventId 7-EventName 8-EventDate 9-TimeStart 10-TimeEnd 11-EventNote 12-SkillListXmlFragment
	' 13-FileListXmlFragment
		
	' set current memberId to first returned ..
	currentMemberId = ""
	For i = 0 To UBound(list,2)
	
		' new member row so ..
		If CStr(currentMemberId) <> CStr(list(0,i)) Then
		
			' .. reset current member			
			currentMemberId = list(0,i)
			isMemberEnabled = False
			If list(4,i) = 1 Then isMemberEnabled = True
			isProgramMemberEnabled = False
			If list(5,i) = 1 Then isProgramMemberEnabled = True
			
			' .. set some header info
			subject = "[" & Application.Value("APPLICATION_NAME") & "] **Schedule Information for " & list(2,i) & " " & list(1,i) & "**"
			toAddress = list(3,i)
			
			text = ""
			eventText = ""
			
			text = text & "Hello " & list(2,i) & " " & list(1,i)
			text = text & vbCrLf & vbCrLf & "You are receiving this message from " & Application.Value("APPLICATION_NAME") & " on behalf of "
			text = text & page.Member.NameFirst & " " & page.Member.NameLast & ". "
			text = text & "To view your complete schedule, login to your " & Application.Value("APPLICATION_NAME") & " account at http://" & Request.ServerVariables("SERVER_NAME") & "/. "
			text = text & "Your event information for " & page.Schedule.ProgramName & " (" & page.Schedule.ScheduleName & ") follows. "
			If Len(page.include_text) > 0 Then
				text = text & vbCrLf & vbCrLf & "[Additional Info From " & page.Member.NameFirst & " " & page.Member.NameLast & "]: "
				text = text & vbCrLf & vbCrlf & page.include_text
			End If
			text = text & vbCrLf & vbCrLf & "Schedule/Event team information: "
		End If
		
		' append event info ..
		skillListText = MemberSkillListToString(list(12,i), xml)
		If Len(skillListText) > 0 Then
			eventText = eventText & vbCrLf & vbCrLf & String(60, "-")
			eventText = eventText & vbCrLf & list(7,i)
			eventText = eventText & " [" & dateTime.Convert(list(8,i), "DDD MMM dd, YYYY")
			If Len(list(9,i) & "") > 0 Then 
				eventText = eventText & " - " & dateTime.Convert(list(9,i), "hh:nn PP")
			End If
			eventText = eventText & "]"
			If Len(list(11,i) & "") > 0 Then
				eventText = eventText & vbCrLf & "Notes: " & list(11,i)
			End If
			eventText = eventText & vbCrLf & "Scheduled as: " & skillListText
			If Len(list(13,i) & "") > 0 Then
				eventText = eventText & vbCrLf & "Files for download: " & EventFileListToString(list(13,i), xml)
			End If
		End If
		
		' check if it is time to send the message 
		' if some events were returned ..
		sendNow = False
		If Len(eventText) > 0 Then
			If i = UBound(list,2) Then
				' next row is end of list ..
				sendNow = True
			ElseIf CStr(list(0, i+1)) <> CStr(currentMemberId) Then
				' next row starts new member ..
				sendNow = True
			End If
		End If
		
		If sendNow And isMemberEnabled And IsProgramMemberEnabled Then
			text = text & eventText & EmailDisclaimerToString(page.Member.ClientName)
			Call emailSender.SendMessage(toAddress, fromAddress, subject, text)
			
			' archive it ..
			email.MemberId = page.Member.MemberId
			email.ClientId = page.Member.ClientId
			email.RecipientAddressList = toAddress
			email.Subject = subject
			email.Text = text
			email.IsSent = 1
			email.DateSent = Now()
			Call email.Add("")
		End If
	Next
End Sub

Function ValidFormSendSchedule(page)
	ValidFormSendSchedule = True
	
	If Not ValidData(page.include_text, False, 0, 5000, "Additional text", "") Then ValidFormSendSchedule = False
End Function

Function FormEmailScheduleToString(page, memberCount)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim memberListFragment
	Call SetMemberList(page.Schedule.ScheduleId, memberListFragment, memberCount)
	
	Dim publishMessage
	
	If memberCount = 0 Then Exit Function
	
	If page.Schedule.PublishStatus = 1 Then
		publishMessage = publishMessage & "One or more event teams on this schedule have changes that you have not yet published. "
		publishMessage = publishMessage & "To make sure that this schedule is sent to the members that you expect, "
		publishMessage = publishMessage & "you should publish this schedule first. "
		publishMessage = CustomApplicationMessageToString("Unpublished changes!", publishMessage, "Error")
	End If
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3><p>"
	str = str & "This page sends a schedule summary to each event team member for this schedule. </p></div>"

	str = str & "<div class=""tip-box""><h3>I want to ..</h3><ul>"
	pg.Action = SEND_MESSAGE
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>"
	str = str & "Send a message without including schedule</a>.</li></ul></div>"
	
	' refresh pg ..
	Set pg = page.Clone()
	
	str = str & publishMessage
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-send-schedule-by-email"">"
	str = str & "<input type=""hidden"" name=""form_send_schedule_by_email_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table>"
	str = str & "<tr><td class=""label"">Recipients</td>"
	str = str & "<td><input type=""text"" class=""large disabled"" disabled=""disabled"" value=""Team members for " & html(page.Schedule.ScheduleName) & """ /></td></tr>"
	str = str & "<tr id=""member-list""><td>&nbsp;</td><td><div class=""large"">" 
	str = str & memberListFragment & "</div></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">An individualized copy of the event team schedule will be sent <br />by email "
	str = str & "to your members (" & memberCount & ") who belong to an event team <br />for this "
	str = str & "schedule. <a href=""#"" id=""member-list-trigger"">Show members</a>.</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"

	str = str & "<tr><td class=""label"">Additional&nbsp;text</td>"
	str = str & "<td><textarea name=""include_text"" class=""large"" style=""height:100px;""></textarea></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">Any message you leave here will be included with the schedule <br />you are sending to each team member. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"

	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""submit"" value=""Send"" />"
	str = str & "&nbsp;&nbsp;<a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Cancel</a></td></tr>"
	str = str & "</table></form></div>"
	
	FormEmailScheduleToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim scheduleLink
	pg.FilterScheduleId = page.Schedule.ScheduleId: pg.ProgramId = page.Schedule.ProgramId: pg.Action = "": pg.ScheduleId = ""
	scheduleLink = "<a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>" & html(page.Schedule.ScheduleName) & "</a> / "
	
	Dim programLink
	pg.ProgramId = page.Schedule.ProgramId: pg.Action = "": pg.ScheduleId = "": pg.FilterScheduleId = ""
	programLink = "<a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>" & html(page.Schedule.ProgramName) & "</a> / "

	Dim scheduleRootLink
	pg.Action = "": pg.ScheduleId = "": pg.ProgramId = "": pg.FilterScheduleId = ""
	scheduleRootLink = "<a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Schedules</a> / "
	
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case Else
			str = str & scheduleRootLink & programLink & scheduleLink & "Email schedule"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Dim masterTeamViewButton
	pg.Action = DISPLAY_MASTER_SCHEDULE
	href = "/schedule/schedules.asp" & pg.UrlParamsToString(True)
	masterTeamViewButton = "<li><a href=""" & href & """><img src=""/_images/icons/group.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Team View</a></li>"
	
	Dim schedulesButton
	pg.Action = "": pg.ScheduleId = ""
	href = "/schedule/schedules.asp" & pg.UrlParamsToString(True)
	schedulesButton = "<li><a href=""" & href & """><img src=""/_images/icons/calendar.png"" alt="""" class=""icon"" /></a><a href=""" & href & """>Schedules</a></li>"
	
	Select Case page.Action
		Case Else
			str = str & schedulesButton & masterTeamViewButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_RemoveDupesFromStringList.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->

<%
Class cPage
	' unencrypted
	Public MessageID
	Public Year
	Public Month
	Public Day
	
	' form data
	Public include_text
	
	' encrypted
	Public Action
	Public ProgramID
	Public ScheduleID
	Public FilterScheduleId
	Public EmailId

	' objects
	Public Member
	Public Client
	Public Program
	Public Schedule
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(Year) > 0 Then str = str & "y=" & Year & amp
		If Len(Month) > 0 Then str = str & "m=" & Month & amp
		If Len(Day) > 0 Then str = str & "d=" & Day & amp
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(ScheduleID) > 0 Then str = str & "scid=" & Encrypt(ScheduleID) & amp
		If Len(FilterScheduleId) > 0 Then str = str & "fscid=" & Encrypt(FilterScheduleId) & amp
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
		c.Year = Year
		c.Month = Month
		c.Day = Day

		c.Action = Action
		c.ProgramID = ProgramID
		c.ScheduleID = ScheduleID
		c.FilterScheduleId = FilterScheduleId
		c.EmailId = EmailId
				
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Schedule = Schedule
		
		Set Clone = c
	End Function
End Class
%>

