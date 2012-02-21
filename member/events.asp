<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "availability"
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
	page.ProgramID = Decrypt(Request.QueryString("pid"))
	page.ScheduleID = Decrypt(Request.QueryString("scid"))
	page.EventId = Decrypt(Request.QueryString("eid"))
	page.Action = Decrypt(Request.QueryString("act"))
	page.EventAvailabilityID = Decrypt(Request.QueryString("eaid"))
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	If Request.Form("FormScheduleDropdownIsPostback") = IS_POSTBACK Then
		page.ScheduleID = Request.Form("ScheduleID")
	End If
	
	If Request.Form("FormProgramDropdownIsPostback") = IS_POSTBACK Then
		page.ProgramID = Request.Form("ProgramID")
		page.ScheduleID = ""
	End If
	
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then page.Program.Load()
	
	Set page.Schedule = New cSchedule
	page.Schedule.ScheduleID = page.ScheduleID
	If Len(page.Schedule.ScheduleID) > 0 Then page.Schedule.Load()
	
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
		<link rel="stylesheet" type="text/css" href="events.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=m_pageTitleText%></title>
		<style type="text/css">
			.message, #event-grid, #event-grid p.instructions	{width:622px;}
		</style>
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/jquery-ui-1.6.custom/js/jquery-ui-1.6.custom.min.js"></script>
		<script type="text/javascript" src="/_incs/script/jquery/plugins/dimensions/jquery.dimensions.min.js"></script>
		<script type="text/javascript" src="events.js"></script>
		<script type="text/javascript">
			var SET_MEMBER_TO_AVAILABLE				= <%=SET_MEMBER_TO_AVAILABLE %>
			var SET_MEMBER_TO_NOT_AVAILABLE			= <%=SET_MEMBER_TO_NOT_AVAILABLE %>
			var UPDATE_RECORD						= <%=UPDATE_RECORD %>
			var DELETE_RECORD						= <%=DELETE_RECORD %>
			var GET_AVAILABILITY_FORM				= <%=GET_AVAILABILITY_FORM %>
		</script>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage

	Call OnPageLoad(page)
	
	Select Case page.Action
		case EMAIL_SINGLE_EVENT_TO_MEMBER
			Call EmailEventToMember(page)
			page.MessageID = 6051: page.Action = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))

		Case DOWNLOAD_EVENT_AS_ICAL_FILE
			Call StreamICalToBrowser(page.EventId)
			Response.End

		Case Else
			str = str & EventGridToString(page.Client.NameClient, page.Member.MemberId, page)
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub EmailEventToMember(page)
	Dim str, i
	Dim pg			: Set pg = page.Clone()
	Dim dateTime	: Set dateTime = New cFormatDate
	Dim skills		: skills = ""
	Dim subject		: subject = ""

	Dim ev            : Set ev = New cEvent
	ev.EventID = pg.EventID
	ev.Load()
	Dim member        : Set member = New cMember
	member.MemberID = pg.Member.MemberID
	member.Load()
	Dim skillList     : skillList = GetMemberSkillsByEventID(page)
	Dim email         : Set email = New cEmailSender

	str = str & "Hello " & member.NameFirst & " " & member.NameLast
	str = str & vbCrLf & vbCrLf & "Find below the event information you requested."
	str = str & vbCrLf & vbCrLf & String(60, "-")
	str = str & vbCrLf & "Event: " & ev.EventName & " (" & ev.ProgramName & ")"
	str = str & vbCrLf & "Date: " & dateTime.Convert(ev.EventDate, "DDDD MMM dd, YYYY")
	' time
	str = str & vbCrLf & "Time: "
	if Len(ev.TimeStart) > 0 Then
	str = str & dateTime.Convert(ev.TimeStart, "hh:nn PP") & " - "
	if Len(ev.TimeEnd) > 0 Then
	str = str & dateTime.Convert(ev.TimeEnd, "hh:nn PP")
	Else
	str = str & "??"
	End If
	Else
	str = str & "No start time provided."
	End if
	' skill list and hasFiles
	if IsArray(skillList) Then
	str = str & vbCrLf 
	for i = 0 To UBound(skillList,2)
	skills = skills & skillList(0,i) & ", "
	next
	if Len(skills) > 0 Then skills = Left(skills, Len(skills) - 2)
	str = str & vbCrLf & "Scheduled For: " & skills

	if ev.HasFiles Then
	str = str & vbCrLf & "Files: This event has files for download."
	End If
	End If
	if Len(ev.EventNote) > 0 Then 
	str = str & vbCrLf & vbCrLf
	str = str & "Notes for This Event: " & ev.EventNote
	End If
	str = str & vbCrLf & String(60, "-")
	str = str & EmailDisclaimerToString(member.ClientName)

	subject = "** [" & Application.Value("APPLICATION_NAME") & "] " & member.ClientName & " Event Info for " & ev.ProgramName & " **"

	Call email.SendMessage(member.Email, member.Email, subject, str)

	Set email = Nothing
	Set dateTime = Nothing
	Set ev = Nothing
End Sub

Function GetMemberSkillsByEventID(p)
   Dim cnn           : Set cnn = Server.CreateObject("ADODB.Connection")
   Dim rs            : Set rs = Server.CreateObject("ADODB.Recordset")

   cnn.Open Application.Value("CNN_STR")
   cnn.up_eventGetMemberSkillsForEvent CLng(p.EventID), CLng(p.Member.MemberID), rs
   If Not rs.EOF Then GetMemberSkillsByEventID = rs.GetRows()

   if rs.State = adStateOpen Then rs.Close(): Set rs = Nothing
   Set cnn = nothing
End Function

Sub StreamICalToBrowser(eventId)
	Dim str, i
	
	Dim evnt			: Set evnt = New cEvent
	evnt.EventID = eventID
	evnt.Load()
	Dim memberList		: memberList = evnt.XmlScheduleList()
	Dim description		: description = ""
	Dim team			: team = ""
	Dim path			: path = Application.Value("CREATE_PDF_FILE_DIRECTORY") & CleanFileName(evnt.EventName, "_") & ".ics"
	Dim fso				: Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim stream			: Set stream = Server.CreateObject("ADODB.Stream")
	Dim file			: Set file = fso.CreateTextFile(path, 2)
	Dim dateTime		: Set dateTime = New cFormatDate
	Dim xmlDoc			: Set xmlDoc = Server.CreateObject("Microsoft.XMLDom")
	
	' HACK:
	' -------------
	' - set events without a start/end time to midnight, otherwise outlook converts date
	' to the day before. There doesn't seem to be a way to set an event as 'All Day' in outlook
	' using ics formatting (there should be ..)
	
	' format .ics file
	file.Write("BEGIN:VCALENDAR" & vbCrLf)
	file.Write("VERSION:1.0" & vbCrLf)
	file.Write("PRODID:WORSHIPSCHEDULE v1.0" & vbCrLf)
	file.Write("BEGIN:VEVENT" & vbCrLf)
	' start time
	file.Write("DTSTART:" & dateTime.Convert(evnt.EventDate, "YYYYMMDD"))
	If Len(evnt.TimeStart) > 0 Then
		file.Write("T" & dateTime.Convert(evnt.TimeStart, "HHHnn") & "00")
	Else
		file.Write("T000000")
	End If
	file.Write(vbCrLf)
	' end time
	If Len(evnt.TimeEnd) > 0 Then
		file.Write("DTEND:" & dateTime.Convert(evnt.EventDate, "YYYYMMDD") & "T" & dateTime.Convert(evnt.TimeEnd, "HHHnn") & "00" & vbCrLf)
	End If
	' event name
	file.Write("SUMMARY:" & evnt.EventName & vbCrLf)
	' notes/description
	If Len(evnt.EventNote) > 0 Then
		description = description & "Notes: " & evnt.EventNote
	End If
	' event team (included in description field)
	If IsArray(memberList) Then
		For i = 0 To UBound(memberList,2)
			If Len(memberList(2,i)) > 0 Then
				team = team & memberList(0,i) & ": "
				team = team & XmlFragmentToList(memberList(2,i), ", ", xmlDoc) & "\n"
			End If
		Next
		If Len(team) > 0 Then
			If Len(description) > 0 Then
				description = description & "\n" & "\n"
			End If
			description = description & "Event Team" & "\n" & String(60, "-") & "\n"
			description = description & team
		End If
	End If
	If Len(description) > 0 Then
		file.Write("DESCRIPTION:" & description & vbCrLf)
	End If
	file.Write("END:VEVENT" & vbCrLf)
	file.Write("END:VCALENDAR" & vbCrLf)
	file.Close
	Set file = Nothing
	
	Set file = fso.GetFile(path)
	
	' stream it to the browser
	Response.Clear
	Response.AddHeader "Content-Disposition", "attachment; filename=" & file.Name
	Response.AddHeader "Content-Length", file.Size
	Response.ContentType = "application/octet-stream"
	stream.Open
	stream.Type = 1
	Response.CharSet = "UTF-8"
	stream.LoadFromFile(path)
	Response.BinaryWrite(stream.Read)
	stream.Close
	
	fso.DeleteFile(path)

	Set fso = Nothing
	Set stream = Nothing
	Set evnt = Nothing
	Set dateTime = Nothing
	Set xmlDoc = Nothing
End Sub

Function EventGridToString(clientName, memberId, page)
	Dim str, i
	Dim pg					: set pg = page.Clone()
	
	Dim eventAvailability	: Set eventAvailability = New cEventAvailability
	eventAvailability.MemberID = memberID
	
	Dim list				: list = eventAvailability.AvailabilityList(page.Program.ProgramId, page.Schedule.ScheduleId, "EventDate,TimeStart")
	Dim count				: count = 0
	
	Dim availableStatusClass
	Dim availableButtonClass
	Dim notAvailableButtonClass
	Dim availableLinkClass
	Dim notAvailableLinkClass
	Dim buttonExclamation
	Dim description
	Dim isScheduled
	Dim isVisible
	
	Dim timeParts
	
	' 0-EventAvailabilityID 1-MemberId 2-MemberNote 3-IsAvailable 4-IsViewedByMember
	' 5-EventAvailabilityDateCreated 6-EventAvailabilityDateCreated 7-EventId 8-EventName
	' 9-EventDate 10-TimeStart 11-TimeEnd 12-EventDescription 13-ScheduleId 
	' 14-ScheduleName 15-ScheduleIsVisible 16-ProgramId 17-ProgramName
	' 18-ProgramIsEnabled 19-IsScheduled
	
	str = str & m_appMessageText
	str = str & "<h3>" & Server.HTMLEncode(clientName) & " events</h3>"
	If Len(page.Program.ProgramId) > 0 Then 
		str = str & "<h4 class=""first"">" & Server.HTMLEncode(page.Program.ProgramName)
		If Len(page.Schedule.ScheduleId) > 0 Then
			str = str & " | " & Server.HTMLEncode(page.Schedule.ScheduleName)
		End If	
		str = str & "</h4>"	
	Else
		str = str & "<h4 class=""first"">All programs</h4>"
	End If
	
	If IsArray(list) Then
		str = str & "<div id=""event-grid"">"
		str = str & "<p class=""instructions"">Use this listing to indicate your availability for " & SErver.HTMLEncode(clientName) & " events. "
		str = str & "If you leave a note with an event, whoever is creating the schedule will see that note when they are assigning members to the event team. </p>"
		
		For i = 0 To UBound(list,2)
			buttonExclamation = ""
			
			description = server.HTMLEncode(list(12,i) & "")		: If Len(description) = 0 Then description = "No description has been provided for this event. "
			isScheduled = False										: If list(19,i) = 1 Then isScheduled = True
			isVisible = False										: If list(15,i) = 1 Then isVisible = True
			
			availableButtonClass = "available"
			notAvailableButtonClass = "not-available"
			availableLinkClass = ""
			notAvailableLinkClass = ""
			
			If list(4,i) = 0 Then
				availableLinkClass = "unknown"
				notAvailableLinkClass = "unknown"
				
				availableStatusClass = " unknown"
				
				availableButtonClass = "unknown"
				notAvailableButtonClass = "unknown"

				buttonExclamation = "?"
			Else
				If list(3,i) = 0 Then 
					notAvailableLinkClass = "selected"
					availableStatusClass = " not-available"
				Else
					availableLinkClass = "selected"
					availableStatusClass = " available"
				End If
			End If
			
			If isVisible Then
				str = str & "<div class=""event-item"">"
				
				str = str & "<ul class=""hover-toolbar"" style=""display:none;"">"
				pg.EventID = list(7,i): pg.Action = EMAIL_SINGLE_EVENT_TO_MEMBER
				str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Send me details by email"">"
				str = str & "<img src=""/_images/icons/email_date.png"" alt="""" /></a></li>"
				pg.EventID = list(7,i): pg.Action = DOWNLOAD_EVENT_AS_ICAL_FILE
				str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Download as iCal"">"
				str = str & "<img src=""/_images/icons/disk.png"" alt="""" /></a></li></ul>"
				
				str = str & "<div class=""left"">"
				str = str & "<div class=""date" & availableStatusClass & """>"
				str = str & "<p class=""day"">" & Day(list(9,i)) & "</p>"
				str = str & "<p class=""month"">" & Left(MonthName(Month(list(9,i)), True), 3) & "</p>"
				str = str & "<p class=""year"">" & Year(list(9,i)) & "</p></div>"
				If Len(list(10,i)) > 0 Then
					timeParts = Split(TimeValue(list(10,i)), ":")
					str = str & "<p class=""time"">" & timeParts(0) & ":" & timeParts(1) & LCase(Right(TimeValue(list(10,i)), 2)) & "</p>"
				End If
				
				str = str & "</div>"
				
				str = str & "<h2>"
				str = str & server.HTMLEncode(list(8,i))
				str = str & "<span class=""program"">&nbsp;&nbsp;[" & Server.HTMLEncode(list(17,i)) & "]</span>"
				str = str & "</h2>"
				str = str & "<p class=""description"">" & description & "</p>"
				
				str = str & "<div class=""button-container"" id=""eaid-" & list(0,i) & """>"
				If isScheduled Then str = str & "<span class=""scheduled"">I'm scheduled!</span>"
				str = str & "<span class=""available""><a href=""#"" class=""" & availableLinkClass & """>Available" & buttonExclamation & "</a></span>"
				str = str & "<span class=""not-available""><a href=""#"" class=""" & notAvailableLinkClass & """>Not available" & buttonExclamation & "</a></span>"
				
				pg.EventId = list(7,i): pg.Action = SHOW_EVENT_DETAILS
				str = str & "<span class=""link details"">View <a href=""/member/schedules.asp" & pg.UrlParamsToString(True) & """>details</a></span>"
				str = str & "<span class=""link note"">Leave a <a href=""#"">note</a></span></div>"
				
				str = str & AvailabilityNoteToString(list(2,i), list(6,i))
				str = str & "</div>"
				
				count = count + 1
			End If
		Next 
		str = str & "</div>"
	End If
	
	If count = 0 Then
		str = NoEventsDialogToString(page)
	End If
		
	EventGridToString = str
End Function

Function NoEventsDialogToString(page)
	Dim dialog				: Set dialog = New cDialog
	
	dialog.Headline = "Ok, where is the list of events for " & html(page.Program.ProgramName) & "?"
	
	dialog.Text = dialog.Text & "<p>It looks like "
	If Len(page.Program.ProgramId) = 0 Then
		dialog.Text = dialog.Text & "none of the programs you belong to have any events set up, "
	Else
		If Len(page.Schedule.ScheduleId) = 0 Then
			dialog.Text = dialog.Text & " the program you selected (" & html(page.Program.ProgramName) & ") doesn't have any events set up, "
		Else
			dialog.Text = dialog.Text & " the schedule you selected (" & html(page.Schedule.ScheduleName) & ") for the " & html(page.Program.ProgramName) & " program doesn't have any events set up, "
		End If
	End If
	dialog.Text = dialog.Text & "all the events occur in the past, or perhaps the person who creates your schedules has set them to hidden. </p>"
	
	dialog.SubText = dialog.SubText & "<p>Once the programs you belong to have some events (or they are unhidden by your scheduler), "
	dialog.SubText = dialog.SubText & "this page is where you'll tell " & Application.Value("APPLICATION_NAME") & " which events you are or aren't available for. </p>"
	
	If Len(page.Program.ProgramId) > 0 Then
		dialog.LinkList = dialog.LinkList & "<li><a href=""/member/events.asp"">Show all program events</a></li>"
	End If
	dialog.LinkList = dialog.LinkList & "<li><a href=""/member/contacts.asp"">Email an account administrator</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/member/programs.asp"">Back to my program list</a></li>"
	
	NoEventsDialogToString = dialog.ToString()
End Function

Function ScheduleDropdownToString(page)
	Dim str, i
	Dim pg			: Set pg = page.Clone()
	Dim list		: If Len(page.Program.ProgramID) > 0 then list = page.Program.ScheduleList()
	Dim isSelected	: isSelected = ""
	
	Dim defaultText		: defaultText = "< Select a schedule >"
	If Len(page.Schedule.ScheduleID) > 0 Then defaultText = "< Show all schedules >"
	
	If Not page.Program.HasSchedules Then Exit Function
	If Not IsArray(list) Then Exit Function
	
	' 0-ScheduleID 1-ScheduleName 5-IsVisible
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-schedule-dropdown"">"
	str = str & "<input type=""hidden"" name=""FormScheduleDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""ScheduleID"" id=""schedule-dropdown"">"
	str = str & "<option value="""">" & Html(defaultText) & "</option>"
	For i = 0 To UBound(list,2) 
		' schedule isVisible
		If list(5,i) = 1 Then
			isSelected = ""
			If CStr(list(0,i)) = CStr(page.Schedule.ScheduleID) Then isSelected = " selected=""selected"""
			str = str & "<option value=""" & list(0,i) & """" & isSelected & ">" & html(list(1,i)) & "</option>"
		End If
	Next
	str = str & "</select></form></li>"
	
	ScheduleDropdownToString = str
End Function

Function ProgramDropdownToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = page.Member.ProgramList()
	Dim isSelected		: isSelected = ""
	
	Dim defaultText		: defaultText = "< Select a program >"
	If Len(page.Program.ProgramID) > 0 Then defaultText = "< Show all programs >"
	
	' 0-ProgramID 1-ProgramName 5-EnrollStatusID 10-IsActive 18-ProgramIsEnabled
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-program-dropdown"">"
	str = str & "<input type=""hidden"" name=""FormProgramDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""ProgramID"" id=""program-dropdown"">"
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
	Dim dateTime	: Set dateTime = New cFormatDate
	Dim pg			: Set pg = page.Clone()
	
	Dim eventAvailability	: Set eventAvailability = New cEventAvailability
	eventAvailability.EventAvailabilityID = page.EventAvailabilityID
	If Len(eventAvailability.EventAvailabilityId) > 0 Then eventAvailability.Load()
	
	Dim programLink
	If Len(page.Program.ProgramId) > 0 Then
		pg.Action = SHOW_PROGRAM_DETAILS
		programLink = "<a href=""/member/programs.asp" & pg.UrlParamsToString(True) & """>" & html(page.Program.ProgramName) & "</a> / "
	End If
	
	Dim availabilityLink
	pg.Action = ""
	availabilityLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Availability</a> / "
	
	Dim eventLink	
	pg.Action = SHOW_EVENT_DETAILS: pg.EventId = eventAvailability.EventId
	eventLink = "<a href=""/member/schedules.asp" & pg.UrlParamsToString(True) & """>" & html(eventAvailability.EventName) & " (" & dateTime.Convert(eventAvailability.EventDate, "mm/dd/YYYY") & ")</a> / "
	
	str = str & "<a href=""/member/overview.asp"">Member Home</a> / "
	Select Case page.Action
		Case UPDATE_RECORD
			str = str & programLink & availabilityLink & eventLink & "Note"
		Case Else
			str = str & programLink & "Availability"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim href
	
	Dim unviewedEventsCheckbox
	unviewedEventsCheckbox = "<li><label>"
	unviewedEventsCheckbox = unviewedEventsCheckbox & "<input type=""checkbox"" class=""checkbox"" name=""show_unviewed"" "
	unviewedEventsCheckbox = unviewedEventsCheckbox & "id=""unviewed-event-switch-checkbox"" />Show new events!</label></li>"
	
	Select Case page.Action
		Case Else
			str = str & unviewedEventsCheckbox
			str = str & ScheduleDropdownToString(page)
			str = str & ProgramDropdownToString(page)
	End Select
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_AvailabilityNoteToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_CleanFileName.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_GetListFromXmlFragment.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_availability_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	Public ProgramID
	Public ScheduleID
	Public EventAvailabilityID
	Public EventID
	
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
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(EventID) > 0 Then str = str & "eid=" & Encrypt(EventID) & amp
		If Len(EventAvailabilityID) > 0 Then str = str & "eaid=" & Encrypt(EventAvailabilityID) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(ScheduleID) > 0 Then str = str & "scid=" & Encrypt(ScheduleID) & amp
		
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

		c.Action = Action
		c.ProgramID = ProgramID
		c.ScheduleID = ScheduleID
		c.EventID = EventID
		c.EventAvailabilityID = EventAvailabilityID
		c.MessageID = MessageID
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Schedule = Schedule
		
		Set Clone = c
	End Function
End Class
%>

