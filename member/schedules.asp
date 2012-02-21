<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Const SHOW_EVENT_LIST_FOR_DATE = "10"
Const SHOW_EVENT_LIST = "11"
Const SHOW_TEAM_GRID = "12"
Const OTHER_ACTION_FIRST_EVENT = "20"
Const OTHER_ACTION_LAST_EVENT ="21"
Const SHOW_ALL_EVENTS = "22"
Const SHOW_SCHEDULED_EVENTS = "23"

' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "calendar"
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
	
	page.Action = Decrypt(Request.QueryString("act"))
	page.ReturnAction = Decrypt(Request.QueryString("ract"))
	page.ProgramID = DeCrypt(Request.QueryString("pid"))
	page.EventID = DeCrypt(Request.QueryString("eid"))
	page.FileID = DeCrypt(Request.QueryString("fid"))
	page.ScheduleID = DeCrypt(Request.QueryString("scid"))
	
	page.MessageID = Request.QueryString("msgid")
	page.SortBy = Request.QueryString("sb")
	page.OtherAction = Request.QueryString("oa")
	page.ScheduledEvents = Request.QueryString("se")
	page.Day = Request.QueryString("d")
	page.Month = Request.QueryString("m")
	page.Year = Request.QueryString("y")
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	If Request.Form("FormProgramDropdownIsPostback") = IS_POSTBACK Then
		page.ProgramID = Request.Form("ProgramID")
		page.ScheduleID = ""
	End If
	Set page.Program = New cProgram
	page.Program.ProgramID = page.ProgramID
	If Len(page.Program.ProgramID) > 0 Then Call page.Program.Load()
	
	Set page.Evnt = New cEvent
	page.Evnt.EventID = page.EventID
	If Len(page.Evnt.EventID) > 0 Then Call page.Evnt.Load()
	
	Set page.File = New cFile
	page.File.FileID = page.FileID
	If Len(page.File.FileID) > 0 Then Call page.File.Load()
	
	If Request.Form("FormScheduleDropdownIsPostback") = IS_POSTBACK Then
		page.ScheduleID = Request.Form("ScheduleID")
	End If
	Set page.Schedule = New cSchedule
	page.Schedule.ScheduleID = page.ScheduleID
	If Len(page.Schedule.ScheduleID) > 0 Then Call page.Schedule.Load()
	
	If Request.Form("FormSortByDropdownIsPostback") = IS_POSTBACK Then
		page.SortBy = Request.Form("SortBy")
	End If
	
	If Request.Form("FormScheduledEventsCheckboxIsPostback") = IS_POSTBACK Then
		page.ScheduledEvents = Request.Form("ScheduledEvents")
	End If
	
	If Request.Form("FormOtherActionDropdownIsPostback") = IS_POSTBACK Then
		page.OtherAction = Request.Form("OtherAction")
	End If
	
	' load event array and set start date here as TabList view token should
	' be generated after setting calendar start date ..
	page.EventList = GetEventListForMember(page.Member.MemberID, page.ProgramID, page.ScheduleID, False, LookupSortParam(page.SortBy))
	Call SetCalendarStartDate(page)
	
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
		<link rel="stylesheet" type="text/css" href="/_incs/style/inside_member_schedule.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script language="javascript" type="text/javascript">
			$(document).ready(function(){
				// wire up my events checkbox ..
				$("#my-events").click(function(){
					$("#form-scheduled-events-checkbox").submit();
				});
			});
		</script>
		<style type="text/css">
			.message-width {width:644px;}
			.summary	{width:650px;}
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
		case EMAIL_SINGLE_EVENT_TO_MEMBER
			Call EmailEventToMember(page)
			page.MessageID = 6051: page.Action = page.ReturnAction: page.ReturnAction = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))

		Case DOWNLOAD_EVENT_AS_ICAL_FILE
			Call StreamICalToBrowser(page)
			Response.End

		Case SHOW_EVENT_DETAILS
			str = str & EventDetailsToString(page)
			
		Case SHOW_EVENT_LIST_FOR_DATE
			str = str & EventListToString(page)

		Case SHOW_EVENT_LIST
			str = str & EventListToString(page)
	
		Case SHOW_TEAM_GRID
			str = str & TeamGridToString(page)
						
		Case STREAM_FILE_TO_BROWSER
			Call DoStreamFileToBrowser(page.FileId, page.Member.MemberId, rv)
			Response.End
			
		Case Else
			str = str & CalendarToString(page)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoStreamFileToBrowser(fileId, memberId, outError)
	Dim file			: Set file = New cFile
	file.FileId = fileId
	
	Call file.StreamFile(memberId, outError)
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

Sub StreamICalToBrowser(p)
	Dim str, i
	
	Dim ev				: Set ev = New cEvent
	ev.EventID = p.EventID
	ev.Load()
	Dim memberList		: memberList = ev.XmlScheduleList()
	Dim description		: description = ""
	Dim team			: team = ""
	Dim path			: path = Application.Value("CREATE_PDF_FILE_DIRECTORY") & CleanFileName(ev.EventName, "_") & ".ics"
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
	file.Write("DTSTART:" & dateTime.Convert(ev.EventDate, "YYYYMMDD"))
	If Len(ev.TimeStart) > 0 Then
		file.Write("T" & dateTime.Convert(ev.TimeStart, "HHHnn") & "00")
	Else
		file.Write("T000000")
	End If
	file.Write(vbCrLf)
	' end time
	If Len(ev.TimeEnd) > 0 Then
		file.Write("DTEND:" & dateTime.Convert(ev.EventDate, "YYYYMMDD") & "T" & dateTime.Convert(ev.TimeEnd, "HHHnn") & "00" & vbCrLf)
	End If
	' event name
	file.Write("SUMMARY:" & ev.EventName & vbCrLf)
	' notes/description
	If Len(ev.EventNote) > 0 Then
		description = description & "Notes: " & ev.EventNote
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
	Set ev = Nothing
	Set dateTime = Nothing
	Set xmlDoc = Nothing
End Sub

Function FileListForSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim fileDisplay			: Set fileDisplay = New cFileDisplay
	
	Dim list				: list = page.Evnt.FileDetailsList()
	
	Dim fileName
	Dim count				: count = 0
	Dim isPublic			: isPublic = True
	Dim style				: style = ""
	
	Dim isScheduled			: isScheduled = False
	If IsArray(GetMemberSkillsByEventID(page)) Then isScheduled = True
	
	' 0-FileID 1-FileName 2-FriendlyName 3-Description 4-FileExtension 5-FileSize
	' 6-DownloadCount 7-IsPublic
	
	count = 0
	If IsArray(list) Then
	
		str = str & "<ul class=""file-list"">"
		For i = 0 To UBound(list,2)
			isPublic = True			: If list(7,i) = 0 Then isPublic = False
			
			If isScheduled Or isPublic Then
				style = " style=""background-image:url('" & fileDisplay.GetIconPath(list(4,i)) & "');"""
				fileName = list(2,i) & "." & list(4,i)
				
				pg.Action = STREAM_FILE_TO_BROWSER: pg.FileId = list(0,i)
				
				str = str & "<li" & style & ">"
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(fileName) & "</a>"
				str = str & "</li>"

				count = count + 1
			End If
			
		Next
		str = str & "</ul>"
		
	End If
	
	If count = 0 Then
		str = "<p class=""alert"">No files are linked to this event. </p>"
	End If
		
	FileListForSummaryToString = str
End Function

Function EventTeamGridForSummaryToString(page)
	Dim str, i
	
	Dim node
	Dim xml					: Set xml = Server.CreateObject("MSXML2.DOMDocument")
	xml.Async = False

	Dim list				: list = page.Evnt.EventTeamDetailsList()
	
	Dim count				: count = 0
	Dim rows				: rows = ""
	Dim alt					: alt = ""
	
	Dim skillStringList
	
	Dim isSkillEnabled		
	Dim isSkillGroupEnabled
	Dim isPublished
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-IsMemberEnabled 4-IsProgramMemberEnabled 5-SkillListingXmlFragment
	
	If IsArray(list) Then
	
		For i = 0 To UBound(list,2)
			isMemberEnabled = True			: If list(3,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True	: If list(4,i) = 0 Then isProgramMemberEnabled = False
			
			skillStringList = ""
			If isMemberEnabled And isProgramMemberEnabled Then
		
				xml.LoadXml(list(5,i))
				For Each node In xml.DocumentElement.ChildNodes
					isPublished = True
					If node.Attributes.GetNamedItem("PublishStatus").Text = CStr(IS_MARKED_FOR_PUBLISH) Then isPublished = False
					isSkillEnabled = True
					If node.Attributes.GetNamedItem("IsSkillEnabled").Text = "0" Then isSkillEnabled = False
					isSkillGroupEnabled = True
					If node.Attributes.GetNamedItem("IsSkillGroupEnabled").Text = "0" Then isSkillGroupEnabled = False
					
					If isPublished And isSkillEnabled And isSkillGroupEnabled Then
						skillStringList = skillStringList & node.Attributes.GetNamedItem("SkillName").Text & ", "
					End If
				Next
			End If
			If Len(skillStringList) > 0 Then 
				alt = ""		: If count Mod 2 = 0 Then alt = " class=""alt"""
				skillStringList = Left(skillStringList, Len(skillStringList) - 2)
				
				rows = rows & "<tr" & alt & "><td><img src=""/_images/icons/user.png"" class=""icon"" alt="""" />"
				rows = rows & "<strong>" & html(list(1,i) & ", " & list(2,i)) & "</strong></td>"
				rows = rows & "<td>" & html(skillStringList) & "</td></tr>"
				
				count = count + 1
			End If
		Next
	End If
	
	If count > 0 Then
		str = str & "<p>Members in this list belong to the event team for this event. </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Member Name</th><th>Skills</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = "<p class=""alert"">No members are assigned to this event team. </p>"
	End If
	
	EventTeamGridForSummaryToString = str
End Function

Function EventDetailsToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3>"
	pg.ReturnAction = pg.Action: pg.Action = EMAIL_SINGLE_EVENT_TO_MEMBER
	str = str & "<ul><li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Send me this event by email</a></li>"
	pg.Action = DOWNLOAD_EVENT_AS_ICAL_FILE
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Download as iCal</a></li>"
	str = str & "</ul></div>"
	
	str = str & m_appMessageText
	str = str & "<h3>" & html(page.Evnt.EventName) & "</h3>"
	str = str & "<h4 class=""first"">" & dateTime.Convert(page.Evnt.EventDate, "DDD MMMM dd, YYYY")
	If Len(page.Evnt.TimeStart & "") > 0 Then str = str & " at " & dateTime.Convert(page.Evnt.TimeStart, "hh:nn pp")
	str = str & "</h4>"
	
	str = str & "<div class=""summary"">"
		
	' notes/description
	str = str & "<h5 class=""description"">Description</h5>"
	If Len(page.Evnt.EventNote & "") > 0 Then
		str = str & "<p>" & html(page.Evnt.EventNote) & "</p>"
	Else
		str = str & "<p class=""alert"">No notes are included with this event. </p>"
	End If

	' files
	str = str & "<h5 class=""files"">Event files</h5>"
	str = str & FileListForSummaryToString(page)
	
	' event team 
	str = str & "<h5 class=""event-team"">Event Team</h5>"
	str = str & EventTeamGridForSummaryToString(page)
	
	' meta data
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul>"
	str = str & "<li>Created on " & dateTime.Convert(page.Evnt.DateCreated, "DDD MMMM dd, YYYY at hh:nn pp") & ". </li>"
	str = str & "<li>Last modified on " & dateTime.Convert(page.Evnt.DateModified, "DDD MMMM dd, YYYY at hh:nn pp") & ". </li>"
	str = str & "</ul>"
	
	
	str = str & "</div>"
	
	EventDetailsToString = str
End Function

Function TeamGridToString(page)
	Dim str, i, j, k
	
	' handle programs without skills ..
	If Not page.Program.HasSkills Then
		str = "The team view cannot be displayed for the program or schedule you selected. "
		str = str & "The program does not have any skills configured. "
		str = CustomApplicationMessageToString("The team view cannot be displayed!", str, "Error")
		TeamGridToString = str
		Exit Function
	End If
	
	' handle programs without events ..
	If Not page.Program.HasCurrentEvents Then
		str = "The team view cannot be displayed for the program or schedule you selected. "
		str = str & "The program does not have any current events. "
		str = CustomApplicationMessageToString("The team view cannot be displayed!", str, "Error")
		TeamGridToString = str
		Exit Function
	End If
	
	Dim rowCount			: rowCount = 0
	Dim eventCount			: eventCount = 0
	Dim columns				: columns = 4
	Dim previousCell		: previousCell = 0
	Dim previousGroupIdx	
	
	Dim scheduleList		: scheduleList = GetScheduleDetailsForTeamGrid(page.ProgramID, page.ScheduleID)
	Dim programSkillList	: programSkillList = page.Program.SkillList("")
	Dim eventIDList			: eventIDList = ArrayDimensionToList(scheduleList, 0)
	Dim skillList			: skillList = ArrayDimensionToList(programSkillList, 1)
	Dim groupList			: groupList = ArrayDimensionToList(programSkillList, 5)
	Dim currentGroup		: currentGroup = ""
	Dim isFirstSkill	
		
	' 0-eventID 1-eventName 2-eventDate 3-timeStart 4-skillGroupName 5-skillName
	' 6-nameLast 7-nameFirst 8-memberID 9-ProgramIsEnabled 10-ScheduleIsVisible

	str = str & "<div id=""team-grid"">"
	Do While previousCell <= UBound(eventIDList)
		str = str & "<table class=""grid"">"
		
		' header row ..
		For i = previousCell To previousCell + columns - 1
			
			If i = previousCell Then
				str = str & "<tr class=""header""><td class=""skill-header-column"">&nbsp;</td>"
			End If
			If i <= UBound(eventIDList) Then
				eventCount = eventCount + 1
				str = str & "<td>" & TeamGridHeaderToString(scheduleList, eventIDList(i)) & "</td>"
			Else
				str = str & "<td>&nbsp;</td>"
			End If
		Next
		str = str & "</tr>"
		
		rowCount = 0 
		For j = 0 To UBound(groupList)
		
			' put in spacer if not first group ..
			If (currentGroup <> groupList(j)) And (j <> 0) Then
				str = str & "<tr><td style=""border:none;height:2px;padding:0;background-color:#4284de;"" colspan=""10""></td></tr>"
			End If
			
			isFirstSkill = True
			For k = 0 To UBound(skillList)
				If IsSkillFromGroup(skillList(k), groupList(j), programSkillList) Then
					rowCount = rowCount + 1 
					
					If isFirstSkill Then
						str = str & "<tr class=""first-skill-row"">"
					ElseIf rowCount Mod 2 = 0 Then
						str = str & "<tr class=""alt"">"
					Else
						str = str & "<tr>"
					End If
					isFirstSkill = False
					 
					' vertical skill header column ..
					str = str & "<td class=""skill-header-column"">"
					str = str & "<strong>" & html(skillList(k)) & "</strong>"
					str = str & "</td>"
					
					For i = previousCell To previousCell + columns - 1
					
						' member list
						If i <= UBound(eventIDList) Then
							str = str & "<td>"
							str = str & MemberListForEventToString(page, scheduleList, eventIDList(i), skillList(k))
							str = str & "</td>"
						Else
							str = str & "<td>&nbsp;"
							str = str & "</td>"
						End If
					Next
					
					str = str & "</tr>"
				End If	
			Next
			
			currentGroup = groupList(j)
		Next
		
		previousCell = previousCell + columns
		str = str & "</table>"
	Loop
	str = str & "</div>"
	
	TeamGridToString = str
End Function

Function MemberListForEventToString(page, scheduleList, eventID, skillName)
	Dim str, i
	Dim name		: name = ""
	
	For i = 0 To UBound(scheduleList,2)
	
		' check for this event ..
		If CLng(eventID) = CLng(scheduleList(0,i)) Then
		
			' check for this skill
			If skillName = scheduleList(5,i) Then
				name = scheduleList(6,i) & ", " & scheduleList(7,i)
				If CLng(page.Member.MemberID) = CLng(scheduleList(8,i)) Then
					name = "<span class=""highlight"">" & html(name) & "</span>"
				End If
				str = str & "<li>" & name & "</li>"
			End If
			
		End If
	Next
	
	If Len(str) > 0 Then
		str = "<ul>" & str & "</ul>"
	Else
		str = "&nbsp;"
	End If
	
	MemberListForEventToString = str
End Function

Function TeamGridHeaderToString(scheduleList, eventID)
	Dim str, i
	Dim dateTime		: Set dateTime = New cFormatDate

	' 0-eventID 1-eventName 2-eventDate 3-timeStart 
	
	For i = 0 To UBound(scheduleList,2)
		If CStr(scheduleList(0,i)) = CStr(eventID) Then
			str = str & "<strong>" & html(scheduleList(1,i)) & "</strong>"
			str = str & "<br />" & dateTime.Convert(scheduleList(2,i), "DDD MMM dd, YYYY")
			Exit For
		End If
	Next
	
	TeamGridHeaderToString = str
End Function

Function IsSkillFromGroup(skillName, groupName, skillList)
   Dim i
   IsSkillFromGroup = False
   
   For i = 0 To UBound(skillList,2)
      If skillList(1,i) = skillName And skillList(5,i) = groupName Then
         IsSkillFromGroup = True
         Exit Function
      End If
   Next
End Function

Function CalendarToString(page)
	Dim str, i, rv
	Dim pg				: Set pg = page.Clone()
	Dim cal				: Set cal = New gtdCalendar
	Dim count			: count = 0 
	Dim displayEvent	: displayEvent = False
	Dim calendar		: calendar = ""
	Dim list			: list = page.EventList
	Dim highlightClass	: highlightClass = ""
	Dim availableIcon	: availableIcon = ""
	
	cal.DisplayTopNav = True
	cal.DisplayDatePicker = True
	cal.DisplayHeader = True
	
	' set qs parms for nav ..
	Call cal.AddNavUrlParams("pid", Encrypt(pg.ProgramID))
	Call cal.AddNavUrlParams("act", Encrypt(pg.Action))
	Call cal.AddNavUrlParams("eid", Encrypt(pg.EventID))
	Call cal.AddNavUrlParams("fid", Encrypt(pg.FileID))
	Call cal.AddNavUrlParams("scid", Encrypt(pg.ScheduleID))
	Call cal.AddNavUrlParams("se", pg.ScheduledEvents)
	
	' set qs parms for date links ..
	Call cal.AddDateUrlParams("act", Encrypt(SHOW_EVENT_LIST_FOR_DATE))
	Call cal.AddDateUrlParams("pid", Encrypt(pg.ProgramID))
	Call cal.AddDateUrlParams("eid", Encrypt(pg.EventID))
	Call cal.AddDateUrlParams("fid", Encrypt(pg.FileID))
	Call cal.AddDateUrlParams("scid", Encrypt(pg.ScheduleID))
	Call cal.AddDateUrlParams("se", pg.ScheduledEvents)
	
	' 0-EventID 1-EventName 2-EventNote 3-EventDate 4-TimeStart 5-TimeEnd 6-ClientID
	' 7-MemberID 8-ProgramName 9-ProgramID 10-ScheduleID 11-ScheduleName 12-ScheduleDesc
	' 13-ScheduleIsVisible 14-IsScheduled 15-SkillListXmlFrag 16-FileListXmlFrag 17-IsAvailable
	' 18-EventAvailabilityID 19-ProgramIsEnabled 20-ProgramMemberIsActive
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			
			' filter for scheduled events ..
			displayEvent = False
			If Len(page.ScheduledEvents) > 0 Then
				If list(14,i) = 1 Then
					displayEvent = True
				End If
			Else
				displayEvent = True
			End If
			
			str = ""
			highlightClass = ""
			If list(14,i) = 1 Then
				highlightClass = " class=""highlight"""
			End If
			availableIcon = "clock_add.png"
			If list(17,i) = 0 Then availableIcon = "clock_delete.png"
			
			If displayEvent Then
				count = count + 1
				str = str & "<div class=""gtdCalItem"">"
				pg.EventID = list(0,i): pg.Action = SHOW_EVENT_DETAILS
				str = str & "<a" & highlightClass & " href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Event Details"">"
				str = str & "<strong>" & html(list(8,i)) & "</strong>"
				str = str & "<br />" & html(list(1,i)) & "</a>"
				
				str = str & "<ul class=""toolbar"">"
				pg.EventID = list(0,i): pg.Action = SHOW_EVENT_DETAILS
				str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Event Details""><img class=""icon"" src=""/_images/icons/magnifier.png"" alt="""" /></a></li>"
				pg.EventID = list(0,i): pg.Action = DOWNLOAD_EVENT_AS_ICAL_FILE: pg.ReturnAction = page.Action
				str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""iCal File""><img class=""icon"" src=""/_images/icons/disk.png"" alt="""" /></a></li>"
				pg.EventID = list(0,i): pg.Action = EMAIL_SINGLE_EVENT_TO_MEMBER: pg.ReturnAction = page.Action
				str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email Event Info""><img class=""icon"" src=""/_images/icons/email_date.png"" alt="""" /></a></li>"
				str = str & "</ul></div>"
				Call cal.AddItem(list(3,i), str, "")
			End If
		Next
	End If
	
	Call cal.SetDate(page.Year, page.Month)

	calendar = calendar & m_appMessageText
	calendar = calendar & cal.ToString()
	
	CalendarToString = calendar
End Function

Function EventListToString(page)
	Dim str, i
	Dim dateTime		: Set dateTime = New cFormatDate
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = page.EventList
	
	Dim displayEvent	: displayEvent = True
	Dim includeThisDate	: includeThisDate = True
	
	Dim count			: count = 0
	Dim eventHref		: eventHref = ""
	Dim altClass		: altClass = ""
	Dim itemClass		: itemClass = ""
	Dim isChecked		: isChecked = ""
	Dim eventIcon		: eventIcon = ""
	Dim availableIcon	: availableIcon = ""
	
	str = str & "<div class=""message-width"">" & m_appMessageText & "</div>"
	
	' 0-EventID 1-EventName 2-EventNote 3-EventDate 4-TimeStart 5-TimeEnd 6-ClientID
	' 7-MemberID 8-ProgramName 9-ProgramID 10-ScheduleID 11-ScheduleName 12-ScheduleDesc
	' 13-ScheduleIsVisible 14-IsScheduled 15-SkillListXmlFrag 16-FileListXmlFrag 17-IsAvailable
	' 18-EventAvailabilityID 19-ProgramIsEnabled 20-ProgramMemberIsActive
	
	str = str & "<div class=""grid"">"
	str = str & "<table><tr><th scope=""col"" style=""width:1%;""><input type=""checkbox"" name=""master"" checked=""checked"" disabled=""disabled"" /></th>"
	str = str & "<th scope=""col"">Event</th><th scope=""col"">Schedule</th><th scope=""col""></th></tr>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)

			' filter for scheduled events ..
			displayEvent = False
			If Len(page.ScheduledEvents) > 0 Then
				If list(14,i) = 1 Then
					displayEvent = True
				End If
			Else
				displayEvent = True
			End If
			
			' filter for day number link clicked ..
			includeThisDate = True
			If page.Action = SHOW_EVENT_LIST_FOR_DATE Then
				If DateValue(list(3,i)) <> DateValue(page.Year & "-" & page.Month & "-" & page.Day) Then
					includeThisDate = False
				End If
			End If
			
			If displayEvent And includeThisDate Then
				count = count + 1
				altClass = ""
				If count Mod 2 <> 0 Then altClass = " class=""alt"""

				isChecked = ""
				eventIcon = "date.png"
				itemClass = " class=""data"""
				If list(14,i) = 1 Then
					isChecked = " checked=""checked"""
					itemClass = " class=""data highlight"""
					eventIcon = "date_check.png"
				End If
				availableIcon = "clock_add.png"
				If list(17,i) = 0 Then  availableIcon = "clock_delete.png"
				
				str = str & "<tr" & altClass & ">"
				str = str & "<td><input type=""checkbox"" disabled=""disabled""" & isChecked & " /></td>"
				str = str & "<td><img class=""icon"" src=""/_images/icons/" & eventIcon & """ alt="""" />"
				str = str & "<div" & itemClass & ">"
				pg.Action = SHOW_EVENT_DETAILS: pg.EventID = list(0,i)
				eventHref = pg.Url & pg.UrlParamsToString(True)
				str = str & "<strong>" & html(list(8,i)) & " | <a href=""" & eventHref & """ title=""Event Details"">" & html(list(1,i)) & "</a></strong>"
				str = str & "<br />" & dateTime.Convert(list(3,i), "DDDD MMMM dd, YYYY")
				If Len(list(4,i)) > 0 Then
					str = str & " at " & dateTime.Convert(list(4,i), "hh:nn PP")
				End If
				str = str & "</div></td>"
				str = str & "<td>" & html(list(11,i)) & "</td>"
				str = str & "<td class=""toolbar"" style=""text-align:right;"">"
				str = str & "<a href=""" & eventHref & """ title=""Event Details""><img src=""/_images/icons/magnifier.png"" alt="""" /></a>"
				pg.EventID = list(0,i): pg.Action = DOWNLOAD_EVENT_AS_ICAL_FILE: pg.ReturnAction = page.Action
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""iCal File""><img src=""/_images/icons/disk.png"" alt="""" /></a>"
				pg.EventID = list(0,i): pg.Action = EMAIL_SINGLE_EVENT_TO_MEMBER: pg.ReturnAction = page.Action
				str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email Event Info""><img src=""/_images/icons/email_date.png"" alt="""" /></a>"
				str = str & "</td></tr>"
			End If
		Next
	End If
	str = str & "</table></div>"
	
	' no events returned
	If count = 0 Then
		str = "No events were returned for the program, schedule, or calendar date you selected. "
		str = "<div class=""message-width"">" & CustomApplicationMessageToString("No events were returned!", str, "Error") & "</div>"
	End If
	
	EventListToString = str
End Function

Function GetEventListForMember(memberID, programID, scheduleID, hidePastEvents, sortColumn)
	Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs			: Set rs = Server.CreateObject("ADODB.Recordset")
	Dim cmd			: Set cmd = Server.CreateObject("ADODB.Command")
	Dim today		: if hidePastEvents Then today = Now()
	
	' 0-EventID 1-EventName 2-EventNote 3-EventDate 4-TimeStart 5-TimeEnd 6-ClientID
	' 7-MemberID 8-ProgramName 9-ProgramID 10-ScheduleID 11-ScheduleName 12-ScheduleDesc
	' 13-ScheduleIsVisible 14-IsScheduled 15-SkillListXmlFrag 16-FileListXmlFrag 17-IsAvailable
	' 18-EventAvailabilityID 19-ProgramIsEnabled 20-ProgramMemberIsActive
	
	cnn.Open Application.Value("CNN_STR")
	cnn.CursorLocation = adUseClient
	
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "dbo.up_memberGetEventList"
	cmd.ActiveConnection = cnn
	
	cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(memberID))
	If hidePastEvents Then
		cmd.Parameters.Append cmd.CreateParameter("@Today", adDate, adParamInput, 0, today)
	End If
	If Len(programID) > 0 Then
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, CLng(programID))
	End If
	If Len(scheduleID) > 0 Then
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleID", adBigInt, adParamInput, 0, CLng(scheduleID))
	End If
	
	Set rs = cmd.Execute
	rs.Sort = sortColumn
	If Not rs.EOF Then GetEventListForMember = rs.GetRows()
	
	Set cmd = Nothing
	Set rs = Nothing
	Set cnn = Nothing	
End Function

Function GetMemberSkillsByEventID(p)
   Dim cnn           : Set cnn = Server.CreateObject("ADODB.Connection")
   Dim rs            : Set rs = Server.CreateObject("ADODB.Recordset")

   cnn.Open Application.Value("CNN_STR")
   cnn.up_eventGetMemberSkillsForEvent CLng(p.EventID), CLng(p.Member.MemberID), rs
   If Not rs.EOF Then GetMemberSkillsByEventID = rs.GetRows()

   if rs.State = adStateOpen Then rs.Close(): Set rs = Nothing
   Set cnn = nothing
End Function

Function GetScheduleDetailsForTeamGrid(programID, scheduleID)
	Dim cnn           : Set cnn = Server.CreateObject("ADODB.Connection")
	Dim rs            : Set rs = Server.CreateObject("ADODB.Recordset")
	  
	' 0-eventID 1-eventName 2-eventDate 3-timeStart 4-skillGroupName 5-skillName
	' 6-nameLast 7-nameFirst 8-memberID 9-ProgramIsEnabled 10-ScheduleIsVisible
	
	cnn.Open Application.Value("CNN_STR")
	If Len(scheduleID) = 0 Then
		cnn.up_scheduleGetProgramMemberSkillListByProgramID CLng(programID), rs
	Else
		cnn.up_scheduleGetProgramMemberSkillListByProgramID CLng(programID), CLng(scheduleID), rs
	End If
	If Not rs.EOF Then GetScheduleDetailsForTeamGrid = rs.GetRows()

	If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
End Function

Sub SetCalendarStartDate(page)
	Dim i, thisIdx
	Dim list		: list = page.EventList
	Dim count		: count = 0
	Dim testThis	: testThis = False
	
	If Not IsArray(list) Then Exit Sub
	
	Select Case page.OtherAction
		Case OTHER_ACTION_FIRST_EVENT
			For i = 0 To UBound(list,2)
			
				testThis = False
				If Len(page.ScheduledEvents) > 0 Then
					If list(14,i) = 1 Then
						testThis = True
					End If
				Else
					testThis = True
				End If
				
				If testThis Then
					If count = 0 Then thisIdx = i
					
					If list(3,i) < list(3,thisIdx) Then
						thisIdx = i
					End If
					
					count = count + 1
				End If
			Next
			If count > 0 Then
				page.Month = Month(list(3,thisIdx))
				page.Year = Year(list(3,thisIdx))
			End If
			
		Case OTHER_ACTION_LAST_EVENT
			For i = 0 To UBound(list,2)
			
				testThis = False
				If Len(page.ScheduledEvents) > 0 Then
					If list(14,i) = 1 Then
						testThis = True
					End If
				Else
					testThis = True
				End If
				
				If testThis Then
					If count = 0 Then thisIdx = i
					
					If list(3,i) > list(3,thisIdx) Then
						thisIdx = i
					End If
					
					count = count + 1
				End If
			Next
			If count > 0 Then
				page.Month = Month(list(3,thisIdx))
				page.Year = Year(list(3,thisIdx))
			End If
			
		Case Else
			' do nothing
			
	End Select
End Sub

Function LookupSortParam(val)
	Dim str
	
	Select Case val
		Case SORT_BY_EVENTDATE
			str = "EventDate"
		Case SORT_BY_PROGRAM_NAME 
			str = "ProgramName, EventDate"
		Case SORT_BY_SCHEDULENAME
			str = "ScheduleName, EventDate"
		Case SORT_BY_IS_SCHEDULED
			str = "IsScheduled DESC, EventDate"
		Case SORT_BY_IS_AVAILABLE
			str = "IsAvailable DESC, EventDate"
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
	str = str & "<option value=""" & SORT_BY_EVENTDATE & """>Event Date</option>"
	str = str & "<option value=""" & SORT_BY_PROGRAM_NAME & """>Program</option>"
	str = str & "<option value=""" & SORT_BY_SCHEDULENAME & """>Schedule</option>"
	str = str & "<option value=""" & SORT_BY_IS_SCHEDULED & """>My Events</option>"
	str = str & "<option value=""" & SORT_BY_IS_AVAILABLE & """>My Availability</option>"
	str = str & "</select></form></li>"	
	
	SortByDropdownToString = str
End Function

Function ScheduledEventsCheckboxToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim isChecked	: isChecked = ""
	If Len(page.ScheduledEvents) > 0 Then isChecked = " checked=""checked"""
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-scheduled-events-checkbox"">"
	str = str & "<input type=""hidden"" name=""FormScheduledEventsCheckboxIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<label><input type=""checkbox"" name=""ScheduledEvents""" & isChecked & " class=""checkbox"" id=""my-events""/>"
	str = str & "My Events!</label>"
	str = str & "</form></li>"	
	
	ScheduledEventsCheckboxToString = str
End Function

Function OtherActionsDropdownToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()

	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formOtherActionDropdown"">"
	str = str & "<input type=""hidden"" name=""FormOtherActionDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""OtherAction"" onchange=""document.forms.formOtherActionDropdown.submit();"">"
	str = str & "<option value="""">" & html("< Other actions >") & "</option>"
	str = str & "<option value="""">" & html("--") & "</option>"
	str = str & "<option value=""" & OTHER_ACTION_FIRST_EVENT & """>" & html("< Move to first event >") & "</option>"
	str = str & "<option value=""" & OTHER_ACTION_LAST_EVENT & """>" & html("< Move to last event >") & "</option>"
	str = str & "</select></form></li>"	
	
	OtherActionsDropdownToString = str
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
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formScheduleDropdown"">"
	str = str & "<input type=""hidden"" name=""FormScheduleDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""ScheduleID"" onchange=""document.forms.formScheduleDropdown.submit();"">"
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
	
	str = str & "<li><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""formProgramDropdown"">"
	str = str & "<input type=""hidden"" name=""FormProgramDropdownIsPostback"" value=""" & IS_POSTBACK & """/>"
	str = str & "<select name=""ProgramID"" onchange=""document.forms.formProgramDropdown.submit();"">"
	If page.Action <> SHOW_TEAM_GRID Then
		str = str & "<option value="""">" & Html(defaultText) & "</option>"
	End If
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
	Dim pg
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim calendarLink
	Set pg = page.Clone()
	pg.ProgramID = "": pg.ScheduleID = "": pg.Action = ""
	calendarLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Calendar</a>"
	
	Dim scheduleLink
	Set pg = page.Clone()
	pg.ProgramId = pg.Evnt.ProgramId: pg.ScheduleId = pg.Evnt.ScheduleId: pg.Action = ""
	scheduleLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.Evnt.ScheduleName) & "</a>"
	
	Dim programLink
	Set pg = page.Clone()
	pg.ScheduleID = "": pg.Action = "": pg.ProgramId = pg.Evnt.ProgramId
	programLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>" & html(page.Evnt.ProgramName) & "</a>"
	
	Dim eventsLink
	Set pg = page.Clone()
	pg.ScheduleID = "": pg.ProgramID = ""
	eventsLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Events</a>"
	
	str = str & "<a href=""/member/overview.asp"">Member Home</a> / "
	Select Case page.Action
		Case SHOW_EVENT_DETAILS
			str = str & calendarLink & " / " & programLink & " / " & scheduleLink & " / " & html(page.Evnt.EventName) & " (" & dateTime.Convert(page.Evnt.EventDate, "mm/dd/YYYY") & ")"
						
		Case SHOW_TEAM_GRID
			str = str & calendarLink & " / "
			If Len(page.Schedule.ScheduleID) > 0 Then 
				str = str & programLink & " / "
				str = str & html(page.Schedule.ScheduleName) & " :: Team View"
			Else
				str = str & html(page.Program.ProgramName) & " :: Team View"
			End If

		Case SHOW_EVENT_LIST_FOR_DATE
			str = str & calendarLink & " / "
			str = str & dateTime.Convert(page.Month & "/" & page.Day & "/" & page.Year, "DDDD MMMM dd, YYYY")
			
		Case SHOW_EVENT_LIST
			str = str & calendarLink & " / "
			If Len(page.ProgramID) > 0 Then
				str = str & eventsLink & " / "
				If Len(page.Schedule.ScheduleID) > 0 Then 
					str = str & programLink & " / "
					str = str & html(page.Schedule.ScheduleName)
				Else
					str = str & html(page.Program.ProgramName)
				End If
			Else
				str = str & "Events"
			End If
			
		Case Else
			If Len(page.ProgramID) > 0 Then 
				str = str & calendarLink & " / "
				If Len(page.ScheduleID) > 0 Then
					str = str & programLink & " / "
					str = str & html(page.Schedule.ScheduleName)
				Else
					str = str & html(page.Program.ProgramName)
				End If
			Else
				str = str & "Calendar"
			End If
			
	End Select
	
	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str, href
	Dim pg
	
	Dim calendarButton
	Set pg = page.Clone()
	pg.Action = "": pg.SortBy = "": pg.Day = ""
	href = pg.Url & pg.UrlParamsToString(True)
	calendarButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/calendar.png"" alt="""" /></a><a href=""" & href & """>Calendar</a></li>"
	
	Dim eventListButton
	Set pg = page.Clone()
	pg.Action = SHOW_EVENT_LIST
	href = pg.Url & pg.UrlParamsToString(True)
	eventListButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/date.png"" alt="""" /></a><a href=""" & href & """>Event List</a></li>"
	
	Dim teamViewButton
	Set pg = page.Clone()
	pg.Action = SHOW_TEAM_GRID
	href = pg.Url & pg.UrlParamsToString(True)
	teamViewButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/group.png"" alt="""" /></a><a href=""" & href & """>Team View</a></li>"
	
	Dim availabilityButton
	Set pg = page.Clone()
	pg.Action = ""
	href = "/member/events.asp" & pg.UrlParamsToString(True)
	availabilityButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/clock.png"" alt="""" /></a><a href=""" & href & """>Availability</a></li>"

	Select Case page.Action
		Case SHOW_EVENT_DETAILS
			str = str & calendarButton & eventListButton & availabilityButton			
		Case SHOW_TEAM_GRID
			str = str & ScheduleDropdownToString(page)
			str = str & ProgramDropdownToString(page)
			str = str & eventListButton
			str = str & calendarButton
		Case SHOW_EVENT_LIST_FOR_DATE
			str = str & ScheduledEventsCheckboxToString(page)
			str = str & SortByDropdownToString(page)
			str = str & ScheduleDropdownToString(page)
			str = str & ProgramDropdownToString(page)
			If Len(page.ProgramID) > 0 Then str = str & teamViewButton
			str = str & calendarButton
		Case SHOW_EVENT_LIST
			str = str & ScheduledEventsCheckboxToString(page)
			str = str & SortByDropdownToString(page)
			str = str & ScheduleDropdownToString(page)
			str = str & ProgramDropdownToString(page)
			If Len(page.ProgramID) > 0 Then str = str & teamViewButton
			str = str & calendarButton
		Case Else
			str = str & ScheduledEventsCheckboxToString(page)
			str = str & OtherActionsDropdownToString(page)
			str = str & ScheduleDropdownToString(page)
			str = str & ProgramDropdownToString(page)
			If Len(page.ProgramID) > 0 Then str = str & teamViewButton
			str = str & eventListButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/gtdCalendar/gtdCal.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_CleanFileName.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_ArrayDimensionToList.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_GetListFromXmlFragment.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_availability_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_displayer_cls.asp"-->

<%
Class cPage
	' unencrypted
	Public MessageID
	Public SortBy
	Public ScheduledEvents
	Public Day
	Public Month
	Public Year
	
	' encrypted
	Public Action
	Public ReturnAction
	Public ProgramID
	Public EventID
	Public FileID
	Public ScheduleID
	
	' objects
	Public Member
	Public Client
	Public Program
	Public Evnt
	Public File
	Public Schedule
	
	' don't persist
	Public OtherAction
	Public EventList
		
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(Day) > 0 Then str = str & "d=" & Day & amp
		If Len(Month) > 0 Then str = str & "m=" & Month & amp
		If Len(Year) > 0 Then str = str & "y=" & Year & amp
		If Len(ScheduledEvents) > 0 Then str = str & "se=" & ScheduledEvents & amp
		If Len(SortBy) > 0 Then str = str & "sb=" & SortBy & amp

		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(EventID) > 0 Then str = str & "eid=" & Encrypt(EventID) & amp
		If Len(FileID) > 0 Then str = str & "fid=" & Encrypt(FileID) & amp
		If Len(ScheduleID) > 0 Then str = str & "scid=" & Encrypt(ScheduleID) & amp
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ReturnAction) > 0 Then str = str & "ract=" & Encrypt(ReturnAction) & amp
		
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
		c.ScheduledEvents = ScheduledEvents
		c.OtherAction = OtherAction
		c.Day = Day
		c.Month = Month
		c.Year = Year

		c.ProgramID = ProgramID
		c.EventID = EventID
		c.FileID = FileID
		c.ScheduleID = ScheduleID
		c.Action = Action
		c.ReturnAction = Action
		
		c.EventList = EventList
		
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Evnt = Evnt
		Set c.File = File
		Set c.Schedule = Schedule
		
		Set Clone = c
	End Function
End Class
%>

