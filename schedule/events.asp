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
	page.EventID = Decrypt(Request.QueryString("eid"))
	page.FileID = Decrypt(Request.QueryString("fid"))
	
	page.RemoveEventIdList = page.Uploader.Form("remove_event_id_list")

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
	Set page.Evnt = New cEvent
	page.Evnt.EventID = page.EventID
	If Len(page.Evnt.EventID) > 0 Then Call page.Evnt.Load()

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
		<link href="/_incs/script/jquery/plugins/asmselect/jquery.asmselect.css" rel="stylesheet" type="text/css" />
		<link rel="stylesheet" type="text/css" href="/_incs/script/jquery/jquery-ui-1.6.custom/development-bundle/themes/start/ui.all.css" />	
		<link rel="stylesheet" type="text/css" href="events.css" />	

		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>

		<script type="text/javascript" src="/_incs/script/jquery/jquery-ui-1.6.custom/js/jquery-ui-1.6.custom.min.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/fileupload/jquery.MultiFile.pack.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/asmselect/jquery.asmselect.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/getUrlParam/jquery.getUrlParam.js"></script>
		<script language="javascript" type="text/javascript" src="events.js"></script>

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
		Case SHOW_EVENT_DETAILS
			str = str & EventSummaryToString(page)
			
		Case UPDATE_RECORD
			If page.Uploader.Form("form_event_is_postback") = IS_POSTBACK Then
				Call LoadFormFromRequest(page)
				If ValidEvent(page.Evnt) Then
					Call DoUpdateEvent(page, rv)
					Select Case rv
						Case 0
							page.MessageID = 5000
						Case Else
							page.MessageID = 5001
					End Select
					page.Action = "": page.EventID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormEventToString(page)
				End If
			Else
				str = str & FormEventToString(page)
			End If
			
		Case ADDNEW_RECORD
			If page.Uploader.Form("form_event_is_postback") = IS_POSTBACK Then
				Call LoadFormFromRequest(page)
				If ValidEvent(page.Evnt) Then
					Call DoInsertEvent(page, rv)
					Select Case rv
						Case 0
							page.MessageID = 5000
						Case Else
							page.MessageID = 5001
					End Select
					page.Action = "": page.EventID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormEventToString(page)
				End If
			Else
				' set the scheduleId for this event when first hitting the form ..
				page.Evnt.ScheduleId = page.ScheduleId
				str = str & FormEventToString(page)
			End If
			
		Case DUPLICATE_EVENT
			If page.Uploader.Form("form_event_is_postback") = IS_POSTBACK Then
				Call LoadFormFromRequest(page)
				If ValidEvent(page.Evnt) Then
					Call DoInsertEvent(page, rv)
					If page.IncludeTeam = "1" Then
						Call CopyEventSchedule(page.Evnt.EventId, page.EventId, rv)
					End If
					Select Case rv
						Case 0
							If page.IncludeTeam = "1" Then
								page.MessageID = 5013
							Else
								page.MessageId = 5014
							End If	
						Case Else
							page.MessageID = 5001
					End Select
					page.Action = "": page.EventID = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormEventToString(page)
				End If
			Else
				str = str & FormEventToString(page)
			End If
			
		Case DELETE_RECORD
			If page.Uploader.Form("form_confirm_delete_event_is_postback") = IS_POSTBACK Then
				Call DoDeleteEvent(page.Evnt, rv)
				Select Case rv
					Case 0
						page.MessageID = 5003
					Case Else
						page.MessageID = 5006
				End Select
				page.Action = "": page.EventID = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
			Else
				str = str & FormConfirmDeleteEventToString(page)
			End If
			
		Case DELETE_MULTIPLE_EVENTS
			If page.Uploader.Form("form_confirm_delete_multiple_events") = IS_POSTBACK Then
				Call DoDeleteMultipleEvents(page.RemoveEventIdList, rv)
				Select Case rv
					Case 0
						page.MessageId = 5019
					Case Else
						page.MessageId = 5020
				End Select
				page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
				
			Else
				' no events selected ..
				If Len(page.RemoveEventIdList) = 0 Then
					page.Action = "": page.MessageId = 5021
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				End If
				str = str & FormConfirmRemoveMultipleEvents(page)
			End If
			
		Case STREAM_FILE_TO_BROWSER
			Call DoStreamFile(page.Member.MemberId, page.FileId, rv)
			Response.End
			
		Case SEND_MESSAGE
			page.EmailId = InsertEmailMessage(page.EventId, page.Member, rv)
			page.Action = ""
			Response.Redirect("/email/email.asp" & page.UrlParamsToString(False))
			
		Case Else
			If Not page.Schedule.HasEvents Then
				str = str & NoEventsDialogToString(page)
				Call DoClearTabLinkBar()
			Else
				str = str & EventGridToString(page)
			End If
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoClearTabLinkBar()
	m_tabLinkBarText = "<li>&nbsp;</li>"
End Sub

Function InsertEmailMessage(eventId, member, rv)
	Dim email			: Set email = New cEmail
	
	email.MemberId = member.MemberId
	email.ClientId = member.ClientId
	email.GroupList = "event|" & eventId
	Call email.Add(rv)
	
	InsertEmailMessage = email.EmailId
End Function

Sub DoDeleteMultipleEvents(idList, outError)
	Dim i
	
	Dim cacheError			: cacheError = 0
	outError = 0
	
	Dim evnt				: Set evnt = New cEvent
	Dim list				: list = Split(idList, ",")
	
	If IsArray(list) Then
		For i = 0 To UBound(list)
			If Len(list(i)) > 0 Then
				evnt.EventId = list(i)
				Call evnt.Delete(cacheError)
				outError = cacheError + outError
			End If
		Next
	End If
	
End Sub

Sub DoStreamFile(memberId, fileId, outError)
	Dim file			: Set file = New cFile
	
	file.FileID = fileId
	Call file.StreamFile(memberId, outError)
End Sub

Sub DoInsertEvent(page, outError)
	Call page.Evnt.Add(outError)
	Call SaveFiles(page, outError)
End Sub

Sub DoUpdateEvent(page, outError)
	Call page.Evnt.Save(outError)
	Call SaveFiles(page, outError)
End Sub

Sub DoDeleteEvent(evnt, outError)
	Call evnt.Delete(outError)
End Sub

Sub SaveFiles(page, outError)
	Dim i
	Dim file			: Set file = New cFile
	Dim eventFile		: Set eventFile = New cEventFile
	Dim tempError		: tempError = 0
	
	' handle any uploaded files ..
	For i = 1 To page.Uploader.Files.Count
		If page.Uploader.Files(i).Size > 0 Then
			
			' add to filestore ..
			file.ClientID = page.Client.ClientId
			file.ProgramID = page.Schedule.ProgramId
			file.FileOwnerID = page.Member.MemberId
			file.IsPublic = 1
			Call file.ToFilestore(page.Uploader.Files(i), tempError)
			outError = outError + tempError
			
			' add as eventfile ..
			eventFile.EventID = page.Evnt.EventID
			eventFile.FileID = file.FileID
			Call eventFile.Add(tempError)
			outError = outError + tempError
		End If
	Next
End Sub

Function SkillListForEventTeamGridForSummaryToString(fragment, xml)
	Dim str
	
	xml.Async = False
	xml.LoadXml(fragment)
	
	Dim node
	
	Dim isPublished
	Dim isSkillEnabled
	Dim isSkillGroupEnabled
	
	For Each node In xml.DocumentElement.ChildNodes
		isPublished = True
		If node.Attributes.GetNamedItem("PublishStatus").Text = CStr(IS_MARKED_FOR_PUBLISH) Then isPublished = False
		isSkillEnabled = True
		If node.Attributes.GetNamedItem("IsSkillEnabled").Text = "0" Then isSkillEnabled = False
		isSkillGroupEnabled = True
		If node.Attributes.GetNamedItem("IsSkillGroupEnabled").Text = "0" Then isSkillGroupEnabled = False
		
		If isPublished And isSkillEnabled And isSkillGroupEnabled Then
			str = str & node.Attributes.GetNamedItem("SkillName").Text & ", "
		End If
	Next
	If Len(str) > 0 Then str = Left(str, Len(str) - 2)
	
	SkillListForEventTeamGridForSummaryToString = str
End Function

Function EventTeamGridForSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim xml					: Set xml = Server.CreateObject("MSXML2.DOMDocument")
	
	Dim list				: list = page.Evnt.EventTeamDetailsList()
	Dim rows				: rows = ""
	Dim alt					: alt = ""
	Dim count				: count = 0
	
	Dim availabilityText
	Dim href
	
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-IsMemberEnabled 4-IsProgramMemberEnabled
	' 5-SkillListingXmlFragment 6-IsAvailable 7-IsAvailabilityViewedByMember 8-ProgramMemberId

	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isMemberEnabled = True			: If list(3,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True	: If list(4,i) = 0 Then isProgramMemberEnabled = False
		
			If isMemberEnabled And isProgramMemberEnabled Then
				alt = "":					: If count Mod 2 > 0 Then alt = " class=""alt"""
				
				availabilityText = "Yes"	: If list(6,i) = 0 Then availabilityText = "<span style=""color:red"">No</span>"
				If list(7,i) = 0 Then availabilityText = "??"
				
				pg.MemberId = list(0,i): pg.ProgramMemberId = list(8,i): pg.Action = SHOW_PROGRAM_MEMBER_DETAILS: pg.ScheduleId = "": pg.EventId = ""
				href = "/admin/profile.asp" & pg.UrlParamsToString(True)
				
				rows = rows & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/user.png"" alt="""" />"
				rows = rows & "<a href=""" & href & """ title=""Details""><strong>" & html(list(1,i) & ", " & list(2,i)) & "</strong></a></td>"
				rows = rows & "<td>" & availabilityText & "</td>"
				rows = rows & "<td>" & SkillListForEventTeamGridForSummaryToString(list(5,i), xml) & "</td>"
				rows = rows & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
				rows = rows & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
				
				count = count + 1
			End If
		Next
	End If	
	
	If count > 0 Then	
		str = str & "<div class=""grid"">"
		str = str & "<table><thead><tr><th>Member</th><th>Available</th><th>Skills</th><th>&nbsp;</th></tr></thead>"
		str = str & "<tbody>" & rows & "</tbody></table></div>"
	Else
		str = "<p class=""alert"">There are no members assigned (or published) to the event team for this event. </p>"
	End If	
	
	EventTeamGridForSummaryToString = str
End Function

Function FileGridForSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	Dim fileDisplay			: Set fileDisplay = New cFileDisplay
	
	Dim list				: list = page.Evnt.FileDetailsList()
	Dim items				: items = ""
	Dim alt					: alt = ""
	Dim count				: count = 0
	
	Dim style				: style = ""
	Dim href				: href = ""
	
	' 0-FileID 1-FileName 2-FriendlyName 3-Description 4-FileExtension 5-FileSize
	' 6-DownloadCount 7-IsPublic

	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			style = " style=""background-image:url('" & fileDisplay.GetIconPath(list(4,i)) & "');"""
			
			pg.Action = STREAM_FILE_TO_BROWSER: pg.FileId = list(0,i): pg.ScheduleId = "": pg.EventId = "": pg.Month = "": pg.Year = "": pg.Day = ""
			href = pg.Url & pg.UrlParamsToString(True)

			items = items & "<li" & style & "><a href=""" & href & """ title=""Download"">" 
			items = items & html(list(2,i) & "." & list(4,i)) & "</a></li>"
			
			count = count + 1
		Next
	End If	
	
	If count > 0 Then
		str = str & "<ul class=""file-list"">" & items & "</ul>"	
	Else
		str = "<p class=""alert"">There are no files linked to this event. </p>"
	End If	
	FileGridForSummaryToString = str
End Function

Function AvailabilityGridForEventSummaryToString(page)
	Dim str, i
	
	Dim eventAvailability	: Set eventAvailability = New cEventAvailability
	eventAvailability.EventId = page.Evnt.EventId
	
	Dim list				: list = eventAvailability.EventAvailabilityList()
	Dim count				: count = 0
	
	Dim memberName			: memberName = ""
	
	Dim isAvailable
	Dim hasSavedAvailability
	Dim isMemberEnabled
	Dim isProgramMemberEnabled
	
	Dim availableItems		
	Dim notAvailableItems
	Dim unknownItems		
	
	' 0-MemberId 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberIsActive 
	' 5-EventAvailabilityID 6-IsAvailable 7-IsViewedByMember 8-DateAvailabilityModified

	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isMemberEnabled = True						: If list(3,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True				: If list(4,i) = 0 Then isProgramMemberEnabled = False
			
			If isMemberEnabled And isProgramMemberEnabled Then
				memberName = list(1,i) & ", " & list(2,i)
				
				isAvailable = True						: If list(6,i) = 0 Then isAvailable = False
				hasSavedAvailability = True				: If list(7,i) = 0 Then hasSavedAvailability = False
				
				If Not hasSavedAvailability	Then
					unknownItems = unknownItems & "<li>" & html(memberName) & "</li>"
				Else
					If isAvailable Then
						availableItems = availableItems & "<li>" & html(memberName) & "</li>"
					Else
						notAvailableItems = notAvailableItems & "<li>" & html(memberName) & "</li>"
					End If
				End If
						
				count = count + 1
			End If
		Next
		
		If Len(unknownItems) > 0 Then unknownItems = "<ul class=""negative"">" & unknownItems & "</ul>"
		If Len(availableItems) > 0 Then availableItems = "<ul>" & availableItems & "</ul>"
		If Len(notAvailableItems) > 0 Then notAvailableItems = "<ul>" & notAvailableItems & "</ul>"
		
	End If
	
	If count > 0 Then
		str = str & "<p>Members in the 'not known' column have not yet logged in to indicate their availability for this event. </p>"
		str = str & "<div class=""grid"">"
		str = str & "<table class=""three-column""><thead><tr><th>Available</th><th>Not Available</th><th>Not Known</th></tr></thead>"
		str = str & "<tbody><tr><td class=""alt"">" & availableItems & "</td>" 
		str = str & "<td>" & notAvailableItems & "</td>" 
		str = str & "<td class=""alt"">" & unknownItems & "</td></tr></tbody></table></div>"
	Else
		str = "<p class=""alert"">Availability information is not available for this event because no members belong to this program. </p>"
	End If
	
	AvailabilityGridForEventSummaryToString = str
End Function

Function EventSummaryToString(page)
	Dim str
	Dim dateTime				: Set dateTime = New cFormatDate
	
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & html(page.Evnt.EventName) & "</h3>"
	str = str & "<h4 class=""first"">" & dateTime.Convert(page.Evnt.EventDate, "DDDD MMM dd, YYYY")
	If Len(page.Evnt.TimeStart & "") > 0 Then str = str & " at " & dateTime.Convert(page.Evnt.TimeStart	, "hh:nn pp")
	str = str & "</h4>"
	
	If page.Evnt.HasUnpublishedChanges = 1 Then 
		str = str & "<h5 class=""not-published"">Unpublished changes</h5>"
		str = str & "<p class=""alert"">The event team (schedule) for this event has changes that you have not yet published to your member calendar. </p>"
	End If
	
	str = str & "<h5 class=""description"">Notes for this event</h5>"
	If Len(page.Evnt.EventNote & "") > 0 Then 
		str = str & "<p>" & html(page.Evnt.EventNote) & ". </p>"
	Else
		str = str & "<p class=""alert"">There are no notes for this event. </p>"
	End If
	
	str = str & "<h5 class=""event-team"">Event team</h5>"
	str = str & EventTeamGridForSummaryToString(page)
	
	str = str & "<h5 class=""availability"">Member availability</h5>"
	str = str & AvailabilityGridForEventSummaryToString(page)
	
	str = str & "<h5 class=""files"">File list</h5>"
	str = str & FileGridForSummaryToString(page)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & "<ul><li>This event was created on " & dateTime.Convert(page.Evnt.DateCreated, "DDD MMMM dd, YYYY around hh:nn pp") & ". </li>"
	str = str & "</ul>"
	
	str = str & "</div>"
	
	EventSummaryToString = str
End Function

Function EventGridToString(page)
	Dim str, i
	Dim dateTime	: Set dateTime = New cFormatDate
	Dim pg			: Set pg = page.Clone()
	
	Dim list		: list = page.Schedule.EventList("")
	Dim count		: count = 0
	
	Dim tipbox
	tipBox = tipBox & "<div class=""tip-box""><h3>I want to ..</h3><ul>"
	pg.Action = ADDNEW_RECORD
	tipBox = tipBox & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Add a new event</a></li>"
	tipBox = tipBox & "</ul></div>"
	
	str = str & tipbox
	str = str & m_appMessageText
	str = str & "<div class=""grid"">"
	pg.Action = DELETE_MULTIPLE_EVENTS
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-delete-multiple-events"" enctype=""multipart/form-data"">"
	str = str & "<input type=""hidden"" name=""form_remove_multiple_events_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><thead><tr><th scope=""col"" style=""width:1%;""><input type=""checkbox"" name=""master"" id=""master"" /></th>"
	str = str & "<th scope=""col"" style=""width:50%;"">Event</th>"
	str = str & "<th scope=""col"">Date</th>"
	str = str & "<th scope=""col"">Time</th>"
	str = str & "<th scope=""col"">&nbsp;</th>"
	str = str & "</tr></thead><tbody>"

	' 0-EventID 1-EventName 2-EventDate 3-EventNote 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified
	' 8-ProgramID 9-ProgramName 10-ScheduleID 11-ScheduleName 12-ScheduleEventCount
	' 13-ScheduleIsVisible 14-HasUnpublishedChanges 15-EventFileCount 16-ScheduledMemberCount
	' 17-HtmlBackgroundColor
	
	' refresh pg as I set action above ..
	Set pg = page.Clone()
	
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			count = count + 1
			
			str = str & "<tr><td><input type=""checkbox"" name=""remove_event_id_list"" value=""" & list(0,i) & """ /></td>"
			str = str & "<td><img class=""icon"" src=""/_images/icons/date.png"" alt="""" />"
			str = str & "<strong>" & html(list(11,i)) & " | </strong>"
			
			pg.EventId = list(0,i): pg.Action = SHOW_EVENT_DETAILS
			str = str & "<strong><a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Details"">" & html(list(1,i)) & "</a></strong></td>"

			str = str & "<td>" & dateTime.Convert(list(2,i), "mm-dd-YYYY") & "</td>"
			str = str & "<td>"
			If Len(list(4,i)) > 0 Then
				str = str & dateTime.Convert(list(4,i), "hh:nn&nbsp;pp")
			Else
				str = str & "&nbsp;"
			End If
			str = str & "</td>"
			str = str & "<td class=""toolbar"">"
			pg.EventId = list(0,i): pg.Action = SHOW_EVENT_DETAILS
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Event Details"">"
			str = str & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a>"
			pg.EventID = list(0,i): pg.Action = UPDATE_RECORD
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Edit Event"">"
			str = str & "<img src=""/_images/icons/pencil.png"" alt="""" /></a>"
			pg.EventID = list(0,i): pg.Action = DUPLICATE_EVENT
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Copy Event"">"
			str = str & "<img src=""/_images/icons/paste_date.png"" alt="""" /></a>"
			pg.EventID = list(0,i): pg.Action = UPDATE_RECORD
			str = str & "<a href=""/schedule/teams.asp" & pg.UrlParamsToString(True) & """ title=""Event Team"">"
			str = str & "<img src=""/_images/icons/group_edit.png"" alt="""" /></a>"
			
			pg.EventId = list(0,i): pg.Action = SEND_MESSAGE
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Email"">"
			str = str & "<img src=""/_images/icons/email.png"" alt="""" /></a>"
			pg.EventID = list(0,i): pg.Action = DELETE_RECORD
			str = str & "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """ title=""Remove Event"">"
			str = str & "<img src=""/_images/icons/cross.png"" alt="""" /></a></td></tr>"
		Next
	End If
	str = str & "</tbody></table></form></div>"
	
	EventGridToString = str
End Function

Sub LoadFormFromRequest(page)
	page.Evnt.ScheduleID = page.Uploader.Form("schedule_id")
	page.Evnt.EventName = page.Uploader.Form("event_name")
	page.Evnt.EventNote = page.Uploader.Form("event_note")
	page.Evnt.EventDate = page.Uploader.Form("event_date")
	page.Evnt.TimeStart = page.Uploader.Form("time_start")
	page.Evnt.TimeEnd = page.Uploader.Form("time_end")
	page.Evnt.FileList = page.Uploader.Form("file_id_list")
	
	page.IncludeTeam = page.Uploader.Form("include_team")
End Sub

Function ValidEvent(evnt)
	ValidEvent = True
	If Application.Value("FORM_VALIDATION_OFF") Then Exit Function	
	
	Dim isTime		: isTime = True
	
	If Len(evnt.ScheduleID) = 0 Then
		AddCustomFrmError("You must select a schedule for the event to belong to.")
		ValidEvent = False
	End If
	If Not ValidData(evnt.EventName, True, 0, 200, "Event Name", "") Then ValidEvent = False
	If Not ValidData(evnt.EventNote, False, 0, 1000, "Event Notes", "") Then ValidEvent = False
	If Not ValidData(evnt.EventDate, True, 0, 0, "Date", "date") Then ValidEvent = False
	
	If Not ValidData(evnt.TimeStart, False, 0, 0, "Start Time", "time") Then 
		ValidEvent = False
		isTime = False
	End If
	If Not ValidData(evnt.TimeEnd, False, 0, 0, "End Time", "time") Then 
		ValidEvent = False
		isTime = False
	End If

	' only run this validation if start/end time are valid
	If isTime Then
		If Len(evnt.TimeEnd & evnt.TimeStart) <> 0 Then
			If Len(evnt.TimeStart) = 0 Then
				'can't provide only TimeEnd
				AddCustomFrmError("A Start Time is required if an End Time is provided (both are optional).")
				ValidEvent = False
			ElseIf Len(evnt.TimeEnd) = 0 Then
				'do nothing - I need this condition so I don't hit next test
			ElseIf DateDiff("s", CDate(evnt.TimeStart), CDate(evnt.TimeEnd)) <= 0 Then
				'Start time must be before end time
				AddCustomFrmError("An End Time cannot occur before a Start Time (both are optional).")
				ValidEvent = False
			End If
		End If
	End If
End Function

Function ScheduleDropdownToString(list, id)
	Dim str, i
	Dim selected		: selected = ""
	
	str = str & "<select name=""schedule_id"">"
	str = str & "<option value="""">" & html("< Select a schedule >") & "</option>"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			selected = ""
			If CStr(list(0,i) & "") = CStr(id & "") Then selected = " selected=""selected"""
			str = str & "<option value=""" & list(0,i) & """" & selected & ">" & html(list(1,i)) & "</option>"
		Next
	End If	
	str = str & "</select>"
	
	ScheduleDropdownToString = str
End Function

Function IsSelected(stringList, id)
	Dim str, i
	
	If Len(stringList) = 0 Then Exit Function
	
	Dim list			: list = Split(stringList, ",")
	
	If IsArray(list) Then
		For i = 0 To UBound(list)
			If CStr(list(i) & "") = CStr(id & "") Then
				str = " selected=""selected"""
			End If
		Next
	End If
	
	IsSelected = str
End Function

Function FileDropdownToString(programFileList, selectedFiles)
	Dim str, i
	
	' 0-FileID 1-FriendlyName 2-FileExtension 3-FileName

	str = str & "<select name=""file_id_list"" id=""file-list"" multiple=""multiple"" title=""Include a file .."">"
	If IsArray(programFileList) Then
		For i = 0 To UBound(programFileList,2)
			str = str & "<option value=""" & programFileList(0,i) & """" & IsSelected(selectedFiles, programFileList(0,i)) & ">" 
			str = str & html(programFileList(1,i) & "." & programFileList(2,i)) & "</option>"
		Next
	End If
	str = str & "</select>"
	
	FileDropdownToString = str 
End Function

Function FormEventToString(page)
	Dim str
	
	Dim pg				: Set pg = page.Clone()
	
	Dim cancelUrl
	pg.Action = "": pg.EventId = ""
	cancelUrl = "/schedule/events.asp" & pg.UrlParamsToString(True)
	
	' reset pg object as it was just modified ..
	Set pg = page.Clone()
	
	Dim program				: Set program = New cProgram
	program.ProgramID = page.Schedule.ProgramID
	
	Dim programFileList		: programFileList = program.FileList()
		
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-event"" enctype=""multipart/form-data"">"
	str = str & "<input type=""hidden"" name=""form_event_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><tbody>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Schedule") & "</td>"
	str = str & "<td>" & ScheduleDropdownToString(program.ScheduleList, page.Evnt.ScheduleID) & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">The schedule that this event will belong to.</td></tr>"
	
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Event Name") & "</td>"
	str = str & "<td><input type=""text"" name=""event_name"" value=""" & html(page.Evnt.EventName) & """ class=""large"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Event Note") & "</td>"
	str = str & "<td><textarea name=""event_note"" class=""large"">" & html(page.Evnt.EventNote) & "</textarea></td></tr>"

	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Date") & "</td>"
	str = str & "<td><input type=""text"" id=""event-date"" name=""event_date"" class=""small"" value=""" & html(page.Evnt.EventDate) & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Start Time") & "</td>"
	str = str & "<td><input type=""text"" name=""time_start"" class=""small"" value=""" & html(page.Evnt.TimeStart) & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "End Time") & "</td>"
	str = str & "<td><input type=""text"" name=""time_end"" class=""small"" value=""" & html(page.Evnt.TimeEnd) & """ /></td></tr>"
	
	If page.Action = DUPLICATE_EVENT Then
		str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
		str = str & "<tr><td class=""label top"">" & RequiredElementToString(False, "Include team?") & "</td>"
		str = str & "<td>" & YesNoDropdownToString(page.IncludeTeam, "include_team") & "</td></tr>"
		str = str & "<tr><td>&nbsp;</td><td class=""hint"">If yes, then the event team for this event will also <br />be copied into your new event. </td></tr>"
	End If
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label top"">" & RequiredElementToString(False, "Files") & "</td>"
	str = str & "<td>" & FileDropdownToString(programFileList, page.Evnt.FileList)
	str = str & "<div class=""hint"">"
	If IsArray(programFileList) Then
		str = str & "Your members will be able to download any files <br />you select from this list when they view this event <br />on their calendar. "
	Else
		str = str & "You haven't uploaded any files for this program. <br /> "
		str = str & "Any files you save with this event will be automatically<br /> added to this file list. "
	End If
	str = str & "</div></td></tr>"

	str = str & "<tr><td>&nbsp;</td><td><div id=""upload-trigger"" style=""margin-top:10px;""><a href=""#"">Upload</a> more files</div>"
	str = str & "<input type=""file"" id=""event-file"" name=""event_file_"" class=""multi"" /></td></tr>"

	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""submit"" value=""Save"" />"
	str = str & "&nbsp;&nbsp;<a href=""" & cancelUrl & """>Cancel</a></td></tr>"
	str = str & "</tbody></table></form></div>"
	
	FormEventToString = str
End Function

Function FormConfirmDeleteEventToString(page)
	Dim str, msg
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim pg				: Set pg = page.Clone()
	Dim cancelUrl
	pg.Action = "": pg.EventId = ""
	cancelUrl = "/schedule/events.asp" & pg.UrlParamsToString(True)
	
	' reset pg object as it was just modified ..
	Set pg = page.Clone()
	
	Dim returnPage		: returnPage = "/schedule/events.asp"
	
	msg = msg & "You will permanently remove the <strong>" & html(page.Evnt.EventName) & "</strong> event (on " & dateTime.Convert(page.Evnt.EventDate, "DDDD MMM dd, YYYY") & ") from the " & html(page.Evnt.ScheduleName) & " schedule. "
	msg = msg & "You will lose any calendar and or schedule information associated with this event. "
	msg = msg & "This action cannot be reversed. "
	
	str = str & CustomApplicationMessageToString("Please confirm this action!", msg, "Confirm")
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-confirm-delete-event"" enctype=""multipart/form-data"">"
	str = str & "<input type=""hidden"" name=""form_confirm_delete_event_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = "": pg.ScheduleId = "": pg.EventID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & cancelUrl & """>Cancel</a></td></tr>"
	str = str & "</p></form>"
	
	FormConfirmDeleteEventToString = str
End Function

Function FormConfirmRemoveMultipleEvents(page)
	Dim str, msg
	Dim pg				: Set pg = page.Clone()
	
	Dim count		
	Dim countText
	Dim countText2
	
	If Len(page.RemoveEventIdList) > 0 Then
		count = UBound(Split(Replace(page.RemoveEventIdList, " ", ""), ",")) + 1
	
		If count = 1 Then 
			countText = "this event"
			countText2 = "this event"
		Else
			countText = "these " & count & " events"
			countText2 = "these events"
		End If 
	Else
		countText = "these 0 events"
		countText2 = "these events"
	End If
	
	Dim cancelUrl
	pg.Action = "": pg.EventId = ""
	cancelUrl = "/schedule/events.asp" & pg.UrlParamsToString(True)
	
	' reset pg object as it was just modified ..
	Set pg = page.Clone()

	msg = msg & "You will permanently remove " & countText & " from the <strong>" & html(page.Schedule.ScheduleName) & "</strong> schedule. "
	msg = msg & "You will lose any calender, schedule, or event team information associated " & countText2 & ". "
	msg = msg & "This action cannot be reversed. "
	
	str = str & CustomApplicationMessageToString("Please confirm this action!", msg, "Confirm")
	
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-confirm-delete-multiple-events"" enctype=""multipart/form-data"">"
	str = str & "<input type=""hidden"" name=""form_confirm_delete_multiple_events"" value=""" & IS_POSTBACK & """ />"
	str = str & "<input type=""hidden"" name=""remove_event_id_list"" value=""" & page.RemoveEventIdList & """ />"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = "": pg.ScheduleId = "": pg.EventID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & cancelUrl & """>Cancel</a></td></tr>"
	str = str & "</p></form>"
	
	FormConfirmRemoveMultipleEvents = str
End Function

Function NoEventsDialogToString(page)
	Dim dialog			: Set dialog = New cDialog
	Dim pg				: Set pg = page.Clone()
	
	dialog.HeadLine = "Whoa, something's missing here ..!"
	
	dialog.Text = dialog.Text & "<p>You are in the Manage Schedules section of your account, "
	dialog.Text = dialog.Text & "trying to view the event list for your " & html(page.Schedule.ScheduleName) & " schedule. "
	dialog.Text = dialog.Text & "However, there are no events in this schedule yet. </p>"
	dialog.Text = dialog.Text & "<p>To get started on fixing this, click <strong>Create your first event</strong>. </p>"
	
	dialog.SubText = dialog.SubText & "<p>Once you have created an event, "
	dialog.SubText = dialog.SubText & "this page is where you'll add, remove or copy the events for this schedule. "
	dialog.SubText = dialog.SubText & "This is also where you'll be setting up your event teams (the members you've scheduled to work at or attend your events). </p>"

	pg.Action = ADDNEW_RECORD
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Create your first event</a></li>"
	pg.Action = ""
	dialog.LinkList = dialog.LinkList & "<li><a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Back to my schedules</a></li>"

	NoEventsDialogToString = dialog.ToString()
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim dateTime	: Set dateTime = New cFormatDate
	
	Dim eventGridLink
	pg.Action = "": pg.EventId = ""
	eventGridLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Events</a> / "
	
	Dim scheduleLink
	pg.FilterScheduleId = pg.ScheduleId
	pg.ScheduleId = "": pg.ProgramId = ""
	scheduleLink = "<a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>" & html(page.Schedule.ScheduleName) & "</a> / "

	Dim programLink
	Set pg = page.Clone()
	pg.ProgramId = page.Schedule.ProgramId
	pg.Action = "": pg.FilterScheduleId = "": pg.ScheduleId = ""
	programLink = "<a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>" & html(page.Schedule.ProgramName) & "</a> / "
	
	Dim scheduleHomeLink
	pg.ScheduleId = "": pg.FilterScheduleId = ""
	scheduleHomeLink = "<a href=""/schedule/schedules.asp" & pg.UrlParamsToString(True) & """>Schedules</a> / "

	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case SHOW_EVENT_DETAILS
			str = str & scheduleHomeLink & programLink & scheduleLink & html(page.Evnt.EventName) & " on " & dateTime.Convert(page.Evnt.EventDate, "DDD MMM dd, YYYY")
		
		Case DUPLICATE_EVENT
			str = str & scheduleHomeLink & programLink & scheduleLink & eventGridLink & "Copy Event"
		Case ADDNEW_RECORD
			str = str & scheduleHomeLink & programLink & scheduleLink & eventGridLink & "Add Event"
		Case UPDATE_RECORD
			str = str & scheduleHomeLink & programLink & scheduleLink & eventGridLink & "Edit Event '" & html(page.Evnt.EventName) & " (" & dateTime.Convert(page.Evnt.EventDate, "MM/dd/YYYY") & ")'"
		Case DELETE_RECORD
			str = str & scheduleHomeLink & programLink & scheduleLink & eventGridLink & "Remove Event '" & html(page.Evnt.EventName) & " (" & dateTime.Convert(page.Evnt.EventDate, "MM/dd/YYYY") & ")'"
		Case Else
			str = str & scheduleHomeLink & programLink & scheduleLink & "Event List"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	If Len(pg.ScheduleId & "") = 0 Then
		pg.ScheduleId = pg.Evnt.ScheduleId
	End If
	
	Dim eventTeamViewButton
	pg.EventId = "": pg.Action = ""
	href = "/schedule/teams.asp" & pg.UrlParamsToString(True)
	eventTeamViewButton = "<li><a href=""" & href & """><img src=""/_images/icons/group.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Event Teams</a></li>"
	
	Dim deleteMultipleButton
	pg.EventId = "": pg.Action = DELETE_MULTIPLE_EVENTS
	href = pg.Url & pg.UrlParamsToString(True)
	deleteMultipleButton = deleteMultipleButton & "<li><a href=""" & href & """ class=""delete-multiple-events""><img src=""/_images/icons/date_delete.png"" class=""icon"" alt="""" /></a><a href=""" & href & """ class=""delete-multiple-events"">Delete Checked</a></li>"
	
	Dim eventListButton
	pg.EventID = "": pg.Action = ""
	href = pg.Url & pg.UrlParamsToString(True)
	eventListButton = EventListButton & "<li><a href=""" & href & """><img src=""/_images/icons/event_multiple_2.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Event List</a></li>"
	
	Dim newEventButton
	Set pg = page.Clone()
	pg.EventID = "": pg.Action = ADDNEW_RECORD
	href = pg.Url & pg.UrlParamsToString(True)
	newEventButton = newEventButton & "<li><a href=""" & href & """><img src=""/_images/icons/date_add.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>New Event</a></li>"
	
	Dim schedulesButton
	If CStr(pg.FilterScheduleId) <> CStr(pg.ScheduleId) Then pg.FilterScheduleId = ""
	pg.EventID = "": pg.Action = "": pg.ScheduleID = ""
	href = "/schedule/schedules.asp" & pg.UrlParamsToString(True)
	schedulesButton = schedulesButton & "<li><a href=""" & href & """><img src=""/_images/icons/calendar.png"" class=""icon"" alt="""" /></a><a href=""" & href & """>Schedules</a></li>"
	
	Select Case page.Action
		Case SHOW_EVENT_DETAILS
			str = str & eventTeamViewButton & eventListButton & schedulesButton
		
		Case UPDATE_RECORD
			str = str & eventListButton

		Case ADDNEW_RECORD
			str = str & eventListButton
			
		Case DELETE_RECORD
			str = str & eventListButton
			
		Case Else
			str = str & schedulesButton & eventTeamViewButton & deleteMultipleButton & newEventButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_YesNoDropdownToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_CopyEventSchedule.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_availability_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_file_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_build_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/dialog_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/file_displayer_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	Public Month
	Public Day
	Public Year
	Public IncludeTeam
	Public RemoveEventIdList
	
	' encrypted
	Public Action
	Public ProgramID
	Public ScheduleID
	Public FilterScheduleId
	Public EventID
	Public FileID
	Public EmailId
	Public MemberId
	Public ProgramMemberId

	' objects
	Public Member
	Public Client
	Public Program
	Public Schedule
	Public Evnt
	Public Uploader	
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(Month) > 0 Then str = str & "m=" & Month & amp
		If Len(Year) > 0 Then str = str & "y=" & Year & amp
		If Len(Day) > 0 Then str = str & "d=" & Day & amp
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(ProgramID) > 0 Then str = str & "pid=" & Encrypt(ProgramID) & amp
		If Len(ScheduleID) > 0 Then str = str & "scid=" & Encrypt(ScheduleID) & amp
		If Len(FilterScheduleId) > 0 Then str = str & "fscid=" & Encrypt(FilterScheduleId) & amp
		If Len(EventID) > 0 Then str = str & "eid=" & Encrypt(EventID) & amp
		If Len(FileID) > 0 Then str = str & "fid=" & Encrypt(FileID) & amp
		If Len(EmailId) > 0 Then str = str & "emid=" & Encrypt(EmailId) & amp
		If Len(MemberId) > 0 Then str = str & "mid=" & Encrypt(MemberId) & amp
		If Len(ProgramMemberId) > 0 Then str = str & "pmid=" & Encrypt(ProgramMemberId) & amp
		
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
		c.Month = Month
		c.Year = Year
		c.Day = Day

		c.Action = Action
		c.ProgramID = ProgramID
		c.ScheduleID = ScheduleID
		c.FilterScheduleId = FilterScheduleId
		c.EventID = EventID
		c.FileID = FileID
		c.EmailId = EmailId
		c.MemberId = MemberId
		c.ProgramMemberId = ProgramMemberId
				
		Set c.Member = Member
		Set c.Client = Client
		Set c.Program = Program
		Set c.Schedule = Schedule
		Set c.Evnt = Evnt
		Set c.Uploader = Uploader
		
		Set Clone = c
	End Function
End Class
%>

