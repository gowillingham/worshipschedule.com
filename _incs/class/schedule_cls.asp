<%
Class cSchedule

	Private m_iScheduleID		'as long int
	Private m_iProgramID		'as long int
	Private m_sProgramName		'as string
	Private m_sScheduleName		'as string
	Private m_sScheduleDesc		'as string
	Private m_iIsVisible		'as tiny int
	Private m_dDateCreated		'as date
	Private m_dDateModified		'as date
	Private m_dDatePublished	'as date
	Private m_iPublicationStatus	'as tiny int
	Private m_iEventCount		'as int
	Private m_iScheduledMemberCount		' as int
	Private m_htmlBackgroundColor	' str
	
	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get ScheduleID() 'As long int
		ScheduleID = m_iScheduleID
	End Property

	Public Property Let ScheduleID(val) 'As long int
		m_iScheduleID = val
	End Property
	
	Public Property Get ProgramID() 'As long int
		ProgramID = m_iProgramID
	End Property

	Public Property Let ProgramID(val) 'As long int
		m_iProgramID = val
	End Property
	
	Public Property Get ProgramName()
		ProgramName = m_sProgramName
	End Property
	
	Public Property Get ScheduleName() 'As string
		ScheduleName = m_sScheduleName
	End Property

	Public Property Let ScheduleName(val) 'As string
		m_sScheduleName = val
	End Property
	
	Public Property Get ScheduleDesc() 'As string
		ScheduleDesc = m_sScheduleDesc
	End Property

	Public Property Let ScheduleDesc(val) 'As string
		m_sScheduleDesc = val
	End Property
	
	Public Property Get IsVisible() 'As small int
		IsVisible = m_iIsVisible
	End Property

	Public Property Let IsVisible(val) 'As small int
		m_iIsVisible = val
	End Property
	
	Public Property Get EventCount()
		EventCount = m_iEventCount
	End Property

	Public Property Get HasEvents() ' as bool
		HasEvents = False
		If m_iEventCount > 0 Then HasEvents = True
	End Property
	
	Public Property Get ScheduledMemberCount()
		ScheduledMemberCount = m_iScheduledMemberCount
	End Property
	
	Public Property Get HasScheduledMembers()
		HasScheduledMembers = False
		If m_iScheduledMemberCount > 0 Then HasScheduledMembers = True
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_dDateCreated
	End Property

	Public Property Let DateCreated(val) 'As small date
		m_dDateCreated = val
	End Property
	
	Public Property Get DateModified() 'As date
		DateModified = m_dDateModified
	End Property

	Public Property Let DateModified(val) 'As small date
		m_dDateModified = val
	End Property
	
	Public Property Get DatePublished() 'As date
		DatePublished = m_dDatePublished
	End Property
	
	Public Property Get PublishStatus() 'As tiny int
		PublishStatus = m_iPublicationStatus
	End Property
	
	Public Property Get HtmlBackgroundColor()
		HtmlBackgroundColor = m_htmlBackgroundColor
	End Property
	
	Public Property Let HtmlBackgroundColor(val)
		m_htmlBackgroundColor = val
	End Property
	
	Private Sub Class_Initialize()
		m_sScheduleName = ""
		m_sScheduleDesc = ""
		m_iIsVisible = 1
		m_dDateCreated = ""
		m_dDateModified = ""
		m_iPublicationStatus = 0
			
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cSchedule"
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(m_rs) Then
			If m_rs.State = adStateOpen Then m_rs.Close
			Set m_rs = Nothing
		End If
		If IsObject(m_cnn) Then
			If m_cnn.State = adStateOpen Then m_cnn.Close
			Set m_cnn = Nothing
		End If
	End Sub
	
	Public Function List()
		If Len(m_iProgramID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".List()", "Required parameter ProgramID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
	
		' 0-ScheduleId 1-ScheduleName 2-ScheduleDesc 3-IsVisible 4-DateCreated 5-DateModified
		' 6-ProgramID 7-ProgramName 8-HtmlBackgroundColor 9-EventCount
		
		m_cnn.up_scheduleGetScheduleListForProgram CLng(m_iProgramID), m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function EventTeamDetailsList()
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(5 + vbObjectError, CLASS_NAME & ".EventTeamDetailList()", "Required parameter ScheduleID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberIsEnabled 5-ProgramMemberIsEnabled 
		' 6-EventId 7-EventName 8-EventDate 9-TimeStart 10-TimeEnd 11-EventNote 12-SkillListXmlFragment
		' 13-FileListXmlFragment 14-ProgramMemberId
				
		m_cnn.up_scheduleGetEventTeamDetailsByScheduleId CLng(m_iScheduleID), m_rs
		If Not m_rs.EOF Then EventTeamDetailsList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function MemberList()	
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ProgramMemberSkillList()", "Required parameter ScheduleID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-EventID 5-IsAvailable 6-IsViewedByMember
		' 7-SkillName 8-email 9-ProgramMemberIsActive
		
		m_cnn.up_scheduleGetMemberList CLng(m_iScheduleID), m_rs
		If Not m_rs.EOF Then MemberList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function AvailabilityList()
		' returns list of members missing availability ..
		
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".AvailabilityList()", "Required parameter ScheduleID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-MemberID 1-NameLast 2-NameFirst 3-LastLogin 4-Email 5-IsMemberAccountEnabled
		' 6-IsMissingAvailabilityInfo 7-IsProgramMemberEnabled 8-ProgramMemberId
		
		m_cnn.up_scheduleGetAvailabilityUpToDateList CLng(m_iScheduleID), m_rs
		If Not m_rs.EOF Then AvailabilityList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function ScheduleBuildList()
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ScheduleBuildList()", "Required parameter ScheduleID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-MemberId 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberIsActive 
		' 5-PublishStatus 6-SkillID 7-SkillName 8-SkillIsEnabled 9-SkillGroupID 
		' 10-SkillGroupName 11-SkillGroupIsEnabled 12-ProgramMemberSkillID 13-EventId
		' 14-EventName 15-EventDate 16-TimeStart 17-TimeEnd 18-IsAvailable 
		' 19-IsAvailabilityViewedByMember 20-AvailabilityDateModified 21-ScheduleID
		' 22-ScheduleName

		m_cnn.up_scheduleGetScheduleBuildForScheduleID CLng(m_iScheduleID), m_rs
		If Not m_rs.EOF Then ScheduleBuildList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function EventListForPeriod(startDate, endDate, outError)
		Dim cmd
		
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(5 + vbObjectError, CLASS_NAME & ":EventListForPeriod()", "Required parameter m_iScheduleID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_scheduleGetEventListForPeriod"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleId", adBigInt, adParamInput, 0, CLng(m_iScheduleId))
		cmd.Parameters.Append cmd.CreateParameter("@Start", adDate, adParamInput, 0, startDate)
		cmd.Parameters.Append cmd.CreateParameter("@End", adDate, adParamInput, 0, endDate)

		'0-EventID 1-EventName 2-EventNote 3-EventDate 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified 8-HasFiles

		Set m_rs = cmd.Execute
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If Not m_rs.EOF Then EventListForPeriod = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close()
		Set cmd = Nothing
	End Function
	
	Public Function EventList(sortColumn)
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ":GetEventArrayList()", "Required parameter m_iScheduleID not provided.")
		
		' 0-EventID 1-EventName 2-EventDate 3-EventNote 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified
		' 8-ProgramID 9-ProgramName 10-ScheduleID 11-ScheduleName 12-ScheduleEventCount
		' 13-ScheduleIsVisible 14-HasUnpublishedChanges 15-EventFileCount 16-ScheduledMemberCount
		' 17-HtmlBackgroundColor 18-FileListXMLFragment
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		m_rs.CursorLocation = adUseClient	' makes sortable
		
		m_cnn.up_eventGetEventListByScheduleID CLng(m_iScheduleID), m_rs
		
		m_rs.Sort = sortColumn
		If Not m_rs.EOF Then EventList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Load() 'As Boolean
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ":Load()", "Required parameter m_iScheduleID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_scheduleGetSchedule CLng(m_iScheduleID), m_rs
		If Not m_rs.EOF Then
			m_iScheduleID = m_rs("ScheduleID").Value
			m_iProgramID = m_rs("ProgramID").Value
			m_sProgramName = m_rs("ProgramName").Value
			m_sScheduleName = m_rs("ScheduleName").Value
			m_sScheduleDesc = m_rs("ScheduleDesc").Value
			m_iIsVisible = m_rs("IsVisible").Value
			m_dDateCreated = m_rs("DateCreated").Value
			m_dDateModified = m_rs("DateModified").Value
			m_dDatePublished = m_rs("DateLastPublished").Value
			m_iPublicationStatus = m_rs("PublicationStatus").Value
			m_iEventCount = m_rs("EventCount").Value
			m_iScheduledMemberCount = m_rs("ScheduledMemberCount").Value
			m_htmlBackgroundColor = m_rs("HtmlBackgroundColor").Value
			
			Load = True
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd

		If Len(m_iProgramID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ":Add()", "Required parameter m_iProgramID not provided.")
		If Len(m_sScheduleName) = 0 Then Call Err.Raise(3 + vbObjectError, CLASS_NAME & ":Add()", "Required parameter m_sScheduleName not provided.")
		
		If Not IsDate(m_dDateCreated) Then m_dDateCreated = Now()
		If Not IsDate(m_dDateModified) Then m_dDateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_scheduleInsert"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, CLng(m_iProgramID))
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleName", adVarChar, adParamInput, 100, CStr(m_sScheduleName))
		If Len(m_sScheduleDesc) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@ScheduleDesc", adVarChar, adParamInput, 1000, CStr(m_sScheduleDesc))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@ScheduleDesc", adVarChar, adParamInput, 1000, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@IsVisible", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsVisible))
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, CDate(m_dDateCreated))
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_dDateModified)
		cmd.Parameters.Append cmd.CreateParameter("@HtmlBackgroundcolor", adVarChar, adParamInput, 10, CStr(m_htmlBackgroundColor))

		cmd.Parameters.Append cmd.CreateParameter("@NewScheduleID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_iScheduleID = cmd.Parameters("@NewScheduleID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd

		If Len(m_iScheduleID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ":Save()", "Required parameter m_iScheduleID not provided.")
		If Len(m_iProgramID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ":Save()", "Required parameter m_iProgramID not provided.")
		If Len(m_iProgramID) = 0 Then Call Err.Raise(3 + vbObjectError, CLASS_NAME & ":Save()", "Required parameter m_sScheduleName not provided.")
		
		If Not IsDate(m_dDateModified) Then m_dDateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_scheduleUpdate"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleID", adBigInt, adParamInput, 0, CLng(m_iScheduleID))
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleName", adVarChar, adParamInput, 100, CStr(m_sScheduleName))
		If Len(m_sScheduleDesc) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@ScheduleDesc", adVarChar, adParamInput, 1000, CStr(m_sScheduleDesc))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@ScheduleDesc", adVarChar, adParamInput, 1000, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@IsVisible", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsVisible))
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_dDateModified)
		cmd.Parameters.Append cmd.CreateParameter("@HtmlBackgroundcolor", adVarChar, adParamInput, 10, CStr(m_htmlBackgroundColor))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Save = True
		Else
			Save = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Function Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ":Save()", "Required parameter m_iScheduleID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_scheduleDelete"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleID", adBigInt, adParamInput, 0, CLng(m_iScheduleID))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Delete = True
		Else
			Delete = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Function Publish(publishedBy, outError)
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ":Save()", "Required parameter m_iScheduleID not provided.")
		
		If UpdateSchedulePublish(publishedBy, "PUBLISH", outError) Then
			Publish = True
		Else
			Publish = False
		End If
	End Function
	
	Public Function RemovePublish(outError)
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ":Save()", "Required parameter m_iScheduleID not provided.")

		If UpdateSchedulePublish("", "UNPUBLISH", outError) Then
			RemovePublish = True
		Else
			RemovePublish = False
		End If
	End Function
	
	Private Function UpdateSchedulePublish(publishedBy, action, outError)
		Dim cmd

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_schedulePublishByScheduleID"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleID", adBigInt, adParamInput, 0, CLng(m_iScheduleID))
		cmd.Parameters.Append cmd.CreateParameter("@Action", adVarChar, adParamInput, 25, CStr(action))
		cmd.Parameters.Append cmd.CreateParameter("@DatePublished", adDBTimeStamp, adParamInput, 0, Now())
		cmd.Parameters.Append cmd.CreateParameter("@Publisher", adVarChar, adParamInput, 102, CStr(publishedBy))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		If outError = 0 Then
			UpdateSchedulePublish = True
		Else
			UpdateSchedulePublish = False
		End If
		
		Set cmd = Nothing
	End Function
End Class
%>
