<%
Class cProgram

	Private m_ProgramID		'as long int
	Private m_ClientID		'as long int
	Private m_ClientName	'as str
	Private m_ProgramName		'as string
	Private m_ProgramDesc		'as string
	Private m_EnrollmentType		'as small int
	Private m_IsEnabled		'as small int
	Private m_DateCreated		'as date
	Private m_DateModified		'as date
	Private m_DefaultAvailability		'as small int
	Private m_MemberCanEnroll		' as tinyint
	Private m_MemberCanEditSkills	' as tinyint
	Private m_HasSkills		' as tinyint
	Private m_HasSkillGroups	' as tinyint
	Private m_MemberCount		' as int
	Private m_ScheduleCount		' as int
	Private m_EventCount		' as int
	Private m_CurrentEventCount	' as int
	Private m_PublicationStatus ' as int
	Private m_HasEnabledSkills	' tinyint
	Private m_HasEnabledMemberSkills ' tinyint

	Private m_SQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	Private m_Error	'as int
	
	Private CLASS_NAME	'as string
	
	Public Property Get ProgramID() 'As long int
		ProgramID = m_ProgramID
	End Property

	Public Property Let ProgramID(val) 'As long int
		m_ProgramID = val
	End Property
	
	Public Property Get ClientID() 'As long int
		ClientID = m_ClientID
	End Property

	Public Property Let ClientID(val) 'As long int
		m_ClientID = val
	End Property
	
	Public Property Get ClientName()
		ClientName = m_ClientName
	End Property
	
	Public Property Get ProgramName() 'As string
		ProgramName = m_ProgramName
	End Property

	Public Property Let ProgramName(val) 'As string
		m_ProgramName = val
	End Property
	
	Public Property Get ProgramDesc() 'As string
		ProgramDesc = m_ProgramDesc
	End Property

	Public Property Let ProgramDesc(val) 'As string
		m_ProgramDesc = val
	End Property
	
	Public Property Get EnrollmentType() 'As small int
		EnrollmentType = m_EnrollmentType
	End Property

	Public Property Let EnrollmentType(val) 'As small int
		m_EnrollmentType = val
	End Property
	
	Public Property Get IsEnabled() 'As small int
		IsEnabled = m_IsEnabled
	End Property

	Public Property Let IsEnabled(val) 'As small int
		m_IsEnabled = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	Public Property Get DateModified() 'As date
		DateModified = m_DateModified
	End Property

	Public Property Get DefaultAvailability() 'As small int
		DefaultAvailability = m_DefaultAvailability
	End Property

	Public Property Let DefaultAvailability(val) 'As small int
		m_DefaultAvailability = val
	End Property
	
	Public Property Get MemberCanEnroll() 'As small int
		MemberCanEnroll = m_MemberCanEnroll
	End Property

	Public Property Let MemberCanEnroll(val) 'As small int
		m_MemberCanEnroll = val
	End Property
	
	Public Property Get MemberCanEditSkills() 'As small int
		MemberCanEditSkills = m_MemberCanEditSkills
	End Property

	Public Property Let MemberCanEditSkills(val) 'As small int
		m_MemberCanEditSkills = val
	End Property
	
	Public Property Get HasSkills()
		HasSkills = False
		If CInt(m_HasSkills) = 1 Then HasSkills = True
	End Property
	
	Public Property Get HasSkillGroups()
		HasSkillGroups = False
		If CInt(m_HasSkillGroups) = 1 Then HasSkillGroups = True
	End Property
	
	Public Property Get HasSchedules()
		HasSchedules = False
		If m_ScheduleCount > 0 Then HasSchedules = True
	End Property
	
	Public Property Get MemberCount()
		MemberCount = m_MemberCount
	End Property
	
	Public Property Get HasMembers()
		HasMembers = False
		If m_MemberCount > 0 Then HasMembers = True
	End Property
	
	Public Property Get ScheduleCount()
		ScheduleCount = m_ScheduleCount
	End Property
	
	Public Property Get EventCount()
		EventCount = m_EventCount
	End Property
	
	Public Property Get HasEvents()
		HasEvents = False
		If m_EventCount > 0 Then HasEvents = True
	End Property
	
	Public Property Get CurrentEventCount()
		CurrentEventCount = m_CurrentEventCount
	End Property
	
	Public Property Get HasCurrentEvents()
		HasCurrentEvents = False
		If m_CurrentEventCount > 0 Then HasCurrentEvents = True
	End Property
	
	Public Property Get PublishStatus()
		PublishStatus = m_PublicationStatus
	End Property
	
	Public Property Get HasEnabledSkills()
		HasEnabledSkills = True
		If m_HasEnabledSkills = 0 Then HasEnabledSkills = False
	End Property
	
	Public Property Get HasEnabledMemberSkills()
		HasEnabledMemberSkills = True
		If m_HasEnabledMemberSkills = 0 Then HasEnabledMemberSkills = False
	End Property
	
	Private Sub Class_Initialize()
		m_ProgramID = 0
		m_ClientID = 0
		m_ProgramName = ""
		m_ProgramDesc = ""
		m_EnrollmentType = 0
		m_DateCreated = ""
		m_DateModified = ""
		m_DefaultAvailability = 0
	
		m_Error = 0
		m_SQL = Application.Value("CNN_STR")
		CLASS_NAME = "cProgram"
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
	
	Public Function PublishEvents(outError)
		Dim cmd
		
		If Len(m_programId) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".PublishEvents()", "Required parameter ProgramId not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_programPublishEventsByProgramID"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, m_programId)
		cmd.Parameters.Append cmd.CreateParameter("@PublishDate", adDate, adParamInput, 0, Now())
		
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Function
	
	Public Function ScheduleBuildList()
		If Len(m_programId) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ScheduleBuildList()", "Required parameter ProgramId not provided.")

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

		m_cnn.up_scheduleGetScheduleBuildForProgramID CLng(m_programId), m_rs
		If Not m_rs.EOF Then ScheduleBuildList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function FileList()
		If Len(m_ProgramID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".FileList(): ", "Required parameter ProgramID not provided.")  'm_iProgramID Required.

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-FileID 1-FriendlyName 2-FileExtension 3-FileName

		m_cnn.up_filesGetFileListByProgramID CLng(m_programId), m_rs
		If Not m_rs.EOF Then FileList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function MemberList()
		If Len(m_ProgramID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".MemberList(): ", "Required parameter ProgramID not provided.")  'm_iProgramID Required.
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If

		' HACK: create new rs because global/class level rs might have been sorted ..
		Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 6-ProgramMemberIsEnabled 7-MemberActiveStatus 8-DateCreated 9-DateModified
		' 10-ProgramMemberID 11-IsApproved 12-HasMissingAvailability 13-Email

		m_cnn.up_programGetProgramMemberList CLng(m_ProgramID), m_rs
		If Not m_rs.EOF Then MemberList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function SkillList(sortColumn)		'as array
		If Len(m_ProgramID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".SkillList(): ", "Required parameter ProgramID not provided.")  'm_iProgramID Required.
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		m_rs.CursorLocation = adUseClient	' makes sortable
		
		' 0-SkillID 1-SkillName 2-SkillDesc 3-IsSkillEnabled 4-SkillGroupID 5-GroupName
		' 6-GroupDesc 7-IsGroupEnabled 8-LastModified 9-DateCreated 10-ActiveMemberCount
		
		m_cnn.up_skillGetSkillListByProgramID CLng(m_ProgramID), m_rs
		
		m_rs.Sort = sortColumn
		If Not m_rs.EOF Then SkillList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function SkillGroupList()
		If Len(m_ProgramID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".SkillGroupList(): ", "Required parameter ProgramID not provided.")  'm_iProgramID Required.
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-SkillGroupID 1-GroupName 2-GroupDesc 3-IsEnabled 4-AllowMultiple 5-DateModified 6-DateCreated
		
		m_cnn.up_skillGetSkillGroupDetailsByProgramID CLng(m_ProgramID), m_rs
		If Not m_rs.EOF Then SkillGroupList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function ScheduleList()
		If Len(m_ProgramID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".ScheduleList(): ", "Required parameter ProgramID not provided.")  'm_iProgramID Required.
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-ScheduleID 1-ScheduleName 2-ScheduleDesc 3-DateCreated 4-DateModified 5-IsVisible
		' 6-DatePublished 7-HasUnpublishedChanges 8-EventCount
		
		m_cnn.up_scheduleGetScheduleDetailsByProgramID CLng(m_ProgramID), m_rs
		If Not m_rs.EOF Then ScheduleList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function EventList(sortColumn)
		If Len(m_ProgramID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".EventList(): ", "Required parameter ProgramID not provided.")  'm_iProgramID Required.
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		m_rs.CursorLocation = adUseClient	' makes sortable
		
		' 0-EventID 1-EventName 2-EventDate 3-EventNote 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified
		' 8-ProgramID 9-ProgramName 10-ScheduleID 11-ScheduleName 12-ScheduleEventCount
		' 13-ScheduleIsVisible 14-HasUnpublishedChanges 15-EventFileCount 16-ScheduledMemberCount

		m_cnn.up_eventGetEventListForProgram CLng(m_ProgramID), m_rs
		
		m_rs.Sort = sortColumn
		If Not m_rs.EOF Then EventList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close()
	End Function
	
	Public Function Load() 'As Boolean
		If Len(m_ProgramID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load(): ", "Required parameter ProgramID not provided.")  'm_iProgramID Required.
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_programGetProgramDetails CLng(m_ProgramID), m_rs
		If Not m_rs.EOF Then
			m_ClientID = m_rs("ClientID").Value
			m_ClientName = m_rs("ClientName").Value
			m_ProgramName = m_rs("ProgramName").Value
			m_ProgramDesc = m_rs("ProgramDesc").Value
			m_EnrollmentType = m_rs("EnrollmentType").Value
			m_IsEnabled = m_rs("IsEnabled").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateModified = m_rs("DateModified").Value
			m_DefaultAvailability = m_rs("DefaultAvailability").Value
			m_HasSkills = m_rs("HasSkills").Value
			m_HasSkillGroups = m_rs("HasSkillGroups").Value
			m_MemberCount = m_rs("MemberCount").Value
			m_ScheduleCount = m_rs("ScheduleCount").Value
			m_EventCount = m_rs("EventCount").Value
			m_CurrentEventCount = m_rs("CurrentEventCount").Value
			m_MemberCanEnroll = m_rs("MemberCanEnroll").Value
			m_MemberCanEditSkills = m_rs("MemberCanEditSkills").Value
			m_PublicationStatus = m_rs("PublicationStatus").Value
			m_HasEnabledSkills = m_rs("HasEnabledSkills").Value
			m_HasEnabledMemberSkills = m_rs("HasEnabledMemberSkills").Value
		
			Load = True
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd
		
		m_DateCreated = Now()
		m_DateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_programInsertProgram"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, m_ClientID)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramName", adVarChar, adParamInput, 100, m_ProgramName)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramDesc", adVarChar, adParamInput, 2000, m_ProgramDesc)
		cmd.Parameters.Append cmd.CreateParameter("@EnrollmentType", adUnsignedTinyInt, adParamInput, 0, m_EnrollmentType)
		cmd.Parameters.Append cmd.CreateParameter("@IsEnabled", adUnsignedTinyInt, adParamInput, 0, m_IsEnabled)
		cmd.Parameters.Append cmd.CreateParameter("@DefaultAvailability", adUnsignedTinyInt, adParamInput, 0, m_DefaultAvailability)
		cmd.Parameters.Append cmd.CreateParameter("@MemberCanEnroll", adUnsignedTinyInt, adParamInput, 0, m_MemberCanEnroll)
		cmd.Parameters.Append cmd.CreateParameter("@MemberCanEditSkills", adUnsignedTinyInt, adParamInput, 0, m_MemberCanEditSkills)
		cmd.Parameters.Append cmd.CreateParameter("@NewProgramID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		m_ProgramID = cmd.Parameters("@NewProgramID").Value
		
		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		m_DateModified = Now()
		
		If Len(m_ProgramID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save(): ", "Required parameter ProgramID not provided.")  'm_iProgramID Required.

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_programUpdateProgram"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, m_ProgramID)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, m_ClientID)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramName", adVarChar, adParamInput, 100, m_ProgramName)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramDesc", adVarChar, adParamInput, 2000, m_ProgramDesc)
		cmd.Parameters.Append cmd.CreateParameter("@EnrollmentType", adUnsignedTinyInt, adParamInput, 0, m_EnrollmentType)
		cmd.Parameters.Append cmd.CreateParameter("@IsEnabled", adUnsignedTinyInt, adParamInput, 0, m_IsEnabled)
		cmd.Parameters.Append cmd.CreateParameter("@DefaultAvailability", adUnsignedTinyInt, adParamInput, 0, m_DefaultAvailability)
		cmd.Parameters.Append cmd.CreateParameter("@MemberCanEnroll", adUnsignedTinyInt, adParamInput, 0, m_MemberCanEnroll)
		cmd.Parameters.Append cmd.CreateParameter("@MemberCanEditSkills", adUnsignedTinyInt, adParamInput, 0, m_MemberCanEditSkills)
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
		
		If Len(m_ProgramID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete(): ", "Required parameter ProgramID not provided.")  'm_iProgramID Required.

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_SQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_programDeleteProgram"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, m_ProgramID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Delete = True
		Else
			Delete = False
		End If
		
		Set cmd = Nothing
	End Function

End Class
%>