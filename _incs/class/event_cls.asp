<script runat="server" language="vbscript" type="text/vbscript">
Class cEvent

	Private m_iEventID		'as long int
	Private m_iScheduleID		'as long int
	Private m_sEventName		'as string
	Private m_sEventNote		'as string
	Private m_dEventDate		'as date
	Private m_dTimeStart		'as date
	Private m_dTimeEnd		'as date
	Private m_sFileList		'as string
	Private m_aFileDetailsList	'as array
	Private m_iHasFiles			'as tinyint
	Private m_dDateModified		'as date
	Private m_dDateCreated		'as date
	Private m_ScheduleID		'as long int
	Private m_ProgramID			'as long int
	Private m_ProgramName		'as string
	Private m_ScheduleName		'as string
	Private m_ScheduledMemberCount	' as int
	Private m_htmlBackgroundColor
	Private m_hasUnpublishedChanges		' as tinyint
	
	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	Private m_iError	'as int
	Private m_bIsDirty  'as bool
	
	Private IDS_AS_STRING		'int
	Private DETAILS_AS_ARRAY	'int
	Private CLASS_NAME	'as string
	
	Public Property Get EventID() 'As long int
		EventID = m_iEventID
	End Property

	Public Property Let EventID(val) 'As long int
		m_iEventID = val
		m_bIsDirty = True
	End Property
	
	Public Property Get ScheduleID() 'As long int
		ScheduleID = m_iScheduleID
	End Property

	Public Property Let ScheduleID(val) 'As long int
		m_iScheduleID = val
		m_bIsDirty = True
	End Property
	
	Public Property Get EventName() 'As string
		EventName = m_sEventName
	End Property

	Public Property Let EventName(val) 'As string
		m_sEventName = val
		m_bIsDirty = True
	End Property
	
	Public Property Get EventNote() 'As string
		EventNote = m_sEventNote
	End Property

	Public Property Let EventNote(val) 'As string
		m_sEventNote = val
		m_bIsDirty = True
	End Property
	
	Public Property Get EventDate() 'As date
		EventDate = m_dEventDate
	End Property

	Public Property Let EventDate(val) 'As date
		m_dEventDate = val
		m_bIsDirty = True
	End Property
	
	Public Property Get TimeStart() 'As date
		TimeStart = m_dTimeStart
	End Property

	Public Property Let TimeStart(val) 'As date
		m_dTimeStart = val
		m_bIsDirty = True
	End Property
	
	Public Property Get TimeEnd() 'As date
		TimeEnd = m_dTimeEnd
	End Property

	Public Property Let TimeEnd(val) 'As date
		m_dTimeEnd = val
		m_bIsDirty = True
	End Property
	
	Public Property Get DateModified() 'As date
		DateModified = m_dDateModified
	End Property

	Public Property Get DateCreated() 'As date
		DateCreated = m_dDateCreated
	End Property
	
	Public Property Get HtmlBackgroundColor()
		HtmlBackgroundColor =m_htmlBackgroundColor
	End Property
	
	Public Property Get HasFiles() 'as boolean
		HasFiles = False
		If CInt(m_iHasFiles) = 1 Then
			HasFiles = True
		End If
	End Property
	
	Public Property Get HasScheduledMembers() 'as bool
		HasScheduledMembers = False
		If ScheduledMemberCount > 0 Then HasScheduledMembers = True
	End Property
	
	Public Property Get ScheduledMemberCount() ' as int
		ScheduledMemberCount = m_ScheduledMemberCount
	End Property
	
	Public Property Get HasUnpublishedChanges() ' tinyint
		HasUnpublishedChanges = m_hasUnpublishedChanges
	End Property
	
	Public Property Get FileList() 'As string
		' returns comma delim list of fileIDs
		
		FileList = m_sFileList
	End Property
	
	Public Property Let FileList(val) 'As string
		' takes comma-delim list of fileIDs
		' and cleans them of spaces ..
		
		m_sFileList = Replace(val, " ", "")
	End Property
	
	Public Property Get FileDetailsList() 'as array
		' returns list of file details for eventID
		
		FileDetailsList = GetFileList(DETAILS_AS_ARRAY)
	End Property
	
	Public Property Get IsDirty()
		IsDirty = m_bIsDirty
	End Property
	
	Public Property Get ProgramID()
		ProgramID = m_ProgramID
	End Property

	Public Property Get ProgramName()
		ProgramName = m_ProgramName
	End Property

	Public Property Get ScheduleName()
		ScheduleName = m_ScheduleName
	End Property
	
	Private Sub Class_Initialize()
		m_sEventName = ""
		m_sEventNote = ""
		m_dEventDate = ""
		m_dTimeStart = ""
		m_dTimeEnd = ""
		m_sFileList = ""
		m_dDateModified = ""
		m_dDateCreated = ""
	
		m_iError = 0
		m_sSQL = Application.Value("CNN_STR")
		m_bIsDirty = False
		
		CLASS_NAME = "cEvent"
		IDS_AS_STRING = 0
		DETAILS_AS_ARRAY = 1
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
	
	Public Function EventTeamDetailsList()
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".EventTeamDetailsList();", "EventID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		if Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-MemberID 1-NameLast 2-NameFirst 3-IsMemberEnabled 4-IsProgramMemberEnabled
		' 5-SkillListingXmlFragment 6-IsAvailable 7-IsAvailabilityViewedByMember
		
		m_cnn.up_scheduleGetEventTeamDetailsByEventID CLng(m_iEventID), m_rs
		If Not m_rs.EOF Then EventTeamDetailsList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function

   Public Function XmlScheduleList()
      If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".XmlScheduleList();", "EventID not provided.")
      
      If Not IsObject(m_cnn) Then 
         Set m_cnn = Server.CreateObject("ADODB.Connection")
         m_cnn.Open m_sSQL
      End If
      if Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
      
      m_cnn.up_scheduleGetXmlSchedulePublish CLng(m_iEventID), m_rs
	  If Not m_rs.EOF Then XmlScheduleList = m_rs.GetRows()
      
      If m_rs.State = adStateOpen Then m_rs.Close
   End Function
   
	Public Function ScheduledMemberList() ' as array
		' pulls from dbo.ScheduleBuild
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ScheduledMemberList();", "EventID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-ProgramMemberSkillID 1-SkillID 2-SkillName 3-SkillIsEnabled 4-SkillGroupID
		' 5-SkillGroupName 6-SkillGroupIsEnabled 7-NameLast 8-NameFirst 9-MemberActiveStatus
		' 10-IsAvailable 11-MemberNote 12-AvailabilityDateModified 13-PublishStatus
		' 14-ProgramMemberIsActive
		
		m_cnn.up_scheduleGetScheduleBuildByEventID CLng(eventID), m_rs
		If Not m_rs.EOF Then ScheduledMemberList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function AvailableMemberList()
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".AvailableMemberList();", "EventID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-ProgramMemberSkillID 1-NameLast 2-NameFirst 3-MemberEnabled 4-ProgramMemberEnabled 5-SkillName
		' 6-SkillGroupName 7-SkillEnabled 8-SkillGroupEnabled 9-IsAvailable 10-AvailabilityNote 
		' 11-IsViewedByMember 12-DateAvailabilityModified 13-ProgramMemberID 14-MemberID 15-SkillID 
		' 16-SkillGroupID
	
		m_cnn.up_eventGetAvailableMemberSkillsByEventID CLng(m_iEventID), m_rs
		If Not m_rs.EOF Then AvailableMemberList = m_rs.GetRows

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function MemberList() 'as array
		' pulls members from dbo.SchedulePublish
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".MemberList();", "EventID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-SkillName 5-IsAvailable 
		' 6-IsViewedByMember 7-SchedulePublishID 8-email 9-ProgramMemberIsActive
	
		m_cnn.up_eventGetMemberList CLng(m_iEventID), m_rs
		If Not m_rs.EOF Then MemberList = m_rs.GetRows

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Private Function GetFileList(returnType)
		Dim arr, i, sDelim, str
		sDelim = ","
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-FileID 1-FileName 2-FriendlyName 3-Description 4-FileExtension 5-FileSize
		' 6-DownloadCount 7-IsPublic

		m_cnn.up_filesGetFileDetailsByEventID CLng(m_iEventID), m_rs
		If Not m_rs.EOF Then
			arr = m_rs.GetRows()
			Select Case returnType
				Case DETAILS_AS_ARRAY
					GetFileList = arr
				Case IDS_AS_STRING
					If IsArray(arr) Then
						For i = 0 To UBound(arr,2)
							str = str & arr(0,i) & sDelim
						Next
						If Len(str) > 0 Then str = Left(str, Len(str) - Len(sDelim))
					End If
					GetFileList = str
				Case Else
					' do nothing
			End Select
		End If

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function ClearFiles(outError)
		Dim cmd
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".List():", "Required parameter EventID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_filesDeleteEventFilesByEventID"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
	End Function
	
	Public Function Publish(memberName, outError)
		Dim cmd
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".Publish():", "Required parameter EventID not provided.")

		Dim modifiedDate		: modifiedDate = Now()
		If Len(memberName) = 0 Then memberName = "Unknown"
			
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_schedulePublishByEventID"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		cmd.Parameters.Append cmd.CreateParameter("@DatePublished", adDBTimeStamp, adParamInput, 0, modifiedDate)
		cmd.Parameters.Append cmd.CreateParameter("@Publisher", adVarChar, adParamInput, 102, CStr(memberName))
		
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Function
	
	Public Function Load() 'As Boolean
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Load();", "EventID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_eventGetEvent CLng(m_iEventID), m_rs
		If Not m_rs.EOF Then
			m_iEventID = m_rs("EventID").Value
			m_iScheduleID = m_rs("ScheduleID").Value
			m_sEventName = m_rs("EventName").Value
			m_sEventNote = m_rs("EventNote").Value
			m_dEventDate = m_rs("EventDate").Value
			m_dTimeStart = m_rs("TimeStart").Value
			m_dTimeEnd = m_rs("TimeEnd").Value
			m_dDateModified = m_rs("DateModified").Value
			m_dDateCreated = m_rs("DateCreated").Value
			m_ProgramID = m_rs("ProgramID").Value
			m_ProgramName = m_rs("ProgramName").Value
			m_ScheduleName = m_rs("ScheduleName").Value
			m_iHasFiles = m_rs("HasFiles").Value
			m_ScheduledMemberCount = m_rs("ScheduledMemberCount").Value
			m_htmlBackgroundColor = m_rs("htmlBackgroundColor").Value
			m_HasUnpublishedChanges = m_rs("HasUnpublishedChanges").Value
			
			If m_rs.State = adStateOpen Then m_rs.Close
			
			' load the file list if HasFiles
			If m_iHasFiles = 1 Then
				m_sFileList = GetFileList(IDS_AS_STRING)
			End If
		
			Load = True
			m_bIsDirty = False
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd

		If Len(m_iScheduleID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Add();", "ScheduleID not provided.")
		m_dDateCreated = Now()
		m_dDateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_eventInsert"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleID", adBigInt, adParamInput, 0, CLng(m_iScheduleID))
		cmd.Parameters.Append cmd.CreateParameter("@EventName", adVarChar, adParamInput, 200, CStr(m_sEventName))
		If Len(m_sEventNote) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@EventNote", adVarChar, adParamInput, 1000, CStr(m_sEventNote))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@EventNote", adVarChar, adParamInput, 1000, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@EventDate", adDate, adParamInput, 0, CDate(m_dEventDate))
		If Len(m_dTimeStart) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@TimeStart", adDate, adParamInput, 0, CDate(m_dEventDate & " " & m_dTimeStart))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@TimeStart", adDate, adParamInput, 0, Null)
		End If
		If Len(m_dTimeEnd) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@TimeEnd", adDate, adParamInput, 0, CDate(m_dEventDate & " " & m_dTimeEnd))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@TimeEnd", adDate, adParamInput, 0, Null)
		End If
		If Len(m_sFileList) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@FileList", adVarChar, adParamInput, 8000, CStr(m_sFileList))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@FileIDList", adVarChar, adParamInput, 8000, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, CDate(m_dDateCreated))
		cmd.Parameters.Append cmd.CreateParameter("@NewEventID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_iEventID = cmd.Parameters("@NewEventID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		m_dDateModified = Now()
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Save();", "EventID not provided.")
		If Len(m_iScheduleID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Save();", "ScheduleID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_eventUpdate"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleID", adBigInt, adParamInput, 0, CLng(m_iScheduleID))
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		cmd.Parameters.Append cmd.CreateParameter("@EventName", adVarChar, adParamInput, 200, CStr(m_sEventName))
		If Len(m_sEventNote) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@EventNote", adVarChar, adParamInput, 1000, CStr(m_sEventNote))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@EventNote", adVarChar, adParamInput, 1000, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@EventDate", adDate, adParamInput, 0, CDate(m_dEventDate))
		If Len(m_dTimeStart) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@TimeStart", adDate, adParamInput, 0, CDate(m_dEventDate & " " & m_dTimeStart))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@TimeStart", adDate, adParamInput, 0, Null)
		End If
		If Len(m_dTimeEnd) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@TimeEnd", adDate, adParamInput, 0, CDate(m_dEventDate & " " & m_dTimeEnd))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@TimeEnd", adDate, adParamInput, 0, Null)
		End If
		If Len(m_sFileList) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@FileList", adVarChar, adParamInput, 8000, CStr(m_sFileList))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@FileIDList", adVarChar, adParamInput, 8000, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, CDate(m_dDateModified))
	
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
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Delete();", "EventID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_eventDelete"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))

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
</script>