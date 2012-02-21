<script runat="server" type="text/vbscript" language="vbscript">

Class cScheduleBuild

	Private m_iEventID		'as long int
	Private m_iProgramMemberSkillID		'as long int
	Private m_iPublishStatus		'as small int
	Private m_dDateCreated		'as date
	Private m_dDateModified		'as date

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	Private m_iError	'as int
	
	Private CLASS_NAME	'as string
	
	Public Property Get EventID() 'As long int
		EventID = m_iEventID
	End Property

	Public Property Let EventID(val) 'As long int
		m_iEventID = val
	End Property
	
	Public Property Get ProgramMemberSkillID() 'As long int
		ProgramMemberSkillID = m_iProgramMemberSkillID
	End Property

	Public Property Let ProgramMemberSkillID(val) 'As long int
		m_iProgramMemberSkillID = val
	End Property
	
	Public Property Get PublishStatus() 'As small int
		PublishStatus = m_iPublishStatus
	End Property

	Public Property Let PublishStatus(val) 'As small int
		m_iPublishStatus = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_dDateCreated
	End Property

	Public Property Get DateModified() 'As date
		DateModified = m_dDateModified
	End Property

	Private Sub Class_Initialize()
'		m_iPublishStatus = 1		' 0=IS_PUBLISHED 1=IS_MARKED_FOR_PUBLISH 2=IS_MARKED_FOR_UNPUBLISH
		m_dDateCreated = ""
		m_dDateModified = ""
	
		m_iError = 0
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cScheduleBuild"
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
	
	Public Function TeamList(scheduleId)
		If Len(scheduleId) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".TeamList();", "Required parameter EventID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-EventID 1-EventName 2-EventDate 3-TimeStart 4-TimeEnd 5-MemberId 6-NameLast 7-NameFirst 
		' 8-MemberIsEnabled 9-ProgramMemberId 10-ProgramMemberIsEnabled 11-IsAvailable 
		' 12-AvailabilityIsViewedByMember 13-SkillGroupId 14-SkillGroupName 15-SkillGroupIsEnabled
		' 16-SkillId 17-SkillName 18-SkillIsEnabled 19-PublishStatus
		
		m_cnn.up_scheduleGetProgramMemberSkillListByScheduleID CLng(scheduleId), m_rs
		If Not m_rs.EOF Then TeamList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function ScheduleList(skillID)
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ScheduleList();", "Required parameter EventID not provided.")
		If Len(skillID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ScheduleList();", "Required parameter SkillID not provided.")
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-ProgramMemberSkillID 1-FullName 2-NameLast 3-NameFirst 4-MemberID 5-AvailabilityNote
		' 6-AvailabilityNoteDate 7-IsAvailable 8-MemberActiveStatus 9-ProgramMemberIsActive
		' 10-IsScheduled 11-PublishStatus

		m_cnn.up_scheduleGetMemberListBySkillIDEventID CLng(skillID), CLng(m_iEventID), m_rs
		If Not m_rs.EOF Then ScheduleList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Load() 'As Boolean
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Load();", "Required parameter EventID not provided.")
		If Len(m_iProgramMemberSkillID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Load();", "Required parameter ProgramMemberSkillID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_scheduleGetScheduleBuild CLng(m_iEventID), CLng(m_iProgramMemberSkillID), m_rs
		If Not m_rs.EOF Then
			m_iEventID = m_rs("EventID").Value
			m_iProgramMemberSkillID = m_rs("ProgramMemberSkillID").Value
			m_iPublishStatus = m_rs("PublishStatus").Value
			m_dDateCreated = m_rs("DateCreated").Value
			m_dDateModified = m_rs("DateModified").Value
		
			Load = True
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Add();", "Required parameter EventID not provided.")
		If Len(m_iProgramMemberSkillID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Add();", "Required parameter ProgramMemberSkillID not provided.")

		m_dDateCreated = Now()
		m_dDateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_scheduleInsertScheduleBuild"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberSkillID", adBigInt, adParamInput, 0, CLng(m_iProgramMemberSkillID))
		cmd.Parameters.Append cmd.CreateParameter("@PublishStatus", adUnsignedTinyInt, adParamInput, 0, CInt(m_iPublishStatus))
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_dDateCreated)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		m_dDateModified = Now()
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Save();", "Required parameter EventID not provided.")
		If Len(m_iProgramMemberSkillID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Save();", "Required parameter ProgramMemberSkillID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_scheduleUpdateScheduleBuild"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberSkillID", adBigInt, adParamInput, 0, CLng(m_iProgramMemberSkillID))
		cmd.Parameters.Append cmd.CreateParameter("@PublishStatus", adUnsignedTinyInt, adParamInput, 0, CInt(m_iPublishStatus))
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_dDateModified)
	
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
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Delete();", "Required parameter EventID not provided.")
		If Len(m_iProgramMemberSkillID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Delete();", "Required parameter ProgramMemberSkillID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_scheduleDeleteScheduleBuild"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberSkillID", adBigInt, adParamInput, 0, CLng(m_iProgramMemberSkillID))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Delete = True
		Else
			Delete = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Function CopyFromEvent(ByVal fromEventID, ByRef outError)
		' accept EventID, copy all rows in ScheduleBuild from that EventID
		' to m_iEventID.
		Dim cmd
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".CopyFromEvent();", "Required parameter EventID not provided.")
		If Len(fromEventID) = 0 Then Call Err.Raise(3 + vbObjectError, CLASS_NAME & ".CopyFromEvent();", "Required parameter FromEventID not provided.")
		
		' don't execute if same dates ..
		If CLng(m_iEventID) = CLng(fromEventID) Then
			outError = -2
			Exit Function
		End If
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_scheduleCopyEventToScheduleBuild"
			.ActiveConnection = m_cnn
		End With

		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		cmd.Parameters.Append cmd.CreateParameter("@EventToCopyID", adBigInt, adParamInput, 0, CLng(fromEventID))
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDBTimeStamp, adParamInput, 0, Now())

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			CopyFromEvent = True
		Else
			CopyFromEvent = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Function ClearAllByEventID(outError)
		' accept eventID, clear all dbo.ScheduleBuild associated with that eventID
		Dim cmd
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ClearAllByEventID();", "Required parameter EventID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_scheduleDeleteScheduleBuildByEventID"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Function
	
	Public Function ClearAllByEventIDSkillID(skillID, outError)
		' accept eventID, skillID and clear dbo.ScheduleBuild associated with that event/skill
		Dim cmd
		
		If Len(m_iEventID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ClearAllByEventIDSkillID();", "Required parameter EventID not provided.")
		If Len(skillID) = 0 Then Call Err.Raise(4 + vbObjectError, CLASS_NAME & ".ClearAllByEventIDSkillID();", "Required parameter SkillID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.commandText = "up_scheduleDeleteScheduleBuildByEventIDSkillID"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		cmd.Parameters.Append cmd.CreateParameter("@SkillID", adBigInt, adParamInput, 0, CLng(skillID))
		
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
		
	End Function 
End Class
</script>