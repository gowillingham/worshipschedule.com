<script language="vbscript" type="text/vbscript" runat="server">
Class cSkill

	Private m_iSkillID		'as long int
	Private m_iProgramID		'as long int
	Private m_sProgramName		'as string
	Private m_iSkillGroupID		'as long int
	Private m_sSkillName		'as string
	Private m_sSkillDesc		'as string
	Private m_iIsEnabled		'as small int
	Private m_dDateModified		'as date
	Private m_dDateCreated		'as date
	Private m_sGroupName		'as string
	Private m_sGroupDesc		'as string
	Private m_iSkillGroupIsEnabled ' as tinyint
	Private m_iMembercount		'as int

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	Private CLASS_NAME	'as string
	
	Public Property Get SkillID() 'As long int
		SkillID = m_iSkillID
	End Property

	Public Property Let SkillID(val) 'As long int
		m_iSkillID = val
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
	
	Public Property Get SkillGroupID() 'As long int
		SkillGroupID = m_iSkillGroupID
	End Property

	Public Property Let SkillGroupID(val) 'As long int
		m_iSkillGroupID = val
	End Property
	
	Public Property Get SkillName() 'As string
		SkillName = m_sSkillName
	End Property

	Public Property Let SkillName(val) 'As string
		m_sSkillName = val
	End Property
	
	Public Property Get SkillDesc() 'As string
		SkillDesc = m_sSkillDesc
	End Property

	Public Property Let SkillDesc(val) 'As string
		m_sSkillDesc = val
	End Property
	
	Public Property Get IsEnabled() 'As small int
		IsEnabled = m_iIsEnabled
	End Property

	Public Property Let IsEnabled(val) 'As small int
		m_iIsEnabled = val
	End Property
	
	Public Property Get GroupName() 'As str
		GroupName = m_sGroupName
	End Property

	Public Property Get GroupDesc() 'As str
		GroupDesc = m_sGroupDesc
	End Property
	
	Public Property Get SkillGroupIsEnabled()
		SkillGroupIsEnabled = m_iSkillGroupIsEnabled
	End Property

	Public Property Get LastModified() 'As date
		LastModified = m_dLastModified
	End Property

	Public Property Let LastModified(val) 'As date
		m_dLastModified = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_dDateCreated
	End Property
	
	Public Property Get MemberCount()
		MemberCount = m_iMemberCount
	End Property

	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cSkill"
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
	
	Public Function ScheduleInfoList()
		If Len(m_iSkillID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".ScheduleInfoList(): ", "Required parameter SkillID not provided.")  
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-MemberId 1-NameLast 2-NameFirst 3-MemberIsEnabled 4-ProgramMemberIsEnabled 5-IsAvailable 6-AvailabilityNote
		' 7-AvailabilityIsViewedByMember 8-AvailabilityDateModified 9-PublishStatus 10-EventID 11-EventName 
		' 12-EventDescription 13-EventDate 14-TimeStart 15-TimeEnd 16-ScheduleID 17-ScheduleName 
		' 18-IsVisible 19-EventAvailabilityID
	
		m_cnn.up_scheduleGetScheduleInfoForSkillId CLng(m_iSkillID), m_rs
		If Not m_rs.EOF Then ScheduleInfoList = m_rs.GetRows()
	
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function MemberList()
		If Len(m_iSkillID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".MemberList(): ", "Required parameter SkillID not provided.")  
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberSkillID 5-ProgramMemberID
		' 6-IsApproved 7-ProgramMemberIsActive 8-Email 9-IsMissingAvailabilityInfoForSkill
		
		m_cnn.up_skillGetSkillMemberList CLng(m_iSkillID), m_rs
		If Not m_rs.EOF Then MemberList = m_rs.GetRows()
	
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function SkillList(sortColumn)
		If Len(m_iProgramID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".SkillList(): ", "Required parameter ProgramID not provided.")  
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		m_rs.CursorLocation = adUseClient	' makes sortable

		' 0-SkillID 1-SkillName 2-SkillDesc 3-SkillIsEnabled 4-SkillGroupID 5-GroupName
		' 6-GroupDesc 7-GroupIsEnabled 8-DateGroupModified 9-DateGroupCreated 10-DateSkillModified
		' 11-DateSkillCreated 12-MemberCount

		m_cnn.up_skillGetSkillList CLng(m_iProgramID), m_rs

		m_rs.Sort = sortColumn
		If Not m_rs.EOF Then SkillList = m_rs.GetRows()
	
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function SkillGroupList()
		If Len(m_iProgramID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".SkillGroupList(): ", "Required parameter ProgramID not provided.")  

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-SkillGroupID 1-GroupName 2-GroupDesc 3-IsEnabled 4-AllowMultiple 5-DateModified
		' 6-DateCreated 7-SkillCount 8-SkillListXMLFragment

		m_cnn.up_skillGetSkillGroupDetailsByProgramID CLng(m_iProgramID), m_rs
		If Not m_rs.EOF Then SkillGroupList = m_rs.GetRows()
	
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Load() 'As Boolean
		If Len(m_iSkillID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load(): ", "Required parameter SkillID not provided.")  
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_skillGetSkillDetails CLng(m_iSkillID), m_rs
		If Not m_rs.EOF Then
			m_iProgramID = m_rs("ProgramID").Value
			m_sProgramName = m_rs("ProgramName").Value
			m_iSkillGroupID = m_rs("SkillGroupID").Value
			m_sSkillName = m_rs("SkillName").Value
			m_sSkillDesc = m_rs("SkillDesc").Value
			m_iIsEnabled = m_rs("IsEnabled").Value
			m_dDateModified = m_rs("LastModified").Value
			m_dDateCreated = m_rs("DateCreated").Value
			m_sGroupName = m_rs("GroupName").Value
			m_sGroupDesc = m_rs("GroupDesc").Value
			m_iSkillGroupIsEnabled = m_rs("SkillGroupIsEnabled").Value
			m_iMemberCount = m_rs("MemberCount").Value
		
			Load = True
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd

		If Len(m_iProgramID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Add(): ", "Required parameter ProgramID not provided.")  
		
		m_dDateCreated = Now()
		m_dDateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_skillInsertSkill"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0,CLng(m_iProgramID))
		If Len(m_iSkillGroupID) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@SkillGroupID", adBigInt, adParamInput, 0, CLng(m_iSkillGroupID))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@SkillGroupID", adBigInt, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@SkillName", adVarChar, adParamInput, 100, m_sSkillName)
		cmd.Parameters.Append cmd.CreateParameter("@SkillDesc", adVarChar, adParamInput, 2000, m_sSkillDesc)
		cmd.Parameters.Append cmd.CreateParameter("@IsEnabled", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsEnabled))
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_dDateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@NewSkillID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_iSkillID = cmd.Parameters("@NewSkillID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		m_dDateModified = Now()
		
		If Len(m_iSkillID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save(): ", "Required parameter SkillID not provided.")  

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_skillUpdateSkill"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@SkillID", adBigInt, adParamInput, 0, m_iSkillID)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, m_iProgramID)
		If Len(m_iSkillGroupID) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@SkillGroupID", adBigInt, adParamInput, 0, CLng(m_iSkillGroupID))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@SkillGroupID", adBigInt, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@SkillName", adVarChar, adParamInput, 100, m_sSkillName)
		cmd.Parameters.Append cmd.CreateParameter("@SkillDesc", adVarChar, adParamInput, 2000, m_sSkillDesc)
		cmd.Parameters.Append cmd.CreateParameter("@IsEnabled", adUnsignedTinyInt, adParamInput, 0, m_iIsEnabled)
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
		
		If Len(m_iSkillID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete(): ", "Required parameter SkillID not provided.")  

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_skillDeleteSkill"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@SkillID", adBigInt, adParamInput, 0, CLng(m_iSkillID))

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
