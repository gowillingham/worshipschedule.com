<script runat="server" type="text/vbscript" language="vbscript">

Class cProgramMember

	Private m_iProgramMemberID		'as long int
	Private m_iMemberID		'as long int
	Private m_iProgramID		'as long int
	Private m_sProgramName		'as string
	Private m_iEnrollStatusID		'as small int
	Private m_sEnrollStatusText 'as string
	Private m_iIsLeader		'as small int
	Private m_iIsActive		'as small int
	Private m_dLastScheduleView		'as date
	Private m_dDateModified		'as date
	Private m_dDateCreated		'as date
	Private m_ProgramDesc

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset

	Private CLASS_NAME	'as string
	
	Public Property Get ProgramMemberID() 'As long int
		ProgramMemberID = m_iProgramMemberID
	End Property

	Public Property Let ProgramMemberID(val) 'As long int
		m_iProgramMemberID = val
	End Property
	
	Public Property Get MemberID() 'As long int
		MemberID = m_iMemberID
	End Property

	Public Property Let MemberID(val) 'As long int
		m_iMemberID = val
	End Property
	
	Public Property Get ProgramID() 'As long int
		ProgramID = m_iProgramID
	End Property

	Public Property Let ProgramID(val) 'As long int
		m_iProgramID = val
	End Property
	
	Public Property Get ProgramName() 'as str
		ProgramName = m_sProgramName
	End Property
	
	Public Property Get EnrollStatusID() 'As small int
		EnrollStatusID = m_iEnrollStatusID
	End Property

	Public Property Let EnrollStatusID(val) 'As small int
		m_iEnrollStatusID = val
	End Property
	
	Public Property Get EnrollStatusText() 'As str
		EnrollStatusText = m_sEnrollStatusText
	End Property

	Public Property Get IsLeader() 'As small int
		IsLeader = m_iIsLeader
	End Property

	Public Property Let IsLeader(val) 'As small int
		m_iIsLeader = val
	End Property
	
	Public Property Get IsActive() 'As small int
		IsActive = m_iIsActive
	End Property

	Public Property Let IsActive(val) 'As small int
		m_iIsActive = val
	End Property
	
	Public Property Get LastScheduleView() 'As date
		LastScheduleView = m_dLastScheduleView
	End Property

	Public Property Let LastScheduleView(val) 'As date
		m_dLastScheduleView = val
	End Property
	
	Public Property Get DateModified() 'As date
		DateModified = m_dDateModified
	End Property

	Public Property Get DateCreated() 'As date
		DateCreated = m_dDateCreated
	End Property
	
	Public Property Get ProgramDesc()
		ProgramDesc = m_ProgramDesc
	End Property

	Private Sub Class_Initialize()
		m_iEnrollStatusID = 0
		m_iIsLeader = 0
		m_iIsActive = 0
		m_dDateModified = ""
		m_dDateCreated = ""
	
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cProgramMember"
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
	
	Public Function GetMemberList()
		Dim cmd
		If Len(m_iProgramID) = 0 Then Call err.Raise(vbObjectError + 1, CLASS_NAME & ".GetMemberList()", "Missing required parameter ProgramID")
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 
		' 6-ProgramMemberIsEnabled 7-MemberActiveStatus 8-DateCreated 9-DateModified
		' 10-ProgramMemberID 11-IsApproved 12-HasMissingAvailability 13-Email
		
		m_cnn.up_programGetProgramMemberList CLng(m_iProgramID), m_rs
		If Not m_rs.EOF Then GetMemberList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
		Set cmd = Nothing
	End Function
	
	Public Function GetSkillList(sortColumn)
		If Len(m_iProgramMemberID) = 0 Then Call err.Raise(vbObjectError + 2, CLASS_NAME & ".GetSkillList()", "Missing required parameter ProgramMemberID")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.CursorLocation = adUseClient	' makes sortable
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		m_rs.CursorLocation = adUseClient
		
		' 0-SkillID 1-ProgramMemberSkillID 2-ProgramMemberID 3-SkillName 4-SkillDesc 5-SkillGroupName
		' 6-IsProgramMemberSkill 7-DateCreated 8-SkillIsEnabled 9-SkillGroupIsEnabled

		m_cnn.up_programGetProgramMemberSkillList CLng(m_iProgramMemberID), m_rs
		m_rs.Sort = sortColumn
		If Not m_rs.EOF Then GetSkillList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function GetAvailableProgramList()
		If Len(m_iMemberID) = 0 Then Call err.Raise(vbObjectError + 3, CLASS_NAME & ".LoadByMemberProgram()", "Missing required parameter MemberID")

		' 0-ProgramID 1-ProgramName 2-Description 3-EnrollmentType 4-IsEnabled
		' 5-DateCreated 6-DateModified 7-MemberCount 8-SkillCount 9-MemberCanEnroll
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_memberGetAvailableProgramsList CLng(m_iMemberID), m_rs
		If Not m_rs.EOF Then GetAvailableProgramList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function LoadByMemberProgram()
		If Len(m_iMemberID) = 0 Then Call err.Raise(vbObjectError + 3, CLASS_NAME & ".LoadByMemberProgram()", "Missing required parameter MemberID")
		If Len(m_iProgramID) = 0 Then Call err.Raise(vbObjectError + 1, CLASS_NAME & ".LoadByMemberProgram()", "Missing required parameter ProgramID")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		m_cnn.up_memberGetProgramMemberByProgramIDMemberID CLng(m_iProgramID), CLng(m_iMemberID), m_rs
		If Not m_rs.EOF Then
			m_iProgramMemberID = m_rs("ProgramMemberID").Value
			m_iMemberID = m_rs("MemberID").Value
			m_iProgramID = m_rs("ProgramID").Value
			m_sProgramName = m_rs("ProgramName").Value
			m_iEnrollStatusID = m_rs("EnrollStatusID").Value
			m_sEnrollStatusText = m_rs("EnrollStatusText").Value
			m_iIsLeader = m_rs("IsLeader").Value
			m_iIsActive = m_rs("IsActive").Value
			m_dLastScheduleView = m_rs("LastScheduleView").Value
			m_dDateModified = m_rs("DateModified").Value
			m_dDateCreated = m_rs("DateCreated").Value
		
			LoadByMemberProgram = True
		Else
			LoadByMemberProgram = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Load() 'As Boolean
		If Len(m_iProgramMemberID) = 0 Then Call err.Raise(vbObjectError + 2, CLASS_NAME & ".Load()", "Missing required parameter ProgramMemberID")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_memberGetProgramMember CLng(m_iProgramMemberID), m_rs
		If Not m_rs.EOF Then
			m_iProgramMemberID = m_rs("ProgramMemberID").Value
			m_iMemberID = m_rs("MemberID").Value
			m_iProgramID = m_rs("ProgramID").Value
			m_sProgramName = m_rs("ProgramName").Value
			m_iEnrollStatusID = m_rs("EnrollStatusID").Value
			m_sEnrollStatusText = m_rs("EnrollStatusText").Value
			m_iIsLeader = m_rs("IsLeader").Value
			m_iIsActive = m_rs("IsActive").Value
			m_dLastScheduleView = m_rs("LastScheduleView").Value
			m_dDateModified = m_rs("DateModified").Value
			m_dDateCreated = m_rs("DateCreated").Value
			m_ProgramDesc = m_rs("ProgramDesc").Value
		
			Load = True
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd

		If Len(m_iMemberID) = 0 Then Call err.Raise(vbObjectError + 3, CLASS_NAME & ".Add()", "Missing required parameter MemberID")
		If Len(m_iProgramID) = 0 Then Call err.Raise(vbObjectError + 1, CLASS_NAME & ".Add()", "Missing required parameter ProgramID")

		m_dDateCreated = Now()
		m_dDateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_memberInsertProgramMember"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_iMemberID))
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, CLng(m_iProgramID))
		cmd.Parameters.Append cmd.CreateParameter("@EnrollStatusID", adUnsignedTinyInt, adParamInput, 0, CInt(m_iEnrollStatusID))
		cmd.Parameters.Append cmd.CreateParameter("@IsActive", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsActive))
		cmd.Parameters.Append cmd.CreateParameter("@IsLeader", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsLeader))
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, CDate(m_dDateCreated))
		cmd.Parameters.Append cmd.CreateParameter("@NewProgramMemberID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_iProgramMemberID = cmd.Parameters("@NewProgramMemberID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		m_dDateModified = Now()
		
		If Len(m_iProgramMemberID) = 0 Then Call err.Raise(vbObjectError + 2, CLASS_NAME & ".Save()", "Missing required parameter ProgramMemberID")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_memberUpdateProgramMember"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberID", adBigInt, adParamInput, 0, CLng(m_iProgramMemberID))
		cmd.Parameters.Append cmd.CreateParameter("@EnrollStatusID", adUnsignedTinyInt, adParamInput, 0, CInt(m_iEnrollStatusID))
		cmd.Parameters.Append cmd.CreateParameter("@IsActive", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsActive))
		cmd.Parameters.Append cmd.CreateParameter("@IsLeader", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsLeader))
		If Len(m_dLastScheduleView & "") = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@LastScheduleView", adDate, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@LastScheduleView", adDate, adParamInput, 0, CDate(m_dLastScheduleView))
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
		
		If Len(m_iProgramMemberID) = 0 Then Call err.Raise(vbObjectError + 2, CLASS_NAME & ".Delete()", "Missing required parameter ProgramMemberID")
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_memberDeleteProgramMember"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberID", adBigInt, adParamInput, 0, m_iProgramMemberID)

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
