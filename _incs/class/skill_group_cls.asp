<script language="vbscript" type="text/vbscript" runat="server">
Class cSkillGroup

	Private m_SkillGroupID		'as long int
	Private m_ProgramID		'as long int
	Private m_ProgramName	' as string
	Private m_GroupName		'as string
	Private m_GroupDesc		'as string
	Private m_IsEnabled		'as small int
	Private m_AllowMultiple		'as small int
	Private m_DateModified		'as date
	Private m_DateCreated		'as date
	Private m_SkillCount		'as int

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get SkillGroupID() 'As long int
		SkillGroupID = m_SkillGroupID
	End Property

	Public Property Let SkillGroupID(val) 'As long int
		m_SkillGroupID = val
	End Property
	
	Public Property Get ProgramID() 'As long int
		ProgramID = m_ProgramID
	End Property

	Public Property Let ProgramID(val) 'As long int
		m_ProgramID = val
	End Property
	
	Public Property Get ProgramName()
		ProgramName = m_ProgramName
	End Property
	
	Public Property Get GroupName() 'As string
		GroupName = m_GroupName
	End Property

	Public Property Let GroupName(val) 'As string
		m_GroupName = val
	End Property
	
	Public Property Get GroupDesc() 'As string
		GroupDesc = m_GroupDesc
	End Property

	Public Property Let GroupDesc(val) 'As string
		m_GroupDesc = val
	End Property
	
	Public Property Get IsEnabled() 'As small int
		IsEnabled = m_IsEnabled
	End Property

	Public Property Let IsEnabled(val) 'As small int
		m_IsEnabled = val
	End Property
	
	Public Property Get AllowMultiple() 'As small int
		AllowMultiple = m_AllowMultiple
	End Property

	Public Property Let AllowMultiple(val) 'As small int
		m_AllowMultiple = val
	End Property
	
	Public Property Get DateModified() 'As date
		DateModified = m_DateModified
	End Property

	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property
	
	Public Property Get SkillCount() 'as int
		SkillCount = m_SkillCount
	End Property

	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cSkillGroup"
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
	
	Function MemberList()
		If Len(m_SkillGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".MemberList():", "Required parameter SkillGroupID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberActiveStatus 5-ProgramMemberID
		' 6-ProgramMemberIsActive 7-SkillListXmlFragment

		m_cnn.up_skillGetSkillGroupMemberList CLng(m_SkillGroupID), m_rs
		If Not m_rs.EOF Then MemberList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Function UngroupedSkillMemberList(programId)
		If Len(programId) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".UngroupedSkillMemberList():", "Required parameter programId not provided.")

		' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberActiveStatus 5-ProgramMemberID
		' 6-ProgramMemberIsActive 7-SkillListXmlFragment

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-MemberId 1-NameLast 2-NameFirst 3-Email 4-MemberActiveStatus 5-ProgramMemberID
		' 6-ProgramMemberIsActive 7-SkillListXmlFragment

		m_cnn.up_skillGetUngroupedSkillMemberList CLng(programId), m_rs
		If Not m_rs.EOF Then UngroupedSkillMemberList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Load() 'As Boolean
		If Len(SkillGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter SkillGroupID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_skillGetSkillGroupDetails CLng(m_SkillGroupID), m_rs
		If Not m_rs.EOF Then
			m_ProgramID = m_rs("ProgramID").Value
			m_ProgramName = m_rs("ProgramName").Value
			m_GroupName = m_rs("GroupName").Value
			m_GroupDesc = m_rs("GroupDesc").Value
			m_IsEnabled = m_rs("IsEnabled").Value
			m_AllowMultiple = m_rs("AllowMultiple").Value
			m_DateModified = m_rs("DateModified").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_SkillCount = m_rs("SkillCount").Value
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
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_skillInsertSkillGroup"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, m_ProgramID)
		cmd.Parameters.Append cmd.CreateParameter("@GroupName", adVarChar, adParamInput, 100, m_GroupName)
		If Len(m_GroupDesc) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@GroupDesc", adVarChar, adParamInput, 2000, Null)
		Else 
			cmd.Parameters.Append cmd.CreateParameter("@GroupDesc", adVarChar, adParamInput, 2000, m_GroupDesc)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@IsEnabled", adUnsignedTinyInt, adParamInput, 0, m_IsEnabled)
		cmd.Parameters.Append cmd.CreateParameter("@AllowMultiple", adUnsignedTinyInt, adParamInput, 0, m_AllowMultiple)
		cmd.Parameters.Append cmd.CreateParameter("@NewSkillGroupID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_SkillGroupID = cmd.Parameters("@NewSkillGroupID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(SkillGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter SkillGroupID not provided.")

		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_skillUpdateSkillGroup"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@SkillGroupID", adBigInt, adParamInput, 0, m_SkillGroupID)
		cmd.Parameters.Append cmd.CreateParameter("@GroupName", adVarChar, adParamInput, 100, m_GroupName)
		If Len(m_GroupDesc) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@GroupDesc", adVarChar, adParamInput, 2000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@GroupDesc", adVarChar, adParamInput, 2000, m_GroupDesc)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@IsEnabled", adUnsignedTinyInt, adParamInput, 0, m_IsEnabled)
		cmd.Parameters.Append cmd.CreateParameter("@AllowMultiple", adUnsignedTinyInt, adParamInput, 0, m_AllowMultiple)
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
		
		If Len(SkillGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter SkillGroupID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_skillDeleteSkillGroup"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@SkillGroupID", adBigInt, adParamInput, 0, m_SkillGroupID)

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

