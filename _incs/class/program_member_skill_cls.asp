<script runat="server" type="text/vbscript" language="vbscript">

Class cProgramMemberSkill

	Private m_ProgramMemberSkillID		'as long int
	Private m_ProgramMemberID		'as long int
	Private m_SkillID		'as long int
	Private m_DateCreated		'as date

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get ProgramMemberSkillID() 'As long int
		ProgramMemberSkillID = m_ProgramMemberSkillID
	End Property

	Public Property Let ProgramMemberSkillID(val) 'As long int
		m_ProgramMemberSkillID = val
	End Property
	
	Public Property Get ProgramMemberID() 'As long int
		ProgramMemberID = m_ProgramMemberID
	End Property

	Public Property Let ProgramMemberID(val) 'As long int
		m_ProgramMemberID = val
	End Property
	
	Public Property Get SkillID() 'As long int
		SkillID = m_SkillID
	End Property

	Public Property Let SkillID(val) 'As long int
		m_SkillID = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cProgramMemberSkill"
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
	
	Public Sub LoadByMemberIdSkillId(memberId)
		If Len(m_skillId) = 0 Then Call Err.Raise(vbObjectError + 5, CLASS_NAME & ".LoadByMemberIdSkillId():", "Required parameter SkillId not provided.")
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_programGetProgramMemberSkillByMemberIDSkillID CLng(memberId), CLng(m_skillId), m_rs
		If Not m_rs.EOF Then
			m_programMemberSkillID = m_rs("ProgramMemberSkillID").Value
			m_programMemberID = m_rs("ProgramMemberID").Value
			m_skillID = m_rs("SkillID").Value
			m_dateCreated = m_rs("DateCreated").Value
		End If		
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Sub
	
	Public Function Load() 'As Boolean
		If Len(m_programMemberSkillID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter ProgramMemberSkillID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_programGetProgramMemberSkill CLng(m_ProgramMemberSkillID), m_rs
		If Not m_rs.EOF Then
			m_ProgramMemberSkillID = m_rs("ProgramMemberSkillID").Value
			m_ProgramMemberID = m_rs("ProgramMemberID").Value
			m_SkillID = m_rs("SkillID").Value
			m_DateCreated = m_rs("DateCreated").Value
		
			Load = True
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd
		
		m_DateCreated = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_programInsertProgramMemberSkill"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberID", adBigInt, adParamInput, 0, m_ProgramMemberID)
		cmd.Parameters.Append cmd.CreateParameter("@SkillID", adBigInt, adParamInput, 0, m_SkillID)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@NewProgramMemberSkillID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_ProgramMemberSkillID = cmd.Parameters("@NewProgramMemberSkillID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(ProgramMemberSkillID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter ProgramMemberSkillID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_programUpdateProgramMemberSkill"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberSkillID", adBigInt, adParamInput, 0, m_ProgramMemberSkillID)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberID", adBigInt, adParamInput, 0, m_ProgramMemberID)
		cmd.Parameters.Append cmd.CreateParameter("@SkillID", adBigInt, adParamInput, 0, m_SkillID)
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
		
		If Len(ProgramMemberSkillID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter ProgramMemberSkillID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_programDeleteProgramMemberSkill"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberSkillID", adBigInt, adParamInput, 0, m_ProgramMemberSkillID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Delete = True
		Else
			Delete = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Sub DeleteBySkillID(outError)
		Dim cmd
		
		If Len(m_programMemberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".DeleteBySkill():", "Required parameter ProgramMemberID not provided.")
		If Len(m_skillID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".DeleteBySkill():", "Required parameter SkillID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_programDeleteProgramMemberSkillBySkillID"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@SkillID", adBigInt, adParamInput, 0, CLng(m_SkillID))
		cmd.Parameters.Append cmd.CreateParameter("@ProgramMemberID", adBigInt, adParamInput, 0, CLng(m_ProgramMemberID))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
End Class

</script>