<script type="text/vbscript" runat="server" language="vbscript">
Class cEmailGroup

	Private m_EmailGroupID		'as long int
	Private m_Name		'as string
	Private m_Description		'as string
	Private m_MemberID		'as long int
	Private m_DateCreated		'as date
	Private m_DateModified		'as date
	Private m_MemberCount		' int

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get EmailGroupID() 'As long int
		EmailGroupID = m_EmailGroupID
	End Property

	Public Property Let EmailGroupID(val) 'As long int
		m_EmailGroupID = val
	End Property
	
	Public Property Get Name() 'As string
		Name = m_Name
	End Property

	Public Property Let Name(val) 'As string
		m_Name = val
	End Property
	
	Public Property Get Description() 'As string
		Description = m_Description
	End Property

	Public Property Let Description(val) 'As string
		m_Description = val
	End Property
	
	Public Property Get MemberID() 'As long int
		MemberID = m_MemberID
	End Property

	Public Property Let MemberID(val) 'As long int
		m_MemberID = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	Public Property Get DateModified() 'As date
		DateModified = m_DateModified
	End Property

	Public Property Get MemberCount() 'As date
		MemberCount = m_MemberCount
	End Property
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cEmailGroup"
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
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".List():", "Required parameter MemberID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_emailGetEmailGroupList CLng(m_MemberID), m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function MemberList()
		If Len(m_EmailGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".MemberList():", "Required parameter EmailGroupID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-EmailGroupMemberID 1-EmailGroupId 2-MemberID 3-Email 4-NameLast 5-NameFirst 6-DateCreated 7-MemberActiveStatus
		
		m_cnn.up_emailGetEmailGroupMemberList CLng(m_EmailGroupID), m_rs
		If Not m_rs.EOF Then MemberList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Clear(outError)
		Dim cmd
		
		If Len(m_EmailGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Clear():", "Required parameter EmailGroupID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailClearEmailGroupMembers"
			.ActiveConnection = m_cnn
		End With
	
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EmailGroupID", adBigInt, adParamInput, 0, CLng(m_EmailGroupID))
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Function
	
	Public Sub Load() 'As Boolean
		If Len(m_EmailGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter EmailGroupID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_emailGetEmailGroup CLng(m_EmailGroupID), m_rs
		If Not m_rs.EOF Then
			m_EmailGroupID = m_rs("EmailGroupID").Value
			m_Name = m_rs("Name").Value
			m_Description = m_rs("Description").Value
			m_MemberID = m_rs("MemberID").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateModified = m_rs("DateModified").Value
			m_MemberCount = m_rs("MemberCount").Value
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Sub
	
	Public Sub Add(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Add():", "Required parameter MemberID not provided.")
		
		m_DateCreated = Now()
		m_DateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailInsertEmailGroup"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 200, m_Name)
		If Len(m_Description) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@Description", adVarChar, adParamInput, 1000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@Description", adVarChar, adParamInput, 1000, m_Description)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, m_MemberID)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_DateModified)
		cmd.Parameters.Append cmd.CreateParameter("@NewEmailGroupID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		m_EmailGroupID = cmd.Parameters("@NewEmailGroupID").Value
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
	
	Public Sub Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_EmailGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter EmailGroupID not provided.")
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Save():", "Required parameter MemberID not provided.")
		
		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailUpdateEmailGroup"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EmailGroupID", adBigInt, adParamInput, 0, m_EmailGroupID)
		cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 200, m_Name)
		If Len(m_Description) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@Description", adVarChar, adParamInput, 1000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@Description", adVarChar, adParamInput, 1000, m_Description)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_DateModified)
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Sub
	
	Public Sub Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_EmailGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter EmailGroupID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailDeleteEmailGroup"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EmailGroupID", adBigInt, adParamInput, 0, m_EmailGroupID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Sub

End Class
</script>
