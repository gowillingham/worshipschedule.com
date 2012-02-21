<script type="text/vbscript" runat="server" language="vbscript">
Class cEmailGroupMember

	Private m_EmailGroupMemberID		'as long int
	Private m_EmailGroupID		'as long int
	Private m_MemberID		'as long int
	Private m_AddressText		'as string
	Private m_DateCreated		'as date
	Private m_DateModified		' as date
	
	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get EmailGroupMemberID() 'As long int
		EmailGroupMemberID = m_EmailGroupMemberID
	End Property

	Public Property Let EmailGroupMemberID(val) 'As long int
		m_EmailGroupMemberID = val
	End Property
	
	Public Property Get EmailGroupID() 'As long int
		EmailGroupID = m_EmailGroupID
	End Property

	Public Property Let EmailGroupID(val) 'As long int
		m_EmailGroupID = val
	End Property
	
	Public Property Get MemberID() 'As long int
		MemberID = m_MemberID
	End Property

	Public Property Let MemberID(val) 'As long int
		m_MemberID = val
	End Property
	
	Public Property Get AddressText() 'As string
		AddressText = m_AddressText
	End Property

	Public Property Let AddressText(val) 'As string
		m_AddressText = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property
	
	Public Property Get DateModified() 'As date
		DateModified = m_DateModified
	End Property
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cEmailGroupMember"
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
		If Len(m_EmailGroupID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".List():", "Required parameter EmailGroupID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-EmailGroupMemberID 1-EmailGroupId 2-MemberID 3-Email 4-NameLast 5-NameFirst 6-DateCreated 7-MemberActiveStatus
		
		m_cnn.up_emailGetEmailGroupMemberList CLng(m_EmailGroupID), m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Sub Load() 'As Boolean
		If Len(m_EmailGroupMemberID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter EmailGroupMemberID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_emailGetEmailGroupMember CLng(m_EmailGroupMemberID), m_rs
		If Not m_rs.EOF Then
			m_EmailGroupMemberID = m_rs("EmailGroupMemberID").Value
			m_EmailGroupID = m_rs("EmailGroupID").Value
			m_MemberID = m_rs("MemberID").Value
			m_AddressText = m_rs("AddressText").Value
			m_DateCreated = m_rs("DateCreated").Value
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Sub
	
	Public Sub Add(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_EmailGroupID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Add():", "Required parameter EmailGroupID not provided.")
		
		m_DateCreated = Now()
		m_DateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailInsertEmailGroupMember"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EmailGroupID", adBigInt, adParamInput, 0, CLng(m_EmailGroupID))
		If Len(m_MemberID) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_MemberID))
		End If
		If Len(m_AddressText) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@AddressText", adVarChar, adParamInput, 200, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@AddressText", adVarChar, adParamInput, 200, m_AddressText)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@NewEmailGroupMemberID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		m_EmailGroupMemberID = cmd.Parameters("@NewEmailGroupMemberID").Value

		Set cmd = Nothing
	End Sub
	
	Public Sub Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_EmailGroupMemberID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter EmailGroupMemberID not provided.")
		If Len(m_EmailGroupID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Save():", "Required parameter EmailGroupID not provided.")

		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailUpdateEmailGroupMember"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EmailGroupMemberID", adBigInt, adParamInput, 0, CLng(m_EmailGroupMemberID))
		cmd.Parameters.Append cmd.CreateParameter("@EmailGroupID", adBigInt, adParamInput, 0, CLng(m_EmailGroupID))
		If Len(m_MemberID) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_MemberID))
		End If
		If Len(m_AddressText) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@AddressText", adVarChar, adParamInput, 200, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@AddressText", adVarChar, adParamInput, 200, m_AddressText)
		End If
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Sub
	
	Public Sub DeleteByMemberId(memberId, outError)
		Dim cmd
		
		If Len(m_EmailGroupId) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".DeleteByMemberId():", "Required parameter EmailGroupID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailDeleteEmailGroupMemberByMemberId"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EmailGroupId", adBigInt, adParamInput, 0, CLng(m_EmailGroupId))
		cmd.Parameters.Append cmd.CreateParameter("@MemberId", adBigInt, adParamInput, 0, CLng(memberId))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
	
	Public Sub Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_EmailGroupMemberID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter EmailGroupMemberID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailDeleteEmailGroupMember"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EmailGroupMemberID", adBigInt, adParamInput, 0, CLng(m_EmailGroupMemberID))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
End Class
</script>
