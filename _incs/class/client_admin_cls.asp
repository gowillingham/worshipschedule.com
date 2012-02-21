<script type="text/vbscript" runat="server" language="vbscript">
Class cClientAdmin

	Private m_ClientAdminID		'as long int
	Private m_ClientID		'as long int
	Private m_MemberID		'as long int
	Private m_DateModified		'as date
	Private m_DateCreated		'as date

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get ClientAdminID() 'As long int
		ClientAdminID = m_ClientAdminID
	End Property

	Public Property Let ClientAdminID(val) 'As long int
		m_ClientAdminID = val
	End Property
	
	Public Property Get ClientID() 'As long int
		ClientID = m_ClientID
	End Property

	Public Property Let ClientID(val) 'As long int
		m_ClientID = val
	End Property
	
	Public Property Get MemberID() 'As long int
		MemberID = m_MemberID
	End Property

	Public Property Let MemberID(val) 'As long int
		m_MemberID = val
	End Property
	
	Public Property Get DateModified() 'As date
		DateModified = m_DateModified
	End Property

	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cClientAdmin"
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
	
	' return the earliest created clientAdminId
	Public Function GetOldest()
		Dim i
		If Len(m_ClientID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".GetOldest():", "Required parameter ClientID not provided.")
	
		Dim firstAdminDateCreated
		Dim firstAdminMemberId
		
		Dim adminList			: adminList = List()

		' 0-ClientID 1-MemberID 2-NameFirst 3-NameLast 4-NameClient 5-DateCreated 
		' 6-DateModified 7-ClientAdminID
		firstAdminMemberId = adminList(1,0)
		firstAdminDateCreated = adminList(5,0)
		For i = 1 To UBound(adminList,2)
			If adminList(5,i) > firstAdminDateCreated Then
				firstAdminMemberId = adminList(1,i)
				firstAdminDateCreated = adminList(5,i)
			End If
		Next
		
		GetOldest = firstAdminMemberId
	End Function
	
	Public Function List()
		If Len(m_ClientID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".ClientList():", "Required parameter ClientID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-ClientID 1-MemberID 2-NameFirst 3-NameLast 4-NameClient 5-DateCreated 
		' 6-DateModified 7-ClientAdminID
		
		m_cnn.up_clientGetClientAdminList CLng(m_ClientID), m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()

		m_rs.Close()
	End Function
	
	Public Function Load() 'As Boolean
		If Len(m_clientAdminID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter ClientAdminID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_clientGetClientAdmin CLng(m_ClientAdminID), m_rs
		If Not m_rs.EOF Then
			m_ClientID = m_rs("ClientID").Value
			m_MemberID = m_rs("MemberID").Value
			m_DateModified = m_rs("DateModified").Value
			m_DateCreated = m_rs("DateCreated").Value
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd
		If Len(m_ClientID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Add():", "Required parameter ClientID not provided.")
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".Add():", "Required parameter MemberID not provided.")

		m_DateCreated = Now()
		m_DateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_clientInsertClientAdmin"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, m_ClientID)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, m_MemberID)
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_DateModified)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@NewClientAdminID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_ClientAdminID = cmd.Parameters("@NewClientAdminID").Value
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_ClientAdminID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter ClientAdminID not provided.")

		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_clientUpdateClientAdmin"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ClientAdminID", adBigInt, adParamInput, 0, m_ClientAdminID)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, m_ClientID)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, m_MemberID)
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_DateModified)
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Function
	
	Public Function Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_clientAdminID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter ClientAdminID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_clientDeleteClientAdmin"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ClientAdminID", adBigInt, adParamInput, 0, CLng(m_ClientAdminID))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Function

End Class
</script>
