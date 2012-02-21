<script runat="server" language="vbscript" type="text/vbscript">

	Class cHomePage
		Private m_HomePageID		'as int
		Private m_Name		'as string
		Private m_Url		'as string
		Private m_IsEnabled		'as small int
		Private m_IsAdmin		'as small int

		Private m_sSQL		'as string
		Private m_cnn		'as ADODB.Connection
		Private m_rs		'as ADODB.Recordset
		
		Private CLASS_NAME	'as string
		
		Public Property Get HomePageID() 'As int
			HomePageID = m_HomePageID
		End Property

		Public Property Let HomePageID(val) 'As int
			m_HomePageID = val
		End Property
		
		Public Property Get Name() 'As string
			Name = m_Name
		End Property

		Public Property Let Name(val) 'As string
			m_Name = val
		End Property
		
		Public Property Get Url() 'As string
			Url = m_Url
		End Property

		Public Property Let Url(val) 'As string
			m_Url = val
		End Property
		
		Public Property Get IsEnabled() 'As small int
			IsEnabled = m_IsEnabled
		End Property

		Public Property Let IsEnabled(val) 'As small int
			m_IsEnabled = val
		End Property
		
		Public Property Get IsAdmin() 'As small int
			IsAdmin = m_IsAdmin
		End Property

		Public Property Let IsAdmin(val) 'As small int
			m_IsAdmin = val
		End Property
		
		
		Private Sub Class_Initialize()
			m_sSQL = Application.Value("CNN_STR")
			CLASS_NAME = "cHomePage"
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
		
		Public Function List() ' as array
		
			If Not IsObject(m_cnn) Then 
				Set m_cnn = Server.CreateObject("ADODB.Connection")
				m_cnn.Open m_sSQL
			End If
			If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
			
			' 0-HomePageId 1-Name 2-Url 3-IsEnabled 4-IsAdmin
			
			m_cnn.up_memberGetHomePageList m_rs
			If Not m_rs.EOF Then List = m_rs.GetRows()

			If m_rs.State = adStateOpen Then m_rs.Close
		End Function
		
		Public Sub Load() 'As Boolean
			If Len(m_HomePageID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter HomePageID not provided.")
		
			If Not IsObject(m_cnn) Then 
				Set m_cnn = Server.CreateObject("ADODB.Connection")
				m_cnn.Open m_sSQL
			End If
			If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
			
			m_cnn.up_memberGetHomePage CLng(m_HomePageID), m_rs
			If Not m_rs.EOF Then
				m_HomePageID = m_rs("HomePageID").Value
				m_Name = m_rs("Name").Value
				m_Url = m_rs("Url").Value
				m_IsEnabled = m_rs("IsEnabled").Value
				m_IsAdmin = m_rs("IsAdmin").Value
			End If
			
			If m_rs.State = adStateOpen Then m_rs.Close
		End Sub
		
		Public Sub Add(ByRef outError) 'As Boolean
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
				.CommandText = "dbo.up_memberInsertHomePage"
				.ActiveConnection = m_cnn
			End With
			
			cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
			cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, m_Name)
			cmd.Parameters.Append cmd.CreateParameter("@Url", adVarChar, adParamInput, 200, m_Url)
			cmd.Parameters.Append cmd.CreateParameter("@IsEnabled", adUnsignedTinyInt, adParamInput, 0, m_IsEnabled)
			cmd.Parameters.Append cmd.CreateParameter("@IsAdmin", adUnsignedTinyInt, adParamInput, 0, m_IsAdmin)
			cmd.Parameters.Append cmd.CreateParameter("@NewHomePageID", adBigInt, adParamOutput)
		
			cmd.Execute ,,adExecuteNoRecords
			outError = cmd.Parameters("@RETURN_VALUE").Value
			m_HomePageID = cmd.Parameters("@NewHomePageID").Value

			Set cmd = Nothing
		End Sub
		
		Public Sub Save(ByRef outError) 'As Boolean
			Dim cmd
			
			If Len(m_HomePageID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter HomePageID not provided.")
			m_DateModified = Now()
			
			If Not IsObject(m_cnn) Then 
				Set m_cnn = Server.CreateObject("ADODB.Connection")
				m_cnn.Open m_sSQL
			End If
			Set cmd = Server.CreateObject("ADODB.Command")

			With cmd
				.CommandType = adCmdStoredProc
				.CommandText = "dbo.up_memberUpdateHomePage"
				.ActiveConnection = m_cnn
			End With
			
			cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
			cmd.Parameters.Append cmd.CreateParameter("@HomePageID", adInteger, adParamInput, 0, m_HomePageID)
			cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, m_Name)
			cmd.Parameters.Append cmd.CreateParameter("@Url", adVarChar, adParamInput, 200, m_Url)
			cmd.Parameters.Append cmd.CreateParameter("@IsEnabled", adUnsignedTinyInt, adParamInput, 0, m_IsEnabled)
			cmd.Parameters.Append cmd.CreateParameter("@IsAdmin", adUnsignedTinyInt, adParamInput, 0, m_IsAdmin)
			cmd.Execute ,,adExecuteNoRecords
			outError = cmd.Parameters("@RETURN_VALUE").Value
			
			Set cmd = Nothing
		End Sub
		
		Public Sub Delete(ByRef outError) 'As Boolean
			Dim cmd
			
			If Len(m_HomePageID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter HomePageID not provided.")

			If Not IsObject(m_cnn) Then 
				Set m_cnn = Server.CreateObject("ADODB.Connection")
				m_cnn.Open m_sSQL
			End If
			Set cmd = Server.CreateObject("ADODB.Command")

			With cmd
				.CommandType = adCmdStoredProc
				.CommandText = "dbo.up_memberDeleteHomePage"
				.ActiveConnection = m_cnn
			End With
			
			cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
			cmd.Parameters.Append cmd.CreateParameter("@HomePageID", adInteger, adParamInput, 0, m_HomePageID)

			cmd.Execute ,,adExecuteNoRecords
			outError = cmd.Parameters("@RETURN_VALUE").Value

			Set cmd = Nothing
		End Sub
	End Class
	
</script>

