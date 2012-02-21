<script runat="server" language="vbscript" type="text/vbscript">
Class cEventFile
	Private m_EventFileID		'as long int
	Private m_EventID		'as long int
	Private m_FileID		'as long int
	Private m_DateCreated		'as date

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get EventFileID() 'As long int
		EventFileID = m_EventFileID
	End Property

	Public Property Let EventFileID(val) 'As long int
		m_EventFileID = val
	End Property
	
	Public Property Get EventID() 'As long int
		EventID = m_EventID
	End Property

	Public Property Let EventID(val) 'As long int
		m_EventID = val
	End Property
	
	Public Property Get FileID() 'As long int
		FileID = m_FileID
	End Property

	Public Property Let FileID(val) 'As long int
		m_FileID = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cEventFile"
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
		If Len(m_EventID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".List():", "Required parameter EventID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-EventFileID 1-FileName 2-FriendlyName 3-FileExtension 4-FileSize 5-IsPublic 6-DownloadCount
		' 7-Description 8-FileID 9-ClientID 10-ProgramID 11-FileOwnerID 12-DateFileCreated
		' 13-DateFileModified 14-DateEventFileCreated
		
		m_cnn.up_filesGetEventFileList CLng(m_EventID), m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Sub Load() 'As Boolean
		If Len(m_EventFileID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter EventFileID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_filesGetEventFile CLng(m_EventFileID), m_rs
		If Not m_rs.EOF Then
			m_EventFileID = m_rs("EventFileID").Value
			m_EventID = m_rs("EventID").Value
			m_FileID = m_rs("FileID").Value
			m_DateCreated = m_rs("DateCreated").Value
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Sub
	
	Public Sub Add(ByRef outError) 'As Boolean
		Dim cmd
		
		m_DateCreated = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_filesInsertEventFile"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, m_EventID)
		cmd.Parameters.Append cmd.CreateParameter("@FileID", adBigInt, adParamInput, 0, m_FileID)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@NewEventFileID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		m_EventFileID = cmd.Parameters("@NewEventFileID").Value

		Set cmd = Nothing
	End Sub
	
	Public Sub Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_EventFileID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter EventFileID not provided.")
		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_filesUpdateEventFile"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventFileID", adBigInt, adParamInput, 0, m_EventFileID)
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, m_EventID)
		cmd.Parameters.Append cmd.CreateParameter("@FileID", adBigInt, adParamInput, 0, m_FileID)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Sub
	
	Public Sub Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_EventFileID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter EventFileID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_filesDeleteEventFile"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventFileID", adBigInt, adParamInput, 0, m_EventFileID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
End Class
</script>

