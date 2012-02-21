<script type="text/vbscript" runat="server" language="vbscript">

Class cNotification

	Private m_iID		'as long int
	Private m_sSubject		'as string
	Private m_sHrefText		'as string
	Private m_sText		'as string
	Private m_iClientID		'as long int
	Private m_iProgramID		'as long int
	Private m_dDisplayUntil		'as date
	Private m_dDisplayFrom		'as date
	Private m_iShowTo		'as small int
	Private m_dDateCreated		'as date
	Private m_dDateModified		'as date

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get ID() 'As long int
		ID = m_iID
	End Property

	Public Property Let ID(val) 'As long int
		m_iID = val
	End Property
	
	Public Property Get Subject() 'As string
		Subject = m_sSubject
	End Property

	Public Property Let Subject(val) 'As string
		m_sSubject = val
	End Property
	
	Public Property Get HrefText() 'As string
		HrefText = m_sHrefText
	End Property

	Public Property Let HrefText(val) 'As string
		m_sHrefText = val
	End Property
	
	Public Property Get Text() 'As string
		Text = m_sText
	End Property

	Public Property Let Text(val) 'As string
		m_sText = val
	End Property
	
	Public Property Get ClientID() 'As long int
		ClientID = m_iClientID
	End Property

	Public Property Let ClientID(val) 'As long int
		m_iClientID = val
	End Property
	
	Public Property Get ProgramID() 'As long int
		ProgramID = m_iProgramID
	End Property

	Public Property Let ProgramID(val) 'As long int
		m_iProgramID = val
	End Property
	
	Public Property Get DisplayUntil() 'As date
		DisplayUntil = m_dDisplayUntil
	End Property

	Public Property Let DisplayUntil(val) 'As date
		m_dDisplayUntil = val
	End Property
	
	Public Property Get DisplayFrom() 'As date
		DisplayFrom = m_dDisplayFrom
	End Property

	Public Property Let DisplayFrom(val) 'As date
		m_dDisplayFrom = val
	End Property
	
	Public Property Get ShowTo() 'As small int
		ShowTo = m_iShowTo
	End Property

	Public Property Let ShowTo(val) 'As small int
		m_iShowTo = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_dDateCreated
	End Property

	Public Property Get DateModified() 'As date
		DateModified = m_dDateModified
	End Property

	
	Private Sub Class_Initialize()
		m_iID = 0
		m_sSubject = ""
		m_sText = ""
		m_iClientID = 0
		m_iProgramID = 0
		m_dDisplayUntil = ""
		m_dDisplayFrom = ""
		m_iShowTo = 0
		m_dDateCreated = ""
		m_dDateModified = ""
		
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cNotification"
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
	
	Public Function Load() 'As Boolean
		If Len(ID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter ID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_adminGetNotification CLng(m_iID), m_rs
		If Not m_rs.EOF Then
			m_iID = m_rs("ID").Value
			m_sSubject = m_rs("Subject").Value
			m_sHrefText = m_rs("HrefText").Value
			m_sText = m_rs("Text").Value
			m_iClientID = m_rs("ClientID").Value
			m_iProgramID = m_rs("ProgramID").Value
			m_dDisplayUntil = m_rs("DisplayUntil").Value
			m_dDisplayFrom = m_rs("DisplayFrom").Value
			m_iShowTo = m_rs("ShowTo").Value
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
		
		m_dDateCreated = Now()
		m_dDateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_adminInsertNotification"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@Subject", adVarChar, adParamInput, 200, m_sSubject)
		If Len(m_sHrefText) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@HrefText", adVarChar, adParamInput, 200, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@HrefText", adVarChar, adParamInput, 200, m_sHrefText)
		End If
		If Len(m_sText) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 8000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 8000, m_sText)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, m_iClientID)
		If Len(m_iProgramID) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, m_iProgramID)
		End If
		If Len(m_dDisplayUntil) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@DisplayUntil", adDate, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@DisplayUntil", adDate, adParamInput, 0, m_dDisplayUntil)
		End If
		If Len(m_dDisplayUntil) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@DisplayFrom", adDate, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@DisplayFrom", adDate, adParamInput, 0, m_dDisplayFrom)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@ShowTo", adUnsignedTinyInt, adParamInput, 0, m_iShowTo)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_dDateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_dDateModified)
		cmd.Parameters.Append cmd.CreateParameter("@NewNotificationID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_iID = cmd.Parameters("@NewNotificationID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(ID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter ID not provided.")

		m_dDateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_adminUpdateNotification"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@NotificationID", adBigInt, adParamInput, 0, CLng(m_iID))
		cmd.Parameters.Append cmd.CreateParameter("@Subject", adVarChar, adParamInput, 200, m_sSubject)
		If Len(m_sHrefText) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@HrefText", adVarChar, adParamInput, 200, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@HrefText", adVarChar, adParamInput, 200, m_sHrefText)
		End If
		If Len(m_sText) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 8000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 8000, m_sText)
		End If
		If Len(m_iClientID) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, m_iClientID)
		End If
		If Len(m_iProgramID) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, m_iProgramID)
		End If
		If Len(m_dDisplayFrom) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@DisplayFrom", adDate, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@DisplayFrom", adDate, adParamInput, 0, m_dDisplayFrom)
		End If
		If Len(m_dDisplayUntil) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@DisplayUntil", adDate, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@DisplayUntil", adDate, adParamInput, 0, m_dDisplayUntil)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@ShowTo", adUnsignedTinyInt, adParamInput, 0, m_iShowTo)
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
		
		If Len(ID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter ID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_adminDeleteNotification"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@NotificationID", adBigInt, adParamInput, 0, m_iID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Delete = True
		Else
			Delete = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Function List(memberID)
		' return list from dbo.Notification for admin grid
		If Len(memberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".List():", "Required parameter memberID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_adminGetNotificationList CLng(memberID), m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
End Class

</script>