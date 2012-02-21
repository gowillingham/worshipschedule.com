<script runat="server" language="vbscript" type="text/vbscript">

Class cMemberNotification
	Private m_ID		'as long int
	Private m_NotificationID		'as long int
	Private m_MemberID		'as long int
	Private m_DismissStatus		'as small int
	Private m_DateCreated		'as date
	Private m_DateModified		'as date
	Private m_Subject
	Private m_Text
	Private m_HrefText			'as string
	Private m_NotificationDate

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get ID() 'As long int
		ID = m_ID
	End Property

	Public Property Let ID(val) 'As long int
		m_ID = val
	End Property
	
	Public Property Get NotificationID() 'As long int
		NotificationID = m_NotificationID
	End Property

	Public Property Let NotificationID(val) 'As long int
		m_NotificationID = val
	End Property
	
	Public Property Get MemberID() 'As long int
		MemberID = m_MemberID
	End Property

	Public Property Let MemberID(val) 'As long int
		m_MemberID = val
	End Property
	
	Public Property Get DismissStatus() 'As small int
		DismissStatus = m_DismissStatus
	End Property

	Public Property Let DismissStatus(val) 'As small int
		m_DismissStatus = val
	End Property
	
	Public Property Get Subject()
		Subject = m_Subject
	End Property
	
	Public Property Get Text()
		Text = m_Text
	End Property
	
	Public Property Get HrefText()
		HrefText = m_HrefText
	End Property
	
	Public Property Get NotificationDate()
		NotificationDate = m_NotificationDate
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	Public Property Get DateModified() 'As date
		DateModified = m_DateModified
	End Property

	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cMemberNotification"
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
		
		m_cnn.up_memberGetMemberNotification CLng(m_ID), m_rs
		If Not m_rs.EOF Then
			m_ID = m_rs("ID").Value
			m_NotificationID = m_rs("NotificationID").Value
			m_MemberID = m_rs("MemberID").Value
			m_DismissStatus = m_rs("DismissStatus").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateModified = m_rs("DateModified").Value
			m_Subject = m_rs("Subject").Value
			m_Text = m_rs("Text").Value
			m_HrefText = m_rs("HrefText").Value
			m_NotificationDate = m_rs("NotificationDate").Value
		
			Load = True
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd

		If Len(m_NotificationID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Add():", "Required parameter NotificationID not provided.")
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Add():", "Required parameter MemberID not provided.")
		If Len(m_DismissStatus) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".Add():", "Required parameter DismissStatus not provided.")
		
		m_DateCreated = Now()
		m_DateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_memberInsertMemberNotification"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@NotificationID", adBigInt, adParamInput, 0, m_NotificationID)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, m_MemberID)
		cmd.Parameters.Append cmd.CreateParameter("@DismissStatus", adUnsignedTinyInt, adParamInput, 0, m_DismissStatus)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@NewMemberNotificationID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_ID = cmd.Parameters("@NewMemberNotificationID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(ID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter ID not provided.")

		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_memberUpdateMemberNotification"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@MemberNotificationID", adBigInt, adParamInput, 0, m_ID)
		cmd.Parameters.Append cmd.CreateParameter("@NotificationID", adBigInt, adParamInput, 0, m_NotificationID)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, m_MemberID)
		cmd.Parameters.Append cmd.CreateParameter("@DismissStatus", adUnsignedTinyInt, adParamInput, 0, m_DismissStatus)
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_DateModified)
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
			.CommandText = "dbo.up_memberDeleteMemberNotification"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@MemberNotificationID", adBigInt, adParamInput, 0, m_ID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Delete = True
		Else
			Delete = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Function List()
		' return array of current notifications for member
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".List():", "Required parameter MemberID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-ID 1-Subject 2-Text 3-DismissStatus 4-DateCreated 5-DateModified 6-HrefText
		' 7-ClientID 8-ProgramId 9-ProgramIsEnabled 10-ProgramMemberId 11-ProgramMemberIsEnabled
		' 12-ProgramName
		
		m_cnn.up_memberGetMemberNotificationList CLng(m_memberID), Now(),  m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Sub MarkDismissedForDelete(outError)
		Dim cmd
		
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".MarkDismissedForDelete():", "Required parameter MemberID not provided.")
		m_dateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")
		
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_memberNotificationMarkRowsForDelete"
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open Application.Value("CNN_STR")
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_MemberID))
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDBTimeStamp, adParamInput, 0, m_DateModified)
		
		cmd.Execute ,,adExecuteNoRecords
		
		Set cmd = Nothing
	End Sub

End Class

</script>