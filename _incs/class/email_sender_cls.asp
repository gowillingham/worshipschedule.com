<%

Class cEmailSender
	Private m_Email
	Private m_iMsgCount, m_iSentCount, m_iErrCount, m_sErrMsg, m_sAddressWithErrList
	Private m_sPickupFolderPath
	
	Public Property Let ToAddress(val)
		m_Email.To = val
	End Property
	
	Public Property Let CcAddress(val)
		m_Email.CC = val
	End Property
	
	Public Property Let BccAddress(val)
		m_Email.BCC = val
	End Property
	
	Public Property Let Subject(val)
		m_Email.Subject = val
	End Property
	
	Public Property Let Text(val)
		m_Email.TextBody = val
	End Property
	
	Public Property Let From(val)
		m_email.From = val
	End Property
	
	Public Function Send()
		m_email.Send()
	End Function
	
	Public Function SendMessage(sTo, sFrom, sSubject, sBody)
		Dim i, arr
		SendMessage = 0
		With m_Email
			.To = sTo
			.From = sFrom
			.Subject = sSubject
			.TextBody = sBody
		End With

		On Error Resume Next
			m_Email.Send
			If Err.Number <> 0 Then 
				SendMessage = Err.Number
				m_sErrMsg = Err.Description
				'increment error counter
				m_iErrCount = m_iErrCount + 1
				'add this address to list of addresses with errors
				If Len(m_sAddressWithErr) = 0 Then
					m_sAddressWithErrList = sTo
				Else
					m_sAddressWithErrList = m_sAddressWithErrList & " @@ " & sTo
				End If
			Else
				'increment successful send counter
				m_iSentCount = m_iSentCount + 1
			End If
			'increment msg counter
			m_iMsgCount = m_iMsgCount + 1
		On Error GoTo 0
	End Function
	
	Public Function AddAttachment(val)
		m_Email.AddAttachment val
	End Function
	
	Public Property Get ErrorCount()
		ErrorCount = m_iErrCount
	End Property
	
	Public Property Get TotalMessageCount()
		TotalMessageCount = m_iMsgCount
	End Property
	
	Public Property Get SentMessageCount()
		SentMessageCount = m_iSentCount
	End Property
	
	Public Property Get ErrorDescription()
		'this will only return the last error if more than 
		'one message is being sent by the object
		ErrorDescription = m_sErrMsg
		m_sErrMsg = 0 
	End Property
	
	Public Property Get AddressWithErrorList()
		'return list of addresses that threw error as array
		AddressWithErrorlist = Split(m_sAddressWithErrList, " @@ ")
	End Property
	
	Private Sub Class_Initialize()
		m_iMsgCount = 0
		m_iSentCount = 0
		m_iErrCount = 0
	
		If Not IsObject(m_Email) Then 
			Set m_Email = Server.CreateObject("CDO.Message")
		End If
	
		If Application.Value("IsLiveServer") Then
			' config info for remote server
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' send using remote smtp
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.worshipschedule.com"
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
			' remote server authentication credentials
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="webapplication@gtdsolutions"
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="reneepaul"
		Else
			m_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1 ' send using pickup folder
			
			 'pickup folder
			m_sPickupFolderPath = Application.Value("cEmailSender.PICKUP_FOLDER")
			If Len(m_sPickupFolderPath) > 0 Then 
				m_Email.Configuration.Fields.Item(cdoSMTPServerPickupDirectory) = m_sPickupFolderPath
			Else
				m_Email.Configuration.Fields.Item(cdoSMTPServerPickupDirectory) = "c:\inetpub\mailroot\pickup"	' default iis location ..
			End If
		End If
		m_Email.Configuration.Fields.Update
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(m_Email) Then Set m_Email = Nothing
	End Sub
End Class

%>