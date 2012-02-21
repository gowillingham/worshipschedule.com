<script type="text/vbscript" runat="server" language="vbscript">

Class cSession

	Private m_SessionID		'as string
	Private m_MemberID		'as long int
	Private m_ClientID		'as long int
	Private m_IsAdmin		'as small int
	Private m_IsLeader		'as small int
	Private m_IsImpersonated		'as small int
	Private m_SessionKey		'as string
	Private m_DateLastRefresh		'as date
	Private m_DateCreated		'as date
	Private m_IsMemberProfileComplete		'as tinyint
	Private m_IsClientProfileComplete		'as tinyint
	Private m_IsFileStoreEnabled			'as tinyint
	Private m_IsTrialAccount				'as tinyint
	Private m_TrialAccountLength			'as int
	Private m_FileStorage					'as int
	Private m_IsClientEnabled				'tinyint
	Private m_DateClientCreated				'date
	Private m_DateMemberCreated				'date
	Private m_NameFirst						'as str
	Private m_NameLast						'as str
	
	Private m_SubscriptionCount
	Private m_SubscriptionTermStart
	Private m_SubscriptionTermLength
	
	Private m_currentPage
	
	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	Private SESSION_TIMEOUT ' int
	Private SESSION_ABANDON ' int
	
	Public Property Get SessionID() 'As string
		SessionID = m_SessionID
	End Property

	Public Property Let SessionID(val) 'As string
		m_SessionID = val
	End Property
	
	Public Property Get MemberID() 'As long int
		MemberID = m_MemberID
	End Property

	Public Property Let MemberID(val) 'As long int
		m_MemberID = val
	End Property
	
	Public Property Get ClientID() 'As long int
		ClientID = m_ClientID
	End Property

	Public Property Let ClientID(val) 'As long int
		m_ClientID = val
	End Property
	
	Public Property Get IsAdmin() 'As small int
		IsAdmin = False
		If m_IsAdmin = 1 Then IsAdmin = True
	End Property

	Public Property Let IsAdmin(val) 'As small int
		If val Then 
			m_IsAdmin = 1
		Else
			m_IsAdmin = 0
		End If
	End Property
	
	Public Property Get IsLeader() 'As small int
		IsLeader = False
		If m_IsLeader = 1 Then IsLeader = True
	End Property

	Public Property Let IsLeader(val) 'As small int
		If val Then 
			m_IsLeader = 1
		Else
			m_IsLeader = 0
		End If
	End Property
	
	Public Property Get IsImpersonated() 'As small int
		IsImpersonated = m_IsImpersonated
	End Property

	Public Property Let IsImpersonated(val) 'As small int
		m_IsImpersonated = val
	End Property
	
	Public Property Get SessionKey() 'As string
		SessionKey = m_SessionKey
	End Property

	Public Property Let SessionKey(val) 'As string
		m_SessionKey = val
	End Property
	
	Public Property Get DateLastRefresh() 'As date
		DateLastRefresh = m_DateLastRefresh
	End Property

	Public Property Let DateLastRefresh(val) 'As date
		m_DateLastRefresh = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	Public Property Get IsMemberProfileComplete()
		IsMemberProfileComplete = True
		If m_IsMemberProfileComplete = 0 Then IsMemberProfileComplete = False
	End Property

	Public Property Get IsClientProfileComplete()
		IsClientProfileComplete = True
		If m_IsClientProfileComplete = 0 Then IsClientProfileComplete = False
	End Property

	Public Property Get IsFileStoreEnabled()
		IsFileStoreEnabled = True
		If m_IsFileStoreEnabled = 0 Then IsFileStoreEnabled = False
	End Property

	Public Property Get IsTrialAccount()
		IsTrialAccount = True
		If m_IsTrialAccount = 0 Then IsTrialAccount = False
	End Property

	Public Property Get TrialAccountLength()
		TrialAccountLength = m_TrialAccountLength
	End Property
	
	' calculated ..
	Public Function TrialExpiresDate()
		TrialExpiresDate = DateAdd("d", m_trialAccountLength, m_dateClientCreated)
	End Function

	Public Property Get FileStorage()
		FileStorage = m_FileStorage
	End Property
	
	Public Property Get IsClientEnabled()
		IsClientEnabled = True
		If m_IsClientEnabled = 0 Then IsClientEnabled = False
	End Property
	
	Public Property Get DateClientCreated()
		DateClientCreated = m_DateClientCreated
	End Property

	Public Property Get DateMemberCreated()
		DateMemberCreated = m_DateMemberCreated
	End Property
	
	Public Property Get NameFirst()
		NameFirst = m_NameFirst
	End Property
	
	Public Property Get NameLast()
		NameLast = m_NameLast
	End Property
	
	Public Property Get HasSubscription()
		HasSubscription = False
		If m_SubscriptionCount > 0 Then
			HasSubscription = True
		End If
	End Property
	
	Public Property Get SubscriptionCount()
		SubscriptionCount = m_SubscriptionCount
	End Property
	
	Public Property Get SubscriptionTermLength()
		SubscriptionTermLength = m_SubscriptionTermLength
	End Property
	
	Public Property Get SubscriptionTermStart()
		SubscriptionTermStart = m_SubscriptionTermStart
	End Property
	
	Public Property Get AccountExpiresDate() 
		If HasSubscription() Then
			accountExpiresDate = DateAdd("d", -1, DateAdd("m", m_subscriptionTermLength, m_subscriptionTermStart))
		End If
	End Property
	
	Public Property Get CurrentPage()
		CurrentPage = m_currentPage
	End Property
	
	Public Property Let CurrentPage(val)
		m_currentPage = val
	End Property
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		SESSION_TIMEOUT = Application.Value("SESSION_TIMEOUT")
		SESSION_ABANDON = Application.Value("SESSION_ABANDON_TIMEOUT")
		CLASS_NAME = "cSession"
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
	
	Public Sub Refresh(sessionID, outStatus)
		Dim rv
		
		If Len(sessionID) > 0 Then m_SessionID = sessionID
		outStatus = 0
		
		' check for missing SessionID
		If Len(m_SessionID) = 0 Then
			' session cookie not exists, not logged in
			outStatus = -1
			Exit Sub
		End If
		
		Call Load()
		
		' save the page we're trying to get to with the session
		m_currentPage = Request.ServerVariables("URL")
		If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then 
			m_currentPage = m_currentPage & "?" & Request.ServerVariables("QUERY_STRING")
		End If
		Call Save("")
		
		' validate session .. make sure session exists for this member
		If Len(m_MemberID) = 0 Then
			' session row in db not exists, not logged in
			outStatus = -1
			Exit Sub
		End If
		
		' check if session has timed out ..
		If DateDiff("s", m_DateLastRefresh, Now()) > SESSION_TIMEOUT Then
			' return -2
			outStatus = -2
			Exit Sub
		End If
		
		' check if session will be abandoned ..
		If DateDiff("h", m_DateLastRefresh, Now()) > SESSION_ABANDON Then
		    ' return -1
		    outStatus = -1
		    Exit Sub
		End If
		
		m_DateLastRefresh = Now()
		Call Save(rv)
	End Sub
	
	Public Sub Load() 'As Boolean
		If Len(m_SessionID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter SessionID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_sessionGetSession m_SessionID, m_rs
		If Not m_rs.EOF Then
			m_MemberID = m_rs("MemberID").Value
			m_ClientID = m_rs("ClientID").Value
			m_IsAdmin = m_rs("IsAdmin").Value
			m_IsLeader = m_rs("IsLeader").Value
			m_IsImpersonated = m_rs("IsImpersonated").Value
			m_SessionKey = m_rs("SessionKey").Value
			m_IsMemberProfileComplete = m_rs("IsMemberProfileComplete").Value
			m_IsClientProfileComplete = m_rs("IsClientProfileComplete").Value
			m_IsFileStoreEnabled = m_rs("IsFileStoreEnabled").Value
			m_FileStorage = m_rs("FileStorage").Value
			m_IsTrialAccount = m_rs("IsTrialAccount").Value
			m_TrialAccountLength = m_rs("TrialAccountLength").Value
			m_IsClientEnabled = m_rs("IsClientEnabled").Value
			m_DateLastRefresh = m_rs("DateLastRefresh").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateClientCreated = m_rs("DateClientCreated").Value
			m_DateMemberCreated = m_rs("DateMemberCreated").Value
			m_NameFirst = m_rs("NameFirst").Value
			m_NameLast = m_rs("NameLast").Value
			m_SubscriptionCount = m_rs("SubscriptionCount").Value
			m_SubscriptionTermStart = m_rs("SubscriptionTermStart").Value
			m_SubscriptionTermLength = m_rs("SubscriptionTermLength").Value
			m_currentPage = m_rs("CurrentPage").Value
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Sub
	
	Public Sub Add(ByRef outError) 'As Boolean
		Dim cmd
		
		m_DateCreated = Now()
		m_DateLastRefresh = m_DateCreated
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_sessionInsert"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, m_MemberID)
		cmd.Parameters.Append cmd.CreateParameter("@IsImpersonated", adUnsignedTinyInt, adParamInput, 0, m_IsImpersonated)
		cmd.Parameters.Append cmd.CreateParameter("@SessionKey", adVarChar, adParamInput, 50, m_SessionKey)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, CDate(m_DateCreated))
		cmd.Parameters.Append cmd.CreateParameter("@DateLastRefresh", adDate, adParamInput, 0, CDate(m_DateLastRefresh))
		cmd.Parameters.Append cmd.CreateParameter("@CurrentPage", adVarChar, adParamInput, 2000, m_currentPage)
		cmd.Parameters.Append cmd.CreateParameter("@NewSessionID", adGUID, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		m_SessionID = cmd.Parameters("@NewSessionID").Value

		Set cmd = Nothing
	End Sub
	
	Public Sub Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_SessionID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter SessionID not provided.")
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_sessionUpdate"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@SessionID", adGUID, adParamInput, 0, m_SessionID)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, m_MemberID)
		cmd.Parameters.Append cmd.CreateParameter("@IsAdmin", adUnsignedTinyInt, adParamInput, 0, m_IsAdmin)
		cmd.Parameters.Append cmd.CreateParameter("@IsLeader", adUnsignedTinyInt, adParamInput, 0, m_IsLeader)
		cmd.Parameters.Append cmd.CreateParameter("@IsImpersonated", adUnsignedTinyInt, adParamInput, 0, m_IsImpersonated)
		cmd.Parameters.Append cmd.CreateParameter("@SessionKey", adVarChar, adParamInput, 50, m_SessionKey)
		cmd.Parameters.Append cmd.CreateParameter("@DateLastRefresh", adDate, adParamInput, 0, CDate(m_DateLastRefresh))
		cmd.Parameters.Append cmd.CreateParameter("@CurrentPage", adVarChar, adParamInput, 2000, m_currentPage)
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Sub
	
	Public Sub Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_SessionID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter SessionID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_sessionDelete"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@SessionID", adGUID, adParamInput, 0, m_SessionID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
End Class

</script>
