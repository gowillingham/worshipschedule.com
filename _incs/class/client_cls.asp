<script type="text/vbscript" runat="server" language="vbscript">

Class cClient
	' REQUIREMENTS:
	' -------------------
	' this class requires class cMember to call cClient.CreateClientAccount
	' -------------------
	
	Private m_ClientID 'as Int
	Private m_NameClient 'as String
	Private m_AddressLine1 'as String
	Private m_AddressLine2 'as String
	Private m_City 'as String
	Private m_StateID 'as Int
	Private m_stateCode 'as Int
	Private m_stateLongName 'as Int
	Private m_PostalCode 'as String
	Private m_PhoneMain 'as String
	Private m_PhoneAlternate 'as String
	Private m_PhoneFax 'as String
	Private m_Email 'as String
	Private m_EmailRetype 'as string
	Private m_HomePage 'as String
	Private m_IsProfileComplete 'as String
	Private m_IsActive 'as String
	Private m_DateCreated 'as String
	Private m_DateModified 'as String
	Private m_FileStorage 'as Int
	Private m_IsTrialAccount 'as String
	Private m_TrialAccountLength 'as int
	Private m_GUID 'as string
	Private m_MemberCount 'as int
	Private m_ProgramCount 'as int

	Private m_SubscriptionCount ' int
	Private m_SubscriptionTermLength
	Private m_SubscriptionTermStart
	
	Private m_EventCount	' int
	Private m_ScheduleCount ' int
	
	Private m_ContactNameFirst ' str
	Private m_ContactNameLast ' str
	Private m_ContactAddressLine1 ' str
	Private m_ContactAddressLine2' str
	Private m_ContactCity ' str
	Private m_ContactStateID ' str
	Private m_ContactStateCode ' str
	Private m_ContactStateLongName ' str
	Private m_ContactPostalCode ' str
	Private m_ContactPhone ' str
	Private m_ContactEmail ' str
	Private m_ContactEmailRetype ' str
	
	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	Private m_error		'as int
	Private m_bIsLoaded	'as bool
	Private CLASS_NAME	'as string
	Private TYPE_OF		'str
	
'	// prop let/get
	Public Property Get ClientID() 'As Int
		ClientID = m_ClientID
	End Property

	Public Property Let ClientID(val) 'As Int
		m_ClientID = val
	End Property 

	Public Property Get NameClient() 'As String
		NameClient = m_NameClient
	End Property

	Public Property Let NameClient(val) 'As String
		m_NameClient = val
	End Property 

	Public Property Get AddressLine1() 'As String
		AddressLine1 = m_AddressLine1
	End Property

	Public Property Let AddressLine1(val) 'As String
		m_AddressLine1 = val
	End Property 

	Public Property Get AddressLine2() 'As String
		AddressLine2 = m_AddressLine2
	End Property

	Public Property Let AddressLine2(val) 'As String
		m_AddressLine2 = val
	End Property 

	Public Property Get City() 'As String
		City = m_City
	End Property

	Public Property Let City(val) 'As String
		m_City = val
	End Property 

	Public Property Get StateID() 'As Int
		StateID = m_StateID
	End Property

	Public Property Let StateID(val) 'As Int
		m_StateID = val
	End Property 

	Public Property Get StateCode() 'As Int
		StateCode = m_StateCode
	End Property

	Public Property Let StateCode(val) 'As Int
		m_StateCode = val
	End Property 

	Public Property Get StateLongName() 'As Int
		StateLongName = m_StateLongName
	End Property

	Public Property Let StateLongName(val) 'As Int
		m_StateLongName = val
	End Property 

	Public Property Get PostalCode() 'As String
		PostalCode = m_PostalCode
	End Property

	Public Property Let PostalCode(val) 'As String
		m_PostalCode = val
	End Property 

	Public Property Get PhoneMain() 'As String
		PhoneMain = m_PhoneMain
	End Property

	Public Property Let PhoneMain(val) 'As String
		m_PhoneMain = val
	End Property 

	Public Property Get PhoneAlternate() 'As String
		PhoneAlternate = m_PhoneAlternate
	End Property

	Public Property Let PhoneAlternate(val) 'As String
		m_PhoneAlternate = val
	End Property 

	Public Property Get PhoneFax() 'As String
		PhoneFax = m_PhoneFax
	End Property

	Public Property Let PhoneFax(val) 'As String
		m_PhoneFax = val
	End Property 

	Public Property Get Email() 'As String
		Email = m_Email
	End Property

	Public Property Let Email(val) 'As String
		m_Email = val
	End Property 

	Public Property Get EmailRetype() 'As String
		EmailRetype = m_EmailRetype
	End Property

	Public Property Let EmailRetype(val) 'As String
		m_EmailRetype = val
	End Property 

	Public Property Get HomePage() 'As String
		HomePage = m_HomePage
	End Property

	Public Property Let HomePage(val) 'As String
		m_HomePage = val
	End Property 

	Public Property Get IsProfileComplete() 'As String
		IsProfileComplete = m_IsProfileComplete
	End Property

	Public Property Let IsProfileComplete(val) 'As String
		m_IsProfileComplete = val
	End Property 

	Public Property Get IsActive() 'As String
		IsActive = m_IsActive
	End Property

	Public Property Let IsActive(val) 'As String
		m_IsActive = val
	End Property 

	Public Property Get DateCreated() 'As String
		DateCreated = m_DateCreated
	End Property

	Public Property Get DateModified() 'As String
		DateModified = m_DateModified
	End Property

	Public Property Get FileStorage() 'As Int
		FileStorage = m_FileStorage
	End Property

	Public Property Let FileStorage(val) 'As Int
		m_FileStorage = val
	End Property 

	Public Property Get IsTrialAccount() 'as int
		IsTrialAccount = m_IsTrialAccount
	End Property

	Public Property Let IsTrialAccount(val) 'as int
		m_IsTrialAccount = val
	End Property 
	
	Public Property Get TrialAccountLength() 'as Int
		TrialAccountLength = m_TrialAccountLength
	End Property

	Public Property Let TrialAccountLength(val) 'as int
		m_TrialAccountLength = val
	End Property 

	Public Property Get GUID() 'as string
		GUID = Replace(Replace(m_GUID, "{", ""), "}", "")
	End Property
	
	Public Property Get MemberCount() 'as int
		MemberCount = m_MemberCount
	End Property

	Public Property Get ProgramCount() 'as int
		ProgramCount = m_ProgramCount
	End Property
	
	Public Property Get EventCount()
		EventCount = m_EventCount
	End Property
	
	Public Property Get ScheduleCount()
		ScheduleCount = m_ScheduleCount
	End Property
	
	Public Property Get HasPrograms()
		HasPrograms = False
		If m_ProgramCount > 0 Then HasPrograms = True
	End Property
	
	Public Property Get HasEvents()
		HasEvents = False
		If m_EventCount > 0 Then HasEvents = True
	End Property
	
	Public Property Get HasSchedules()
		HasSchedules = False
		If m_ScheduleCount > 0 Then HasSchedules = True
	End Property
	
	Public Property Get HasSubscriptions()
		HasSubscriptions = True
		If m_SubscriptionCount = 0 Then HasSubscriptions = False
	End Property
	
	Public Property Get SubscriptionCount()
		SubscriptionCount = m_SubscriptionCount
	End Property
	
	Public Property Get SubscriptionTermStart()
		SubscriptionTermStart = m_SubscriptionTermStart
	End Property
	
	Public Property Get SubscriptionTermLength()
		SubscriptionTermLength = m_SubscriptionTermLength
	End Property

	Public Property Get ContactNameFirst()
		ContactNameFirst = m_ContactNameFirst
	End Property
	
	Public Property Let ContactNameFirst(val)
		m_ContactNameFirst = val
	End Property
	
	Public Property Get ContactNameLast()
		ContactNameLast = m_ContactNameLast
	End Property
	
	Public Property Let ContactNameLast(val)
		m_ContactNameLast = val
	End Property
	
	Public Property Get ContactAddressLine1()
		ContactAddressLine1 = m_ContactAddressLine1
	End Property
	
	Public Property Let ContactAddressLine1(val)
		m_ContactAddressLine1 = val
	End Property
	
	Public Property Get ContactAddressLine2()
		ContactAddressLine2 = m_ContactAddressLine2
	End Property
	
	Public Property Let ContactAddressLine2(val)
		m_ContactAddressLine2 = val
	End Property
	
	Public Property Get ContactCity()
		ContactCity = m_ContactCity
	End Property
	
	Public Property Let ContactCity(val)
		m_ContactCity = val
	End Property
	
	Public Property Get ContactStateID()
		ContactStateID = m_ContactStateID
	End Property
	
	Public Property Let ContactStateID(val)
		m_ContactStateID = val
	End Property
	
	Public Property Let ContactStateCode(val)
		m_ContactStateCode = val
	End Property
	
	Public Property Get ContactStateCode()
		ContactStateCode = m_ContactStateCode
	End Property
	
	Public Property Get ContactStateLongName()
		ContactStateLongName = m_ContactStateLongName
	End Property
	
	Public Property Get ContactPostalCode()
		ContactPostalCode = m_ContactPostalCode
	End Property
	
	Public Property Let ContactPostalCode(val)
		m_ContactPostalCode = val
	End Property
	
	Public Property Get ContactPhone()
		ContactPhone = m_ContactPhone
	End Property
	
	Public Property Let ContactPhone(val)
		m_ContactPhone = val
	End Property
	
	Public Property Get ContactEmail()
		ContactEmail = m_ContactEmail
	End Property
	
	Public Property Let ContactEmail(val)
		m_ContactEmail = val
	End Property
	
	Public Property Get ContactEmailRetype()
		ContactEmailRetype = m_ContactEmailRetype
	End Property
	
	Public Property Let ContactEmailRetype(val)
		m_ContactEmailRetype = val
	End Property
	
	Public Function SubscriptionExpiresDate()
		If m_isTrialAccount = 1 Then Exit Function
		
		SubscriptionExpiresDate = DateAdd("m", m_subscriptionTermLength, m_subscriptionTermStart)
	End Function
	
	Public Function TrialExpiresDate()
		TrialExpiresDate = DateAdd("d", m_trialAccountLength, m_dateCreated)
	End Function
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cClient"
		TYPE_OF = "ws.Client"
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
	
	Public Property Get IsTypeOf()
		IsTypeOf = TYPE_OF
	End Property
	
	Public Function Add(outError) 'As Boolean
		Dim cmd, parm, rv
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_clientInsertClient"
			.ActiveConnection = m_cnn
		End With
		
		m_DateCreated = Now()

		Set parm = cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NameClient", adVarChar, adParamInput, 100, m_NameClient)
		cmd.Parameters.Append parm
		If Len(m_AddressLine1) > 0 Then
			Set parm = cmd.CreateParameter("@AddressLine1", adVarChar, adParamInput, 100, m_AddressLine1)
		Else
			Set parm = cmd.CreateParameter("@AddressLine1", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_AddressLine2) > 0 Then
			Set parm = cmd.CreateParameter("@AddressLine2", adVarChar, adParamInput, 100, m_AddressLine2)
		Else
			Set parm = cmd.CreateParameter("@AddressLine2", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_City) > 0 Then
			Set parm = cmd.CreateParameter("@City", adVarChar, adParamInput, 100, m_City)
		Else
			Set parm = cmd.CreateParameter("@City", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_StateID) > 0 Then
			Set parm = cmd.CreateParameter("@StateID", adInteger, adParamInput, 0, CInt(m_StateID))
		Else
			Set parm = cmd.CreateParameter("@StateID", adInteger, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PostalCode) > 0 Then
			Set parm = cmd.CreateParameter("@PostalCode", adVarChar, adParamInput, 10, m_PostalCode)
		Else
			Set parm = cmd.CreateParameter("@PostalCode", adVarChar, adParamInput, 10, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneMain) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneMain", adVarChar, adParamInput, 14, m_PhoneMain)
		Else
			Set parm = cmd.CreateParameter("@PhoneMain", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneAlternate) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneAlternate", adVarChar, adParamInput, 14, m_PhoneAlternate)
		Else
			Set parm = cmd.CreateParameter("@PhoneAlternate", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneFax) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneFax", adVarChar, adParamInput, 14, m_PhoneFax)
		Else
			Set parm = cmd.CreateParameter("@PhoneFax", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_Email) > 0 Then
			Set parm = cmd.CreateParameter("@Email", adVarChar, adParamInput, 100, m_Email)
		Else
			Set parm = cmd.CreateParameter("@Email", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_HomePage) > 0 Then
			Set parm = cmd.CreateParameter("@HomePage", adVarChar, adParamInput, 100, m_HomePage)
		Else
			Set parm = cmd.CreateParameter("@HomePage", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@IsActive", adUnsignedTinyInt, adParamInput, 0, CInt(m_IsActive))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@IsTrialAccount", adUnsignedTinyInt, adParamInput, 0, CInt(m_IsTrialAccount))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@TrialAccountLength", adUnsignedTinyInt, adParamInput, 0, CInt(m_TrialAccountLength))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@IsProfileComplete", adUnsignedTinyInt, adParamInput, 0, CInt(m_IsProfileComplete))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@FileStorage", adBigInt, adParamInput, 0, m_FileStorage)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append parm
		

		Set parm = cmd.CreateParameter("@ContactNameFirst", adVarChar, adParamInput, 50, m_ContactNameFirst)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactNameLast", adVarChar, adParamInput, 50, m_ContactNameLast)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactAddressLine1", adVarChar, adParamInput, 100, m_ContactAddressLine1)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactAddressLine2", adVarChar, adParamInput, 100, m_ContactAddressLine2)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactCity", adVarChar, adParamInput, 100, m_ContactCity)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactStateID", adInteger, adParamInput, 0, m_ContactStateID)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactPostalCode", adVarChar, adParamInput, 10, m_ContactPostalCode)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactPhone", adVarChar, adParamInput, 15, m_ContactPhone)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactEmail", adVarChar, adParamInput, 100, m_ContactEmail)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NewClientID", adBigInt, adParamOutput, 0)
		cmd.Parameters.Append parm

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		m_ClientId = cmd.Parameters("@NewClientID").Value
		
		Set cmd = Nothing
		Set parm = Nothing	
	End Function
	
	Public Function Update(outError) 'As Boolean
		Dim cmd, parm
		
		If Len(m_ClientID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Update();", "ClientID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		
		Set cmd = Server.CreateObject("ADODB.Command")
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_clientUpdateClient"
			.ActiveConnection = m_cnn
		End With

		Set parm = cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ClientID", adVarChar, adParamInput, 100, CLng(m_ClientID))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NameClient", adVarChar, adParamInput, 100, m_NameClient)
		cmd.Parameters.Append parm
		If Len(m_AddressLine1) > 0 Then
			Set parm = cmd.CreateParameter("@AddressLine1", adVarChar, adParamInput, 100, m_AddressLine1)
		Else
			Set parm = cmd.CreateParameter("@AddressLine1", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_AddressLine2) > 0 Then
			Set parm = cmd.CreateParameter("@AddressLine2", adVarChar, adParamInput, 100, m_AddressLine2)
		Else
			Set parm = cmd.CreateParameter("@AddressLine2", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_City) > 0 Then
			Set parm = cmd.CreateParameter("@City", adVarChar, adParamInput, 100, m_City)
		Else
			Set parm = cmd.CreateParameter("@City", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_StateID) > 0 Then
			Set parm = cmd.CreateParameter("@StateID", adInteger, adParamInput, 0, CInt(m_StateID))
		Else
			Set parm = cmd.CreateParameter("@StateID", adInteger, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PostalCode) > 0 Then
			Set parm = cmd.CreateParameter("@PostalCode", adVarChar, adParamInput, 10, m_PostalCode)
		Else
			Set parm = cmd.CreateParameter("@PostalCode", adVarChar, adParamInput, 10, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneMain) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneMain", adVarChar, adParamInput, 14, m_PhoneMain)
		Else
			Set parm = cmd.CreateParameter("@PhoneMain", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneAlternate) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneAlternate", adVarChar, adParamInput, 14, m_PhoneAlternate)
		Else
			Set parm = cmd.CreateParameter("@PhoneAlternate", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneFax) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneFax", adVarChar, adParamInput, 14, m_PhoneFax)
		Else
			Set parm = cmd.CreateParameter("@PhoneFax", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_Email) > 0 Then
			Set parm = cmd.CreateParameter("@Email", adVarChar, adParamInput, 100, m_Email)
		Else
			Set parm = cmd.CreateParameter("@Email", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_HomePage) > 0 Then
			Set parm = cmd.CreateParameter("@HomePage", adVarChar, adParamInput, 100, m_HomePage)
		Else
			Set parm = cmd.CreateParameter("@HomePage", adVarChar, adParamInput, 100, Null)
		End If
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@IsActive", adUnsignedTinyInt, adParamInput, 0, CInt(m_IsActive))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@IsTrialAccount", adUnsignedTinyInt, adParamInput, 0, CInt(m_IsTrialAccount))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@TrialAccountLength", adUnsignedTinyInt, adParamInput, 0, CInt(m_TrialAccountLength))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@IsProfileComplete", adUnsignedTinyInt, adParamInput, 0, CInt(m_IsProfileComplete))
		cmd.Parameters.Append parm
		If Len(m_FileStorage) > 0 Then
			Set parm = cmd.CreateParameter("@FileStorage", adBigInt, adParamInput, 0, CLng(m_FileStorage))
		Else
			Set parm = cmd.CreateParameter("@FileStorage", adBigInt, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append parm

		Set parm = cmd.CreateParameter("@ContactNameFirst", adVarChar, adParamInput, 50, m_ContactNameFirst)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactNameLast", adVarChar, adParamInput, 50, m_ContactNameLast)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactAddressLine1", adVarChar, adParamInput, 100, m_ContactAddressLine1)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactAddressLine2", adVarChar, adParamInput, 100, m_ContactAddressLine2)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactCity", adVarChar, adParamInput, 100, m_ContactCity)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactStateID", adInteger, adParamInput, 0, m_ContactStateID)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactPostalCode", adVarChar, adParamInput, 10, m_ContactPostalCode)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactPhone", adVarChar, adParamInput, 15, m_ContactPhone)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ContactEmail", adVarChar, adParamInput, 100, m_ContactEmail)
		cmd.Parameters.Append parm

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
		Set parm = Nothing
	End Function
	
	Public Function Delete(outError) 'As Boolean
		Dim cmd, parm
		If Len(m_ClientID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Delete();", "ClientID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		
		Set cmd = Server.CreateObject("ADODB.Command")
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_clientDeleteClient"
			.ActiveConnection = m_cnn
		End With

		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, CLng(m_ClientiD))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value 
		
		Set cmd = Nothing
		Set parm = Nothing
	End Function
	
	Public Function Load() 'as boolean
	
		If Len(m_ClientID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Load();", "ClientID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_clientGetClientByClientID CLng(m_ClientID), m_rs
		
		If Not m_rs.EOF Then
			m_NameClient = m_rs("NameClient").Value
			m_AddressLine1 = m_rs("AddressLine1").Value
			m_AddressLine2 = m_rs("AddressLine2").Value
			m_City = m_rs("City").Value
			m_StateID = m_rs("StateID").Value
			m_stateCode = m_rs("StateCode").Value
			m_stateLongName = m_rs("StateLongName").Value
			m_PostalCode = m_rs("PostalCode").Value
			m_PhoneMain = m_rs("PhoneMain").Value
			m_PhoneAlternate = m_rs("PhoneAlternate").Value
			m_PhoneFax = m_rs("PhoneFax").Value
			m_Email = m_rs("Email").Value
			m_EmailRetype = m_Email
			m_HomePage = m_rs("HomePage").Value
			m_IsProfileComplete = m_rs("IsProfileComplete").Value
			m_IsActive = m_rs("IsActive").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateModified = m_rs("DateModified").Value
			m_FileStorage = m_rs("FileStorage").Value
			m_IsTrialAccount = m_rs("IsTrialAccount").Value
			m_TrialAccountLength = m_rs("TrialAccountLength").Value
			m_TrialAccountLength = m_rs("TrialAccountLength").Value
			m_GUID = m_rs("GUID").Value
			m_MemberCount = m_rs("MemberCount").Value
			m_ProgramCount = m_rs("ProgramCount").Value
			m_ScheduleCount = m_rs("ScheduleCount").Value
			m_EventCount = m_rs("EventCount").Value
			m_SubscriptionCount = m_rs("SubscriptionCount").Value
			m_SubscriptionTermStart = m_rs("SubscriptionTermStart").Value
			m_SubscriptionTermLength = m_rs("SubscriptionTermLength").Value
			m_ContactNameFirst = m_rs("ContactNameFirst").Value
			m_ContactNameLast = m_rs("ContactNameLast").Value
			m_ContactAddressLine1 = m_rs("ContactAddressLine1").Value
			m_ContactAddressLine2 = m_rs("ContactAddressLine2").Value
			m_ContactCity = m_rs("ContactCity").Value
			m_ContactStateID = m_rs("ContactStateID").Value
			m_ContactStateCode = m_rs("ContactStateCode").Value
			m_ContactStateLongName = m_rs("ContactStateLongName").Value
			m_ContactPostalCode = m_rs("ContactPostalCode").Value
			m_ContactPhone = m_rs("ContactPhone").Value
			m_ContactEmail = m_rs("ContactEmail").Value
			m_ContactEmailRetype = m_ContactEmail
			
			Load = True
		Else
			Load = False
		End If
		
		m_rs.Close
	End Function
	
	Public Function LoadByGUID(guid) 'as boolean
		If Len(guid) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".LoadByGuid();", "Client Guid not provided.")
		m_GUID = guid
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_clientGetClientByGUID guid, m_rs

		If Not m_rs.EOF Then
			m_ClientID = m_rs("ClientID").Value
			m_NameClient = m_rs("NameClient").Value
			m_AddressLine1 = m_rs("AddressLine1").Value
			m_AddressLine2 = m_rs("AddressLine2").Value
			m_City = m_rs("City").Value
			m_StateID = m_rs("StateID").Value
			m_PostalCode = m_rs("PostalCode").Value
			m_PhoneMain = m_rs("PhoneMain").Value
			m_PhoneAlternate = m_rs("PhoneAlternate").Value
			m_PhoneFax = m_rs("PhoneFax").Value
			m_Email = m_rs("Email").Value
			m_HomePage = m_rs("HomePage").Value
			m_IsProfileComplete = m_rs("IsProfileComplete").Value
			m_IsActive = m_rs("IsActive").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateModified = m_rs("DateModified").Value
			m_FileStorage = m_rs("FileStorage").Value
			m_IsTrialAccount = m_rs("IsTrialAccount").Value
			m_TrialAccountLength = m_rs("TrialAccountLength").Value
			m_TrialAccountLength = m_rs("TrialAccountLength").Value
			
			LoadByGUID = True
		Else
			LoadByGUID = False
		End If
		
		m_rs.Close
	End Function
	
	Public Function ScheduleList()
		If Len(m_ClientID) = 0 Then Call Err.Raise(5 + vbObjectError, CLASS_NAME & ".ScheduleList();", "ClientID not provided.")
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_clientGetScheduleList CLng(m_clientId), m_rs
		If Not m_rs.EOF Then ScheduleList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function ProgramList(memberID)
		If Len(m_ClientID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ProgramList();", "ClientID not provided.")
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-EnrollmentType 4-IsEnabled 5-DateCreated
		' 6-DateModified 7-DefaultAvailability 8-MemberCanEnroll 9-MemberCanEditSkills
		' 10-MemberCount 11-SkillCount 12-ScheduleCount 113-EventCount
		
		If Len(memberID) = 0 Then
			m_cnn.up_clientGetProgramList CLng(m_ClientID), m_rs
		Else 
			m_cnn.up_clientGetProgramList CLng(m_ClientID), CLng(memberID), m_rs
		End If
		If Not m_rs.EOF Then ProgramList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function MemberList(programID, sortColumn)
		If Len(m_ClientID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".MemberList();", "ClientID not provided.")
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		m_rs.CursorLocation = adUseClient	' makes sortable
		
		' 0-MemberID 1-NameLast 2-NameFirst 3-NameLogin 4-PWord 5-Email 6-DOB 7-Gender
		' 8-AddressLine1 9-AddressLine2 10-City 11-StateID 12-StateCode 13-PostalCode
		' 14-PhoneHome 15-PhoneMobile 16-PhoneAlternate 17-IsProfileComplete 
		' 18-IsProfileUserCertified 19-IsApproved 20-ActiveStatus 21-LastLogin 22-SecretQuestion
		' 23-SecretAnswer 24-DateCreated 25-DateModified 26-IsAdmin 27-MemberProgramListXML
		
		If Len(programID) = 0 Then
			m_cnn.up_clientGetMemberList CLng(m_ClientID), m_rs
		Else 
			m_cnn.up_clientGetMemberList CLng(m_ClientID), CLng(programID), m_rs
		End If

		m_rs.Sort = sortColumn
		If Not m_rs.EOF Then MemberList = m_rs.GetRows()

		m_rs.Close()
	End Function	
End Class

</script>