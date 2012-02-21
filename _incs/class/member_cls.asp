<script runat="server" type="text/vbscript" language="vbscript">

Class cMember

	Private m_MemberID 'as Int
	Private m_ClientID 'as Int
	Private m_ClientName 'as string
	Private m_NameLast 'as String
	Private m_NameFirst 'as String
	Private m_NameLogin 'as String
	Private m_PWord 'as String
	Private m_Email 'as String
	Private m_DOB 'as String
	Private m_Gender 'as String
	Private m_AddressLine1 'as String
	Private m_AddressLine2 'as String
	Private m_City 'as String
	Private m_StateID 'as Int
	Private m_StateCode 'as String
	Private m_StateName 'as String
	Private m_PostalCode 'as String
	Private m_PhoneHome 'as String
	Private m_PhoneMobile 'as String
	Private m_PhoneAlternate 'as String
	Private m_IsProfileComplete 'as String
	Private m_IsProfileUserCertified 'as String
	Private m_isAdmin
	Private m_isLeader
	Private m_hasPrograms
	Private m_IsApproved 'as String
	Private m_IsApprovedText 'as String
	Private m_ActiveStatus 'as String
	Private m_ActiveStatusText 'as String
	Private m_LastLogin 'as String
	Private m_SecretQuestion 'as String
	Private m_SecretAnswer 'as String
	Private m_DateCreated 'as String
	Private m_DateModified 'as String
	Private m_ProgramCount 'as int
	Private m_HomePageId 'as int
	Private m_HomePageName 'as str
	Private m_HomePageUrl 'as str
	Private m_ClientAdminId ' as long
	
	Private m_EmailRetype
	Private m_PWordRetype
	
	Private m_ProgramID 'as lng
	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	Private m_error		'as int
	Private m_bIsLoaded	'as bool
	Private CLASS_NAME	'as string
	
'	// prop let/get
	Public Property Get MemberID() 'As Int
		MemberID = m_MemberID
	End Property

	Public Property Let MemberID(val) 'As Int
		m_MemberID = val
	End Property 

	Public Property Get ClientID() 'As Int
		ClientID = m_ClientID
	End Property

	Public Property Let ClientID(val) 'As Int
		m_ClientID = val
	End Property 
	
	Public Property Get ClientName() 'as string
		ClientName = m_ClientName
	End Property

	Public Property Get NameLast() 'As String
		NameLast = m_NameLast
	End Property

	Public Property Let NameLast(val) 'As String
		m_NameLast = val
	End Property 

	Public Property Get NameFirst() 'As String
		NameFirst = m_NameFirst
	End Property

	Public Property Let NameFirst(val) 'As String
		m_NameFirst = val
	End Property 

	Public Property Get NameLogin() 'As String
		NameLogin = m_NameLogin
	End Property

	Public Property Let NameLogin(val) 'As String
		m_NameLogin = val
	End Property 

	Public Property Get PWord() 'As String
		PWord = m_PWord
	End Property

	Public Property Let PWord(val) 'As String
		m_PWord = val
	End Property 

	Public Property Get Email() 'As String
		Email = m_Email
	End Property

	Public Property Let Email(val) 'As String
		m_Email = val
	End Property 

	Public Property Get DOB() 'As String
		DOB = m_DOB
	End Property

	Public Property Let DOB(val) 'As String
		m_DOB = val
	End Property 

	Public Property Get Gender() 'As String
		Gender = m_Gender
	End Property

	Public Property Let Gender(val) 'As String
		m_Gender = val
	End Property 
	
	Public Property Get GenderText()
		If m_Gender = "M" Then GenderText = "Male"
		If m_Gender = "F" Then GenderText = "Female"
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
	
	Public Property Let StateCode(val)
		m_stateCode = val
	End Property

	Public Property Get StateName() 'As Int
		StateName = m_StateName
	End Property

	Public Property Get PostalCode() 'As String
		PostalCode = m_PostalCode
	End Property

	Public Property Let PostalCode(val) 'As String
		m_PostalCode = val
	End Property 

	Public Property Get PhoneHome() 'As String
		PhoneHome = m_PhoneHome
	End Property

	Public Property Let PhoneHome(val) 'As String
		m_PhoneHome = val
	End Property 

	Public Property Get PhoneMobile() 'As String
		PhoneMobile = m_PhoneMobile
	End Property

	Public Property Let PhoneMobile(val) 'As String
		m_PhoneMobile = val
	End Property 

	Public Property Get PhoneAlternate() 'As String
		PhoneAlternate = m_PhoneAlternate
	End Property

	Public Property Let PhoneAlternate(val) 'As String
		m_PhoneAlternate = val
	End Property 

	Public Property Get IsProfileComplete() 'As String
		IsProfileComplete = m_IsProfileComplete
	End Property

	Public Property Let IsProfileComplete(val) 'As String
		m_IsProfileComplete = val
	End Property 

	Public Property Get IsProfileUserCertified() 'As String
		IsProfileUserCertified = m_IsProfileUserCertified
	End Property

	Public Property Let IsProfileUserCertified(val) 'As String
		m_IsProfileUserCertified = val
	End Property 
	
	Public Property Get IsAdmin()
		IsAdmin = m_IsAdmin
	End Property
	
	Public Property Get IsLeader()
		IsLeader = m_IsLeader
	End Property
	
	Public Property Get HasPrograms()
		HasPrograms = m_HasPrograms
	End Property

	Public Property Get IsApproved() 'As String
		IsApproved = m_IsApproved
	End Property

	Public Property Let IsApproved(val) 'As String
		m_IsApproved = val
	End Property 

	Public Property Get IsApprovedText() 'As String
		IsApprovedText = m_IsApprovedText
	End Property

	Public Property Get ActiveStatus() 'As String
		ActiveStatus = m_ActiveStatus
	End Property

	Public Property Let ActiveStatus(val) 'As String
		m_ActiveStatus = val
	End Property 

	Public Property Get ActiveStatusText() 'As String
		ActiveStatusText = m_ActiveStatusText
	End Property

	Public Property Get LastLogin() 'As String
		LastLogin = m_LastLogin
	End Property

	Public Property Let LastLogin(val) 'As String
		m_LastLogin = val
	End Property 

	Public Property Get SecretQuestion() 'As String
		SecretQuestion = m_SecretQuestion
	End Property

	Public Property Let SecretQuestion(val) 'As String
		m_SecretQuestion = val
	End Property 

	Public Property Get SecretAnswer() 'As String
		SecretAnswer = m_SecretAnswer
	End Property

	Public Property Let SecretAnswer(val) 'As String
		m_SecretAnswer = val
	End Property 

	Public Property Get DateCreated() 'As String
		DateCreated = m_DateCreated
	End Property

	Public Property Get DateModified() 'As String
		DateModified = m_DateModified
	End Property
	
	Public Property Get EmailRetype() 'As String
		EmailRetype = m_EmailRetype
	End Property

	Public Property Let EmailRetype(val) 'As String
		m_EmailRetype = val
	End Property 

	Public Property Get PWordRetype() 'As String
		PWordRetype = m_PWordRetype
	End Property

	Public Property Let PWordRetype(val) 'As String
		m_PWordRetype = val
	End Property
	
	Public Property Get ProgramCount()
		ProgramCount = m_ProgramCount
	End Property
	
	Public Property Let HomePageID(val)
		m_homePageId = val
	End Property
	
	Public Property Get HomePageID()
		HomePageId = m_homePageId
	End Property
	
	Public Property Get HomePageName()
		HomePageName = m_homePageName
	End Property
	
	Public Property Get HomePageUrl()
		HomePageUrl = m_homePageUrl
	End Property
	
	Public Property Get ClientAdminId()
		ClientAdminId = m_ClientAdminId
	End Property
	
	Public Function AddressToString()
		Dim street, theRest
		
		Dim hasLine1		: hasLine1 = False
		Dim hasLine2		: hasLine2 = False
		Dim hasCity			: hasCity = False
		Dim hasPostal		: hasPostal = False
		Dim hasState		: hasState = False
		
		If Len(m_AddressLine1) > 0 Then hasLine1 = True
		If Len(m_AddressLine2) > 0 Then hasLine2 = True
		If Len(m_City) > 0 Then hasCity = True
		If Len(m_PostalCode) > 0 Then hasPostal = True
		If Len(m_StateCode) > 0 Then hasState = True
		
		If Not (hasLine1 Or hasLine2 Or hasCity Or hasPostal Or hasState) Then
			Exit Function
		End If
		
		If hasLine1 Or hasLine2 Then
			If hasLine1 Then
				street = street & Server.HTMLEncode(m_AddressLine1)
			End If
			If hasLine2 Then
				If Len(street) > 0 Then street = street & "<br />"
				street = street & Server.HTMLEncode(m_AddressLine2)
			End If
		Else
			street = street & Server.HTMLEncode("<No Street Address>")
		End If
		
		If hasCity Or hasState Or hasPostal	Or hasLine1 Or hasLine2 Then
			If hasCity Then
				theRest = theRest & Server.HTMLEncode(m_City & ", ")
			Else
				theRest = theRest & Server.HTMLEncode("???, ")
			End If
			If hasState Then
				theRest = theRest & StateCode & " "
			Else
				theRest = theRest & Server.HTMLEncode("??") 
			End If
			If hasPostal Then
				theRest = theRest & " " & m_PostalCode
			Else
				theRest = theRest & " " & Server.HTMLEncode("??")
			End If
		End If	
		
		AddressToString = street & "<br />" & theRest
	End Function 
	
	Function PhoneListToString()
		Dim str 
		
		If Len(m_PhoneHome) > 0 Then
			str = str & m_PhoneHome & " (home)"
		End If
		If Len(m_PhoneMobile) > 0 Then
			If Len(str) > 0 Then str = str & "<br />"
			str = str & m_PhoneMobile & " (mobile)"
		End If
		If Len(m_PhoneAlternate) > 0 Then
			If Len(str) > 0 Then str = str & "<br />"
			str = str & m_PhoneAlternate & " (alternate)"
		End If
		
		PhoneListToString = str
	End Function

	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cMember"
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
	
	Public Function FileList(programID)
		If Len(m_MemberID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".FileList();", "Missing required parameter MemberID.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-FileID 1-FileName 2-FriendlyName 3-Description 4-ClientID 5-ProgramID 6-DateCreated 7-DateModified
		' 8-FileExtension 9-FileSize 10-MimeType 11-MimeSubType 12-IsPublic 13-ProgramName
		If Len(programID) > 0 Then
			m_cnn.up_filesGetFileListForMemberID CLng(m_MemberID), CLng(programID), m_rs
		Else 
			m_cnn.up_filesGetFileListForMemberID CLng(m_MemberID), m_rs
		End If
		If Not m_rs.EOF Then FileList = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close()
	End Function
	
	Public Function ProgramList() 'as array
		Dim cmd
		If Len(m_MemberID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".ProgramList();", "Missing required parameter MemberID.")
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		' 0-ProgramID 1-ProgramName 2-ProgramDesc 3-ProgramMemberID 4-SkillListXmlFrag 5-EnrollStatusID 6-EnrollStatusText
		' 7-EnrollStatusDesc 8-IsLeader 9-IsLeaderText 10-IsActive 11-IsActiveText 12-LastScheduleView 
		' 13-DateModified 14-DateCreated 15-SkillCount 16-IsMissingAvailability 17-IsAdmin 18-ProgramIsEnabled
		
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_memberGetProgramDetails"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_MemberID))
		
		Set m_rs = cmd.Execute
		If Not m_rs.EOF Then ProgramList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
		Set cmd = Nothing
	End Function
	
	Public Function OwnedProgramsList()
		If Len(m_MemberID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".OwnedProgramsList();", "Missing required parameter MemberID.")
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
	
		' 0-ProgramId 1-ProgramName 2-IsEnabled 3-ScheduleCount 4-EventCount
		
		m_cnn.up_memberGetProgramsOwnedByMemberID CLng(m_MemberID), m_rs
		If Not m_rs.EOF Then OwnedProgramsList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close()
	End Function
	
	Public Function EventList(fromDate, programID, scheduleID)
		Dim cmd
		
		If Len(m_MemberID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".EventList();", "Missing required parameter MemberID.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_memberGetEventList"
			.ActiveConnection = m_cnn
		End With

		' 0-EventID 1-EventName 2-EventNote 3-EventDate 4-TimeStart 5-TimeEnd 6-ClientID
		' 7-MemberID 8-ProgramName 9-ProgramID 10-ScheduleID 11-ScheduleName 12-ScheduleDesc
		' 13-ScheduleIsVisible 14-IsScheduled 15-SkillListXmlFrag 16-FileListXmlFrag 17-IsAvailable
		' 18-EventAvailabilityID 19-ProgramIsEnabled 20-ProgramMemberIsActive 
		' 21-AvailabilityViewedByMember 22-MemberActiveStatus 23-AvailabilityNote 
		' 24-EventAvailabilityDateModified

		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_MemberID))
		cmd.Parameters.Append cmd.CreateParameter("@FromDate", adDate, adParamInput, 0, fromDate)
		cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, programID)
		cmd.Parameters.Append cmd.CreateParameter("@ScheduleID", adBigInt, adParamInput, 0, scheduleID)
	
		Set m_rs = cmd.Execute()
		If Not m_rs.EOF Then EventList = m_rs.GetRows()
				
		If m_rs.State = adStateOpen Then m_rs.Close()
	End Function
	
	Public Function AdminEventList(programId, scheduleId)
		Dim cmd
		
		If Len(m_MemberID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".AdminEventList();", "Missing required parameter MemberID.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		' 0-EventID 1-EventName 2-EventDate 3-EventNote 4-TimeStart 5-TimeEnd 6-DateCreated 7-DateModified
		' 8-ProgramID 9-ProgramName 10-ScheduleID 11-ScheduleName 12-ScheduleEventCount
		' 13-ScheduleIsVisible 14-HasUnpublishedChanges 15-EventFileCount 16-ScheduledMemberCount
		' 17-HtmlBackgroundColor
		
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_eventGetEventListForAdmin"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_memberId))
		If Len(programId) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, CLng(programId))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, Null)
		End If
		If Len(scheduleId) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@ScheduleID", adBigInt, adParamInput, 0, CLng(scheduleId))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@ScheduleID", adBigInt, adParamInput, 0, Null)
		End If
		
		Set m_rs = cmd.Execute()
		If Not m_rs.EOF Then AdminEventList = m_rs.GetRows()
				
		If m_rs.State = adStateOpen Then m_rs.Close()
	End Function
	
	Public Function Add(outError)
		Call Insert(outError)
		If outError = 0 Then
			Add = True
		Else
			Add = False
		End If
	End Function
	
	Public Function QuickAdd(programID, outError)
		m_ProgramID = programID
		Call Insert(outError)
		If outError = 0 Then
			QuickAdd = True
		Else
			QuickAdd = False
		End If
	End Function

	Public Function Update(ByRef outError) 'As Boolean
		Dim cmd, parm
		
		If Len(m_MemberID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Update();", "Missing required parameter MemberID.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_memberUpdate"
			.ActiveConnection = m_cnn
		End With

		Set parm = cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_MemberID))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NameLast", adVarChar, adParamInput, 50, m_NameLast)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NameFirst", adVarChar, adParamInput, 50, m_NameFirst)
		cmd.Parameters.Append parm
		If Len(m_NameLogin) > 0 Then
			Set parm = cmd.CreateParameter("@NameLogin", adVarChar, adParamInput, 50, m_NameLogin)
		Else
			Set parm = cmd.CreateParameter("@NameLogin", adVarChar, adParamInput, 50, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PWord) > 0 Then
			Set parm = cmd.CreateParameter("@PWord", adVarChar, adParamInput, 50, m_PWord)
		Else
			Set parm = cmd.CreateParameter("@PWord", adVarChar, adParamInput, 50, Null)
		End If
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@Email", adVarChar, adParamInput, 100, m_Email)
		cmd.Parameters.Append parm
		If Len(m_DOB) > 0 Then
			Set parm = cmd.CreateParameter("@DOB", adDate, adParamInput, 0, m_DOB)
		Else
			Set parm = cmd.CreateParameter("@DOB", adDate, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_Gender) > 0 Then
			Set parm = cmd.CreateParameter("@Gender", adVarChar, adParamInput, 1, m_Gender)
		Else
			Set parm = cmd.CreateParameter("@Gender", adVarChar, adParamInput, 1, Null)
		End If
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
		If Len(m_PhoneHome) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneHome", adVarChar, adParamInput, 14, m_PhoneHome)
		Else
			Set parm = cmd.CreateParameter("@PhoneHome", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneMobile) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneMobile", adVarChar, adParamInput, 14, m_PhoneMobile)
		Else
			Set parm = cmd.CreateParameter("@PhoneMobile", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneAlternate) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneAlternate", adVarChar, adParamInput, 14, m_PhoneAlternate)
		Else
			Set parm = cmd.CreateParameter("@PhoneAlternate", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_SecretQuestion) > 0 Then
			Set parm = cmd.CreateParameter("@SecretQuestion", adVarChar, adParamInput, 200, m_SecretQuestion)
		Else
			Set parm = cmd.CreateParameter("@SecretQuestion", adVarChar, adParamInput, 200, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_SecretAnswer) > 0 Then
			Set parm = cmd.CreateParameter("@SecretAnswer", adVarChar, adParamInput, 200, m_SecretAnswer)
		Else
			Set parm = cmd.CreateParameter("@SecretAnswer", adVarChar, adParamInput, 200, Null)
		End If
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@IsProfileComplete", adUnsignedTinyInt, adParamInput, 0, CInt(m_IsProfileComplete))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@IsProfileUserCertified", adUnsignedTinyInt, adParamInput, 0, CInt(m_IsProfileUserCertified))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@IsApproved", adUnsignedTinyInt, adParamInput, 0, CInt(m_IsApproved))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ActiveStatus", adUnsignedTinyInt, adParamInput, 0, CInt(m_ActiveStatus))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, Now())
		cmd.Parameters.Append parm
		
		cmd.Parameters.Append cmd.CreateParameter("@StartPage", adInteger, adParamInput, 0, m_HomePageId)
		
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Update = True
		Else
			Update = False
		End If

		Set parm = Nothing: Set cmd = Nothing
	End Function
	
	Public Function Delete(outError) 'As Boolean
		Dim cmd
	
		If Len(m_MemberID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Delete();", "Missing required parameter MemberID.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")
		
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_memberDelete"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_MemberID))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Delete = True
		Else
			Delete = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Function Load()
	
		If Len(m_MemberID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".Load();", "Missing required parameter MemberID.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		m_cnn.up_memberGetDetails CLng(m_MemberID), m_rs
		
		If Not m_rs.EOF Then
			m_ClientID = m_rs("ClientID").Value
			m_ClientName = m_rs("ClientName").Value
			m_NameLast = m_rs("NameLast").Value
			m_NameFirst = m_rs("NameFirst").Value
			m_NameLogin = m_rs("NameLogin").Value
			m_PWord = m_rs("PWord").Value
			m_Email = m_rs("Email").Value
			m_DOB = m_rs("DOB").Value
			m_Gender = m_rs("Gender").Value
			m_AddressLine1 = m_rs("AddressLine1").Value
			m_AddressLine2 = m_rs("AddressLine2").Value
			m_City = m_rs("City").Value
			m_StateID = m_rs("StateID").Value
			m_StateCode = m_rs("StateCode").Value
			m_StateName = m_rs("StateName").Value
			m_PostalCode = m_rs("PostalCode").Value
			m_PhoneHome = m_rs("PhoneHome").Value
			m_PhoneMobile = m_rs("PhoneMobile").Value
			m_PhoneAlternate = m_rs("PhoneAlternate").Value
			m_IsProfileComplete = m_rs("IsProfileComplete").Value
			m_IsProfileUserCertified = m_rs("IsProfileUserCertified").Value
			m_IsAdmin = m_rs("IsAdmin").Value
			m_IsLeader = m_rs("IsLeader").Value
			m_HasPrograms = m_rs("HasPrograms").Value
			m_IsApproved = m_rs("IsApproved").Value
			m_ActiveStatus = m_rs("ActiveStatus").Value
			m_LastLogin = m_rs("LastLogin").Value
			m_SecretQuestion = m_rs("SecretQuestion").Value
			m_SecretAnswer = m_rs("SecretAnswer").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateModified = m_rs("DateModified").Value
			m_ActiveStatusText = m_rs("ActiveStatusText").Value
			m_IsApprovedText = m_rs("IsApprovedText").Value
			m_ProgramCount = m_rs("ProgramCount").Value
			m_HomePageID = m_rs("HomePageID").Value
			m_HomePageName = m_rs("StartPageName").Value
			m_HomePageUrl = m_rs("StartPageUrl").Value
			m_ClientAdminId = m_rs("ClientAdminId").Value
			
			m_EmailRetype = m_Email
			m_PWordRetype = m_PWord
		
			Load = True
		Else
			Load = False
		End If
		
		m_rs.Close
	End Function
	
	Private Function Insert(ByRef outError) 'As Boolean
		Dim cmd, parm

		If Len(m_ClientID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Insert();", "Missing required parameter ClientID.")
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")
		
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_memberInsert"
			.ActiveConnection = m_cnn
		End With
		
		m_dateCreated = Now()

		Set parm = cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, CLng(m_ClientID))
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NameLast", adVarChar, adParamInput, 50, m_NameLast)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@NameFirst", adVarChar, adParamInput, 50, m_NameFirst)
		cmd.Parameters.Append parm
		Set parm = cmd.CreateParameter("@Email", adVarChar, adParamInput, 100, m_Email)
		cmd.Parameters.Append parm
		If Len(m_NameLogin) > 0 Then
			Set parm = cmd.CreateParameter("@NameLogin", adVarChar, adParamInput, 50, m_NameLogin)
		Else
			Set parm = cmd.CreateParameter("@NameLogin", adVarChar, adParamInput, 50, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PWord) > 0 Then
			Set parm = cmd.CreateParameter("@PWord", adVarChar, adParamInput, 50, m_PWord)
		Else
			Set parm = cmd.CreateParameter("@PWord", adVarChar, adParamInput, 50, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_DOB) > 0 Then
			Set parm = cmd.CreateParameter("@DOB", adDate, adParamInput, 0, m_DOB)
		Else
			Set parm = cmd.CreateParameter("@DOB", adDate, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_Gender) > 0 Then
			Set parm = cmd.CreateParameter("@Gender", adVarChar, adParamInput, 1, m_Gender)
		Else
			Set parm = cmd.CreateParameter("@Gender", adVarChar, adParamInput, 1, Null)
		End If
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
		If Len(m_PhoneHome) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneHome", adVarChar, adParamInput, 14, m_PhoneHome)
		Else
			Set parm = cmd.CreateParameter("@PhoneHome", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneMobile) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneMobile", adVarChar, adParamInput, 14, m_PhoneMobile)
		Else
			Set parm = cmd.CreateParameter("@PhoneMobile", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_PhoneAlternate) > 0 Then
			Set parm = cmd.CreateParameter("@PhoneAlternate", adVarChar, adParamInput, 14, m_PhoneAlternate)
		Else
			Set parm = cmd.CreateParameter("@PhoneAlternate", adVarChar, adParamInput, 14, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_SecretQuestion) > 0 Then
			Set parm = cmd.CreateParameter("@SecretQuestion", adVarChar, adParamInput, 200, m_SecretQuestion)
		Else
			Set parm = cmd.CreateParameter("@SecretQuestion", adVarChar, adParamInput, 200, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_SecretAnswer) > 0 Then
			Set parm = cmd.CreateParameter("@SecretAnswer", adVarChar, adParamInput, 200, m_SecretAnswer)
		Else
			Set parm = cmd.CreateParameter("@SecretAnswer", adVarChar, adParamInput, 200, Null)
		End If
		cmd.Parameters.Append parm
		If Len(m_ProgramID) > 0 Then
			Set parm = cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, CLng(m_ProgramID))
		Else
			Set parm = cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append parm
		
		Set parm = cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, Now())
		cmd.Parameters.Append parm
		cmd.Parameters.Append cmd.CreateParameter("@StartPage", adInteger, adParamInput, 0, m_HomePageId)
		
		Set parm = cmd.CreateParameter("@NewID", adBigInt, adParamOutput, 0)
		cmd.Parameters.Append parm

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_MemberID = cmd.Parameters("@NewID").Value
			Insert = True
		Else
			Insert = False
		End If

		Set parm = Nothing: Set cmd = Nothing
	End Function
	
End Class
</script>