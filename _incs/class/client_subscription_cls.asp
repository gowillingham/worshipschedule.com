<script type="text/vbscript" runat="server" language="vbscript">

Class cClientSubscription

	Private m_GUID		'as guid
	Private m_ClientID		'as long int
	Private m_NameClient		'as string
	Private m_ContactNameFirst		'as string
	Private m_ContactNameLast		'as string
	Private m_ContactPhone		'as string
	Private m_ContactEmail		'as string
	Private m_ContactEmailRetype	' as string
	Private m_ContactAddressLine1		'as string
	Private m_ContactAddressLine2		'as string
	Private m_ContactCity		'as string
	Private m_ContactStateID		'as int
	Private m_ContactStateCode		'as str
	Private m_ContactStateLongName		'as str
	Private m_ContactPostalCode		'as string
	Private m_TermLength		'as int
	Private m_TermStart		'as date
	Private m_PaymentType		'as small int
	Private m_IsPaymentReceived		'as small int
	Private m_Price		'as string
	Private m_PaymentReceived		'as string
	Private m_DateCreated		'as date
	Private m_DateModified		'as date
	Private m_SubscriptionID	' as int
	Private m_SubscriptionName	' str
	Private m_ContactPhoneRaw	' str
	Private m_PayPalTransactionID	' str
	Private m_PayPalIsSandbox		' tinyint
	Private m_PayPalPaymentStatus		' str
	Private m_PayPalPaymentStatusReason	' str

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	Private TYPE_OF		' str
	
	Public Property Get GUID() 'As long int
		GUID = Replace(Replace(m_GUID, "}", ""), "{", "")
	End Property

	Public Property Let GUID(val) 'As long int
		m_GUID = Replace(Replace(val, "}", ""), "{", "")
	End Property
	
	Public Property Get ClientID() 'As long int
		ClientID = m_ClientID
	End Property

	Public Property Let ClientID(val) 'As long int
		m_ClientID = val
	End Property
	
	Public Property Get NameClient() 'As string
		NameClient = m_NameClient
	End Property

	Public Property Get ContactNameFirst() 'As string
		ContactNameFirst = m_ContactNameFirst
	End Property

	Public Property Get ContactNameLast() 'As string
		ContactNameLast = m_ContactNameLast
	End Property

	Public Property Get ContactPhone() 'As string
		ContactPhone = m_ContactPhone
	End Property

	Public Property Get ContactEmail() 'As string
		ContactEmail = m_ContactEmail
	End Property

	Public Property Get ContactAddressLine1() 'As string
		ContactAddressLine1 = m_ContactAddressLine1
	End Property

	Public Property Get ContactAddressLine2() 'As string
		ContactAddressLine2 = m_ContactAddressLine2
	End Property

	Public Property Get ContactCity() 'As string
		ContactCity = m_ContactCity
	End Property

	Public Property Get ContactStateID() 'As int
		ContactStateID = m_ContactStateID
	End Property

	Public Property Get ContactStateCode() 'As string
		ContactStateCode = m_ContactStateCode
	End Property

	Public Property Get ContactStateLongName() 'As string
		ContactStateLongName = m_ContactStateLongName
	End Property

	Public Property Get ContactPostalCode() 'As string
		ContactPostalCode = m_ContactPostalCode
	End Property

	Public Property Get TermLength() 'As int
		TermLength = m_TermLength
	End Property

	Public Property Let TermLength(val) 'As int
		m_TermLength = val
	End Property
	
	Public Property Get TermStart() 'As date
		TermStart = m_TermStart
	End Property

	Public Property Let TermStart(val) 'As date
		m_TermStart = val
	End Property
	
	Public Property Get PaymentType() 'As small int
		PaymentType = m_PaymentType
	End Property

	Public Property Let PaymentType(val) 'As small int
		m_PaymentType = val
	End Property
	
	Public Property Get IsPaymentReceived() 'As small int
		IsPaymentReceived = m_IsPaymentReceived
	End Property

	Public Property Let IsPaymentReceived(val) 'As small int
		m_IsPaymentReceived = val
	End Property
	
	Public Property Get Price() 'As string
		Price = m_Price
	End Property

	Public Property Get PaymentReceived() 'As string
		PaymentReceived = m_PaymentReceived
	End Property

	Public Property Let PaymentReceived(val) 'As string
		m_PaymentReceived = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	Public Property Get DateModified() 'As date
		DateModified = m_DateModified
	End Property
	
	Public Property Get SubscriptionID() ' as int
		SubscriptionID = m_SubscriptionID
	End Property
	
	Public Property Let SubscriptionID(val)		' as int
		m_SubscriptionID = val
	End Property

	Public Property Get SubscriptionName() ' as int
		SubscriptionName = m_SubscriptionName
	End Property
	
	Public Property Get ContactPhoneRaw() ' as int
		ContactPhoneRaw = m_ContactPhoneRaw
	End Property
	
	Public Property Get PayPalTransactionID() ' as str
		PayPalTransactionID = m_PayPalTransactionID
	End Property
	
	Public Property Let PayPalTransactionID(val) ' as str
		m_PayPalTransactionID = val
	End Property

	Public Property Get PayPalIsSandbox() ' as tinyint
		PayPalIsSandbox = m_PayPalIsSandbox
	End Property
	
	Public Property Let PayPalIsSandbox(val) ' as tinyint
		m_PayPalIsSandbox = val
	End Property

	Public Property Get PayPalPaymentStatus() ' as tinyint
		PayPalPaymentStatus = m_PayPalPaymentStatus
	End Property
	
	Public Property Let PayPalPaymentStatus(val) ' as tinyint
		m_PayPalPaymentStatus = val
	End Property
	
	Public Property Get PayPalPaymentStatusReason() ' as tinyint
		PayPalPaymentStatusReason = m_PayPalPaymentStatusReason
	End Property
	
	Public Property Let PayPalPaymentStatusReason(val) ' as tinyint
		m_PayPalPaymentStatusReason = val
	End Property	

	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cClientSubscription"
		TYPE_OF = "ws.ClientSubscription"
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
	
	Public Function List() ' as array
		If Len(m_ClientID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".List():", "Required parameter ClientID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-ClientSubscriptionID 1-ClientID 2-NameClient 3-ContactNameFirst 4-ContactNameLast 5-ContactAddressLine1
		' 6-ContactAddressLine2 7-ContactCity 8-ContactStateID 9-ContactStateCode 10-ContactStateLongName 11-ContactPostalCode 
		' 12-ContactPhone 13-ContactEmail 14-TermLength 15-TermStart 16-PaymentType 17-IsPaymentReceived 18-Price 
		' 19-PaymentReceived 20-DateCreated 21-DateModified 22-SubscriptionID 23-SubscriptionName
		
		m_cnn.up_clientGetClientSubscriptionList CLng(m_ClientID), m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	' return date a new subscription should start on ..
	Public Function GetNewSubscriptionStartDate(trialExpiresDate)
		Dim i
		
		If Len(m_ClientID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".GetNewSubscriptionStartDate():", "Required parameter ClientID not provided.")
		
		Dim expires
		Dim maxExpires
		
		Dim subscriptionList			: subscriptionList = List()
		
		' initialize maxExpires to trialAccountExpiration
		maxExpires = trialExpiresDate
		
		If IsArray(subscriptionList) Then
			
			' check paid rows to get the most recent (max) expire date
			For i = 0 To UBound(subscriptionList,2)
				If subscriptionList(17,i) = 1 Then
					expires = DateAdd("m", subscriptionList(14,i), subscriptionList(15,i))
					
					' expires is later, reset maxExpires
					If expires > maxExpires Then maxExpires = expires
				End If	
			Next
		End If	
		
		' now() is later, reset maxExpires
		If Now() > maxExpires Then maxExpires = Now()
		
		GetNewSubscriptionStartDate = maxExpires	
	End Function
	
	' return the most recent paid subscription
	Public Function GetLastPaidSubscription()
		Dim i

		If Len(m_ClientID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".GetLastPaidSubscription():", "Required parameter ClientID not provided.")

		Dim subscriptionList				: subscriptionList = List()
		If Not IsArray(subscriptionList) Then Exit Function
		
		Dim currentClientSubscriptionId
		Dim currentDateCreated	
		Dim hasPaidClientSubscription					

		hasPaidClientSubscription = False
		For i = 0 To UBound(subscriptionList,2)
		
			' only check paid clientSubscriptions ..
			If subscriptionList(17,i) = 1 Then
			
				If Not hasPaidClientSubscription Then

					' this will be first paid clientSubscription found
					' so load temp vars ..
					currentClientSubscriptionId = subscriptionList(0,i)
					currentDateCreated = subscriptionList(20,i)
					hasPaidClientSubscription = True
				Else
				
					' check additional clientSubscriptions found against 
					' most recent found ..
					If subscriptionList(20,i) > currentDateCreated Then
						currentClientSubscriptionId = subscriptionList(0,i)
						currentDateCreated = subscriptionList(20,i)
					End If
				End If
				
			End If
		Next
		
		GetLastPaidSubscription = currentClientSubscriptionId
	End Function
	
	Public Sub Load()
		If Len(m_GUID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter GUID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_clientGetClientSubscription m_GUID, m_rs
		If Not m_rs.EOF Then
			m_GUID = m_rs("GUID").Value
			m_ClientID = m_rs("ClientID").Value
			m_NameClient = m_rs("NameClient").Value
			m_ContactNameFirst = m_rs("ContactNameFirst").Value
			m_ContactNameLast = m_rs("ContactNameLast").Value
			m_ContactPhone = m_rs("ContactPhone").Value
			m_ContactEmail = m_rs("ContactEmail").Value
			m_ContactEmailRetype = m_rs("ContactEmail").Value
			m_ContactAddressLine1 = m_rs("ContactAddressLine1").Value
			m_ContactAddressLine2 = m_rs("ContactAddressLine2").Value
			m_ContactCity = m_rs("ContactCity").Value
			m_ContactStateID = m_rs("ContactStateID").Value
			m_ContactStateCode = m_rs("ContactStateCode").Value
			m_ContactStateLongName = m_rs("ContactStateLongName").Value
			m_ContactPostalCode = m_rs("ContactPostalCode").Value
			m_TermLength = m_rs("TermLength").Value
			m_TermStart = m_rs("TermStart").Value
			m_PaymentType = m_rs("PaymentType").Value
			m_IsPaymentReceived = m_rs("IsPaymentReceived").Value
			m_Price = m_rs("Price").Value
			m_PaymentReceived = m_rs("PaymentReceived").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateModified = m_rs("DateModified").Value
			m_SubscriptionID = m_rs("SubscriptionID").Value
			m_SubscriptionName = m_rs("SubscriptionName").Value
			m_ContactPhoneRaw = m_rs("ContactPhoneRaw").Value
			m_PayPalTransactionID = m_rs("PayPalTransactionID").Value
			m_PayPalIsSandbox= m_rs("PayPalIsSandbox").Value
			m_PayPalPaymentStatus= m_rs("PayPalPaymentStatus").Value
			m_PayPalPaymentStatusReason= m_rs("PayPalPaymentStatusReason").Value
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
			.CommandText = "dbo.up_clientInsertClientSubscription"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, CLng(m_ClientID))
		cmd.Parameters.Append cmd.CreateParameter("@SubscriptionID", adInteger, adParamInput, 0, CInt(m_SubscriptionID))
		cmd.Parameters.Append cmd.CreateParameter("@TermStart", adDate, adParamInput, 0, m_TermStart)
		cmd.Parameters.Append cmd.CreateParameter("@PaymentType", adUnsignedTinyInt, adParamInput, 0, m_PaymentType)
		cmd.Parameters.Append cmd.CreateParameter("@IsPaymentReceived", adUnsignedTinyInt, adParamInput, 0, m_IsPaymentReceived)
		cmd.Parameters.Append cmd.CreateParameter("@PaymentReceived", adCurrency, adParamInput, 0, m_PaymentReceived)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_DateModified)
		cmd.Parameters.Append cmd.CreateParameter("@PayPalTransactionID", adVarChar, adParamInput, 256, m_PayPalTransactionID)
		cmd.Parameters.Append cmd.CreateParameter("@PayPalIsSandbox", adUnsignedTinyInt, adParamInput, 0, m_PayPalIsSandbox)
		cmd.Parameters.Append cmd.CreateParameter("@PayPalPaymentStatus", adVarChar, adParamInput, 50, m_PayPalPaymentStatus)
		cmd.Parameters.Append cmd.CreateParameter("@PayPalPaymentStatusReason", adVarChar, adParamInput, 50, m_PayPalPaymentStatusReason)
		cmd.Parameters.Append cmd.CreateParameter("@GUID", adGuid, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		m_GUID = cmd.Parameters("@GUID").Value

		Set cmd = Nothing
	End Sub
	
	Public Sub Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_GUID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter GUID not provided.")
		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_clientUpdateClientSubscription"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@GUID", adGUID, adParamInput, 0, m_GUID)
		cmd.Parameters.Append cmd.CreateParameter("@SubscriptionID", adInteger, adParamInput, 0, CInt(m_SubscriptionID))
		cmd.Parameters.Append cmd.CreateParameter("@TermStart", adDate, adParamInput, 0, m_TermStart)
		cmd.Parameters.Append cmd.CreateParameter("@PaymentType", adUnsignedTinyInt, adParamInput, 0, m_PaymentType)
		cmd.Parameters.Append cmd.CreateParameter("@IsPaymentReceived", adUnsignedTinyInt, adParamInput, 0, m_IsPaymentReceived)
		cmd.Parameters.Append cmd.CreateParameter("@PaymentReceived", adCurrency, adParamInput, 0, m_PaymentReceived)
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_DateModified)
		cmd.Parameters.Append cmd.CreateParameter("@PayPalTransactionID", adVarChar, adParamInput, 256, m_PayPalTransactionID)
		cmd.Parameters.Append cmd.CreateParameter("@PayPalIsSandbox", adUnsignedTinyInt, adParamInput, 0, m_PayPalIsSandbox)
		cmd.Parameters.Append cmd.CreateParameter("@PayPalPaymentStatus", adVarChar, adParamInput, 50, m_PayPalPaymentStatus)
		cmd.Parameters.Append cmd.CreateParameter("@PayPalPaymentStatusReason", adVarChar, adParamInput, 50, m_PayPalPaymentStatusReason)
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Sub
	
	Public Sub Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_GUID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter GUID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_clientDeleteClientSubscription"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@GUID", adGuid, adParamInput, 0, "{" & m_GUID & "}")

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
End Class

</script>
