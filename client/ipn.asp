<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>

<script runat="server" type="text/vbscript" language="vbscript">

	Call Main()
	
	Sub Main()
		Dim http			: Set http = Server.CreateObject("Msxml2.ServerXMLHTTP")
		Dim subscription	: Set subscription = New cClientSubscription
		Dim myResponse		: myResponse = Request.Form & "&cmd=_notify-validate"
		Dim returnError		: returnError = 0
		Dim httpResponse, str, rv

		' repost to paypal for verify
		Call http.Open("POST", Application.Value("PAYPAL_GATEWAY"), False)
		Call http.SetRequestHeader("Content-type", "application/x-www-form-urlencoded")
		Call http.Send(myResponse)
		httpResponse = http.ResponseText
		
		' generate text for logging email
		str = str & "timestamp=" & Now()
		str = str & vbCrLf & vbCrLf & "httpresponse=" & httpResponse
		str = str & vbCrLf & vbCrLf & "test_ipn=" & Request.Form("test_ipn")
		str = str & vbCrLf & "Subscription.Guid=" & Request.Form("invoice")
		str = str & vbCrLf & "Subscription.PaymentReceived=" & Request.Form("mc_gross")
		str = str & vbCrLf & "txn_id=" & Request.Form("txn_id")
		str = str & vbCrLf & "payment_status=" & Request.Form("payment_status")
		str = str & vbCrLf & "payment_type=" & Request.Form("payment_type")
		str = str & vbCrLf & "pending_reason=" & Request.Form("pending_reason")
		str = str & vbCrLf & "business=" & Request.Form("business")
		str = str & vbCrLf & "mc_gross=" & Request.Form("mc_gross")
		str = str & vbCrLf & "mc_currency=" & Request.Form("mc_currency")
		
		' debug
		Response.Write "<p>http response text=" & http.ResponseText
		Response.Write "<p>test_ipn='" & Request.Form("test_ipn") & "'"	
		Response.Write "<p>invoice='" & Request.Form("invoice") & "'"	
		Response.Write "<p>mc_gross='" & Request.Form("mc_gross") & "'"	
		Response.Write "<p>txn_id='" & Request.Form("txn_id") & "'"	
		Response.Write "<p>payment_status='" & Request.Form("payment_status") & "'"	
		Response.Write "<p>pending_reason='" & Request.Form("pending_reason") & "'"	
		Response.Write "<p>business='" & Request.Form("business") & "'"	
		Response.Write "<p>mc_currency='" & Request.Form("mc_currency") & "'"
		Response.Write "<p>paypal gateway='" & Application.Value("PAYPAL_GATEWAY") & "'"
		
		
		If Not IsValidHttpResponse(httpResponse, str) Then
			Call SendLogEmail(str)
			Exit Sub
		End If
		
		' load the subscription object
		subscription.Guid = Request.Form("invoice")
		If Len(subscription.Guid) > 0 Then 
			Call subscription.Load()
		Else
			str = str & vbCrLf & vbCrLf & "Error: Request.Form(invoice) was blank."
		End If
		
		Call CheckPaymentStatus(Request.Form("payment_status"), str, rv)
		returnError = returnError + rv
		
		Call CheckBusinessEmailAddress(Request.Form("business"), str, rv)
		returnError = returnError + rv
		
		Call CheckCurrency(Request.Form("mc_currency"), str, rv)
		returnError = returnError + rv
		
		Call CheckPrice(Request.Form("mc_gross"), subscription.Price, str, rv)
		returnError = returnError + rv
		
		' save registration
		If Len(subscription.Guid) > 0 Then
	
			str = str & vbCrLf & "subscription.Guid=" & subscription.Guid
			
			subscription.IsPaymentReceived = 1
			subscription.PaymentReceived = Request.Form("mc_gross")
			subscription.PayPalTransactionID = Request.Form("txn_id")
			subscription.PaypalPaymentStatus = Request.Form("payment_status")
			subscription.PaypalPaymentStatusReason = Request.Form("status_reason")
			If Len(Request.Form("test_ipn")) = 0 Then
				subscription.PayPalIsSandbox = 0
			Else
				subscription.PayPalIsSandbox = Request.Form("test_ipn")
			End If
			
			str = str & vbCrLf & "save registration here .."
			
			Call subscription.Save(rv)
			str = str & vbCrLf & "subscription.Save(rv)=" & rv
			
			Dim client		: Set client = New cClient
			client.ClientID = subscription.ClientID
			client.Load()
			If client.IsTrialAccount = 1 Then
				client.IsTrialAccount = 0
				Call client.Update(rv)
				str = str & vbCrLf & "client.Save(rv)=" & rv
			End If

			If rv = 0 Then
				Call DoEmailSubscriptionConfirmation(subscription)
			End If
		End If
		
		Call SendLogEmail(str)

		Set http = Nothing
		Set subscription = Nothing
	End Sub	

	Sub SendLogEmail(str)
		Dim email			: Set email = New cEmailSender
		
		Call email.SendMessage(Application.Value("ADMIN_EMAIL_ADDRESS"), "ipn@worshipschedule.com", "[" & Request.ServerVariables("SERVER_NAME") & "] ** IPN Transaction Report **", str)
		Set email = Nothing
	End Sub
	
	Function IsValidHttpResponse(httpResponse, str)
		IsValidHttpResponse = True
		
		If UCase(httpResponse) <> UCase("VERIFIED") Then	
			str = str & vbCrLf & vbCrLf & "Error: Paypal did not return 'VERIFIED'."
			IsValidHttpResponse = False
		End If
	End Function
	
	Sub CheckPaymentStatus(payment_status, str, outError)
		outError = 0
		
		If UCase(payment_status) <> UCase("Completed") Then
		
			If UCase(Request.Form("payment_status")) = UCase("Pending") Then
				' treate pending as paid ..
				str = str & vbCrLf & vbCrLf & "Error: payment_status <> 'Pending'"
			Else
				str = str & vbCrLf & vbCrLf & "Error: payment_status <> 'Completed'"
				outError = -1
			End If
		End If
	End Sub
	
	Sub CheckBusinessEmailAddress(business, str, outError)
		outError = 0
		
		If UCase(business) <> UCase(Application.Value("PAYPAL_BUSINESS_ID")) Then
			str = str & vbCrLf & vbCrLf & "Error: Possible spoofed post to /ipn.asp because business email address does not match a legit email for my paypal acct."
			outError = -1
		End If
	End Sub
	
	Sub CheckCurrency(mc_currency, str, outError)
		outError = 0

		If UCase(mc_currency) <> UCase("USD") Then
			str = str & vbCrLf & vbCrLf & "Error: Possible spoofed post to /ipn.asp because mc_currency <> 'USD'."
			outError = -1
		End If
	End Sub
	
	Sub CheckPrice(mc_gross, price, str, outError)
		outError = 0
		If Len(mc_gross) = 0 Then
			str = str & vbCrLf & vbCrLf & "Error: Could not check mc_gross as Request.Form(invoice) was blank."
			outError = -1
		End If	

		If FormatCurrency(mc_gross, 2) <> FormatCurrency(price, 2) Then
			str = str & vbCrLf & vbCrLf & "Error: Possible spoofed post to /ipn.asp because mc_gross not equal to item price originally passed to PayPal."
			outError = -1
		End If
	End Sub
		
	Sub DoEmailSubscriptionConfirmation(subscription)
		Dim str 
		Dim email			: Set email = New cEmailSender
		Dim fromAddress		: fromAddress = Application.Value("INFO_EMAIL_ADDRESS")
		Dim toAddress		: toAddress = subscription.ContactEmail
		Dim subject			: subject = "[" & Application.Value("APPLICATION_NAME") & "] ** Subscription Renewal Confirmation **"
		
		str = str & "Hello " & subscription.ContactNameFirst & " " & subscription.ContactNameLast
		str = str & vbCrLf & vbCrLf & "Thank you for your " & Application.Value("APPLICATION_NAME") & " subscription renewal! "
		str = str & "This email is to confirm your renewal for the " & subscription.NameClient & " account. "
		str = str & "Please check over your account contact details below to make sure that they are correct. "
		str = str & "If there is an error, please reply to this email message with any problems you wish to report. "
		str = str & vbCrLf & vbCrLf & "Transaction Information" 
		str = str & vbCrLf & String(40, "-")
		str = str & vbCrLf & "ID: " & Replace(Replace(subscription.Guid, "{", ""), "}", "")
		str = str & vbCrLf & subscription.ContactNameFirst & " " & subscription.ContactNameLast
		str = str & vbCrLf & subscription.ContactAddressLine1
		If Len(subscription.ContactAddressLine2) > 0 Then
			str = str & vbCrLf & Subscription.ContactAddressLine2
		End If
		str = str & vbCrLf & subscription.ContactCity & ", " & subscription.ContactStateCode & " " & subscription.ContactPostalCode
		str = str & vbCrLf & vbCrLf & "Phone: " & subscription.ContactPhone
		str = str & vbCrLf & "Email: " & subscription.ContactEmail
		str = str & vbCrLf & "Renewal: " & Trim(subscription.SubscriptionName) & vbCrLf
		str = str & "Paid: " & Trim(FormatCurrency(subscription.PaymentReceived, 2))
		str = str & vbCrLf & String(40, "-")

		str = str & vbCrLf & vbCrLf & "Again, if there are questions or problems with your renewal, please reply to this email or send them to mailto:" & fromAddress & ". "
		str = str & vbCrLf & vbCrLf & "Thanks again for your interest in " & Application.Value("APPLICATION_NAME") & "! " 

		str = str & vbCrLf & vbCrLf & Application.Value("APPLICATION_NAME") & " Sales"
		str = str & vbCrLf & "mailto:" & fromAddress
		str = str & vbCrLf & "http://" & Request.ServerVariables("SERVER_NAME")
		
		' send confirm to registrant
		Call email.SendMessage(toAddress, fromAddress, subject, str)
		' send confirm to admin
		Call email.SendMessage(fromAddress, Application.Value("SALES_EMAIL_ADDRESS"), "[" & Request.ServerVariables("SERVER_NAME") & "] ** Notify Subscription Renewal **", "Timestamp: " & Now() & vbCrLf & vbCrLf & "====Registration Details==========================" & vbCrLf & vbCrLf & str)
	End Sub
</script>

<!--#INCLUDE VIRTUAL="/_incs/class/client_subscription_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->

