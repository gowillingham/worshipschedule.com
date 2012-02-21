<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-account"
Dim m_pageHeaderText	: m_pageHeaderText = "&nbsp;"
Dim m_impersonateText	: m_impersonateText = ""
Dim m_pageTitleText		: m_pageTitleText = ""
Dim m_topBarText		: m_topBarText = "&nbsp;"
Dim m_bodyText			: m_bodyText = ""
Dim m_tabStripText		: m_tabStripText = ""
Dim m_tabLinkBarText	: m_tabLinkBarText = ""
Dim m_appMessageText	: m_appMessageText = ""
Dim m_acctExpiresText	: m_acctExpiresText = ""

Sub OnPageLoad(ByRef page)
	Dim sess			: Set sess = New cSession
	sess.SessionID = Request.Cookies("sid")
	Call CheckSession(sess, PERMIT_ALL)
	
	' don't allow view this page unless logged in ..
	If Len(sess.MemberID) = 0 Then
		Response.Redirect("/member/login.asp?msgid=" & 1011)	
	End If
	
	page.MessageID = Request.QueryString("msgid")
	page.Action = Decrypt(Request.QueryString("act"))

	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	Set page.ClientSubscription = New cClientSubscription
	page.ClientSubscription.Guid = Request.QueryString("csid")
	If Len(page.ClientSubscription.Guid) > 0 Then page.ClientSubscription.Load()
		
	' set the view tokens
	m_appMessageText = ApplicationMessageToString(page.MessageID)
	page.MessageID = ""
	Call SetTopBar(page)
	Call SetPageHeader(page)
	Call SetPageTitle(page)
	Call SetTabLinkBar(page)
	Call SetTabList(m_pageTabLocation, page)
	Call SetImpersonateText(sess)
	Call SetAccountNotifier(sess)
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/inside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<!--#INCLUDE VIRTUAL="/_incs/script/javascript/javascript_server_variable_wrapper.asp"-->
		<script src="http://www.google.com/jsapi" type="text/javascript" language="javascript"></script>
		<script language="javascript" type="text/javascript">
			google.load("jquery", "1.2.6");
		</script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/tablesorter/jquery.tablesorter.min.js"></script>

		<script language="javascript" type="text/javascript">
			$(document).ready(function(){
				// stripe tables on page
				$(".grid table tr:nth-child(even)").addClass("alt");
				
				// pop-up window for paypal graphic ..
				$("#paypal-graphic").click(function(){
					console.log("clicked paypal ..");
					window.open('https://www.paypal.com/us/cgi-bin/webscr?cmd=xpt/cps/popup/OLCWhatIsPayPal-outside','olcwhatispaypal','toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes, width=400, height=350');
				});
			});
		</script>
		<link href="../_incs/script/jquery/plugins/tablesorter/themes/blue/style.css" rel="stylesheet" type="text/css" />		
		<style type="text/css">
			.message, .details, .form, .grid {width:622px;}
		</style>
		<title><%=m_pageTitleText%></title>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)
	
	Select Case page.Action
		Case DELETE_CLIENT_SUBSCRIPTION
			Call DoDeleteSubscription(page.ClientSubscription, rv)
			page.Action = "": page.ClientSubscription.Guid = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case CANCEL_CLIENT_ACCOUNT
			If Request.Form("form_confirm_cancel_account_is_postback") = IS_POSTBACK Then
				Call DoDisableClient(page.Client, rv)
				Call NotifyApplicationAdmin(page)
				Call ConfirmCloseAccountByEmail(page)
				page.Action = "": page.MessageID = 2033
				Response.Redirect("/member/login.asp" & page.UrlParamsToString(False))
			Else
				str = str & ConfirmCancelAccountToString(page)
			End If
						
		Case ACCEPT_PAYMENT
			str = str & ConfirmPaymentGridToString(page)
			str = str & FormPaypalToString(page)
			
		Case UPDATE_CLIENT_SUBSCRIPTION
			If Request.Form("form_subscription_is_postback") = IS_POSTBACK Then
				Call LoadDataFromPost(page)
				If ValidFormSubscription(page) Then
					Call DoUpdateSubscription(page.ClientSubscription, rv)
					Call DoUpdateContactInfo(page.Client, rv)
					page.Action = ACCEPT_PAYMENT
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormSubscriptionToString(page)
				End If
			Else
				str = str & FormSubscriptionToString(page)
			End If
			
		Case INSERT_CLIENT_SUBSCRIPTION
			If Request.Form("form_subscription_is_postback") = IS_POSTBACK Then
				Call LoadDataFromPost(page)
				If ValidFormSubscription(page) Then
					Call DoInsertSubscription(page, rv)
					page.Action = ACCEPT_PAYMENT
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormSubscriptionToString(page)
				End If
			Else
				str = str & FormSubscriptionToString(page)
			End If
			
		Case UPDATE_CLIENT_CONTACT_INFO
			If Request.Form("form_subscription_is_postback") = IS_POSTBACK Then
				Call LoadDataFromPost(page)
				If ValidFormSubscription(page) Then
					Call DoUpdateContactInfo(page.Client, rv)
					page.Action = "": page.ClientSubscription.Guid = ""
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormSubscriptionToString(page)
				End If
			Else
				str = str & FormSubscriptionToString(page)
			End If
			
		Case Else
			Call DoDeleteOrphanSubscriptions(page, rv)
			str = str & AccountSummaryToString(page)
			
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoDisableClient(client, outError)
	client.IsActive = 0 
	Call client.Update(outError)
End Sub

Sub DoDeleteSubscription(clientSubscription, outError)
	Call clientSubscription.Delete(outError)
End Sub

Sub DoUpdateContactInfo(client, outError)
	Call client.Update(outError)
End Sub

Sub DoUpdateSubscription(clientSubscription, outError)
	Call clientSubscription.Save(outError)
End Sub

Sub DoInsertSubscription(page, outError)
	Dim returnError				: returnError = 0
	outError = 0
	
	Dim dateTime				: Set dateTime = New cFormatDate

	' set the subscription properties
	page.ClientSubscription.ClientID = page.Client.ClientID
	page.ClientSubscription.IsPaymentReceived = 0
	page.ClientSubscription.TermStart = page.ClientSubscription.GetNewSubscriptionStartDate(page.Client.TrialExpiresDate)
	
	' update client properties ..
	Call page.Client.Update(returnError)
	If returnError <> 0 Then outError = -1
	
	' update clientSubscription properties
	Call page.ClientSubscription.Add(returnError)
	If returnError <> 0 Then outError = -1
End Sub

Sub DoDeleteOrphanSubscriptions(page, outError)
	Dim i
	outError = 0

	page.ClientSubscription.ClientID = page.Client.ClientID
	Dim list				: list = page.ClientSubscription.List()
	
	' 0-ClientSubscriptionID  17-IsPaymentReceived 18-Price 19-PaymentReceived 
	
	If Not IsArray(list) Then Exit Sub
	For i = 0 To UBound(list,2)
		If list(17,i) <> 1 Then
			page.ClientSubscription.Guid = list(0,i)
			Call page.ClientSubscription.Delete(outError)
		End If
	Next
End Sub

Sub ConfirmCloseAccountByEmail(page)
	Dim str 
	Dim email		: Set email = New cEmailSender
	Dim toAddress	: toAddress = page.Member.Email
	Dim fromAddress	: fromAddress = Application.Value("INFO_EMAIL_ADDRESS")
	Dim subject		: subject = "[" & Application.Value("APPLICATION_NAME") & "] ** Confirm Closing " & page.Client.NameClient & " Account **"
	
	str = str & "Hello " & page.Member.NameFirst & " " & page.Member.NameLast
	str = str & vbCrLf & vbCrLf & "We're sorry to see you go! "
	str = str & vbCrLf & vbCrLf & "This message is to confirm that you have closed the " & page.Client.NameClient & " " & Application.Value("APPLICATION_NAME") & " account. "
	str = str & "Please allow several weeks to receive a refund of any pro-rated time remaining in your existing subscription. "
	str = str & "If you change your mind and wish to re-open your account, you may reply to this email with that request "
	str = str & "(we will save your account information for up to six months in case you wish to re-start your account). "
	str = str & "If you wish to have your account information purged from our servers before that time, you may contact mailto:" & Application.Value("SUPPORT_EMAIL_ADDRESS") & " with that request. "
	str = str & vbCrLf & vbCrLf & "Thanks for your interest in " & Application.Value("APPLICATION_NAME") & ", and please let us know if we can help you again in the future. "
	str = str & vbCrLf & vbCrLf & Application.Value("APPLICATION_NAME") & " Sales"
	str = str & vbCrLf & "mailto:" & Application.Value("INFO_EMAIL_ADDRESS")
	
	str = str & EmailDisclaimerToString(page.Client.NameClient)
	
	Call email.SendMessage(toAddress, fromAddress, subject, str)
End Sub

Sub NotifyApplicationAdmin(page)
	Dim str 
	Dim email		: Set email = New cEmailSender
	Dim toAddress	: toAddress = Application.Value("ADMIN_EMAIL_ADDRESS")
	Dim fromAddress	: fromAddress = Application.Value("APPLICATION_EMAIL_ADDRESS")
	Dim subject		: subject = "[" & Application.Value("APPLICATION_NAME") & "] ** Close " & page.Client.NameClient & " Account Request **"
	
	str = str & "Timestamp: " & Now()
	str = str & vbCrLf & vbCrLf & "Notify Close Worshipschedule Account"
	str = str & vbCrLf & String(60, "-")
	str = str & vbCrLf & "ClientID: " & page.Client.ClientID
	str = str & vbCrLf & "Client Name: " & page.Client.NameClient
	str = str & vbCrLf & "Member Name: " & page.Member.NameLast & ", " & page.Member.NameFirst
	str = str & vbCrLf & "Member Email: " & page.Member.Email
	str = str & vbCrLf & "MemberID: " & page.Member.MemberID
	str = str & vbCrLf & String(60, "-")

	Call email.SendMessage(toAddress, fromAddress, subject, str)
End Sub

Sub SetCurrentAccountInfo(list, hasCurrentSubscription, accountExpiresDate, accountStartsDate)
	Dim i
	
	Dim guidStringList				: guidStringList = ""
	Dim guidList
	
	' get an array of unexpired, paid subscriptions
	' 0-Guid 14-TermLength 15-TermStart 16-PaymentType 17-IsPaymentReceived 18-Price 19-PaymentReceived 
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			
			' check for payment received
			If list(17,i) = 1 Then
			
				' if not expired then add to list
				If DateAdd("m", list(14,i), list(15,i)) > Now() Then
					guidStringList = guidStringList & list(0,i) & ","
				End If
			End If
		Next
		
		If Len(guidStringList) > 0 Then 
			guidStringList = Left(guidStringList, Len(guidStringList) - 1)
			guidList = Split(guidStringList, ",")
		End If
	End If
	
	If IsArray(guidList) Then
		hasCurrentSubscription = True
		
		' get accountStartsDate from first row ..
		For i = 0 To UBound(list,2)
			If CStr(guidList(LBound(guidList))) = CStr(list(0,i)) Then
				accountStartsDate = list(15,i)
			End If
		Next
		
		' get accountExpiresDate from last row ..
		For i = 0 To UBound(list,2)
			If CStr(guidList(UBound(guidList))) = CStr(list(0,i)) Then
				accountExpiresDate = DateAdd("d", -1, DateAdd("m", list(14,i), list(15,i)))
			End If
		Next
	Else
		hasCurrentSubscription = False
		
		' traverse array backwards to get latest row that 
		' is paid and expires in past ..
		If IsArray(list) Then
			For i = UBound(list,2) To 0 Step -1
				If list(17,i) = 1 Then
					accountExpiresDate = DateAdd("d", -1, DateAdd("m", list(14,i), list(15,i)))
					Exit For
				End If
			Next
		End If
	End If
End Sub

Function ExpirationInfoToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate
	
	Dim trialExpiresDate	: trialExpiresDate = DateAdd("d", page.Client.TrialAccountLength, page.Client.DateCreated)
	Dim isTrialAccount		: isTrialAccount = False
	If page.Client.IsTrialAccount = 1 Then isTrialAccount = True
	
	' isTrialExpired = True if account is not a trial 
	' or is trial and is expired
	Dim isTrialExpired		: isTrialExpired = True
	If page.Client.IsTrialAccount = 1 Then
		If trialExpiresDate > Now() Then
			isTrialExpired = False
		End If
	End If
	
	Dim hasCurrentSubscription
	Dim accountExpiresDate
	Dim accountStartsDate	
	
	page.ClientSubscription.ClientID = page.Client.ClientID
	Call SetCurrentAccountInfo(page.ClientSubscription.List(), hasCurrentSubscription, accountExpiresDate, accountStartsDate)
	
	' unexpired trial account
	If isTrialAccount And Not isTrialExpired Then
		str = str & "Your " & Application.Value("APPLICATION_NAME") & " trial account will expire on "
		str = str & dateTime.Convert(trialExpiresDate, "DDDD MMM dd, YYYY") & ". "
	End If
	
	' expired trial account
	If isTrialAccount And isTrialExpired Then
		str = str & "Your " & Application.Value("APPLICATION_NAME") & " trial account expired on "
		str = str & dateTime.Convert(trialExpiresDate, "DDDD MMM dd, YYYY") & ". "
	End If
		
	' unexpired subscription account
	If hasCurrentSubscription And (accountExpiresDate > Now()) Then
		str = str & "Your current " & Application.Value("APPLICATION_NAME") & " subscription expires on "
		str = str & dateTime.Convert(accountExpiresDate, "DDDD MMM dd, YYYY") & ". "
	End If
	
	' expired subscription account
	If hasCurrentSubscription And (accountExpiresDate < Now()) Then
		str = str & "Your " & Application.Value("APPLICATION_NAME") & " subscription expired on "
		str = str & dateTime.Convert(accountExpiresDate, "DDDD MMM dd, YYYY") & ". "
	End If
	
	str = "<p>" & str & "</p>"
	
	ExpirationInfoToString = str
End Function

Function FormExtendCancelToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	
	str = str & "<table><tr><td>"
	pg.Action = INSERT_CLIENT_SUBSCRIPTION
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-extend-account"" style=""display:inline;"">"
	str = str & "<input type=""submit"" name=""submit"" value=""Extend Account"" />"
	str = str & "<input type=""hidden"" name=""form_extend_account_is_postback"" value=""" & IS_POSTBACK & """ /></form>"
	pg.Action = CANCEL_CLIENT_ACCOUNT
	str = str & "&nbsp;<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-cancel-account"" style=""display:inline;"">"
	str = str & "<input type=""submit"" name=""submit"" value=""Cancel Account"" />"
	str = str & "<input type=""hidden"" name=""form_cancel_account_is_postback"" value=""" & IS_POSTBACK & """ /></form>"
	str = str & "</td></tr></table>"
	
	FormExtendCancelToString = str
End Function

Function AccountHistoryGridToString(page)
	Dim str, i
	Dim dateTime				: Set dateTime = New cFormatDate
	
	page.ClientSubscription.ClientID = page.Client.ClientID
	Dim list					: list = page.ClientSubscription.List()
	
	Dim icon
	Dim expireDate
	Dim expireText
	Dim trialExpireDate
	Dim termText
	
	' 0-ClientSubscriptionID 1-ClientID 2-NameClient 3-ContactNameFirst 4-ContactNameLast 5-ContactAddressLine1
	' 6-ContactAddressLine2 7-ContactCity 8-ContactStateID 9-ContactStateCode 10-ContactStateLongName 11-ContactPostalCode 
	' 12-ContactPhone 13-ContactEmail 14-TermLength 15-TermStart 16-PaymentType 17-IsPaymentReceived 18-Price 
	' 19-PaymentReceived 20-DateCreated 21-DateModified 22-SubscriptionID 23-SubscriptionName

	str = str & "<div class=""grid""><table><thead><tr>"
	str = str & "<th scope=""col"" style=""width:1%;"">"
	str = str & "<input type=""checkbox"" class=""checkbox"" checked=""checked"" disabled=""disabled"" /></th>"
	str = str & "<th scope=""col"" style=""width:1%;"">Start&nbsp;date</th>"
	str = str & "<th scope=""col"">History</th>"
	str = str & "<th scope=""col"" style=""width:1%;"">Expiration</th>"
	str = str & "</tr></thead><tbody>"
	
	' trial account row
	trialExpireDate = DateAdd("d", page.Client.TrialAccountLength, page.Client.DateCreated)
	icon = "money.png"
	expireText = ""
	If trialExpireDate < Now() Then
		expireText = "[Expired]"
		icon = "money_delete.png"
	End If
	str = str & "<tr><td><input type=""checkbox"" disabled=""disabled"" checked=""checked"" /></td>"
	str = str & "<td>" & dateTime.Convert(page.Client.DateCreated, "MM/DD/YYYY") & "</td>"
	str = str & "<td><img class=""icon"" src=""/_images/icons/" & icon & """ alt="""" />"
	str = str & "<strong>" & page.Client.TrialAccountLength & " Day Free Trial " & Application.Value("APPLICATION_NAME") & " Subscription</strong> " & expireText & "</td>"
	str = str & "<td>" & dateTime.Convert(trialExpireDate, "MM/DD/YYYY") & "</td></tr>"
	
	' paid subscription rows
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			If list(17,i) <> "0" Then
				expireText = ""
				icon = "money.png"
				expireDate = DateAdd("d", -1, DateAdd("m", list(14,i), list(15,i)))
				If expireDate < Now() Then
					expireText = " <span style=""font-weight:normal;"">[Expired]</span>"
					icon = "money_delete.png"
				End If
				
				str = str & "<tr><td><input type=""checkbox"" disabled=""disabled"" checked=""checked"" /></td>"
				str = str & "<td>" & dateTime.Convert(list(15,i), "MM/DD/YYYY") & "</td>"
				str = str & "<td><img class=""icon"" src=""/_images/icons/" & icon & """ alt="""" />"
				str = str & "<strong>" & list(23,i) & "</strong> " & expireText & "</td>"
				str = str & "<td>" & dateTime.Convert(expireDate, "MM/DD/YYYY") & "</td></tr>"
			End If
		Next	
	End If
	str = str & "</tbody></table></div>"

	AccountHistoryGridToString = str
End Function

' fill passed in member with billing info ..
Sub SetBillingInfoForClient(clientId, member, hasPaidClientSubscription)
	Dim clientSubscription			: Set clientSubscription = New cClientSubscription
	Dim clientAdmin					: Set clientAdmin = New cClientAdmin
	
	' check for existing subscription that has been paid ..
	clientSubscription.ClientID = clientId
	clientSubscription.Guid = clientSubscription.GetLastPaidSubscription()
	
	hasPaidClientSubscription = False
	If Len(clientSubscription.Guid) > 0 Then 
		hasPaidClientSubscription = True
	End If
	
	If hasPaidClientSubscription Then
		clientSubscription.Load()
	
		member.NameLast = clientSubscription.ContactNameLast
		member.NameFirst = clientSubscription.ContactNameFirst
		member.PhoneHome = clientSubscription.ContactPhone
		member.Email = clientSubscription.ContactEmail
		member.AddressLine1 = clientSubscription.ContactAddressLine1
		member.AddressLine2 = clientSubscription.ContactAddressLine2
		member.StateCode = clientSubscription.ContactStateCode
		member.City = clientSubscription.ContactCity
		member.PostalCode = clientSubscription.ContactPostalCode
	Else
	
		' load contact info from oldest admin ..
		clientAdmin.ClientID = clientId
		member.MemberId = clientAdmin.GetOldest()
		Call member.Load()
	End If
End Sub

Function BillingInfoSummaryToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim member			: Set member = New cMember
	Dim hasPaidClientSubscription

	Call SetBillingInfoForClient(page.Client.ClientId, member, hasPaidClientSubscription)
	
	str = str & "<p>This is the person " & Application.Value("APPLICATION_NAME") & " will contact about any billing related issues. "
	str = str & "They do not need to have a " & Application.Value("APPLICATION_NAME") & " account. </p>"

	Dim address			: address = member.AddressToString

	str = str & "<div class=""simple""><table><tbody>"
	str = str & "<tr><td class=""label"" style=""width:1%;""><strong>Full&nbsp;Name</strong></td>"
	str = str & "<td>" & html(member.NameLast & ", " & member.NameFirst) & "</td></tr>"
	str = str & "<tr><td class=""label""><strong>Phone</strong></td>"
	str = str & "<td>" & html(member.PhoneHome) & "</td></tr>"
	str = str & "<tr><td class=""label""><strong>Email</strong></td>"
	str = str & "<td>" & html(member.Email) & "</td></tr>"
	str = str & "<tr><td class=""label""><strong>Address</strong></td>"
	str = str & "<td>" & address & "</td></tr>"
	
	' only allow edit if subscription exists to edit ..
	If hasPaidClientSubscription Then
		pg.Action = UPDATE_CLIENT_CONTACT_INFO
		str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td>"
		str = str & "<td><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-change-contact-info"" style=""display:inline"">"
		str = str & "<input type=""submit"" name=""submit"" value=""Change Info"" />"
		str = str & "<input type=""hidden"" name=""form_change_contact_info_is_postback"" value=""" & IS_POSTBACK & """ /></form></td></tr>"
	End If
	str = str & "</tbody></table></div>"
	
	BillingInfoSummaryToString = str
End Function

Function AccountSummaryToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3><ul>"
	pg.Action = INSERT_CLIENT_SUBSCRIPTION
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Extend my account</a></li>"
	pg.Action = CANCEL_CLIENT_ACCOUNT
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Close my account</a></li>"
	str = str & "</ul></div>"
	
	str = str & "<div class=""details"">"
	str = str & m_appMessageText
	str = str & "<h3>" & Application.Value("APPLICATION_NAME") & " Account History</h3>"
	str = str & ExpirationInfoToString(page)
	str = str & AccountHistoryGridToString(page)
	str = str & FormExtendCancelToString(page)
	str = str & "<p class=""dot-line"">&nbsp;</p>"
	
	str = str & "<h3>Billing Contact Information</h3>"
	str = str & BillingInfoSummaryToString(page)
	
	str = str & "</div>"
	
	AccountSummaryToString = str
End Function

Function ConfirmCancelAccountToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim header
	Dim body
	header = "Please confirm this action!"
	
	str = str & "<div class=""tip-box""><h3>Before closing your account!</h3><ul>"
	str = str & "<li>Backup any files that you have placed into your account's file storage. </li>"
	str = str & "<li>Export and save your member contact list. </li>. "
	str = str & "</ul></div>"
	
	
	body = body & "You are about to cancel your <strong>" & html(page.Client.NameClient) & "</strong> " & Application.Value("APPLICATION_NAME") & " account. "
	body = body & "Please allow several weeks to receive a refund for any pro-rated time remaining in your existing subscription. "
	body = body & "<br /><br />Select <strong>Close Account</strong> to close and disable your " & html(page.Client.NameClient) & " " & Application.Value("APPLICATION_NAME") & " account immediately. "
	str = str & CustomApplicationMessageToString(header, body, "Confirm")
	
	str = str & "<form method=""post"" action=""" & pg.Url & pg.UrlParamsToString(True) & """ name=""form-confirm-cancel-account"">"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Close Account"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""form_confirm_cancel_account_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</p></form>"


	ConfirmCancelAccountToString = str
End Function

Function ConfirmPaymentGridToString(page)
	Dim str
	Dim pg						: Set pg = page.Clone()
	Dim dateTime				: Set dateTime = New cFormatDate
	
	Dim TermExpireDate
	termExpireDate = DateAdd("m", page.ClientSubscription.TermLength, page.ClientSubscription.TermStart)
	termExpireDate = DateAdd("d", -1, termExpireDate)
	
	Dim address
	address = Html(page.ClientSubscription.ContactAddressLine1)
	If Len(page.ClientSubscription.ContactAddressLine2) > 0 Then
		address = address & "<br />" & Html(page.ClientSubscription.ContactAddressLine2)
	End If
	address = address & "<br />" & Html(page.ClientSubscription.ContactCity & ", " & page.ClientSubscription.ContactStateCode & " " & page.ClientSubscription.ContactPostalCode)

	str = str & "<div class=""tip-box"" style=""border:none;text-align:center;"">"	
	str = str & "<a id=""paypal-graphic"" href=""#"">"
	str = str & "<img style=""border:none;"" src=""https://www.paypal.com/en_US/i/bnr/vertical_solution_PPeCheck.gif"" alt=""Solution Graphics"" /></a>"
	str = str & "<p style=""text-align:center;"">Make your " & Application.Value("APPLICATION_NAME") & " payments "
	str = str & "online with either a credit card or your Paypal account. </p></div>"
	
	Dim header
	Dim body
	header = "Thank you - please confirm your contact information!"
	body = body & "Please check that the information you've provided is correct. "
	body = body & "Select <strong>Edit</strong> if you wish to change anything before you provide your payment information. "
	str = str & CustomApplicationMessageToString(header, body, "Confirm")
	
	str = str & "<h3>" & Application.Value("APPLICATION_NAME") & " Account Renewal for " & Html(page.Client.NameClient) & "</h3>"
	str = str & "<div class=""simple""><table>"
	str = str & "<tr><td class=""label"">Church </td>"
	str = str & "<td>" & html(page.ClientSubscription.NameClient) & "</td></tr>"
	str = str & "<tr><td class=""label"">Contact&nbsp;Name </td>"
	str = str & "<td>" & html(page.ClientSubscription.ContactNameLast & ", " & page.ClientSubscription.ContactNameFirst) & "</td></tr>"
	str = str & "<tr><td class=""label"">Address </td>"
	str = str & "<td>" & address & "</td></tr>"
	str = str & "<tr><td class=""label"">Contact&nbsp;Phone </td>"
	str = str & "<td>" & page.ClientSubscription.ContactPhone & "</td></tr>"
	str = str & "<tr><td class=""label"">Contact&nbsp;Email </td>"
	str = str & "<td>" & page.ClientSubscription.ContactEmail & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">Subscription&nbsp;Length </td>"
	str = str & "<td>" & page.ClientSubscription.TermLength & " months (expires " & dateTime.Convert(termExpireDate, "DDDD, MMMM dd, YYYY") & ")</td></tr>"
	str = str & "<tr><td class=""label"">Price </td>"
	str = str & "<td>" & FormatCurrency(page.ClientSubscription.Price, 2) & " </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	
	str = str & "<tr><td>&nbsp;</td>"
	pg.Action = UPDATE_CLIENT_SUBSCRIPTION
	str = str & "<td><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-request-update-subscription"" style=""display:inline;"">"
	str = str & "<input type=""submit"" name=""submit"" value=""Change"" /></form>"
	pg.Action = DELETE_CLIENT_SUBSCRIPTION
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-request-delete-subscription"" style=""display:inline;"">"
	pg.ClientSubscription.Guid = "": pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "</form></td></tr></table></div>"

	ConfirmPaymentGridToString = str 
End Function

Function FormPaypalToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	Dim header
	Dim body
	header = "Final step - authorize your payment!"
	body = body & "When you select the <strong>Pay Now</strong> button you will be redirected by secure server to PayPal to make your payment. "
	body = body & "Once redirected, you will have the option to pay by credit card or with a PayPal account. "
	body = body & "You are not required to have a PayPal account. "
	
	str = str & "<div class=""grid"">"
	str = str & CustomApplicationMessageToString(header, body, "Confirm")
	
	str = str & "<h3>Authorize Payment</h3>"
	str = str & "<table class=""grid invoice""><thead>"
	str = str & "<tr><th>Item</th><th>Quantity</th><th>Total Price</th><th>&nbsp;</th></tr></thead>"
	str = str & "<tr style=""vertical-align:middle;""><td><strong>" & page.ClientSubscription.SubscriptionName & "</strong></td><td>1</td><td>" & FormatCurrency(page.ClientSubscription.Price, 2) & "</td>"
	str = str & "<td style=""width:1%;text-align:right;"">"
	str = str & "<form style=""display:inline;"" action=""" & Application.Value("PAYPAL_GATEWAY") & """ method=""post"">"
	str = str & "<input type=""hidden""  name=""cmd"" value=""_xclick"" />"
	str = str & "<input type=""hidden""  name=""business"" value=""" & Application.Value("PAYPAL_BUSINESS_ID") & """ />"
	str = str & "<input type=""hidden""  name=""item_name"" value=""" & page.ClientSubscription.SubscriptionName & """ />"
	str = str & "<input type=""hidden""  name=""invoice"" value=""" & page.ClientSubscription.GUID & """ />"
	str = str & "<input type=""hidden""  name=""amount"" value=""" & page.ClientSubscription.Price & """ />"
	str = str & "<input type=""hidden""  name=""no_shipping"" value=""0"" />"
	
	' paypal returns paid transactions to accounts page so display confirm message
	pg.Action = "": pg.MessageID = 2023
	str = str & "<input type=""hidden"" name=""return"" value=""http://" & Request.ServerVariables("SERVER_NAME") & "/client/account.asp" & pg.UrlParamsToString(True) & """ />"

	' paypal returns user-cancelled transactions to ACCEPT_PAYMENT screen with no message ..
	pg.Action = ACCEPT_PAYMENT: pg.MessageID = ""
	str = str & "<input type=""hidden""  name=""cancel_return"" value=""http://" & Request.ServerVariables("SERVER_NAME") & "/client/account.asp" & pg.UrlParamsToString(True) & """ />"

	' url for ipn
	str = str & "<input type=""hidden""  name=""notify_url"" value=""http://" & Request.ServerVariables("SERVER_NAME") & "/client/ipn.asp"" />"

	str = str & "<input type=""hidden""  name=""no_note"" value=""1"" />"
	str = str & "<input type=""hidden""  name=""currency_code"" value=""USD"" />"
	str = str & "<input type=""hidden""  name=""lc"" value=""US"" />"
	str = str & "<input type=""hidden""  name=""bn"" value=""PP-BuyNowBF"" />"
	str = str & "<input type=""image"" src=""https://www.sandbox.paypal.com/en_US/i/btn/btn_paynow_SM.gif"" name=""submit"" alt=""Make payments with PayPal - it's fast, free and secure!"" />"
	str = str & "<img alt="""" border=""0"" src=""https://www.sandbox.paypal.com/en_US/i/scr/pixel.gif"" width=""1"" height=""1"" />"
	' pre-populate credit card fields ..
	str = str & "<input type=""hidden"" name=""first_name"" value=""" & page.ClientSubscription.ContactNameFirst & """ />"
	str = str & "<input type=""hidden"" name=""last_name"" value=""" & page.ClientSubscription.ContactNameLast & """ />"
	str = str & "<input type=""hidden"" name=""address1"" value=""" & page.ClientSubscription.ContactAddressLine1 & """ />"
	str = str & "<input type=""hidden"" name=""address2"" value=""" & page.ClientSubscription.ContactAddressLine2 & """ />"
	str = str & "<input type=""hidden"" name=""city"" value=""" & page.ClientSubscription.ContactCity & """ />"
	str = str & "<input type=""hidden"" name=""state"" value=""" & page.ClientSubscription.ContactStateCode & """ />"
	str = str & "<input type=""hidden"" name=""zip"" value=""" & page.ClientSubscription.ContactPostalCode & """ />"
	str = str & "<input type=""hidden"" name=""night_phone_a"" value=""" & Left(page.ClientSubscription.ContactPhoneRaw, 3) & """ />"
	str = str & "<input type=""hidden"" name=""night_phone_b"" value=""" & Right(Left(page.ClientSubscription.ContactPhoneRaw, 6), 3) & """ />"
	str = str & "<input type=""hidden"" name=""night_phone_c"" value=""" & Right(page.ClientSubscription.ContactPhoneRaw, 4) & """ />"
	
	str = str & "</form></td></tr></table></div>"
	
	
	FormPaypalToString = str
End Function

Function StatesDropdownToString(val)
	Dim str
	
	Dim states				: Set states = New cState
	states.CountryID = UNITED_STATES_COUNTRY_CODE
	
	Dim options				: options = states.OptionListToString(val)
	
	str = str & "<select name=""state_id"">"
	str = str & "<option value="""">&nbsp;</option>"
	str = str & options & "</select>"
	
	StatesDropdownToString = str
End Function

Function ValidFormSubscription(page)
	ValidFormSubscription = True
	
	If (page.Action = INSERT_CLIENT_SUBSCRIPTION) Or (page.Action = UPDATE_CLIENT_SUBSCRIPTION) Then
		If Not ValidData(page.ClientSubscription.SubscriptionID, True, 0, 0, "Subscription Package", "numbers") Then ValidFormSubscription = False
	End If
	
	If Not ValidData(page.Client.ContactNameFirst, True, 0, 50, "First name", "") Then ValidFormSubscription = False
	If Not ValidData(page.Client.ContactNameLast, True, 0, 50, "Last name", "") Then ValidFormSubscription = False
	If Not ValidData(page.Client.ContactAddressLine1, True, 0, 100, "Address", "") Then ValidFormSubscription = False
	If Not ValidData(page.Client.ContactAddressLine2, False, 0, 100, "Address(2)", "") Then ValidFormSubscription = False
	If Not ValidData(page.Client.ContactCity, True, 0, 100, "City", "") Then ValidFormSubscription = False
	If Not ValidData(page.Client.ContactStateID, True, 0, 0, "State", "numbers") Then ValidFormSubscription = False
	If Not ValidData(page.Client.ContactPostalCode, True, 0, 10, "Postal code", "zip") Then ValidFormSubscription = False
	If Not ValidData(page.Client.ContactPhone, True, 0, 15, "Phone", "phone") Then ValidFormSubscription = False
	If Not ValidData(page.Client.ContactEmail, True, 0, 100, "Email", "email") Then ValidFormSubscription = False

	'check that email fields match
	If UCase(page.Client.ContactEmail) <> UCase(page.Client.ContactEmailRetype) Then
		AddCustomFrmError("Email and email retype must match exactly.")
		ValidFormSubscription = False
	End If	
End Function

Sub LoadDataFromPost(page)
	page.Client.ContactNameFirst = Trim(Request.Form("name_first"))
	page.Client.ContactNameLast = Trim(Request.Form("name_last"))
	page.Client.ContactAddressLine1 = Trim(Request.Form("address_line_1"))
	page.Client.ContactAddressLine2 = Trim(Request.Form("address_line_2"))
	page.Client.ContactCity = Trim(Request.Form("city"))
	page.Client.ContactStateID = Trim(Request.Form("state_id"))
	page.Client.ContactPostalCode = Trim(Request.Form("postal_code"))
	page.Client.ContactPhone = Trim(Request.Form("phone"))
	page.Client.ContactEmail = Trim(Request.Form("email"))
	page.Client.ContactEmailRetype = Trim(Request.Form("email_retype"))
	
	page.ClientSubscription.SubscriptionID = Trim(Request.Form("subscription_id"))
End Sub

Function TermDropdownToString(subscriptionID)
	Dim str, i
	
	Dim subscription:		: Set subscription = New cSubscription
	Dim list				: list = subscription.List()
	Dim selected			: selected = ""
	Dim disabled			: disabled = True
	
	' 0-SubscriptionID 1-Name 2-Desc 3-TermLength 4-Price 5-IsEnabled
	
	str = str & "<select name=""subscription_id"" style=""width:auto;"">"
	str = str & "<option value="""">" & html("< Select Term >") & "</option>"
	For i = 0 To UBound(list,2)
		selected = ""
		If CStr(list(0,i) & "") = CStr(subscriptionId & "") Then selected = " selected=""selected"""
		disabled = True
		If list(5,i) <> 0 Then disabled = False
		
		If Not disabled Then
			str = str & "<option value=""" & list(0,i) & """" & selected & ">"  & list(1,i) & "</option>"
		End If
	Next
	str = str & "</select>"
	
	TermDropdownToString = str
End Function

Function FormSubscriptionToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>Info!</h3><p>"
	str = str & "Extend your subscription by online payment with a credit card or paypal account. "
	str = str & "</p></div>"
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-subscription"">"
	str = str & "<table><tbody>"
	str = str & "<tr><td class=""label"">Church</td>"
	str = str & "<td><input type=""text"" value=""" & html(page.Client.NameClient) & """ class=""disabled medium"" disabled=""disabled"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	
	If (page.Action = INSERT_CLIENT_SUBSCRIPTION) Or (pg.Action = UPDATE_CLIENT_SUBSCRIPTION) Then
		str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Subscription Package") & "</td>"
		str = str & "<td>" & TermDropdownToString(page.ClientSubscription.SubscriptionID) & "</td></tr>"
		str = str & "<tr><td>&nbsp;</td>"
		str = str & "<td class=""hint"">Select the length for your subscription. </td></tr>"
		str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	End If
	
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "First name") & "</td>"
	str = str & "<td><input type=""text"" name=""name_first"" value=""" & page.Client.ContactNameFirst & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Last name") & "</td>"
	str = str & "<td><input type=""text"" name=""name_last"" value=""" & page.Client.ContactNameLast & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Address") & "</td>"
	str = str & "<td><input type=""text"" name=""address_line_1"" value=""" & page.Client.ContactAddressLine1 & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td class=""label"">Address(2)</td>"
	str = str & "<td><input type=""text"" name=""address_line_2"" value=""" & page.Client.ContactAddressLine2 & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "City") & "</td>"
	str = str & "<td><input type=""text"" name=""city"" value=""" & page.Client.ContactCity & """ /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "State") & "</td>"
	str = str & "<td>" & StatesDropdownToString(page.Client.ContactStateId) & "</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Postal code") & "</td>"
	str = str & "<td><input type=""text"" name=""postal_code"" value=""" & page.Client.ContactPostalCode & """ class=""small"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Phone") & "</td>"
	str = str & "<td><input type=""text"" name=""phone"" value=""" & page.Client.ContactPhone & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Email") & "</td>"
	str = str & "<td><input type=""text"" name=""email"" value=""" & page.Client.ContactEmail & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Email retype") & "</td>"
	str = str & "<td><input type=""text"" name=""email_retype"" value=""" & page.Client.ContactEmailRetype & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""form_subscription_is_postback"" value=""" & IS_POSTBACK & """ /></td></tr>"
	str = str & "</tbody></table></form></div>"
	
	FormSubscriptionToString = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone
	
	Dim accountLink
	pg.Action = ""
	accountLink = "<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Account</a> / "
	
	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case CANCEL_CLIENT_ACCOUNT
			str = str & accountLink & "Close"
		Case INSERT_CLIENT_SUBSCRIPTION
			str = str & accountLink & "Extend"
		Case UPDATE_CLIENT_SUBSCRIPTION
			str = str & accountLink & "Extend"
		Case ACCEPT_PAYMENT
			str = str & accountLink & "Extend"
		Case UPDATE_CLIENT_CONTACT_INFO
			str = str & accountLink & "Billing Contact"
		Case Else
			str = str & "Account"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Select Case page.Action
		Case Else
			str = str & "<li>&nbsp;</li>"
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_subscription_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/subscription_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_admin_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/state_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID

	' encrypted
	Public Action

	' objects
	Public Member
	Public Client
	Public ClientSubscription
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		If Len(ClientSubscription.Guid) > 0 Then str = str & "csid=" & ClientSubscription.Guid & amp
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		
		If Len(str) > 0 Then 
			str = Left(str, Len(str) - Len(amp))
		Else
			' qstring needs at least one param in case more params are appended ..
			str = str & "noparm=true"
		End If
		str = "?" & str
		
		UrlParamsToString = str
	End Function
	
	Public Function Clone()
		Dim c
		Set c = New cPage

		c.MessageID = MessageID
		c.Action = Action
				
		Set c.Member = Member
		Set c.Client = Client
		Set c.ClientSubscription = ClientSubscription
		
		Set Clone = c
	End Function
End Class
%>

