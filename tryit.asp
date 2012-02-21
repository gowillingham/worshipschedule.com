<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Const SAMPLE_PROGRAM_NAME = "A Sample Team"

Dim m_bodyText

Sub OnPageLoad(ByRef page)
	page.MessageID = Request.QueryString("msgid")
	page.Action = Decrypt(Request.QueryString("act"))
	
	Set page.Member = New cMember
	Set page.Client = New cClient
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/outside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=Application.Value("APPLICATION_NAME") & " - Simple Web Scheduling for Volunteer Teams" %></title>
		<style type="text/css">
			.form	{width:535px;}
		</style>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><a href="/">Home</a> / Try It Free</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="column">
					<div class="feature-box">
						<h3>More Info ..</h3>
						<ul>
							<li><a href="/requirements.asp"><strong>System Requirements</strong></a></li>
							<li><a href="/pricing.asp"><strong>Pricing</strong></a></li>
						</ul>
					</div>
					<img style="" src="_images/hand_at.jpg" alt="Crazy Woman" />
				</div>
				<%=m_bodyText %>
			</div>
		</div>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_footer.asp"-->
	</body>
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Dim newMemberId
	Dim newClientId
	
	Call OnPageLoad(page)

	str = str & "<div style=""width:540px;"">" & ApplicationMessageToString(page.MessageID) & "</div>"
	page.MessageID = ""
	
	Select Case page.Action
		Case CONFIRM_ADDNEW
			str = str & PageCopyToString(page.Action)
			
		Case ADDNEW_RECORD
			
			If Request.Form("form_tryit_is_postback") = IS_POSTBACK Then
				Call LoadFormDataFromRequest(page)
				If ValidFormTryIt(page) Then
					str = str & FormConfirmTryItToString(page)
				Else
					str = str & FormTryItToString(page)
				End If
				
			ElseIf Request.Form("form_confirm_tryit_is_postback") = IS_POSTBACK Then
				Call LoadFormDataFromRequest(page)
				Call DoInsertTrialAccount(page, newMemberId, rv)
				Call SendNewAcountCredentials(newMemberId)
				Call NotifyApplicationAdmin(newMemberId)
						
				page.Action = CONFIRM_ADDNEW
				Response.Redirect(Request.ServerVariables("URL") & page.UrlParamsToString(False))
			End If
			
		Case Else
			str = str & FormTryItToString(page)
			
	End Select 

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoInsertTrialAccount(page, outMemberId, outError)
	Dim client						: Set client = New cClient
	Dim member						: Set member = New cMember
	Dim clientAdmin					: Set clientAdmin = New cClientAdmin
	
	' create client and make member the admin ..
	client.NameClient = page.name_client
	If Len(client.NameClient & "") = 0 Then client.NameClient = Application.Value("NEW_CLIENT_NAME")
	client.IsActive = 1
	client.IsTrialAccount = 1
	client.IsProfileComplete = 0
	client.TrialAccountLength = Application.Value("TRIAL_ACCOUNT_LENGTH")
	client.FileStorage = Application.Value("INITIAL_FILE_STORAGE")
	Call client.Add(outError)
	
	member.ClientId = client.ClientId
	member.NameLast = page.name_last
	member.NameFirst = page.name_first
	member.Email = page.email
	member.HomePageID = ADMIN_HOME_PAGE_ID
	Call member.QuickAdd("", outError)
	
	clientAdmin.ClientId = client.ClientId
	clientAdmin.MemberId = member.MemberId
	Call clientAdmin.Add(outError)
	
	Dim member1					: Set member1 = New cMember
	Dim member2					: Set member2 = New cMember
	Dim member3					: Set member3 = New cMember
	
	member1.NameLast = "Tufnel"				: member1.NameFirst = "Nigel"
	member2.NameLast = "St. Hubbins"		: member2.NameFirst = "David"
	member3.NameLast = "Smalls"				: member3.NameFirst = "Derek"
	
	' create a sample account
	Call DoInsertSampleProgram(client.ClientId, SAMPLE_PROGRAM_NAME, member1, member2, member3, outError) 

	outMemberId = member.MemberId
End Sub

Sub LoadFormDataFromRequest(page)
	page.name_last = Request.Form("name_last")
	page.name_first = Request.Form("name_first")
	page.email = Request.Form("email")
	page.email_retype = Request.Form("email_retype")
	
	page.has_terms = Request.Form("has_terms")
	
	page.name_client = Request.Form("name_client")
End Sub

Function ValidFormTryIt(page) 
	ValidFormTryIt = True

	If Not ValidData(page.name_first, True, 0, 50, "First Name", "") Then ValidFormTryIt = False
	If Not ValidData(page.name_last, True, 0, 50, "Last Name", "") Then ValidFormTryIt = False
	If Not ValidData(page.email, True, 0, 100, "Email Address", "email") Then ValidFormTryIt = False

	'check that email fields match
	If UCase(page.email) <> UCase(page.email_retype) Then
		AddCustomFrmError("The email fields must match.")
		ValidFormTryIt = False
	End If	
	
	'check that terms have been read
	If Len(page.has_terms) = 0 Then
		AddCustomFrmError("You need to have read the Terms of Service (or at least check the box!).")
		ValidFormTryIt = False
	End If
End Function

Function GetStartedWithYourEmailTextToString()
	Dim str
	
	str = str & "<h1>Get Started with Your Email Address</h1>"
	str = str & "<p>All you need is an email address to begin. "
	str = str & "Don't worry, we hate spam and unwanted emails! "
	str = str & "We'll be sending you only email associated with your " & Application.Value("APPLICATION_NAME") & " trial account. "
	str = str & "We will not sell or share your email address with anyone (our <a href=""/policies.asp"">privacy policy</a>). </p>"
	
	GetStartedWithYourEmailTextToString = str
End Function

Function Free45DayTrialTextToString()
	Dim str
	
	str = str & "<h1>Free 45 Day Trial Account</h1>"
	str = str & "<p>" & Application.Value("APPLICATION_NAME") & " makes it easy to manage and schedule all of your volunteers or volunteer teams. "
	str = str & "Sign up for a no obligation 45 day trial account with all features fully enabled. "
	str = str & "At the end of your trial you may upgrade to a full acccount at your option. </p>"
	
	Free45DayTrialTextToString = str
End Function

Function FormConfirmTryitToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & Free45DayTrialTextToString()
	
	str = str & "<div class=""form"">"
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-confirm-tryit"">"
	str = str & "<input type=""hidden"" name=""form_confirm_tryit_is_postback"" value= """ & IS_POSTBACK & """ />"
	str = str & "<input type=""hidden"" name=""name_first"" value=""" & html(page.name_first) & """ />"
	str = str & "<input type=""hidden"" name=""name_last"" value=""" & html(page.name_last) & """ />"
	str = str & "<input type=""hidden"" name=""email"" value=""" & html(page.email) & """ />"
	str = str & "<input type=""hidden"" name=""has_terms"" value=""" & page.has_terms & """ />"
	str = str & "<div class=""confirm-message"">"
	str = str & "<h3>Ok .. almost there!</h3>"
	str = str & "<p>Just one more click to create your free " & Application.Value("APPLICATION_NAME") & " account for managing your team's schedules "
	str = str & "(if you were intending to login to an existing account, our login page is <a href=""/member/login.asp"" title=""Login"">here</a>). </p>"
	str = str & "</div>"
	str = str & "<table><tbody>"
	str = str & "<tr><td class=""label"">Account&nbsp;name</td>"
	str = str & "<td><input type=""text"" name=""name_client"" value=""" & html(page.name_client) & """ class=""large"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">Provide a short name (optional) for your new account, "
	str = str & "usually <br />the name of the church or team you will be scheduling. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td><input type=""submit"" name=""submit"" value=""Save"" /></td></tr>"
	str = str & "</tbody></table></form></div>"
	
	FormConfirmTryitToString = str
End Function

Function FormTryItToString(page)
	Dim str
	
	Dim isChecked		: isChecked = ""
	If Len(page.has_terms) > 0 Then isChecked = " checked=""checked"""
	
	page.Action = ADDNEW_RECORD
	
	str = str & Free45DayTrialTextToString()
	str = str & GetStartedWithYourEmailTextToString()
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & Request.QueryString("URL") & page.UrlParamsToString(True) & """ method=""post"" id=""form-tryit"">"
	str = str & "<input type=""hidden"" name=""form_tryit_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><tr><td class=""label"">First Name</td>"
	str = str & "<td><input type=""text"" name=""name_first"" value=""" & page.name_first & """ /></td></tr>"
	str = str & "<tr><td class=""label"">Last Name</td>"
	str = str & "<td><input type=""text"" name=""name_last"" value=""" & page.name_last & """ /></td></tr>"
	str = str & "<tr><td class=""label"">Email</td>"
	str = str & "<td><input type=""text"" name=""email"" value=""" & page.email & """ /></td></tr>"
	str = str & "<tr><td class=""label"">Retype Email</td>"
	str = str & "<td><input type=""text"" name=""email_retype"" value=""" & page.email_retype & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>"
	str = str & "<input class=""checkbox"" type=""checkbox"" name=""has_terms""" & isChecked & " />"
	str = str & "I have read and agree to the <a href=""/terms.asp""><strong>Terms of Service</strong></a>.</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Signup"" /></td></tr>"
	str = str & "</table></form></div>"
		
	FormTryItToString = str
End Function

Function PageCopyToString(action)
	Dim str
	
	Select Case action
		Case CONFIRM_ADDNEW
			str = str & "<h1>Thank You - Welcome to " & Application.Value("APPLICATION_NAME") & "!</h1>"
			str = str & "<p>A no obligation, fully functioning " & Application.Value("APPLICATION_NAME") & " trial account has been created just for you. "
			str = str & "Please check your email in the next few minutes for your login information "
			str = str & "(you may need to adjust any spam blocking software to accept email from addresses "
			str = str & "that end in '" & Application.Value("ROOT_EMAIL_DOMAIN") & "'). </p>"
			str = str & "<p>If you have any questions, or need help getting around the service, " 
			str = str & "please shoot a note to <a href=""mailto:" & Application.Value("SUPPORT_EMAIL_ADDRESS") & """>" & Application.Value("SUPPORT_EMAIL_ADDRESS") & "</a> "
			str = str & "Thanks again for your interest - we're looking forward to working with you! </p>"

		Case Else
			str = str & Free45DayTrialTextToString()
			str = str & GetStartedWithYourEmailTextToString()

	End Select
	
	PageCopyToString = str
End Function

Sub NotifyApplicationAdmin(memberId)
	Dim member				: Set member = New cMember
	member.MemberId = memberId
	Call member.Load()
	
	Dim body
	Dim emailSender			: Set emailSender = New cEmailSender
	Dim subject				: subject = "[" & Application.Value("APPLICATION_NAME") & "] ** Trial Account Notification **"
	Dim fromAddress			: fromAddress = Application.Value("ADMIN_EMAIL_ADDRESS")
	
	body = body & "New Client Trial Registration - " & Now()
	body = body & vbCrLf & vbCrLf & "----------------------------------------------------"
	body = body & vbCrLf & "ClientID: " & member.ClientID
	body = body & vbCrLf & "Full Name: " & member.NameFirst & " " & member.NameLast
	body = body & vbCrLf & "Email: " & member.Email
	body = body & vbCrLf & "----------------------------------------------------"
	
	Call emailSender.SendMessage(Application.Value("ADMIN_EMAIL_ADDRESS"), Application.Value("ADMIN_EMAIL_ADDRESS"), subject, body)
End Sub

Sub SendNewAcountCredentials(memberId)
	Dim body
	
	Dim member			: Set member = New cMember
	member.MemberId = memberId
	Call member.Load()
	
	Dim emailSender		: Set emailSender = New cEmailSender
	Dim appName			: appName = Application.Value("APPLICATION_NAME")
	
	Dim subject			: subject = "** " & appName & " Login Information **"
	Dim fromAddress		: fromAddress = Application.Value("NO_REPLY_EMAIL_ADDRESS")
	
	body = body & "Hello " & member.NameFirst & " " & member.NameLast
	
	body = body & vbCrLf & vbCrLf
	body = body & "Welcome to " & appName & "! "
	body = body & "Your login information is included below (your password is case sensitive). You should change your login name "
	body = body & "or password to something easier to remember - it can be done at any time by editing your " & appName & " member profile. Please save "
	body = body & "or print out this email and store it in a secure place."
	body = body & vbCrLf & vbCrLf
	body = body &  "---------------------------------"
	body = body & vbCrLf 
	body = body & "Login Name: " & member.NameLogin
	body = body & vbCrLf 
	body = body & "Password: " & member.PWord
	body = body & vbCrLf 
	body = body & "---------------------------------"
	body = body & vbCrLf & vbCrLf
	body = body & "If you experience any problems or have questions about your account, please contact " & appName & " support at "
	body = body & "mailto:" & Application.Value("SUPPORT_EMAIL_ADDRESS") & ". " 
	body = body & "Additional support information may be found on the " & appName & " support page here .."
	body = body & vbCrLF
	body = body & "http://" & Request.ServerVariables("SERVER_NAME") & "/support.asp"
	body = body & vbCrLf & vbCrLf
	body = body & "Before you begin with " & Application.Value("APPLICATION_NAME") & ", take a look at an overview of how "
	body = body & Application.Value("APPLICATION_NAME") & " organizes your team members and schedules here .."
	body = body & vbCrLf
	body = body & "http://" & Request.ServerVariables("SERVER_NAME") & "/help/help.asp?hid=12"
	body = body & vbCrLf & vbCrLf
	body = body & "A guide to getting started with " & Application.Value("APPLICATION_NAME") & " including a walkthrough of how to "
	body = body & "set up your first program is here .."
	body = body & vbCrLf
	body = body & "http://" & Request.ServerVariables("SERVER_NAME") & "/help/help.asp?hid=14"
	body = body & vbCrLf & vbCrLf
	body = body & appName & " online help is available here .."
	body = body & vbCrLF
	body = body & "http://" & Request.ServerVariables("SERVER_NAME") & "/help/help.asp"
	body = body & vbCrLf & vbCrLf
	body = body & "To get started, login here .."
	body = body & vbCrLf
	body = body & "http://" & Request.ServerVariables("SERVER_NAME") & "/member/login.asp"
	body = body & vbCrLf & vbCrLf
	body = body & "Upon first login, you will be asked to review your " & appName & " account profile and your " & appName & " "
	body = body & "member profile. "
	body = body & "Thanks again for your interest in " & appName & ", and we look forward to working with you over your trial period."
	body = body & vbCrLf & vbCrLf
	body = body & "The " & appName & " Team"
	
	body = body & EmailDisclaimerToString("")

	Call emailSender.SendMessage(member.Email, fromAddress, subject, body)
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_SendCredentials.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_DoInsertSampleProgram.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_build_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_admin_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_sender_cls.asp"-->
<%
Class cPage
	Public MessageID
	Public Action
	
	' form data
	Public HasTerms
	
	' objects
	Public Member
	Public Client
	
	' form data
	Public name_last
	Public name_first
	Public email
	Public email_retype
	Public has_terms
	Public name_client
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		
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
		
		c.Action = Action
		c.MessageID = MessageID
		c.HasTerms = HasTerms
		
		Set c.Member = Member
		Set c.Client = Client
		
		Set Clone = c
	End Function
End Class
%>

