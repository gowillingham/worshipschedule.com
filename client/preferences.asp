<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "admin-settings"
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
	Call CheckSession(sess, PERMIT_ADMIN)
	
	page.MessageID = Request.QueryString("msgid")
	page.Action = Decrypt(Request.QueryString("act"))

	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()

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
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		
		<script language="javascript" type="text/javascript">
		
			$(document).ready(function(){
				// attach submit() to save button for form
				$("#save-form-client a").each(function(){
					$(this).click(function(){
						$("#form-client").submit();
						return false;
					});
				});
				
				// set click event for all checkboxes in form ..
				$("#form-required-info .checkbox").each(function(){
					$(this).click(function(){
						var checkbox = this
						var qs = "key=" + checkbox.name + "&value=" + checkbox.checked;
						qs = qs + "&id=" + $("input#id").val();
						
						// show progress indicator for relevant td ..
						$("#form-required-info td .checkbox").each(function(){
							if (this.name == checkbox.name){
								$(this).parent("td").addClass("loading");
							}
						})
						
						// post to worker page ..
						$.ajax({
							type: "POST",
							url: "/_incs/script/ajax/_update_admin_preferences.asp",
							data: qs,
							success: function(){
								$("#form-required-info td .checkbox").each(function(){
									$(this).parent("td").removeClass("loading");
								})
							},
							error: function(){
								$("#form-required-info td .checkbox").each(function(){
									$(this).parent("td").removeClass("loading");
								})
								$("#form-required-info td .checkbox").each(function(){
									if (this.name == checkbox.name){
										$(this).parent("td").addClass("error");
									}
								})
							}
						});	
					});
				});
				
			});
			
		</script>
		<style type="text/css">
			.form, .message, .details, .summary {width:650px;}
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
		Case INSERT_SAMPLE_PROGRAM
			If Request.Form("form_sample_program_is_postback") = IS_POSTBACK Then
				Call LoadSampleProgramFromRequest(page)
				If ValidSampleProgram(page) Then
					Call DoCreateSampleProgram(page, rv)
					Response.Redirect("/admin/programs.asp" & page.UrlParamsToString(False))
				Else
					str = str & FormSampleProgram(page)
				End If
			Else
				str = str & FormSampleProgram(page)
			End If
		
		Case UPDATE_RECORD
			If Request.Form("form_client_is_postback") = IS_POSTBACK Then
				Call LoadClientFromForm(page.Client)
				If ValidFormClient(page) Then
					Call UpdateClient(page.Client, rv)
					Select Case rv
						Case 0
							page.MessageID = 2004
						Case Else
							page.MessageID = 2003
					End Select
					Response.Redirect(page.Url & page.UrlParamsToString(False))
				Else
					str = str & FormEditClientToString(page)
				End If
			Else
				str = str & FormEditClientToString(page)
			End If
			
		Case Else
			str = str & ClientSummaryToString(page)
	End Select

	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoCreateSampleProgram(page, outError)
	Dim member1					: Set member1 = New cMember
	Dim member2					: Set member2 = New cMember
	Dim member3					: Set member3 = New cMember
	
	member1.NameFirst = page.first_name_1
	member1.NameLast = page.last_name_1
	
	member2.NameFirst = page.first_name_2
	member2.NameLast = page.last_name_2
	
	member3.NameFirst = page.first_name_3
	member3.NameLast = page.last_name_3
	
	Call DoInsertSampleProgram(page.Client.ClientId, page.program_name, member1, member2, member3, outError)
End Sub

Sub UpdateClient(client, outError)
	Call client.Update(outError)
End Sub

Function FormEditClientToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>Info</h3><p>"
	str = str & "Your church profile is info that " & Application.Value("APPLICATION_NAME") & " will display to account members about your church. "
	str = str & "</p></div>"
	
	str = str & m_appMessageText
	str = str & "<div class=""form"">"
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-client"">"
	str = str & ErrorToString()
	str = str & "<table>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "Church name") & "</td>"
	str = str & "<td><input type=""text"" name=""name_client"" value=""" & html(page.Client.NameClient) & """ class=""large"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">The way " & Application.Value("APPLICATION_NAME") & " should display the name of your church <br />to your members. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Address line 1") & "</td>"
	str = str & "<td><input type=""text"" name=""address_line_1"" value=""" & html(page.Client.AddressLine1) & """ class=""large"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Address line 2") & "</td>"
	str = str & "<td><input type=""text"" name=""address_line_2"" value=""" & html(page.Client.AddressLine2) & """ class=""large"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "City") & "</td>"
	str = str & "<td><input type=""text"" name=""city"" value=""" & html(page.Client.City) & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "State") & "</td>"
	str = str & "<td>" & StateDropdownToString(page.Client.StateID, UNITED_STATES_COUNTRY_CODE) & "</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Zip code") & "</td>"
	str = str & "<td><input type=""text"" name=""postal_code"" value=""" & html(page.Client.PostalCode) & """ class=""small"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Main phone") & "</td>"
	str = str & "<td><input type=""text"" name=""phone_main"" value=""" & page.Client.PhoneMain & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Alternate phone") & "</td>"
	str = str & "<td><input type=""text"" name=""phone_alternate"" value=""" & page.Client.PhoneAlternate & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Fax") & "</td>"
	str = str & "<td><input type=""text"" name=""phone_fax"" value=""" & page.Client.PhoneFax & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Email") & "</td>"
	str = str & "<td><input type=""text"" name=""email"" value=""" & html(page.Client.Email) & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Retype email") & "</td>"
	str = str & "<td><input type=""text"" name=""email_retype"" value=""" & html(page.Client.EmailRetype) & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(False, "Home page") & "</td>"
	str = str & "<td><input type=""text"" name=""home_page"" value=""" & html(page.Client.HomePage) & """ class=""medium"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">Your church's home page address <br />like this - http://www.example.com. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""form_client_is_postback"" value=""" & IS_POSTBACK & """ /></td></tr>"
	str = str & "</table></form></div>"
	
	FormEditClientToString = str
End Function

Sub LoadClientFromForm(client)
	client.NameClient = Request.Form("name_client")
	client.AddressLine1 = Request.Form("address_line_1")
	client.AddressLine2 = Request.Form("address_line_2")
	client.Email = Request.Form("email")
	client.EmailRetype = Request.Form("email_retype")
	client.City = Request.Form("city")
	client.StateId = Request.Form("state_id")
	client.PostalCode = Request.Form("postal_code")
	client.PhoneMain = Request.Form("phone_main")
	client.PhoneAlternate = Request.Form("phone_alternate")
	client.PhoneFax = Request.Form("phone_fax")
	client.HomePage = Request.Form("home_page")
End Sub

Function StateDropdownToString(id, countryID)
	Dim str, i
	Dim state		: Set state = New cState
	state.countryID = countryId
	Dim arr			: arr = state.List()
	Dim selected	: selected = ""
	
	If Not IsArray(arr) Then Exit Function
	
	Dim list()
	Redim list(1,UBound(arr,2))
	For i = 0 To UBound(arr,2)
		list(0,i) = arr(0,i)
		list(1,i) = Html(arr(1,i) & " - " & arr(2,i))
	Next
	
	str = str & "<select name=""state_id"">"
	str = str & "<option value="""">&nbsp;</option>"
	For i = 0 To UBound(list,2)
		selected = ""
		If CStr(id & "") = CStr(list(0,i)) Then selected = " selected=""selected"""
		str = str & "<option value=""" & list(0,i) & """" & selected & ">" & Server.HTMLEncode(list(1,i)) & "</option>"
	Next
	str = str & "</select>"
	
	StateDropdownToString = str
End Function

Function ValidFormClient(page)
	ValidFormClient = True
	If Application.Value("FORM_VALIDATION_OFF") Then Exit Function	

	If Not ValidData(page.Client.NameClient, True, 0, 100, "Church Name", "") Then ValidFormClient = False
	If Not ValidData(page.Client.AddressLine1, False, 0, 100, "Address Line 1", "") Then ValidFormClient = False
	If Not ValidData(page.Client.AddressLine2, False, 0, 100, "Address Line 2", "") Then ValidFormClient = False
	If Not ValidData(page.Client.City, False, 0, 100, "City", "") Then ValidFormClient = False
	If Not ValidData(page.Client.StateID, False, 0, 3, "State", "") Then ValidFormClient = False
	If Not ValidData(page.Client.PostalCode, False, 0, 10, "Postal Code", "zip") Then ValidFormClient = False
	
	If Not ValidData(page.Client.PhoneMain, False, 0, 14, "Main Phone", "phone") Then ValidFormClient = False
	If Not ValidData(page.Client.PhoneAlternate, False, 0, 14, "Alternate Phone", "phone") Then ValidFormClient = False
	If Not ValidData(page.Client.PhoneFax, False, 0, 14, "Fax", "phone") Then ValidFormClient = False
	
	If Not ValidData(page.Client.Email, False, 0, 100, "Email", "email") Then ValidFormClient = False
	If UCase(page.Client.Email) <> UCase(page.Client.EmailRetype) Then
		AddCustomFrmError("Email and Retype Email must match exactly.")
		ValidFormClient = False
	End If

	If Not ValidData(page.Client.HomePage, False, 0, 100, "Home Page", "") Then ValidFormClient = False
End Function

Function OtherStuffForSummaryToString(page)
	Dim str
	Dim dateTime					: Set dateTime = New cFormatDate
	
	Dim memberText		: memberText = page.Client.MemberCount & " member"
	If page.Client.MemberCount <> 1 Then memberText = memberText & "s"
	
	str = str & "<ul><li>Account created on " & dateTime.Convert(page.Client.DateCreated, "DDDD MMMM dd, YYYY") & ". </li>"
	If page.Client.IsTrialAccount = 1 Then
		str = str & "<li>Trial account expires " & dateTime.Convert(page.Client.TrialExpiresDate, "DDDD MMMM dd, YYYY") & ". </li>"
	Else
		str = str & "<li>Account expires "  & dateTime.Convert(page.Client.SubscriptionExpiresDate, "DDDD MMMM dd, YYYY") & ". </li>"
	End If
	str = str & "<li>Your account has " & memberText & ". </li></ul>"

	OtherStuffForSummaryToString = str
End Function

Function ContactDetailsForSummaryToString(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	Dim dateTime			: Set dateTime = New cFormatDate

	Dim href				: href = ""
	Dim address				: address = ""
	Dim phoneList			: phoneList = ""

	If Len(page.Client.AddressLine1) > 0 Then address = address & page.Client.AddressLine1
	If Len(page.Client.AddressLine2) > 0 Then 
		If Len(address) > 0 Then address = address & "<br />"
		address = address & page.Client.AddressLine2
	End If
	If Len(page.Client.City & page.Client.StateCode & page.Client.PostalCode) > 0 Then
		If Len(page.Client.City) = 0 Then page.Client.City = "????"
		If Len(page.Client.StateCode) = 0 Then page.Client.StateCode = "??"
		If Len(page.Client.PostalCode) = 0 Then page.Client.PostalCode = "??"
		
		If Len(address) > 0 Then address = address & "<br />"
		address = address & page.Client.City & ", " & page.Client.StateCode & " " & page.Client.PostalCode
	End If
	
	If Len(page.Client.PhoneMain) > 0 Then phoneList = page.Client.PhoneMain & " (main)"
	If Len(page.Client.PhoneAlternate) > 0 Then 
		If Len(phoneList) > 0 Then phoneList = phoneList & "<br />"
		phoneList = phoneList & page.Client.PhoneAlternate & " (alternate)"
	End If
	If Len(page.Client.PhoneFax) > 0 Then
		If Len(phoneList) > 0 Then phoneList = phoneList & "<br />"
		phoneList = phoneList & page.Client.PhoneFax & " (fax)"
	End If
	
	If Len(address) > 0 Then 
		str = str & "<p>" & html(page.Client.NameClient) & "<br />" & address & "</p>"
	Else
		pg.Action = UPDATE_RECORD
		href = pg.Url & pg.UrlParamsToString(True)
		
		str = str & "<p class=""alert"">A contact address is missing from your profile. "
		str = str & "You can fix that <a href=""" & href & """>here</a>. </p>"
	End If

	If Len(page.Client.Email) > 0 Then
		str = str & "<p><a href=""mailto:" & html(page.Client.Email) & """><strong>" & html(page.Client.Email) & "</strong></a></p>"
	Else
		pg.Action = UPDATE_RECORD
		href = pg.Url & pg.UrlParamsToString(True)
		
		str = str & "<p class=""alert"">Your account email address is missing from your profile. "
		str = str & "You can fix that <a href=""" & href & """>here</a>. </p>"
	End If

	If Len(phoneList) > 0 Then
		str = str & "<p>" & phoneList & "</p>"
	Else
		pg.Action = UPDATE_RECORD
		href = pg.Url & pg.UrlParamsToString(True)
		
		str = str & "<p class=""alert"">Contact phone numbers are missing from your profile. "
		str = str & "You can fix that <a href=""" & href & """>here</a>. </p>"
	End If

	ContactDetailsForSummaryToString = str
End Function

Function AdministratorGridForSummaryToString(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	
	Dim clientAdmin	: Set clientAdmin = New cClientAdmin
	clientAdmin.ClientId = page.Client.ClientId
	
	Dim list		: list = clientAdmin.List()
	Dim rows		: rows = ""
	Dim href		: href = "#"
	Dim alt			: alt = ""
	Dim count		: count = 0
		
	' 0-ClientID 1-MemberID 2-NameFirst 3-NameLast 4-NameClient 5-DateCreated 
	' 6-DateModified 7-ClientAdminID

	str = str & "<p>Members in this list have administrator permissions over your " & Application.Value("APPLICATION_NAME") & " account. "
	str = str & "They can make changes to any of your programs, events, schedules, or members. </p>"
	str = str & "<div class=""grid""><table><thead><tr>"
	str = str & "<th>Member</th><th>&nbsp;</th></tr></thead>"
	str = str & "<tbody>"
	For i = 0 To UBound(list,2)
		alt = ""					: If count Mod 2 > 0 Then alt = " class=""alt"""
		
		pg.Action = SHOW_MEMBER_DETAILS: pg.MemberId = list(1,i)
		Href = "/admin/profile.asp" & pg.UrlParamsToString(True)
		
		str = str & "<tr" & alt & "><td><img class=""icon"" src=""/_images/icons/user_red.png"" alt="""" />"
		str = str & "<strong>" & html(list(4,i)) & "</strong> | "
		str = str & "<a href=""" & href & """ title=""Details""><strong>" & html(list(3,i) & ", " & list(2,i)) & "</strong></a></td>"
		str = str & "<td class=""toolbar""><a href=""" & href & """ title=""Details"">"
		str = str & "<img src=""/_images/icons/magnifier.png"" alt="""" /></a></td></tr>"
	
		count = count + 1
	Next
	str = str & "</tbody></table></div>"	

	AdministratorGridForSummaryToString = str		
End Function

Function FormRequiredMemberInfoToString(page)
	Dim str
	Dim pg						: Set pg = page.Clone()
	
	Dim userSettings			: Set userSettings = New cUserSettings
	Call userSettings.Load(page.Client.ClientID)
	
	str = str & "<p>Here you can set what information is required when members create or edit their accounts. "
	str = str & Application.Value("APPLICATION_NAME") & " requires each member to provide their email address and name, "
	str = str & "but the rest of your member's profile information can be set as optional or required at your discretion. </p>"
	
	str = str & "<form method=""post"" action=""" & pg.Url & pg.UrlParamsToString(True) & """ id=""form-required-info"">"
	str = str & "<input type=""hidden"" id=""id"" value=""" & page.Client.ClientID & """ />"
	
	' required fields so non-editable ..
	str = str & "<table id=""required-info"">"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""first_name"" checked=""checked"" disabled=""disabled"" />"
	str = str & "First name</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""last_name"" checked=""checked"" disabled=""disabled"" />"
	str = str & "Last name</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""email"" checked=""checked"" disabled=""disabled"" />"
	str = str & "Email</td></tr>"
	
	str = str & "<tr><td>&nbsp;</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""PhoneHome"" "
	str = str & "value=""" & userSettings.GetSetting("PhoneHome") & """" & userSettings.IsChecked("PhoneHome") & " />"
	str = str & "Home phone</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""PhoneMobile"" "
	str = str & "value=""" & userSettings.GetSetting("PhoneMobile") & """" & userSettings.IsChecked("PhoneMobile") & " />"
	str = str & "Mobile phone</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""PhoneAlternate"" "
	str = str & "value=""" & userSettings.GetSetting("PhoneAlternate") & """" & userSettings.IsChecked("PhoneAlternate") & " />"
	str = str & "Alternate phone</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""PhoneMultiple"" "
	str = str & "value=""" & userSettings.GetSetting("PhoneMultiple") & """" & userSettings.IsChecked("PhoneMultiple") & " />"
	str = str & "Require members to provide more than one phone number	</td></tr>"
	
	str = str & "<tr><td>&nbsp;</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""AddressLine1"" "
	str = str & "value=""" & userSettings.GetSetting("AddressLine1") & """" & userSettings.IsChecked("AddressLine1") & " />"
	str = str & "Street address</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""City"" "
	str = str & "value=""" & userSettings.GetSetting("City") & """" & userSettings.IsChecked("City") & " />"
	str = str & "City</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""State"" "
	str = str & "value=""" & userSettings.GetSetting("State") & """" & userSettings.IsChecked("State") & " />"
	str = str & "State</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""PostalCode"" "
	str = str & "value=""" & userSettings.GetSetting("PostalCode") & """" & userSettings.IsChecked("PostalCode") & " />"
	str = str & "Zip code</td></tr>"
	str = str & "<tr><td>&nbsp;</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""Gender"" "
	str = str & "value=""" & userSettings.GetSetting("Gender") & """" & userSettings.IsChecked("Gender") & " />"
	str = str & "Gender</td></tr>"
	str = str & "<tr><td><input type=""checkbox"" class=""checkbox"" name=""DateOfBirth"" "
	str = str & "value=""" & userSettings.GetSetting("DateOfBirth") & """" & userSettings.IsChecked("DateOfBirth") & " />"
	str = str & "Date of birth</td></tr>"
	str = str & "</table></form>"
	
	FormRequiredMemberInfoToString = str
End Function

Function CodeSnippetsForSummaryToString(page)
	Dim str
	Dim href			: href="http://" & Request.ServerVariables("SERVER_NAME") & "/newmember.asp?gid=" & page.Client.Guid 
	
	str = str & "<p>Use this html code snippet if you wish to have a page on your own site where your team members add themselves to your " & Application.Value("APPLICATION_NAME") & " account. "
	str = str & "Just place this html code in any page on your church's own website to allow your members to register for a "
	str = str & Application.Value("APPLICATION_NAME") & " account by clicking the link. </p>"
	str = str & "<p>You may modify any part of this snippet to suit your own needs except the portion between the quotes for the link href attribute. </p>"

	str = str & "<p><code>"
	str = str & html("<p>") & "<br />"
	str = str & "&nbsp;&nbsp;" & html("<h3>Create a " & page.Client.NameClient & " account with " & Application.Value("APPLICATION_NAME") & "</h3>") & "<br />"
	str = str & "&nbsp;&nbsp;" & html("Click <a href=""" & href & """>here!</a>") & "<br />"
	str = str & html("</p>")
	str = str & "</code></p>"
	
	str = str & "<p>Alternately, you could copy and paste the following text link into an email message and send it to anyone that you would like to have create a " & html(page.Client.NameClient) & " account with " & Application.Value("APPLICATION_NAME") & " "
	str = str & "(watch for line-wrapping, as that may break the link). </p>"
	str = str & "<p><code>" & html(href) & "</code></p>"

	CodeSnippetsForSummaryToString = str
End Function

Function CreateSampleProgramForSummaryToString(page)
	Dim str
	Dim pg						: Set pg = page.Clone()
	
	Dim href					: href = ""
	
	pg.Action = INSERT_SAMPLE_PROGRAM
	href = pg.Url & pg.UrlParamsToString(True)
	
	str = str & "<p>Click <a href=""" & href & """ title=""Sample program"">here</a> to create a sample program "
	str = str & "complete with sample members, schedules and events that you can use to practice with " & Application.Value("APPLICATION_NAME") & ". "
	str = str & "You can delete or recreate this sample program at any time. </p>"
	
	CreateSampleProgramForSummaryToString = str
End Function

Function ClientSummaryToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>I want to ..</h3><ul>"
	pg.Action = UPDATE_RECORD
	str = str & "<li><a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Change my " & html(page.Client.NameClient) & " church profile information</a></li>"
	pg.Action = ""
	str = str & "<li><a href=""/admin/members.asp" & pg.UrlParamsToString(True) & """>Add members to my account</a></li>"
	If page.Member.IsAdmin Then
		pg.Action = ""
		str = str & "<li><a href=""/client/administrators.asp" & pg.UrlParamsToString(True) & """>Add or change administrators for my account</a></li>"
	End If
	str = str & "</ul></div>"
	
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & Application.Value("APPLICATION_NAME") & " account settings</h3>"
	
	str = str & "<h5 class=""contact"">Profile/contact information</h5>"
	str = str & ContactDetailsForSummaryToString(page)
	
	str = str & "<h5 class=""administrator"">Account administrators</h5>"
	str = str & AdministratorGridForSummaryToString(page)
	
	str = str & "<h5 class=""settings"">Required member account information</h5>"
	str = str & FormRequiredMemberInfoToString(page)
	
	str = str & "<h5 class=""settings"">Account sign-up code (html)</h5>"
	str = str & CodeSnippetsForSummaryToString(page)
	
	str = str & "<h5 class=""settings"">Create a sample account</h5>"
	str = str & CreateSampleProgramForSummaryToString(page)
	
	str = str & "<h5 class=""other-stuff"">Other stuff</h5>"
	str = str & OtherStuffForSummaryToString(page)
	
	str = str & "</div>"
	
	ClientSummaryToString = str
End Function

Function ValidSampleProgram(page)
	ValidSampleProgram = True
	
	If Not ValidData(page.program_name, True, 1, 100, "Program name", "") Then ValidSampleProgram = False
	
	If Not ValidData(page.first_name_1, True, 1, 50, "First name 1", "") Then ValidSampleProgram = False
	If Not ValidData(page.last_name_1, True, 1, 50, "Last name 1", "") Then ValidSampleProgram = False
	If Not ValidData(page.first_name_2, True, 1, 50, "First name 2", "") Then ValidSampleProgram = False
	If Not ValidData(page.last_name_2, True, 1, 50, "Last name 2", "") Then ValidSampleProgram = False
	If Not ValidData(page.first_name_3, True, 1, 50, "First name 3", "") Then ValidSampleProgram = False
	If Not ValidData(page.last_name_3, True, 1, 50, "Last name 3", "") Then ValidSampleProgram = False
End Function

Sub LoadSampleProgramFromRequest(page)
	page.first_name_1 = Request.Form("first_name_1")
	page.last_name_1 =  Request.Form("last_name_1")
	page.first_name_2 =  Request.Form("first_name_2")
	page.last_name_2 =  Request.Form("last_name_2")
	page.first_name_3 =  Request.Form("first_name_3")
	page.last_name_3 =  Request.Form("last_name_3")
	
	page.program_name = Request.Form("program_name")
End Sub

Function FormSampleProgram(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	
	Dim href
	
	href = pg.Url & pg.UrlParamsToString(True)
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>Use this form to generate a sample program with members, events, and schedules. "
	str = str & "You can use this made-up account to practice or experiment with " & Application.Value("APPLICATION_NAME") & ". </p></div>"
	
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & href & """ method=""post"" id=""form-sample-program"">"
	str = str & "<input type=""hidden"" name=""form_sample_program_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<table><tbody>"
	
	str = str & "<tr><td class=""label"">Program name</td>"
	str = str & "<td><input type=""text"" name=""program_name"" value=""" & page.program_name & """ class=""large"" /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">Make up a name for your sample program. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	
	str = str & "<tr><td class=""label"">Member first/last name</td>"
	str = str & "<td><input type=""text"" name=""first_name_1"" value=""" & page.first_name_1 & """ class=""small""/>"
	str = str & "&nbsp;&nbsp;<input type=""text"" name=""last_name_1"" value=""" & page.last_name_1 & """ /></td></tr>"
	str = str & "<tr><td class=""label"">Member first/last name</td>"
	str = str & "<td><input type=""text"" name=""first_name_2"" value=""" & page.last_name_2 & """ class=""small""/>"
	str = str & "&nbsp;&nbsp;<input type=""text"" name=""last_name_2"" value=""" & page.last_name_2 & """ /></td></tr>"
	str = str & "<tr><td class=""label"">Member first/last name</td>"
	str = str & "<td><input type=""text"" name=""first_name_3"" value=""" & page.last_name_3 & """ class=""small""/>"
	str = str & "&nbsp;&nbsp;<input type=""text"" name=""last_name_3"" value=""" & page.last_name_3 & """ /></td></tr>"
	str = str & "<tr><td>&nbsp;</td><td class=""hint"">Make up names for the three members of your <br />sample program. </td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	
	pg.Action = ""
	href = pg.Url & pg.UrlParamsToString(True)
	
	str = str & "<tr><td>&nbsp;</td><td><input type=""submit"" name=""submit"" value=""Save"" />"
	str = str & "&nbsp;&nbsp;<a href=""" & href & """>Cancel</a></td></tr>"
	
	str = str & "</tbody></table></form></div>"
	
	FormSampleProgram = str
End Function

Sub SetPageHeader(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	Dim preferencesLink
	preferencesLink = "<a href=""/client/preferences.asp"">Preferences</a> / "

	str = str & "<a href=""/admin/overview.asp"">Admin Home</a> / "
	Select Case page.Action
		Case UPDATE_RECORD
			str = str & preferencesLink & "Edit church profile"
		Case Else
			str = str & "Preferences"
	End Select

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	Dim href
	
	Dim saveButton
	href= pg.Url & pg.UrlParamsToString(True)
	saveButton = "<li id=""save-form-client""><a href=""" & href & """><img class=""icon"" src=""/_images/icons/disk.png"" /></a><a href=""" & href & """>Save</a></li>"
	
	Dim editProfileButton
	pg.Action = UPDATE_RECORD
	href = pg.Url & pg.UrlParamsToString(True)
	editProfileButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/world_edit.png"" /></a><a href=""" & href & """>Edit Profile</a></li>"
	
	Dim administratorsButton
	If page.member.IsAdmin Then
		pg.Action = ""
		href = "/client/administrators.asp" & pg.UrlParamsToString(True) 
		administratorsButton = "<li><a href=""" & href & """><img class=""icon"" src=""/_images/icons/medal_gold_2.png"" alt="""" /></a><a href=""" & href & """>Adminstrators</a></li>"
	End If

	Select Case page.Action
		Case UPDATE_RECORD
			str = str & saveButton
		Case Else
			str = str & administratorsButton & editProfileButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_DoInsertSampleProgram.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_build_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_admin_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/user_settings_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/state_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/format_date_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID

	' encrypted
	Public Action
	Public MemberId

	' objects
	Public Member
	Public Client
	
	' form data
	Public last_name_1
	Public first_name_1
	Public last_name_2
	Public first_name_2
	Public last_name_3
	Public first_name_3
	Public program_name	
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(MemberId) > 0 Then str = str & "mid=" & Encrypt(MemberId) & amp
		
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
		c.MemberId = MemberId
				
		Set c.Member = Member
		Set c.Client = Client
		
		Set Clone = c
	End Function
End Class
%>

