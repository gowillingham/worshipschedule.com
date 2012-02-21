<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "profile"
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
	Call CheckSession(sess, PERMIT_MEMBER)
	
	page.MessageID = Request.QueryString("msgid")
	page.Action = Decrypt(Request.QueryString("act"))
	
	Set page.Client = New cClient
	page.Client.ClientID = sess.ClientID
	Call page.Client.Load()
	Set page.Member = New cMember
	page.Member.MemberID = sess.MemberID
	Call page.Member.Load()
	
	page.Settings = GetApplicationSetting(page.Client.ClientID, "MemberRequiredField")
	
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
		<link type="text/css" href="/_incs/script/jquery/jquery-ui-1.6.custom/development-bundle/themes/start/ui.all.css" rel="stylesheet" />	
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=m_pageTitleText%></title>
		
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script language="javascript" type="text/javascript" src="/_incs/script/jquery/plugins/ui/jquery-ui-personalized-1.5.3.js"></script>
		<script language="javascript" type="text/javascript">
			$(document).ready(function(){
				// wire up save button
				$("#save-button a").click(function(){
					$("#form-member").submit();
				});
				
				// focus to first element in form
				$(".gets-focus").focus();
				
				// wire up datepicker ..
				$("#dob").datepicker({
					showOn: "button", 
					buttonImage: "/_images/icons/calendar_edit.png", 
					buttonImageOnly: true,
					buttonText: 'Choose a date ..',
					yearRange: '-105:+5' 				
				});
			});
		</script>
		<style type="text/css">
			.form, .message	{width:622px;}
		</style>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)

	' hack: ie getting leftmost margin incorrect unless I wrap the message in a div ..
	If Len(m_appMessageText) > 0 Then 
		str = str & "<div>" & m_appMessageText & "</div>"
	End If

	Select Case page.Action
		Case UPDATE_RECORD
			Call LoadMemberFromPost(page.Member)
			If ValidFormMember(page.Member, page.Settings) Then
				Call DoUpdateMember(page.Member, rv)
				Select Case rv
					Case 0
						page.MessageID = 1004
					Case -2
						' dupe member
						page.MessageID = 1001
					Case -3
						' dupe login
						page.MessageID = 1002
					Case Else
						' unknown error
						page.MessageID = 1003
				End Select
				page.Action = ""
				Response.Redirect(page.Url & page.UrlParamsToString(False))
				
			Else
				str = str & FormMemberToString(page, page.Member)
			End If
			
		Case Else
			str = str & FormMemberToString(page, page.Member)
			
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Sub DoUpdateMember(member, outError)
	member.IsProfileComplete = MEMBER_PROFILE_COMPLETE
	Call member.Update(outError)
End Sub

Sub LoadMemberFromPost(member)
	member.NameFirst = Trim(Request.Form("NameFirst")) 
	member.NameLast = Trim(Request.Form("NameLast")) 
	member.NameLogin = Trim(Request.Form("NameLogin")) 
	member.Email = Trim(Request.Form("Email"))
	member.EmailRetype = Trim(Request.Form("EmailRetype")) 
	member.PhoneHome = Trim(Request.Form("PhoneHome")) 
	member.PhoneMobile = Trim(Request.Form("PhoneMobile")) 
	member.PhoneAlternate = Trim(Request.Form("PhoneAlternate"))
	member.AddressLine1 = Trim(Request.Form("AddressLine1")) 
	member.AddressLine2 = Trim(Request.Form("AddressLine2")) 
	member.City = Trim(Request.Form("City")) 
	member.StateID = Trim(Request.Form("StateID")) 
	member.PostalCode = Trim(Request.Form("PostalCode"))
	member.Gender = Trim(Request.Form("Gender"))
	member.DOB = Trim(Request.Form("DOB"))
End Sub

Sub SetPageHeader(page)
	Dim str
	
	str = str & "<a href=""/member/overview.asp"">Member Home</a> / "
	str = str & "<a href=""/member/settings.asp"">Settings</a> / "
	str = str & "Profile"

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim href	
	
	Dim saveButton
	href = "#"
	saveButton = saveButton & "<li id=""save-button""><a href=""" & href & """><img class=""icon"" src=""/_images/icons/disk.png"" /></a><a href=""" & href & """>Save</a></li>"
	
	str = str & saveButton
	
	' use this if no buttons will be displayed ..
	'str = str & "<li>&nbsp;</li>"
	
	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_select_option.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_setting.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/state_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_FormMemberToString.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	
	' objects
	Public Member
	Public Client	
	
	' not persisted
	Public Settings
	
	Public Function Url()
		Url = Request.ServerVariables("URL")
	End Function
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
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
		c.Settings = Settings
		
		Set c.Member = Member
		Set c.Client = Client
		
		Set Clone = c
	End Function
End Class
%>

