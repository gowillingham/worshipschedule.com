<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
' global view tokens
Dim m_pageTabLocation	: m_pageTabLocation = "member-settings"
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
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.2.6/jquery.min.js"></script>
		<script type="text/javascript" language="javascript">
			$(document).ready(function(){
				// wire up enable account dropdown
				$("#form-enable-account").change(function(){
					this.submit();
				});
				
				// wire up home page dropdown 
				$("#form-home-page").change(function(){
					this.submit();
				});
			});
		</script>
		
		
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=m_pageTitleText%></title>
		<style type="text/css">
			.summary	{width:650px;}
			.message	{width:650px;}
		</style>
	</head>
	<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_body.asp"-->
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)

	Select Case page.Action
		Case UPDATE_MEMBER_HOME_PAGE
			page.Member.HomePageId = Request.Form("home_page_id")
			Call page.Member.Update(rv)
			page.Action = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))

		Case UPDATE_RECORD
			page.Member.ActiveStatus = Request.Form("is_enabled")
			Call page.Member.Update(rv)
			page.Action = ""
			Response.Redirect(page.Url & page.UrlParamsToString(False))
			
		Case Else
			str = str & AccountSettingsSummaryToString(page)

	End Select
		
	m_bodyText = str
	Set page = Nothing
End Sub

Function AccountSettingsSummaryToString(page)
	Dim str
	
	str = str & "<div class=""tip-box""><h3>Tip!</h3>"
	str = str & "<p>The settings section of your account is where you can change your account preferences and your personal information. "
	str = str & "</p></div>"
	
	str = str & "<div class=""summary"">"
	str = str & "<h3 class=""first"">" & Application.Value("APPLICATION_NAME") & " Settings for " & html(page.Client.NameClient) & "</h3>"
	str = str & "<h4>Home page</h4>"
	str = str & FormHomePage(page)
	str = str & "<h4>Enable/Disable Account</h4>"
	str = str & FormEnableAccount(page)
	str = str & "</div>"
	
	AccountSettingsSummaryToString = str
End Function

Function FormHomePage(page)
	Dim str, i
	Dim pg					: Set pg = page.Clone()
	
	Dim homePage			: Set homePage = New cHomePage
	Dim list				: list = homePage.List()
	
	' 0-HomePageId 1-Name 2-Url 3-IsEnabled 4-IsAdmin
	
	Dim isAdminPage
	Dim selected			: selected = ""
	
	str = str & "<p>You may select your start page for " & Application.Value("APPLICATION_NAME") & ". "
	str = str & "You'll go directly to this page whenever you login to your account. </p>"
	pg.Action = UPDATE_MEMBER_HOME_PAGE
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-home-page"">"
	str = str & "<input type=""hidden"" name=""form_home_page_is_postback"" value= """ & IS_POSTBACK & """ />"

	str = str & "<p><select name=""home_page_id"" id=""home-page"">"
	If IsArray(list) Then
		For i = 0 To UBound(list,2)
			isAdminPage = False		: If list(4,i) = 1 Then isAdminPage = True
			
			If (Not isAdminPage) Or (page.Member.IsAdmin = 1) Then
				selected = ""			: If CStr(page.Member.HomePageID & "") = CStr(list(0,i) & "") Then selected = " selected=""selected"""
				str = str & "<option value=""" & list(0,i) & """" & selected & ">" & list(1,i)  & "</option>"
			End If
		Next
	End If
	str = str & "</select></p></form>"
	
	FormHomePage = str
End Function

Function FormEnableAccount(page)
	Dim str
	Dim pg					: Set pg = page.Clone()
	
	Dim selected			: selected = ""
	
	str = str & "<p>You may set your " & Application.Value("APPLICATION_NAME") & " member account to disabled. "
	str = str & "This will remove your account from all " & html(page.Member.ClientName) & " programs, schedules, members lists, and email lists. "
	str = str & "Use this to disable your account if you wish to discontinue using " & Application.Value("APPLICATION_NAME") & " "
	str = str & "for a period of time without losing all of your profile settings or programs. </p>"
	
	If page.Member.ActiveStatus = 0 Then
		str = str & "<ul class=""other-stuff"">"
		str = str & "<li class=""error"">" ' <strong class=""warning"">Account disabled. </strong>"
		str = str & "This account has been set to disabled. "
		str = str & "</li></ul>"
	End If
	
	pg.Action = UPDATE_RECORD
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-enable-account"">"
	str = str & "<input type=""hidden"" name=""form_enable_account_is_postback"" value=""" & IS_POSTBACK & """ />"
	str = str & "<p><select name=""is_enabled"" id=""is-enabled"">"
	
	selected = "": If CStr(page.Member.ActiveStatus) = "1" Then selected = " selected=""selected"""
	str = str & "<option value=""1""" & selected & ">Account enabled</option>"
	
	selected = "": If CStr(page.Member.ActiveStatus) = "0" Then selected = " selected=""selected"""
	str = str & "<option value=""0""" & selected & ">Account disabled</option>"
	str = str & "</select></p></form>"
	
	FormEnableAccount = str
End Function

Sub SetPageHeader(page)
	Dim str
	
	str = str & "<a href=""/member/overview.asp"">Member Home</a> / "
	str = str & "Settings"

	m_pageHeaderText = str
End Sub

Sub SetTabLinkBar(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href

	Dim editProfileButton
	pg.Action = ""
	href = "/member/profile.asp" & pg.UrlParamsToString(True)
	editProfileButton = editProfileButton & "<li><a href=""" & href & """><img src=""/_images/icons/user_red_edit.png"" class=""icon"" alt="""" /></a>"
	editProfileButton = editProfileButton & "<a href=""" & href & """>Edit Profile</a></li>"
	
	Dim editPasswordButton	
	pg.Action = ""
	href = "/member/password.asp" & pg.UrlParamsToString(True)
	editPasswordButton = editPasswordButton & "<li><a href=""" & href & """><img src=""/_images/icons/key.png"" class=""icon"" alt="""" /></a>"
	editPasswordButton = editPasswordButton & "<a href=""" & href & """>Password</a></li>"
	
	Select Case page.Action
		Case Else
			str = str & editProfileButton & editPasswordButton
	End Select

	m_tabLinkBarText = str
End Sub
%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/home_page_cls.asp"-->
<%
Class cPage
	' unencrypted
	Public MessageID
	
	' encrypted
	Public Action
	
	' objects
	Public Member
	Public Client	
	
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
		
		Set c.Member = Member
		Set c.Client = Client
		
		Set Clone = c
	End Function
End Class
%>

