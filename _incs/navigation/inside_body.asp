	<body id="page-<%=m_pageTabLocation%>">
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_topbar.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><%=m_pageHeaderText%></h1>
			</div>
		</div>
		<%=m_acctExpiresText %>
		<div id="container">
			<%=m_impersonateText %>
			<div class="contentwrap">
				<div class="tabstrip">
					<ul class="tablist">
						<%=m_tabStripText %>
					</ul>
					<ul class="linkbar">
						<%=m_tabLinkBarText %>
					</ul>
				</div>
				<div class="content">
					<%=m_bodyText %>
					<!-- following keeps content inside .content div -->
					<div style="clear:both;height:1px;margin:0;padding:0;">&nbsp;</div>
				</div>
			</div>
		</div>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/inside_footer.asp"-->
	</body>
	
<script type="text/vbscript" runat="server" language="vbscript">

Sub SetAccountNotifier(sess)
	Dim str
	
	' only show to admin or leader
	If Not (sess.IsAdmin Or sess.IsLeader) Then Exit Sub
	
	Dim expired				: expired = False
	If sess.IsTrialAccount Then
		If sess.TrialExpiresDate < Now() Then expired = True
	Else
		If sess.AccountExpiresDate < Now() Then expired = True
	End If
	
	Dim daysRemaining		
	Dim banner
	
	Dim extendText
	extendText = "Extend your account <a href=""/client/account.asp"">here!</a>"
	
	If sess.IsTrialAccount Then
		daysRemaining = DateDiff("d", Now(), sess.TrialExpiresDate)
		If daysRemaining < 0 Then daysRemaining = 0
		
		If expired Then
			str = str & "<strong class=""banner expired"">Your free trial account has expired! </strong>"
			str = str & extendText
		Else
			str = str & "<strong class=""banner"">Free " & Application.Value("TRIAL_ACCOUNT_LENGTH") & " day trial account! </strong>"
			str = str & daysRemaining & " day" 
			If daysRemaining <> 1 Then str = str & "s"
			str = str & " remaining. "
			str = str & extendText
		End If
	Else
		daysRemaining = DateDiff("d", Now(), sess.AccountExpiresDate)
		If daysRemaining < 0 Then daysRemaining = 0
		
		If daysRemaining = 0 Then
			str = str & "<strong class=""banner expired"">Your " & Application.Value("APPLICATION_NAME") & " subscription has expired! </strong>"
			str = str & extendText
		ElseIf daysRemaining < 31 Then
			str = str & "<strong class=""banner"">Your " & Application.Value("APPLICATION_NAME") & " subscription expires in "
			str = str & daysRemaining & " day"
			If daysRemaining <> 1 Then str = str & "s"
			str = str & ". </strong>"
			str = str & extendText
		Else
			' don't notify if more than 30 days remain ..
			Exit Sub
		End If
	End If
	If Len(str) > 0 Then str = "<div id=""account-notifier"">" & str & "</div>"
	
	m_acctExpiresText = str
End Sub

Sub SetImpersonateText(sess)
	Dim str
	
	If sess.IsImpersonated <> 1 Then Exit Sub
	
	Dim member			: Set member = New cMember
	member.MemberID = sess.MemberID
	member.Load()
	
	str = str & "<div id=""impersonate-box"">Logged In as Member '" & Html(member.NameLast & ", " & member.NameFirst) & "'</div>"
	
	m_impersonateText = str
	Set member = Nothing
End Sub

Sub SetPageTitle(page)
	Dim str
	
	str = Server.HTMLEncode(Application.Value("APPLICATION_NAME") & " Web Scheduling for " & page.Client.NameClient)
	
	m_pageTitleText = str
End Sub

Sub SetTopBar(page)
	Dim str
	Dim pg		: Set pg = page.Clone()
	
	str = str & "&nbsp;<span id=""personname"">" & html(page.Member.NameLogin) & "</span>"
	str = str & "&nbsp;|&nbsp;<span id=""overview""><a href=""/member/overview.asp"">Member Home</a></span>"
	If (page.Member.IsAdmin) = 1 Or (page.Member.IsLeader = 1) Then
		str = str & "&nbsp;|&nbsp;<span id=""admin""><a href=""/admin/overview.asp"">Administration</a></span>"
	End If
	str = str & "&nbsp;|&nbsp;<span id=""settings""><a href=""/member/settings.asp"">Settings</a></span>"
	str = str & "&nbsp;|&nbsp;<span id=""help""><a href=""/help/help.asp"" target=""_blank"">Help</a></span>"
	pg.Action = LOGOFF_USER
	str = str & "&nbsp;|&nbsp;<span id=""logoff""><a href=""/member/login.asp" & pg.UrlParamsToString(True) & """>Logoff</a></span>"

	m_topBarText = str
End Sub

Sub SetTabList(tabLocation, page)
	Dim str
	
	Dim adminTabList
	adminTabList = adminTabList & "<li class=""admin-overview"" ><a href=""/admin/overview.asp"">Overview</a></li>"
	adminTabList = adminTabList & "<li class=""admin-programs"" ><a href=""/admin/programs.asp"">Programs</a></li>"
	adminTabList = adminTabList & "<li class=""admin-members"" ><a href=""/admin/members.asp"">Members</a></li>"
	adminTabList = adminTabList & "<li class=""admin-schedules"" ><a href=""/schedule/schedules.asp"">Schedules</a></li>"
	adminTabList = adminTabList & "<li class=""admin-email"" ><a href=""/email/email.asp"">Email</a></li>"
	adminTabList = adminTabList & "<li class=""admin-files"" ><a href=""/admin/files.asp"">Files</a></li>"
	adminTabList = adminTabList & "<li class=""admin-reports"" ><a href=""/reports/default.asp"">Reports</a></li>"

	' don't show these tabs to leader ..
	If page.Member.IsAdmin Then
		adminTabList = adminTabList & "<li class=""admin-settings"" ><a href=""/client/preferences.asp"">Admin Settings</a></li>"
	End If
	adminTabList = adminTabList & "<li class=""admin-account"" ><a href=""/client/account.asp"">Account</a></li>"
	
	Dim memberSettingTabList
	memberSettingTabList = memberSettingTabList & "<li class=""member-settings"" ><a href=""/member/settings.asp"">Settings</a></li>"
	memberSettingTabList = memberSettingTabList & "<li class=""profile"" ><a href=""/member/profile.asp"">Profile</a></li>"
	memberSettingTabList = memberSettingTabList & "<li class=""password"" ><a href=""/member/password.asp"">Password</a></li>"
	memberSettingTabList = memberSettingTabList & "<li class=""contacts"" ><a href=""/member/contacts.asp"">Account Info</a></li>"
	
	Dim memberOverviewTabList
	memberOverviewTabList = memberOverviewTabList & "<li class=""overview"" ><a href=""/member/overview.asp"">Overview</a></li>"
	memberOverviewTabList = memberOverviewTabList & "<li class=""calendar""><a href=""/member/schedules.asp"">Calendar</a></li>"
	memberOverviewTabList = memberOverviewTabList & "<li class=""availability""><a href=""/member/events.asp"">Availability</a></li>"
	memberOverviewTabList = memberOverviewTabList & "<li class=""programs""><a href=""/member/programs.asp"">Programs</a></li>"
	memberOverviewTabList = memberOverviewTabList & "<li class=""files""><a href=""/member/files.asp"">Files</a></li>"

	Select Case tabLocation
		Case "overview"
			str = str & memberOverviewTabList
		Case "calendar"
			str = str & memberOverviewTabList
		Case "availability"
			str = str & memberOverviewTabList
		Case "programs"
			str = str & memberOverviewTabList
		Case "files"
			str = str & memberOverviewTabList
		Case "admin"
			str = str & memberOverviewTabList
			
		Case "profile"
			str = str & memberSettingTabList
		Case "password"
			str = str & memberSettingTabList
		Case "member-settings"
			str = str & memberSettingTabList
		Case "contacts"
			str = str & memberSettingTabList
			
		Case "admin-overview"
			str = str & adminTabList
		Case "admin-programs"
			str = str & adminTabList
		Case "admin-members"
			str = str & adminTabList
		Case "admin-schedules"
			str = str & adminTabList
		Case "admin-files"
			str = str & adminTabList
		Case "admin-reports"
			str = str & adminTabList
		Case "admin-email"
			str = str & adminTabList
		Case "admin-account"
			str = str & adminTabList
		Case "admin-settings"
			str = str & adminTabList

	End Select
	
	m_tabStripText = str
End Sub

</script>
