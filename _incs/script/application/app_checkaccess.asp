<script runat="server" type="text/vbscript" language="vbscript">

Sub LoginMember(sess, member, outError)
	Dim rv
	
	' get member id
	Dim cmd			: Set cmd = Server.CreateObject("ADODB.Command")
	Dim rs			: Set rs = Server.CreateObject("ADODB.Recordset")
	Dim cnn			: Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.Open Application.Value("CNN_STR")

	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "up_memberGetLoginAccountCredentials"
	cmd.ActiveConnection = cnn
		
	cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
	cmd.Parameters.Append cmd.CreateParameter("@NameLogin", adVarChar, adParamInput, 25, CStr(member.nameLogin))
	cmd.Parameters.Append cmd.CreateParameter("@PWord", adVarChar, adParamInput, 25, CStr(member.PWord))

	Set rs = cmd.Execute
	If Not rs.EOF Then 
		sess.MemberID = rs("MemberID").Value
		
		member.MemberId = sess.MemberId
		Call member.Load()
	End If
	
	rs.Close()
	outError = cmd.Parameters("@RETURN_VALUE")
	
	' -1:unknown error 
	' -2:login not found 
	' -3:multiple rows returned
	
	If outError <> 0 Then Exit Sub
	
	' look for existing session to use ..
	
	sess.IsImpersonated = 0
	sess.SessionKey = "mykey"
	Call sess.Add(rv)
	Call sess.Load()
	
	Response.Cookies("sid") = sess.SessionID
	
	If Len(sess.MemberID) > 0 Then
		Call UpdateLastLogin(sess.MemberID)
	End If
End Sub

Sub CheckSession(sess, accessLevel)
	Dim rv
	Dim pg			: Set pg = New cPage
	
	Call sess.Refresh("", rv)
	
	' let everyone in ..
	If accessLevel = PERMIT_ALL Then 
		Exit Sub
	End If
	
	' if we're here, check logins
	If rv = -1 Then
		pg.Action = LOGOFF_SESSION_ABANDON
		Response.Redirect("/member/login.asp" & pg.UrlParamsToString(False))
	ElseIf rv = -2 Then
		pg.Action = LOGOFF_SESSION_TIMEOUT
		pg.MessageID = 1064
		Response.Redirect("/member/login.asp" & pg.UrlParamsToString(False))
	End If
	
	' disabled client
	If Not sess.IsClientEnabled Then
		Call LogoutMember()
		pg.MessageID = 2033
		Response.Redirect("/member/login.asp" & pg.UrlParamsToString(False))
	End If
	
	' expired trial 
	If sess.IsTrialAccount Then
		If (sess.TrialAccountLength - DateDiff("d", sess.DateClientCreated, Now())) < 0 Then
			If sess.IsAdmin Then
				If Request.ServerVariables("REFERER") <> "/client/account.asp" Then
					Response.Redirect("/client/account.asp" & pg.UrlParamsToString(False))
				End If
			Else
				Call LogoutMember()
				pg.MessageID = 2036
				Response.Redirect("/member/login.asp" & pg.UrlParamsToString(False))
			End If
		End If
	End If
	
	' expired account
	If DateAdd("m", sess.SubscriptionTermLength, sess.SubscriptionTermStart) < Now() Then
		If sess.IsAdmin Then
			If Request.ServerVariables("REFERER") <> "/client/account.asp" Then
				Response.Redirect("/client/account.asp" & pg.UrlParamsToString(False))
			End If
		Else
			Call LogoutMember()
			pg.MessageID = 2036
			Response.Redirect("/member/login.asp" & pg.UrlParamsToString(False))
		End If
	End If
	
	' let in all the members 
	If accessLevel = PERMIT_MEMBER Then 
		Exit Sub
	End If
	
	If accessLevel = PERMIT_LEADER Then
		If (Not sess.IsLeader) And (Not sess.IsAdmin) Then
			pg.MessageID = 1016
			Response.Redirect("/member/programs.asp" & pg.UrlParamsToString(False))
		End If
	End If
	
	If accessLevel = PERMIT_ADMIN Then
		If Not sess.IsAdmin Then
			pg.MessageID = 1015
			Response.Redirect("/member/programs.asp" & pg.UrlParamsToString(False))
		End If
	End If
End Sub

Sub LogoutMember()
	Dim rv
	Dim sess		: Set sess = New cSession
	
	sess.SessionID = Request.Cookies("sid")
	If Len(sess.SessionID) > 0 Then
		Call sess.Delete(rv)
	End If
	
	Response.Cookies("sid") = ""
	Response.Cookies("sid").Expires = Now() - 1
	Session.Abandon
	
	Set sess = Nothing
End Sub

Sub UpdateLastLogin(memberID)
	Dim cnn
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.Open Application.Value("CNN_STR")
	cnn.up_memberUpdateLastLogin CLng(memberID), Now()
	
	cnn.Close: Set cnn = Nothing
End Sub

Function GetURLEncryptKey()
	'http://4guysfromrolla.com/webtech/tips/t060500-1.shtml
	'generates a random six-char string
	
	Dim NUMLOWER:				NUMLOWER		= 48  ' 48 = 0
	Dim NUMUPPER:				NUMUPPER		= 57  ' 57 = 9
	Dim LOWERBOUND:			LOWERBOUND	= 65  ' 65 = A
	Dim UPPERBOUND:			UPPERBOUND  = 90  ' 90 = Z
	Dim LOWERBOUND1:			LOWERBOUND1 = 97  ' 97 = a
	Dim UPPERBOUND1:			UPPERBOUND1 = 122 ' 122 = z
	Dim PASSWORD_LENGTH:	PASSWORD_LENGTH = 6

	Dim pwd, count, sNewPassword
	
	' initialize the random number generator
	Randomize()
	sNewPassword = ""
	count = 0
	
	Do Until count = PASSWORD_LENGTH
		' generate a num between 2 and 10 ;
		' if num > 4 create an uppercase else create lowercase
		If Int( ( 10 - 2 + 1 ) * Rnd + 2 ) > 4 Then
			pwd = Int( ( UPPERBOUND - LOWERBOUND + 1 ) * Rnd + LOWERBOUND )
		Else
			pwd = Int( ( UPPERBOUND1 - LOWERBOUND1 + 1 ) * Rnd + LOWERBOUND1 )
		End If

		sNewPassword = sNewPassword + Chr( pwd )
		count = count + 1
	Loop

	GetURLEncryptKey = sNewPassword
End Function

Function GetServerIndicator(sContainer)
	'returns message identifying development server
	Dim str, server, sColorStyle, sBackgroundStyle
	
	If Not Application.Value("IS_DEVELOPMENT_SERVER") Then Exit Function
	
	server = Request.ServerVariables("SERVER_NAME")
	Select Case server
		Case "worshipschedule.local"
			sColorStyle = "red"
			sBackgroundStyle = "yellow"
		Case "worshipschedule.maint.local"
			sColorStyle = "yellow"
			sBackgroundStyle = "red"
		Case "beta.worshipschedule.com"
			sColorStyle = "yellow"
			sBackgroundStyle = "purple"
		Case Else
	End Select 
	
	Select Case sContainer
		Case "DIV"
			
			str = "<div style=""width:100%;color:" & sColorStyle & ";font-size:1.1em;padding:2px 0px;font-weight:bold;text-align:center;background-color:" & sBackgroundStyle & ";"">Development Server - " & UCase(server) & "</div>"
		Case "TABLE_ROW"
			str = "<tr><td colspan=""10""><div style=""width:100%;color:" & sColorStyle & ";font-weight:bold;text-align:center;background-color:" & sBackgroundStyle & ";"">Development Server - " & UCase(server) & "</div></td></tr>"
		Case Else
	End Select
	
	GetServerIndicator = str
End Function



</script>
