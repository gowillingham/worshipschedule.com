<script language="vbscript" runat="server" type="text/vbscript">

Function SendCredentials(memberID, fromAddress)
	Dim subject, body
	
	Dim member				: Set member = New cMember
	Dim email				: Set email = New cEmailSender
	
	member.MemberID = memberID
	member.Load()
	
	subject = "** [" & Application.Value("APPLICATION_NAME") & "] " & member.ClientName & " Login Information for " & member.NameFirst & " " & member.NameLast & " **"

	body = body & "Hello " & member.NameFirst & " " & member.NameLast & ":" & vbCrLf & vbCrLf
	body = body & "Below find the login information for your " & Application.Value("APPLICATION_NAME") & " account with " & member.ClientName & ". "
	body = body & "To login, go to http://" & Request.ServerVariables("SERVER_NAME") & "/member/login.asp. "
	body = body & "Remember that your password is case sensitive. " & vbCrLf & vbCrLf
	body = body & String(60, "-") & vbCrLf
	body = body & "Login Name: " & member.NameLogin & vbCrLf
	body = body & "Password: " & member.PWord & vbCrLf
	body = body & String(60, "-")
	body = body & EmailDisclaimerToString(member.ClientName)
	
	SendCredentials = email.SendMessage(member.Email, fromAddress, subject, body)

	Set email = Nothing
	Set member = Nothing
End Function

Function EmailDisclaimerToString(clientName)
	Dim str 
	
	If Len(clientName) = "" Then clientName = Application.Value("APPLICATION_NAME")
	
	str = str & vbCrLf & vbCrLf & "--" & vbCrLf & vbCrLf
	str = str & "This message was sent by " & Application.Value("APPLICATION_NAME") & " on behalf of " & html(clientName) & ". "
	str = str & "If you believe that you have received this message in error, please contact "
	str = str & Application.Value("APPLICATION_NAME") & " support at mailto:" & Application.Value("SUPPORT_EMAIL_ADDRESS") & ". " & vbCrLf & vbCrLf
	str = str & "All Rights Reserved - Copyright " & Application.Value("APPLICATION_NAME") & " " & Year(Now()) & vbCrLf
	str = str & "(Timestamp: " & Now() & ")"
	
	EmailDisclaimerToString = str
End Function

</script>
