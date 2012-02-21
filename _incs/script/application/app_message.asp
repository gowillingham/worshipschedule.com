<script runat="server" language="vbscript">

Function CustomApplicationMessageToString(header, message, importance)
	Dim str
	Dim alertText			: alertText = ""
	Dim alertClass			: alertClass = ""
	Dim iconPath			: iconPath = ""
	
	Call SetApplicationMessageStyles(importance, alertText, alertClass, iconPath)
	str = str & "<div class=""message"">"
	str = str & "<h3 class=""" & alertClass & """>"
	str = str & "<img class=""icon"" src=""" & iconPath & """ alt="""" />"
	str = str & header & "</h3>"
	str = str & "<div class=""listing"">" & message & "</div>"
	str = str & "</div>"
	
	Call ReplaceTokens(str)
	
	CustomApplicationMessageToString = str
End Function

Function ApplicationMessageToString(id)
	Dim str
	Dim msg
	Dim alertText			: alertText = ""
	Dim alertClass			: alertClass = ""
	Dim iconPath			: iconPath = ""
	
	' missing or illegal id
	If Len(id) = 0 Then Exit Function
	If Not IsNumeric(id) Then Exit Function
	
	' 0-MessageID 1-Text 2-Importance 3-Path to graphic file
	msg = GetMessageByMessageID(id)
	If Not IsArray(msg) Then Exit Function
	
	Call SetApplicationMessageStyles(msg(2,0), alertText, alertClass, iconPath)
	str = str & "<div class=""message"">"
	str = str & "<h3 class=""" & alertClass & """>"
	str = str & "<img class=""icon"" src=""" & iconPath & """ alt="""" />"
	str = str & alertText & "</h3>"
	str = str & "<div class=""listing"">" & msg(1,0) & "</div>"
	str = str & "</div>"
	
	Call ReplaceTokens(str)
	
	ApplicationMessageToString = str
End Function

Sub SetApplicationMessageStyles(importance, alertText, alertClass, iconPath)
	' set default first line, styling, icon
	
	Select Case importance
		Case "Error"
			alertText = "Sorry, but there was a problem! "
			alertClass = "alert-error"
			iconPath = Application.Value("ICON_IMAGE_DIRECTORY") & "alert.png"
		Case "Critical Error"
			alertText = "Sorry, but there was a problem! "
			alertClass = "alert-error"
			iconPath = Application.Value("ICON_IMAGE_DIRECTORY") & "critical.png"
		Case "Critical Info"
			alertText = "Sorry, but there was a problem! "
			alertClass = "alert-message"
			iconPath = Application.Value("ICON_IMAGE_DIRECTORY") & "critical.png"
		Case "Confirm"
			alertText = "Thanks! "
			alertClass = "alert-message"
			iconPath = Application.Value("ICON_IMAGE_DIRECTORY") & "confirm.png"
		Case "Info"
			alertText = "Thanks! "
			alertClass = "alert-message"
			iconPath = Application.Value("ICON_IMAGE_DIRECTORY") & "confirm.png"
		Case Else
			alertText = ""
			alertClass = "alert-message"
			iconPath = Application.Value("ICON_IMAGE_DIRECTORY") & "alert.png"
	End Select

End Sub

Sub ReplaceTokens(ByRef message)
	' application name token
	message = Replace(message, "[[application name]]", Application.Value("APPLICATION_NAME"))
	' global support email address token
	message = Replace(message, "[[support email]]", Application.Value("SUPPORT_EMAIL_ADDRESS"))
	' global support email link token
	message = Replace(message, "[[support email link]]", "<a href=""mailto:" & Application.Value("SUPPORT_EMAIL_ADDRESS") & """ title=""Email Support"">support</a>")
	' global support page token
	message = Replace(message, "[[support page]]", "<a href=""/support.asp"">support</a>")
End Sub

Function GetMessageByMessageID(iMessageID)
	'return multi-dimensional, one item (row) array of message information
	Dim cnn, rs
	
	GetMessageByMessageID = ""
	If Len(iMessageID) = 0 Then Exit Function
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.Open Application.Value("CNN_STR")
	Set rs = Server.CreateObject("ADODB.Recordset")
	cnn.up_adminGetMessageByMessageID CInt(iMessageID), rs
	If Not rs.EOF Then GetMessageByMessageID = rs.GetRows
	
	rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
End Function

</script>