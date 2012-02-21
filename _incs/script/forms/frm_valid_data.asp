<script type="text/vbscript" runat="server" language="vbscript">
'page scope variable to hold user error message
Dim msError

Function ErrorToString()
	Dim str, i
	Dim errors			: errors = ""
	Dim list
	
	If Len(msError) = 0 Then Exit Function
	
	' HACK:
	' ----------------------
	' remove the first '<br />' and convert to array ..
	msError = Right(msError, Len(msError) - 6)
	list = Split(msError, "<br />")
	
	If IsArray(list) Then 
		errors = errors & "<ol>"
		For i = 0 To UBound(list)
			errors = errors & "<li>" & list(i) & "</li>"
		Next
		errors = errors & "</ol>"
	End If
	str = str & "<div class=""error"">"
	str = str & "<img class=""icon"" src=""/_images/icons/critical.png"" alt=""Error"" />"
	str = str & "<div class=""alert"">Sorry! Some information was missing or rejected. Please recheck your info. </div>"
	str = str & "<div class=""listing"">" & errors & "</div>"
	str = str & "</div>"
	
	ErrorToString = str
	msError = ""
End Function

Function FormErrorToString()
	Dim str
	
	If Len(msError) > 0 Then
		str = str & "<div class=""wsErrorBody"">" & vbCrLf
		str = str & "<div class=""wsMessageImage"">" & vbCrLf
		str = str & "<img src=""/_images/icons/alert.png"" alt=""Error""/>"	& vbCrLf
		str = str & "</div>"
		str = str & "<div style=""margin:0 0 0 25px;"">" & vbCrLf
		str = str & "<div class=""wsErrorFirstLine"">Sorry! Some information was missing or was rejected. Please recheck your info.</div>" & vbCrLf
		str = str & "<div class=""wsErrorText"">" & msError & "</div>" & vbCrLf
		str = str & "</div>" & vbCrLf
		str = str & "</div>" & vbCrLf
	End If
	
	FormErrorToString = str
	
	'clear the error
	msError = ""
End Function


Sub ShowFrmError()
	Dim str
	
	If Len(msError) > 0 Then
		str = str & "<div class=""wsErrorBody"">" & vbCrLf
		str = str & "<div class=""wsMessageImage"">" & vbCrLf
		str = str & "<img src=""/_images/icons/alert.png"" alt=""Error""/>"	& vbCrLf
		str = str & "</div>"
		str = str & "<div style=""margin:0 0 0 25px;"">" & vbCrLf
		str = str & "<div class=""wsErrorFirstLine"">Sorry! Some information was missing or was rejected. Please recheck your info.</div>" & vbCrLf
		str = str & "<div class=""wsErrorText"">" & msError & "</div>" & vbCrLf
		str = str & "</div>" & vbCrLf
		str = str & "</div>" & vbCrLf
	End If
	Response.Write str
	
	'clear the error
	msError = ""
End Sub

Sub AddCustomFrmError(sErrorText)
	'use to add a custom message to the error message string
	If Len(sErrorText) > 0 Then
		msError = msError & "<br />" & sErrorText
	End If
End Sub

Function ValidData(ByVal str, bRequired, iMinLength, iMaxLength, sAlias, sValidationType)
	'Check several validation conditions of a string. Generate a string error message
	'to display to the user if string does not validate
	ValidData = True
	str = Trim(str)
	'validate input parameters
	If Len(sAlias) = 0 Then
		Response.Write "Error|Function ValidData. Expected string value for parameter sAlias. Program terminated."
		Response.End
	End If
	If Not IsNumeric(iMinLength) Then
		Response.Write "Error|Function ValidData. Expected integer value for parameter iMinLength. Program terminated."
		Response.End
	End If
	If Not IsNumeric(iMaxLength) Then
		Response.Write "Error|Function ValidData. Expected integer value for parameter iMaxLength. Program terminated."
		Response.End
	End If
	
	'if required, determine if str has been supplied
	Select Case bRequired
		Case True
			If Len(str) = 0 Then 
				msError = msError & "<br />" & sAlias & " is required."
				ValidData = False
				Exit Function
			End If
		Case False
			'do nothing
		Case Else
			Response.Write "Error|Function ValidData. Expected boolean value for parameter bRequired. Program terminated."
			Response.End
	End Select
	
	'validate min length - iMinLength <= 0 will always return true, so this turns iMinLength off
	If Len(str) < iMinLength Then
		msError = msError & "<br />" & sAlias & " must contain at least " & iMinLength & " characters."
		ValidData = False
		Exit Function
	End If
	'validate max length - iMaxLength <= 0 turns this checking off
	If iMaxLength > 0 Then
		If Len(str) > iMaxLength Then
			msError = msError & "<br />" & sAlias & " cannot exceed " & iMaxLength & " characters in length."
			ValidData = False
			Exit Function
		End If
	End If
	
	'if len is 0 then exit - don't need to check against form data types
	If len(str & "") = 0 Then Exit Function
	
	Select Case sValidationType
		Case ""
			'do nothing
		Case "letters"
			If Not IsCharacter(str, "letters") Then
				msError = msError & "<br />" & sAlias & " may only contain letter characters."
				ValidData = False
			End If
		Case "letters_space"
			If Not IsCharacter(str, "letters_space") Then
				msError = msError & "<br />" & sAlias & " may only contain letter characters or spaces."
				ValidData = False
			End If
		Case "letters_space_apostrophe"
			If Not IsCharacter(str, "letters_space_apostrophe") Then
				msError = msError & "<br />" & sAlias & " may contain only letter, space, or apostrophe characters."
				ValidData = False
			End If
		Case "numbers"
			If Not IsCharacter(str, "numbers") Then
				msError = msError & "<br />" & sAlias & " may contain only number characters."
				ValidData = False
			End If
		Case "letters_numbers"
			If Not IsCharacter(str, "letters_numbers") Then
				msError = msError & "<br />" & sAlias & " may contain only number or letter characters."
				ValidData = False
			End If
		Case "email"
			If Not IsEmail(str) Then
				msError = msError & "<br />" & sAlias & " does not appear to be in the expected email address format."
				ValidData = False
			End If
		Case "date"
			If Not IsValidDate(str) Then
				msError = msError & "<br />" & sAlias & " is not in the expected date format of mm/dd/yyyy. Your date should not include any time information. "
				ValidData = False
			End If
		Case "phone"
			If Not IsPhone(str) Then
				msError = msError & "<br />" & sAlias & " is not in an expected phone number format like xxx-xxx-xxxx or (xxx)xxx-xxxx. Please include the area code."
				ValidData = False
			End If
		Case "zip"
			If Not IsZip(str) Then
				msError = msError & "<br />" & sAlias & " is not in the expected zip code format of ##### or #####-####."
				ValidData = False
			End If
		Case "time"
			If Not IsTime(str) Then
				msError = msError & "<br />" & sAlias & " is not in the expected time format of hh:mm am/pm or 24-hour time."
				ValidData = False
			End If
		Case "money"
			If Not IsMoney(str) Then
				msError = msError & "<br>" & sAlias & " is not in the expected format for currency of dddd." & Chr(162) & Chr(162) & " or $dddd." & Chr(162) & Chr(162) & "."
				ValidData = False
			End If
		Case "url"
			If Not IsURL(str) Then
				msError = msError & "<br>" & sAlias & " is not a valid URL address. Make sure the address begins with ""http://""."
				ValidData = False
			End If
		Case Else
			Response.Write "Error|Function ValidData. Unexpected type for parameter sType. Program terminated."
			Response.End
	End Select
End Function

'---------------------------------------------------------
'Helper Functions for Form Validation
'---------------------------------------------------------
Function IsPhone(ByVal str)
	'check that phone can be converted to proper format
	IsPhone = True
	str = Trim(str)
	
	'replace ( ) - . and spaces with zero-length string
	str = Replace(Replace(Replace(Replace(Replace(str & "", "(", ""), ")", ""), "-", ""), ".", ""), " ", "")

	'check for at least 10 digits
	If Len(str) <> 10 Then IsPhone = False
	
	'check that all digits are numeric
	If Not IsNumber(str) Then IsPhone = False
	
	'check for first digit non-zero
	If Left(str, 1) = "0" Then IsPhone = False
End Function

Function IsCharacter(ByVal sString, sTestType)
	Dim nChar
	Dim i
	IsCharacter = True
	Select Case sTestType
		'allow only letters
		Case "letters"
			For i = 1 To Len(sString)
				nChar = Asc(LCase(Mid(sString, i, 1)))
				If Not ((nChar > 96 And nChar < 123))Then
					IsCharacter = False
				End If
			Next
		'allow letters and space - ascii code 32 = space
		Case "letters_space"
			For i = 1 To Len(sString)
				nChar = Asc(LCase(Mid(sString, i, 1)))
				If Not ((nChar > 96 And nChar < 123) Or (nChar = 32))Then
					IsCharacter = False
				End If
			Next
		'allow letters and space and apostrophe - ascii code 39 = apostrophe
		Case "letters_space_apostrophe"
			For i = 1 To Len(sString)
				nChar = Asc(LCase(Mid(sString, i, 1)))
				If Not ((nChar > 96 And nChar < 123) Or (nChar = 32) Or (nChar = 39))Then
					IsCharacter = False
				End If
			Next
		Case "letters_numbers"
			For i = 1 To Len(sString)
				nChar = Asc(LCase(Mid(sString, i, 1)))
				If Not ((nChar > 47 And nChar < 58) or (nChar > 96 And nChar < 123)) Then
					IsCharacter = False
				End If
			Next
		Case "numbers"
			For i = 1 To Len(sString)
				nChar = Asc(LCase(Mid(sString, i, 1)))
				If Not (nChar > 47 And nChar < 58) Then 
					IsCharacter = False
				End If
			Next
		Case Else
			Response.Write "Error|Function IsCharacter. Unexpected type for parameter sTestType. Program terminated."
			Response.End
		End Select
End Function

Function IsNumber(sString)
	Dim nChar
	Dim i
	IsNumber = True
	For i = 1 To Len(sString)
		nChar = Asc(LCase(Mid(sString, i, 1)))
		If Not (nChar > 47 And nChar < 58) Then 
			IsNumber = False
		End If
	Next
End Function

Function IsValidDate(sDate)
	'checking for date of form mm/dd/yyyy or mm/dd/yy
	Dim arrDate
	IsValidDate = True
	
	'do this first and exit in case date looks like this 3//2001 or something
	If Not IsDate(sDate) Then
		IsValidDate = False
		Exit Function
	End If
	
	'further checking for proper format
	arrDate = Split(sDate, "/")
	If UBound(arrDate) <> 2 Then
		IsValidDate = False
	ElseIf ((arrDate(0) < 1) Or (arrDate(0) > 12)) Then
		IsValidDate = False
	ElseIf ((arrDate(1) < 1) Or (arrDate(1) > 31)) Then
		IsValidDate = False
	End If
	
	' test for extra information after like time ..
	Dim parts			: parts = Split(Trim(sDate), " ")
	If UBound(parts) > 0 Then
'	If IsArray(parts) Then
		IsValidDate = False
		Exit Function
	End If
	
End Function

Function IsEmail(sCheckEmail)
	'string meets certain email requirements. Does not guarantee valid email
    Dim sEmail, nAtLoc
    IsEmail = True
    sEmail = Trim(sCheckEmail)
    nAtLoc = InStr(1, sEmail, "@") 'Location of "@"
    
    'check for all the legal email characters - letters, numbers, 
    'underscore, dash, ampersand, and dot
        
    If Not (nAtLoc > 1 And (InStrRev(sEmail, ".") > nAtLoc + 1)) Then
        '"@" must exist, and last "." in string must follow the "@"
        IsEmail = False
    ElseIf InStr(nAtLoc + 1, sEmail, "@") > nAtLoc Then
        'String can't have more than one "@"
        IsEmail = False
    ElseIf Mid(sEmail, nAtLoc + 1, 1) = "." Then
        'String can't have "." immediately following "@"
        IsEmail = False
    ElseIf InStr(1, Right(sEmail, 2), ".") > 0 Then
        'String must have at least a two-character top-level domain.
        IsEmail = False
    End If
End Function

Function IsZip(ByVal sZip)
	'string is in expected zip code format of ##### or #####-####
	Dim arrZip
	IsZip = True

	'length must be five or ten
	If (Len(sZip) <> 5) And (len(sZip) <> 10) Then
		IsZip = False
	ElseIf Not IsNumber(Replace(sZip, "-", "")) then
		IsZip = False
	End If
	
	arrZip = Split(sZip, "-")
	If Not IsArray(arrZip) Then
		IsZip = False
		Exit Function
	End If
	
	'first term doesn't have five numbers
	If UBound(arrZip) > 1 Then
		IsZip = False
	ElseIf Len(arrZip(0)) <> 5 Then
		IsZip = False
	End If

	If UBound(arrZip) = 1 Then
		If Len(arrZip(1)) <> 4 Then
			IsZip = False
		End If
	End If
End Function

Function IsTime(ByVal sTime)
	'string is in expected time format of hh:mm am/pm or military time
	Dim arrTime, arrMinutes
	IsTime = True
	
	'must be a time
	If Not IsDate(sTime) Then
		IsTime = False
		Exit Function
	End If

	'only one colon
	arrTime = Split(sTime, ":")	
	If UBound(arrTime) <> 1 Then
		IsTime = False
		Exit Function
	End If
End Function

Function IsMoney(ByVal sVal)
	'string meets certain requirements for currency - $#####.##
	Dim arrTest
	IsMoney = True
	'trim leading $ if it is there
	If Left(sVal, 1) = "$" Then	sVal = Right(sVal, Len(sVal) - 1)
	'if 0 then exit right away
	If CStr(sVal) = "0" Then Exit Function
	'no decimal point, so exit
	If InStr(1, sVal, ".") = 0 Then 
		IsMoney = False
		Exit Function
	End If
	arrTest = Split(sVal, ".")
	If UBound(arrTest) > 1 Then 
		IsMoney = False
		Exit Function
	End If
'	'test digits right of decimal
	If UBound(arrTest) = 1 Then
		If Len(arrTest(1)) <> 2 Then IsMoney = False
		If Not IsNumber(arrTest(1)) Then IsMoney = False
	End If
	'if these are ok, then test to left
	If CStr(arrTest(0)) = "0" Then Exit Function
	If Len(arrTest(0)) = 0 Then IsMoney = False
	If Not IsNumber(arrTest(0)) Then IsMoney = False
End Function

Function IsURL(sURL)
	Dim objHTTP, iStatus
	IsURL = False
	If Len(trim(sURL)) = 0 Then Exit Function
	iStatus = 0
	'Set objHTTP = Server.CreateObject("Microsoft.XMLHTTP")
	'Set objHTTP = CreateObject("Msxml2.XMLHTTP.3.0")
	Set objHTTP = CreateObject("Msxml2.XMLHTTP.4.0")
	
	On Error Resume Next
	objHTTP.Open "GET", sURL, False
	objHTTP.Send
	if Err.Number = 0 then
		iStatus = objHTTP.Status
	End If
	On Error GoTo 0
	If iStatus = 200 Then IsURL = True
	Set objHTTP = Nothing
End Function

</script>