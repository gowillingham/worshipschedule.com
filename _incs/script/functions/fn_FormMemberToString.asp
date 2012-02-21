<%

Function FormMemberToString(page, member)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	str = str & "<div class=""tip-box""><h3>Tip</h3>"
	str = str & "<p>Required information is marked with a red asterisk! </p></div>"
	
	pg.Action = UPDATE_RECORD
	str = str & "<div class=""form""><form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-member"">"
	str = str & ErrorToString()
	str = str & "<table>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("NameFirst", "First Name", page.Settings) & "</td>"
	str = str & "<td><input class=""medium gets-focus"" type=""text"" name=""NameFirst"" value=""" & HTML(member.NameFirst) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("NameLast", "Last Name", page.Settings) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""NameLast"" value=""" & HTML(member.NameLast) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("Email", "Email", page.Settings) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""Email"" value=""" & HTML(member.Email) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("EmailRetype", "Email Retype", page.Settings) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""EmailRetype"" value=""" & HTML(member.EmailRetype) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("NameLogin", "Login Name", page.Settings) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""NameLogin"" value=""" & HTML(member.NameLogin) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("AddressLine1", "Adress Line 1", page.Settings) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""AddressLine1"" value=""" & HTML(member.AddressLine1) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td class=""label"">Address Line 2</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""AddressLine2"" value=""" & HTML(member.AddressLine2) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("City", "City", page.Settings) & "</td>"
	str = str & "<td><input class=""medium"" type=""text"" name=""City"" value=""" & HTML(member.City) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("State", "State", page.Settings) & "</td>"
	str = str & "<td>" & StateDropdownToString(UNITED_STATES_COUNTRY_CODE, member.StateID)
	str = str & "</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("PostalCode", "Postal Code", page.Settings) & "</td>"
	str = str & "<td><input class=""small"" type=""text"" name=""PostalCode"" value=""" & HTML(member.PostalCode) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("PhoneHome", "Home Phone", page.Settings) & "</td>"
	str = str & "<td><input type=""text"" name=""PhoneHome"" value=""" & HTML(member.PhoneHome) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("PhoneMobile", "Mobil Phone/Pager", page.Settings) & "</td>"
	str = str & "<td><input type=""text"" name=""PhoneMobile"" value=""" & HTML(member.PhoneMobile) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("PhoneAlternate", "Alternate Phone", page.Settings) & "</td>"
	str = str & "<td><input type=""text"" name=""PhoneAlternate"" value=""" & HTML(member.PhoneAlternate) & """ />"
	str = str & "</td></tr>"
	str = str & PhoneHintToString(page.Settings)
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("Gender", "Gender", page.Settings) & "</td>"
	str = str & "<td>" & GenderDropdownToString(member.Gender)
	str = str & "</td></tr>"
	str = str & "<tr><td class=""label"">" & FormatRequiredElement("DateOfBirth", "Birth Date", page.Settings) & "</td>"
	str = str & "<td><input class=""small"" type=""text"" name=""DOB"" id=""dob"" value=""" & HTML(member.DOB) & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>"
	str = str & "<tr><td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormMemberIsPostback"" value=""" & IS_POSTBACK & """ /></td></tr>"
	str = str & "</table></form>"
	str = str & "</div>"
	
	FormMemberToString = str
End Function

Function ValidFormMember(m, settings)
	ValidFormMember = True
	
	If Application.Value("FORM_VALIDATION_OFF") Then Exit Function	

	If Not ValidData(m.NameFirst, IsRequiredElement("NameFirst", settings), 0, 50, "First Name", "") Then ValidFormMember = False
	If Not ValidData(m.NameLast, IsRequiredElement("NameLast", settings), 0, 50, "Last Name", "") Then ValidFormMember = False
	If Not ValidData(m.Email, IsRequiredElement("Email", settings), 0, 100, "Email Address", "email") Then ValidFormMember = False

	'check that email fields match
	If UCase(m.Email) <> UCase(m.EmailRetype) Then
		AddCustomFrmError("Email and Retype Email must match exactly.")
		ValidFormMember = False
	End If	

	If Not ValidData(m.NameLogin, IsRequiredElement("NameLogin", settings), 0, 25, "Login Name", "") Then ValidFormMember = False
	
	If Not ValidData(m.AddressLine1, IsRequiredElement("AddressLine1", settings), 0, 100, "Address Line 1", "") Then ValidFormMember = False
	If Not ValidData(m.AddressLine2, IsRequiredElement("AddressLine2", settings), 0, 100, "Address Line 2", "") Then ValidFormMember = False
	If Not ValidData(m.City, IsRequiredElement("City", settings), 0, 100, "City", "") Then ValidFormMember = False
	If Not ValidData(m.StateID, IsRequiredElement("StateID", settings), 0, 0, "State", "") Then ValidFormMember = False
	If Not ValidData(m.PostalCode, IsRequiredElement("PostalCode", settings), 0, 10, "Postal Code", "zip") Then ValidFormMember = False
	If Not ValidData(m.PhoneHome, IsRequiredElement("PhoneHome", settings), 0, 14, "Home Phone", "phone") Then ValidFormMember = False
	If Not ValidData(m.PhoneMobile, IsRequiredElement("PhoneMobile", settings), 0, 14, "Mobile Phone/Pager", "phone") Then ValidFormMember = False
	If Not ValidData(m.PhoneAlternate, IsRequiredElement("PhoneAlternate", settings), 0, 14, "Alternate Phone", "phone") Then ValidFormMember = False

	' check for multiple phone numbers if required
	Dim phoneCount
	phoneCount = 0
	If IsRequiredElement("PhoneMultiple", settings) Then
	
		' get a count of phone numbers
		If Len(m.PhoneHome) > 0 Then phoneCount = phoneCount + 1
		If Len(m.PhoneMobile) > 0 Then phoneCount = phoneCount + 1
		If Len(m.PhoneAlternate) > 0 Then phoneCount = phoneCount + 1
		
		If phoneCount < 2 Then
			ValidFormMember = False
			AddCustomFrmError("More than one phone number is required.")
		End If
	End If

	If Not ValidData(m.DOB, IsRequiredElement("DateOfBirth", settings), 0, 10, "Date of Birth", "date") Then ValidFormMember = False
	'if provided, only allow DOB in past
	If IsDate(m.DOB) Then
		If DateDiff("s", Now(), m.DOB) > 0 Then 
			If DatePart("y", Now()) <> DatePart("y", m.DOB) Then
				AddCustomFrmError("Date of Birth cannot occur in the future.")
				ValidFormMember = False
			End If
		End If
	End If
	
	If Not ValidData(m.Gender, IsRequiredElement("Gender", settings), 0, 14, "Gender", "") Then ValidFormMember = False
End Function

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

Function PhoneHintToString(settings)
	' check if phone or multiple phone numbers are required
	Dim val, str, i
	
	If Not IsArray(settings) Then Exit Function
	
	' get the current setting value for 'PhoneMultiple'
	For i = 0 To UBound(settings,2)
		If settings(0,i) = "PhoneMultiple" Then
			val = settings(1,i)
			Exit For
		End If
	Next
	
	If val = 1 Then
		str = str & "<tr><td>&nbsp;</td><td><div class=""hint"">" 
		str = str & "<img class=""icon"" src=""/_images/icons/lightbulb.png"" alt=""Hint"" />"
		str = str & "<strong>Note! </strong>More than one phone number is required."
		str = str & "</div></td></tr>"
	End If
	
	PhoneHintToString = str 
End Function

Function StateDropdownToString(countryID, stateID)
	Dim str, i
	Dim state		: Set state = New cState
	state.countryID = countryID
	Dim arr			: arr = state.List()
	
	If Not IsArray(arr) Then Exit Function
	
	Dim list()
	Redim list(1,UBound(arr,2))
	For i = 0 To UBound(arr,2)
		list(0,i) = arr(0,i)
		list(1,i) = Html(arr(1,i) & " - " & arr(2,i))
	Next
	
	str = str & "<select name=""StateID"">"
	str = str & "<option value="""">&nbsp;</option>"
	str = str & SelectOption(list, stateID)	
	str = str & "</select>"
	
	StateDropdownToString = str
	Set state = Nothing
End Function

Function GenderDropdownToString(gender)
	Dim str, i
	Dim list		: list = GetGender(2)
	
	str = str & "<select name=""Gender"">"
	str = str & "<option value="""">&nbsp;</option>"
	str = str & SelectOption(list, gender)	
	str = str & "</select>"
	
	GenderDropdownToString = str
End Function

Function GetGender(iStyle)
	'iStyle: 0=M/F, 1=long names, 2=both 
	Dim arr
	ReDim arr(1,1)
	If iStyle = "1" Then
		arr(0,0) = "F"
		arr(0,1) = "M"
		arr(1,0) = "Female"
		arr(1,1) = "Male"
	ElseIf iStyle = "2" Then
		arr(0,0) = "F"
		arr(0,1) = "M"
		arr(1,0) = "F - Female"
		arr(1,1) = "M - Male"
	Else
		arr(0,0) = "F"
		arr(0,1) = "M"
		arr(1,0) = "F"
		arr(1,1) = "M"
	End If
	GetGender = arr
End Function

%>
