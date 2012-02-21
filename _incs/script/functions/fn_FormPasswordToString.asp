<%

Sub LoadMemberPasswordFromForm(member)
	member.PWord = Request.Form("PWord")
	member.PWordRetype = Request.Form("PWordRetype")
End Sub

Function ValidFormPassword(member)
	ValidFormPassword = True
	
	If Not ValidData(member.PWord, True, 0, 14, "New Password", "") Then ValidFormPassword = False
	
	'check that password fields match
	If member.PWord <> member.PWordRetype Then
		AddCustomFrmError("New Password and Retype Password must match exactly (case sensitive).")
		ValidFormPassword = False
	End If	
End Function

Function FormPasswordToString(page, member)
	Dim str
	Dim pg			: Set pg = page.Clone()
	
	pg.Action = UPDATE_PASSWORD
	str = str & "<div class=""form"">"
	str = str & ErrorToString()
	str = str & "<form action=""" & pg.Url & pg.UrlParamsToString(True) & """ method=""post"" id=""form-password"">"
	str = str & "<table>"
	str = str & "<tr><td class=""label"">" & RequiredElementToString(True, "New Password") & "</td>"
	str = str & "<td><input type=""password"" class=""gets-focus"" name=""PWord"" value=""" & page.Member.PWord & """ /></td></tr><tr>"

	str = str & "<td class=""label"">" & RequiredElementToString(True, "Retype Password") & "</td>"
	str = str & "<td><input type=""password"" name=""PWordRetype"" value=""" & page.Member.PWordRetype & """ />"
	str = str & "</td></tr>"
	str = str & "<tr><td>&nbsp</td><td class=""hint"">Your password is four to fourteen characters <br />and contains letters, numbers, or symbols. </td></tr>"
	
	str = str & "<tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr>"
	str = str & "<td>&nbsp;</td>"
	str = str & "<td><input type=""submit"" name=""Submit"" value=""Save"" />"
	pg.Action = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.Url & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormPasswordIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</td></tr></table></form>"
	str = str & "</div>"

	FormPasswordToString = str
End Function

%>
