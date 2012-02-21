<%

Sub DoDeleteMember(page, outError)
	' can't delete own account ..
	If CStr(page.Member.MemberID) = CStr(page.ThisMember.MemberID) Then
		outError = -4
		Exit Sub
	End If
	
	' can't delete member you don't own ..
	If Not OwnsMember(page.Member.MemberID, page.ThisMember.MemberID) Then
		outError = -3
		Exit Sub
	End If
	
	Call page.ThisMember.Delete(outError)
End Sub

Function FormConfirmDeleteMemberToString(page)
	Dim str
	Dim pg				: Set pg = page.Clone()
	
	str = str & "You will permanently remove your account member <strong>" & html(page.ThisMember.NameLast & ", " & page.ThisMember.NameFirst) & "</strong> from your " & Application.Value("APPLICATION_NAME") & " account. "
	str = str & "You will lose all calender and program information for this member. "
	str = str & "This action cannot be reversed. "
	str = CustomApplicationMessageToString("Please confirm remove member!", str, "Confirm")
	str = str & "<form action=""" & page.Url & page.UrlParamsToString(True) & """ method=""post"" id=""formConfirmDeleteMember"">"
	str = str & "<p><input type=""submit"" name=""Submit"" value=""Remove"" />"
	pg.Action = "": pg.MemberID = ""
	str = str & "&nbsp;&nbsp;<a href=""" & pg.UrlParamsToString(True) & """>Cancel</a>"
	str = str & "<input type=""hidden"" name=""FormConfirmDeleteMemberIsPostback"" value=""" & IS_POSTBACK & """ />"
	str = str & "</p></form>"
	
	FormConfirmDeleteMemberToString = str
End Function

%>