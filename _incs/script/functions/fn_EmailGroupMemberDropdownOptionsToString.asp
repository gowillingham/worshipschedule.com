<script runat="server" type="text/vbscript" language="vbscript">

Function EmailGroupMemberDropdownOptionsToString(memberId)
	Dim str, i
	
	Dim emailGroup			: Set emailGroup = New cEmailGroup
	emailGroup.MemberId = memberId
	
	Dim list				: list = emailGroup.List()
	Dim actionPlusId
	
	If IsArray(list) Then
		str = str & "<option value=""default"">Members ..</option>"
		str = str & "<option value="""">--</option>"
	
		str = str & "<optgroup id=""add-to-optgroup"" label=""Add to .."">"
		For i = 0 To UBound(list,2)
			actionPlusId = "emgid-" & list(0,i) & "-" & INSERT_EMAIL_GROUP_MEMBERS
			str = str & "<option value=""" & actionPlusId & """>" & Server.HTMLEncode(list(1,i)) & "</option>"
		Next
		str = str & "</optgroup>"
		
		str = str & "<optgroup id=""remove-from-optgroup"" label=""Remove from .."">"
		For i = 0 To UBound(list,2)
			actionPlusId = "emgid-" & list(0,i) & "-" & DELETE_EMAIL_GROUP_MEMBERS
			str = str & "<option value=""" & actionPlusId & """>" & Server.HTMLEncode(list(1,i)) & "</option>"
		Next
		str = str & "</optgroup>"
	Else
		str = str & "<option value="""">Members ..</option>"
	End If

	EmailGroupMemberDropdownOptionsToString = str
End Function

</script>