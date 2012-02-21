<script runat="server" type="text/vbscript" language="vbscript">

Function CustomEmailGroupItemsToString(memberId, emailGroupId)
	Dim str, i
	
	Dim emailGroup		: Set emailGroup = New cEmailGroup
	emailGroup.MemberId = memberId
	
	Dim emailGroups		: emailGroups = emailGroup.List()
	Dim classes
	
	If Not IsArray(emailGroups) Then Exit Function
	For i = 0 To UBound(emailGroups,2)
	
		classes = "email-group-node"
		If CStr(emailGroupId & "") = CStr(emailGroups(0,i)) Then
			classes = classes & " highlight"
		End If
		
		str = str & "<li class=""email-group-node"" title=""" & Server.HtmlEncode(emailGroups(1,i)) & """>" 
		str = str & "<span><a href=""#"" class=""" & classes & """ id=""emgid-" & emailGroups(0,i) & """>" & Server.HTMLEncode(emailGroups(1,i)) & "</a></span></li>"
	Next
	
	CustomEmailGroupItemsToString = str
End Function

</script>