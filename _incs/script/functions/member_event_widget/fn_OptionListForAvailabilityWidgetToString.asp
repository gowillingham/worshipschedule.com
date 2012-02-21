<script runat="server" type="text/vbscript" language="vbscript">

Function OptionListForAvailabilityWidgetToString(members)
	Dim str, i
	
	Dim isAccountEnabled
	Dim isProgramEnabled
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-SkillListXmlFrag 4-IsLeader 5-IsAdmin 6-ProgramMemberIsEnabled 
	' 7-MemberActiveStatus 8-DateCreated 9-DateModified 10-ProgramMemberID 11-IsApproved 
	' 12-HasMissingAvailability 13-Email

	If IsArray(members) Then
		For i = 0 To UBound(members,2)
			isAccountEnabled = True				: If members(7,i) = 0 Then isAccountEnabled = False
			isProgramEnabled = True				: If members(6,i) = 0 Then isProgramEnabled = False
			
			If isAccountEnabled And isProgramEnabled Then
				str = str & "<option value=""" & members(0,i) & """>" & Server.HTMLEncode(members(1,i) & ", " & members(2,i)) & "</option>"
			End If	
		Next
	End If
	
	OptionListForAvailabilityWidgetToString = str
End Function

</script>