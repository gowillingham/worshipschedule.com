<script runat="server" type="text/vbscript" language="vbscript">

Function MemberNotesToString(skillId, memberList)
	Dim str, i
	Dim count				: count = 0
	
	Dim isMemberEnabled			: isMemberEnabled = True
	Dim isProgramMemberEnabled	: isProgramMemberEnabled = True
	Dim isThisSkill				: isThisSkill = True
	
	' 1-NameLast 2-NameFirst 3-MemberEnabled 4-ProgramMemberEnabled 10-AvailabilityNote 
	' 12-DateAvailabilityModified 15-SkillID
	
	If Not IsArray(memberList) Then Exit Function
	For i = 0 To UBound(memberList,2)
		isMemberEnabled = True			: If memberList(3,i) = 0 Then isMemberEnabled = False
		isProgramMemberEnabled = True	: If memberList(4,i) = 0 Then isProgramMemberEnabled = False
		isThisSkill = False				: If CStr(skillId & "") = CStr(memberList(15,i) & "") Then isThisSkill = True
		If Len(memberList(10,i)) > 0 Then
		
			If isMemberEnabled And isProgramMemberEnabled And isThisSkill Then
				count = count + 1
				
				str = str & "<li>"
				str = str & "<strong>" & Server.HTMLEncode(memberList(2,i) & " " & memberList(1,i)) & "</strong> on " & memberList(12,i) & ": "
				str = str & Server.HTMLEncode(memberList(10,i))
				str = str & "</li>"
			End If
		End If
	Next
		
	If count > 0 Then
		str = "<ul class=""notes"">" & str & "</ul>"
		str = "<h6>Notes from members</h6>" & str
	End If

	MemberNotesToString = str
End Function

</script>