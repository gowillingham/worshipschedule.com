<script runat="server" type="text/vbscript" language="vbscript">

Sub SetUnscheduledOptionLists(skillId, memberList, scheduled, outAvailable, outNotAvailable)
	Dim str, i, j
	
	' get an array of scheduled programMemberSkillIDs
	Dim scheduledList				: scheduledList = Split(scheduled, ",")

	Dim isThisSkill					: isThisSkill = True
	Dim isAvailable					: isAvailable = True
	Dim isSkillEnabled				: isSkillEnabled = True
	Dim isSkillGroupEnabled			: isSkillGroupEnabled = True
	Dim isMemberEnabled				: isMemberEnabled = True
	Dim isProgramMemberEnabled		: isProgramMemberEnabled = True
	Dim isScheduled					: isScheduled = False
	
	' clear these before generating
	outAvailable = ""
	outNotAvailable = ""
	
	' provide empty first option
	outAvailable = outAvailable & "<option value="""">&nbsp;</option>"
	outNotAvailable = outNotAvailable & "<option value="""">&nbsp;</option>"
	
	If IsArray(memberList) Then
		For i = 0 To UBound(memberList,2)
			isAvailable = True				: if memberList(9,i) = 0 Then isAvailable = False
			isSkillEnabled = True			: if memberList(7,i) = 0 Then isSkillEnabled = False
			isSkillGroupEnabled = True		: If memberList(8,i) = 0 Then isSkillGroupEnabled = False
			isMemberEnabled = True			: If memberList(3,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True	: If memberList(4,i) = 0 Then isProgramMemberEnabled = False
			
			isThisSkill = True				: If CLng(skillId) <> CLng(memberList(15,i)) Then isThisSkill = False
			
			' test if they are already in the scheduled select list ..
			isScheduled = False
			If IsArray(scheduledList) Then
				For j = 0 To UBound(scheduledList)
					If CStr(scheduledList(j) & "") = CStr(memberList(0,i) & "") Then
						isScheduled = True
						Exit For
					End If
				Next
			End If
			
			If (Not isScheduled) And isThisSkill And isSkillEnabled And isSkillGroupEnabled And isMemberEnabled And isProgramMemberEnabled Then
				If isAvailable Then
					outAvailable = outAvailable & "<option value=""" & memberList(0,i) & """>" & Server.HTMLEncode(memberList(1,i) & ", " & memberList(2,i)) & "</option>"
				Else
					outNotAvailable = outNotAvailable & "<option value=""" & memberList(0,i) & """>" & Server.HTMLEncode(memberList(1,i) & ", " & memberList(2,i)) & "</option>"
				End If
			End If
		Next
	End If
End Sub

</script>