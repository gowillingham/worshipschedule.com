<script runat="server" type="text/vbscript" language="vbscript">

Sub SetScheduleOptionList(skillId, buildList, outScheduled, outOptions)
	Dim str, i
	
	Dim isThisSkill					: isThisSkill = True
	Dim isMemberEnabled				: isMemberEnabled = True
	Dim isProgramMemberEnabled		: isProgramMemberEnabled = True
	Dim isSkillEnabled				: isSkillEnabled = True
	Dim isSkillGroupEnabled			: isSkillGroupEnabled = True
	Dim availableClass				: availableClass = ""
	Dim publishStatus				: publishStatus = 0
	
	' clear these before generating ..
	outScheduled = ""
	outOptions = ""

	' provide empty first option
	outOptions = outOptions & "<option value="""">&nbsp;</option>"
	
	If IsArray(buildList) Then
		For i = 0 To UBound(buildList,2)
			isMemberEnabled = True			: If buildList(9,i) = 0 Then isMemberEnabled = False
			isProgramMemberEnabled = True	: If buildList(14,i) = 0 Then isProgramMemberEnabled = False
			isSkillEnabled = True			: If buildList(3,i) = 0 Then isSkillEnabled = False
			isSkillGroupEnabled = True		: If buildList(6,i) = 0 Then isSkillGroupEnabled = False
			
			availableClass = " class=""not-available"""			: If buildLIst(10,i) = 1 Then availableClass = ""
			
			isThisSkill = True				: If CLng(skillId) <> CLng(buildList(1,i)) Then isThisSkill = False
			publishStatus = buildList(13,i)
			
'			If isThisSkill And isMemberEnabled And isProgramMemberEnabled And isScheduled And isSkillEnabled And isSkillGroupEnabled Then
			If isThisSkill And isMemberEnabled And isProgramMemberEnabled And isSkillEnabled And isSkillGroupEnabled Then
				If (publishStatus = IS_PUBLISHED) Or (publishStatus = IS_MARKED_FOR_PUBLISH) Then
					outOptions = outOptions & "<option value=""" & buildList(0,i) & """" & availableClass & ">"
					outOptions = outOptions & Server.HTMLEncode(buildList(7,i) & ", " & buildList(8,i)) & "</option>"
					outScheduled = outScheduled & buildList(0,i) & ","
				End If
			End If
		Next
		
		If Len(outScheduled) > 0 Then outScheduled = Left(outScheduled, Len(outScheduled) - 1)
	End If
End Sub

</script>