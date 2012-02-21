<script language="vbscript" runat="server" type="text/vbscript">

Function OwnsMember(loginMemberID, memberID)
	Dim i, j
	Dim member						: Set member = New cMember
	member.MemberID = memberID
	member.Load()
	Dim ownsProgram					: ownsProgram = False
	OwnsMember = True
	
	Dim ownedProgramList			: ownedProgramList = GetProgramsOwned(loginMemberID)
	Dim programList					: programList = member.ProgramList()
	If Not IsArray(programList) Then
		' member belongs to no programs
		OwnsMember = True
		Exit Function
	End If
	
	For i = 0 To UBound(programList,2)
		ownsProgram = False
		For j = 0 To UBound(ownedProgramList,2)
			If CLng(programList(0,i)) = CLng(ownedProgramList(0,j)) Then
				ownsProgram = True
			End If 
		Next
		If ownsProgram = False Then
			' found un-owned program
			OwnsMember = False
			Exit Function
		End If
	Next
	
	Set member = Nothing
End Function

</script>
