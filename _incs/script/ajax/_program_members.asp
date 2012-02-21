<%
Option Explicit

Call Main()

Sub Main
	Dim str
	
	Dim programId						: programId = Request.QueryString("pid")
	Dim sessionId						: sessionId = Request.QueryString("sid")
	Dim action							: action = Request.QueryString("act")
	Dim idList							: idList = Request.QueryString("id_list")
	
	Select Case action
		Case REMOVE_PROGRAM_MEMBER
			Call DoDeleteProgramMemberByList(idList)
			str = OptionsToJson(programId, sessionId)
		Case ADD_PROGRAM_MEMBER
			Call DoInsertProgramMemberByList(programId, idList)
			str = OptionsToJson(programId, sessionId)
		Case Else
			str = ProgramMemberWidgetToString(programId, sessionId)		
	End Select 
	
	Response.Write str
End Sub

Sub DoDeleteProgramMemberByList(idList)
	Dim i
	
	If Len(idList & "") = 0 Then Exit Sub
	
	Dim list				: list = Split(Replace(idList, " ", ""), ",")
	Dim programMember		: Set programMember = New cProgramMember
	
	If IsArray(list) Then
		For i = 0 To UBound(list)
			programMember.ProgramMemberID = list(i)
			Call programMember.Delete("")		' hack: discard error code ..
		Next
	End If
End Sub

Sub DoInsertProgramMemberByList(programId, idList)
	Dim i
	If Len(idList & "") = 0 Then Exit Sub
	
	Dim list			: list = Split(Replace(idList, " ", ""), ",")
	Dim programMember		: Set programMember = New cProgramMember

	If IsArray(list) Then
		For i = 0 To UBound(list)
			programMember.ProgramID = programID
			programMember.MemberID = list(i)
			programMember.EnrollStatusID = 3
			programMember.IsActive = 1
			Call programMember.Add("")		' hack: discard error code
		Next
	End If
End Sub

Function OptionsToJson(programId, sessionId)
	Dim str
	
	Dim session				: Set session =  New cSession
	session.SessionID = sessionId
	If Len(session.SessionID) > 0 Then Call session.Load()
	 
	Dim client				: Set client = New cClient
	client.ClientId = session.ClientId
	Dim clientMemberList	: clientMemberList = client.MemberList("", "")
	
	Dim programMember		: Set programMember = New cProgramMember
	programMember.ProgramId = programId
	Dim programMemberList	: If Len(programMember.ProgramId) > 0 Then programMemberList = programMember.GetMemberList()	
	
	Call MembersToAddOptionListToString(programMemberList, clientMemberList)
	
	str = str & "{ "
	str = str & "clientMemberOptions: """ & MembersToAddOptionListToString(programMemberList, clientMemberList) & """, "
	str = str & "programMemberOptions: """ & MembersToRemoveOptionListToString(programMemberList) & """ "
	str = str & "}"

	OptionsToJson = str
End Function

Function ProgramMemberWidgetToString(programId, sessionId)
	Dim str
	
	Dim session				: Set session =  New cSession
	session.SessionID = sessionId
	If Len(session.SessionID) > 0 Then Call session.Load()
	 
	Dim client				: Set client = New cClient
	client.ClientId = session.ClientId
	Call client.Load()
	
	Dim clientMemberList	: clientMemberList = client.MemberList("", "")
	
	Dim program				: Set program = New cProgram
	program.ProgramId = programId
	If Len(program.ProgramId) > 0 Then Call program.Load()
	
	Dim programMember		: Set programMember = New cProgramMember
	programMember.ProgramId = programId
	Dim programMemberList	: If Len(programMember.ProgramId) > 0 Then programMemberList = programMember.GetMemberList()	
	
	str = str & "<form id=""form-set-program-members"">"
	str = str & "<p>Higlight the members you wish to add to or remove from the <strong>" & Server.HTMLEncode(program.ProgramName) & "</strong> program and click the relevant arrow button. "
	str = str & "Click <strong>done</strong> when you are finished. </p>"
	str = str & "<table><tr><td><h5>" & server.HTMLEncode(program.ProgramName) & " members</h5>"
	str = str & "<select multiple=""multiple"" id=""members-to-remove-dropdown"" name=""remove_member"">"
	str = str & MembersToRemoveOptionListToString(programMemberList) & "</select></td>"
	str = str & "<td><input type=""button"" id=""add-members-button"" name=""add_member"" value=""" & Server.HTMLEncode("<<") & """ />"
	str = str & "<br /><input type=""button"" id=""remove-members-button"" name=""remove_member"" value=""" & Server.HTMLEncode(">>") & """ /></td>"
	str = str & "<td><h5>" & server.HTMLEncode(client.NameClient) & " members</h5>"
	str = str & "<select multiple=""multiple"" id=""members-to-add-dropdown"" name=""add_member"">"
	str = str & MembersToAddOptionListToString(programMemberList, clientMemberList) & "</select></td>"
	str = str & "</tr></table>"
	str = str & "<p class=""hint"">Tip! You may shift-click or control-click to select multiple members at once. </p>"
	str = str & "</form>"
	
	ProgramMemberWidgetToString = str
End Function

Function MembersToRemoveOptionListToString(programMemberList)
	Dim str, i

	If IsArray(programMemberList) Then
		For i = 0 To UBound(programMemberList,2)
			str = str & "<option value='" & programMemberList(10,i) & "'>" & Server.HTMLEncode(programMemberList(1,i) & ", " & programMemberList(2,i)) & "</option>"
		Next
	End If
	
	MembersToRemoveOptionListToString = str
End Function

Function MembersToAddOptionListToString(programMemberList, clientMemberList)
	Dim str, i, j
	Dim isProgramMember
	
	If IsArray(clientMemberList) Then
		For i = 0 To UBound(clientMemberList,2)
			isProgramMember = False
			If IsArray(programMemberList) Then
				For j = 0 To UBound(programMemberList,2)
					If CLng(programMemberList(0,j)) = CLng(clientMemberList(0,i)) Then
						isProgramMember = True
						Exit For
					End If
				Next
			End If
			
			If Not isProgramMember Then
				str = str & "<option value='" & clientMemberList(0,i) & "'>" & Server.HTMLEncode(clientMemberList(1,i) & ", " & clientMemberList(2,i)) & "</option>"
			End If
		Next
	End If

	MembersToAddOptionListToString = str
End Function
%>

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_Wait.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->
