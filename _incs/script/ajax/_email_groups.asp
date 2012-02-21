<%
Option Explicit

Dim m_dialogNotifyMessage

Call Main()

Sub Main
	Dim str, rv
	Dim sess					: Set sess = New cSession
	
	Dim action					: action			= Request.Form("act")			' todo: change back to post 		
	Dim emailGroupId			: emailGroupId		= Request.Form("emgid")			' todo: change back to post ..
	Dim member_id_list			: member_id_list	= Request.Form("member_id_list")	' todo: change back to post ..	
	
	
'		Dim action					: action			= Request.querystring("act")			' todo: change back to post 		
'		Dim emailGroupId			: emailGroupId		= Request.querystring("emgid")			' todo: change back to post ..
'		Dim member_id_list			: member_id_list	= Request.querystring("member_id_list")	' todo: change back to post ..	
	Dim sessionId				: sessionId			= Request.Form("sid")							
	Dim email_group_name		: email_group_name	= Request.Form("email_group_name")				
	
	If Len(sessionId) > 0 Then
		sess.SessionID = sessionId
		Call sess.Load()
	End If
	
	Call Wait(0.25)		' debug: wait half second ..
	
	Select Case action
		Case INSERT_EMAIL_GROUP_MEMBERS
			Call DoInsertEmailGroupMembers(emailGroupId, member_id_list, rv)
			
		Case DELETE_EMAIL_GROUP_MEMBERS
			Call DoDeleteEmailGroupMembers(emailGroupId, member_id_list, rv)
			
			str = str & "<p>delete ok: " & member_id_list
					
		Case DELETE_RECORD
			Call DoDeleteEmailGroup(emailGroupId, rv)
			str = JsonToString(action, sess.MemberId, "")
			
		Case UPDATE_RECORD
			If ValidEmailGroup(email_group_name) Then
				Call DoUpdateEmailGroup(email_group_name, emailGroupId, rv)
				str = JsonToString(action, sess.MemberId, emailGroupId)
			Else
				str = JsonToString(action, sess.MemberId, emailGroupId)
			End If
				
		
		Case ADDNEW_RECORD
			If ValidEmailGroup(email_group_name) Then
				Call DoInsertEmailGroup(sess.MemberId, email_group_name, emailGroupId, rv)
				str = JsonToString(action, sess.MemberId, emailGroupId)
			Else
				str = JsonToString(action, sess.MemberId, emailGroupId)
			End If
				
		Case Else
			str = FormEmailGroupToString(emailGroupId)
	End Select
	
	Response.Write str
End Sub

Function JsonToString(action, memberId, emailGroupId)
	Dim str
	
	Dim nodes			: nodes = ""
	Dim optionList		: optionList = ""
	Dim form			: form = ""
	Dim errorMessage	: errorMessage = ""
	
	Select Case action
		Case INSERT_EMAIL_GROUP_MEMBERS
		Case DELETE_EMAIL_GROUP_MEMBERS
		Case DELETE_RECORD
			nodes = EmailGroupNodesToString(memberId, "")
			optionList = EmailGroupMemberDropdownOptionsToString(memberId)

		Case UPDATE_RECORD
			nodes = EmailGroupNodesToString(memberId, emailGroupId)
			optionList = EmailGroupMemberDropdownOptionsToString(memberId)
			form = FormEmailGroupToString(emailGroupId)

		Case ADDNEW_RECORD
			nodes = EmailGroupNodesToString(memberId, emailGroupId)
			optionList = EmailGroupMemberDropdownOptionsToString(memberId)
			form = FormEmailGroupToString(emailGroupId)

		Case Else
		
	End Select
	
	errorMessage = DialogNotifyToString()
	
	str = str & "{"
	str = str & " nodes: '" & nodes & "'"
	str = str & ", optionList: '" & optionList & "'"
	str = str & ", form: '" & form & "'"
	str = str & ", errorMessage: '" & errorMessage & "'"
	str = str & "}"
	
	JsonToString = str
End Function

Sub DoDeleteEmailGroupMembers(emailGroupId, member_id_list, outError)
	Dim i
	
	Dim emailGroupMember				: Set emailGroupMember = New cemailGroupMember
	emailGroupMember.emailGroupId = CLng(emailGroupId)
	
	Dim list							: list = Split(member_id_list, ",")
	
	If Not IsArray(list) Then Exit Sub
	
	For i = 0 To UBound(list)
		Call emailGroupMember.DeleteByMemberId(CLng(list(i)), outError)
		response.write "<p>emailGroupMember.DeleteByMemberId(): " & list(i)	& " outError: " & outError
	Next 
End Sub

Sub DoInsertEmailGroupMembers(emailGroupId, member_id_list, outError)
	Dim i
	
	Dim emailGroupMember				: Set emailGroupMember = New cemailGroupMember
	emailGroupMember.emailGroupId = emailGroupId
	
	Dim idList							: idList = Split(member_id_list, ",")
	Dim list							: list = emailGroupMember.List()
	
	' remove existing members ..
	Dim emailGroup						: Set emailGroup = New cEmailGroup
	emailGroup.EmailGroupId = emailGroupId
	
	If Not IsArray(idList) Then Exit Sub
	For i = 0 To UBound(idList)
		If Not IsGroupMember(list, idList(i)) Then
			emailGroupMember.MemberId = idList(i)
			Call emailGroupMember.Add(outError)
		End If
	Next
End Sub

Function IsGroupMember(list, id)
	Dim i
	IsGroupMember = False
	
	' 0-EmailGroupMemberID 1-EmailGroupId 2-MemberID 3-Email 4-NameLast 
	' 5-NameFirst 6-DateCreated 7-MemberActiveStatus

	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		If CStr(list(2,i) & "") = CStr(id & "") Then
			isGroupMember = True
			Exit For
		End If
	Next
End Function

Sub DoDeleteEmailGroup(emailGroupId, outError)
	Dim emailGroup				: Set emailGroup = New cEmailGroup
	
	emailGroup.EmailGroupId = emailGroupId
	If Len(emailGroup.EmailGroupId) > 0 Then
		Call emailGroup.Delete(outError)
	End If
End Sub

Sub DoUpdateEmailGroup(name, id, outError)
	Dim emailGroup				: Set emailGroup = New cEmailGroup
	
	emailGroup.EmailGroupId = id
	Call emailGroup.Load()
	
	emailGroup.Name = name
	Call emailGroup.Save(outError)
End Sub

Sub DoInsertEmailGroup(memberId, emailGroupName, newId, outError)
	Dim emailGroup				: Set emailGroup = New cEmailGroup
	
	emailGroup.MemberId = memberId
	emailGroup.Name = emailGroupName
	Call emailGroup.Add(outError)
	
	newId = emailGroup.EmailGroupId
End Sub

Function EmailGroupNodesToString(memberId, emailGroupId)
	Dim str
	
	str = str & CustomEmailGroupItemsToString(memberId, emailGroupId)
	
	EmailGroupNodesToString = str
End Function

Function ValidEmailGroup(email_group_name)
	ValidEmailGroup = True
	
	If Not ValidData(email_group_name, True, 0, 200, "Email Group Name", "") Then 
		ValidEmailGroup = False
		Call AddDialogNotifyMessage("You forgot to provide a name for your group. ")
	End If
End Function

Function AddDialogNotifyMessage(message)
	Dim str
	
	str = "<li>" & message & "</li>"
	
	m_dialogNotifyMessage = m_dialogNotifyMessage + str		
End Function

Function DialogNotifyToString()
	Dim str
	
	If Len(m_dialogNotifyMessage) > 0 Then
		str = str & "<div class=""dialog-notify-message""><p>Oops! Please check your info ..</p>"
		str = str & "<ul>" & m_dialogNotifyMessage & "</ul></div>"
	End If
	
	DialogNotifyToString = str
End Function

Function FormEmailGroupToString(id)
	Dim str
	
	Dim emailGroup			: Set emailGroup = New cEmailGroup
	emailGroup.EmailGroupId = id
	If Len(emailGroup.EmailGroupId) > 0 Then Call emailGroup.Load()
	
	Dim action
	If Len(id) = 0 Then
		str = str & "<p>Provide a name for your new email group and click <strong>Save</strong> (you can add members later). </p>"
		action = ADDNEW_RECORD
	Else
		str = str & "<p>Provide a new name for your email group and click <strong>Save</strong>. </p>"
		action = UPDATE_RECORD
	End If
	
	str = str & DialogNotifyToString()
	str = str & "<form method=""post"" action=""/_incs/script/ajax/_email_groups.asp"" id=""form-email-group"">"
	str = str & "<input type=""hidden"" name=""act"" value=""" & action & """ />"
	str = str & "<input type=""hidden"" name=""emgid"" value=""" & emailGroup.EmailGroupId & """ />"
	str = str & "<input type=""hidden"" name=""sid"" value="""" id=""session-id"" />"
	str = str & "<table><tbody>"
	str = str & "<tr><td class=""label"">Group name</td>"
	str = str & "<td><input type=""text"" name=""email_group_name"" value=""" & Server.HTMLEncode(emailGroup.Name) & """ /></td></tr>"
	str = str & "</tbody></table></form>"
	
	FormEmailGroupToString = str
End Function
%>

<!--#INCLUDE VIRTUAL="/_incs/class/email_group_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/email_group_member_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_Wait.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_CustomEmailGroupItemsToString.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_EmailGroupMemberDropdownOptionsToString.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/forms/frm_valid_data.asp"-->
