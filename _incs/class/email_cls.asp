<%
Class cEmail

	Private m_EmailID		'as long int
	Private m_MemberID		'as long int
	Private m_ClientID		'as long int
	Private m_Subject		'as string
	Private m_Text		'as string
	Private m_IsMarkedForDelete		'as small int
	Private m_IsSent		'as small int
	Private m_RecipientIDList		'as string
	Private m_RecipientAddressList		'as string
	Private m_BccAddressList		'as string
	Private m_CcAddressList		'as string
	Private m_GroupList		'as string
	Private m_AttachmentList		'as string
	Private m_DateCreated		'as date
	Private m_DateModified		'as date
	Private m_DateSent		'as date
	
	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	Private m_fso		'as Scripting.FileSystemObject
		
	Private CLASS_NAME	'as string
	Private DAYS_TO_SAVE_ATTACHMENTS
	Private HOURS_TO_SAVE_UNSENT_DRAFTS
		
	Public Property Get EmailID() 'As long int
		EmailID = m_EmailID
	End Property

	Public Property Let EmailID(val) 'As long int
		m_EmailID = val
	End Property
	
	Public Property Get MemberID() 'As long int
		MemberID = m_MemberID
	End Property

	Public Property Let MemberID(val) 'As long int
		m_MemberID = val
	End Property
	
	Public Property Get ClientID() 'As long int
		ClientID = m_ClientID
	End Property

	Public Property Let ClientID(val) 'As long int
		m_ClientID = val
	End Property
	
	Public Property Get Subject() 'As string
		Subject = m_Subject
	End Property

	Public Property Let Subject(val) 'As string
		m_Subject = val
	End Property
	
	Public Property Get Text() 'As string
		Text = m_Text
	End Property

	Public Property Let Text(val) 'As string
		m_Text = val
	End Property
	
	Public Property Get IsMarkedForDelete() 'As small int
		IsMarkedForDelete = m_IsMarkedForDelete
	End Property

	Public Property Let IsMarkedForDelete(val) 'As small int
		m_IsMarkedForDelete = val
	End Property
	
	Public Property Get IsSent() 'As small int
		IsSent = m_IsSent
	End Property

	Public Property Let IsSent(val) 'As small int
		m_IsSent = val
	End Property
	
	Public Property Get RecipientIDList() 'As string
		RecipientIDList = m_RecipientIDList
	End Property

	Public Property Let RecipientIDList(val) 'As string
		m_RecipientIDList = val
	End Property
	
	Public Property Get RecipientAddressList() 'As string
		RecipientAddressList = m_RecipientAddressList
	End Property

	Public Property Let RecipientAddressList(val) 'As string
		m_RecipientAddressList = val
	End Property
	
	Public Property Get BccAddressList() 'As string
		BccAddressList = m_BccAddressList
	End Property

	Public Property Let BccAddressList(val) 'As string
		m_BccAddressList = val
	End Property
	
	Public Property Get CcAddressList() 'As string
		CcAddressList = m_CcAddressList
	End Property

	Public Property Let CcAddressList(val) 'As string
		m_CcAddressList = val
	End Property
	
	Public Property Get GroupList() 'As string
		GroupList = m_GroupList
	End Property

	Public Property Let GroupList(val) 'As string
		m_GroupList = val
	End Property
	
	Public Property Get AttachmentList() 'As string
		AttachmentList = m_AttachmentList
	End Property

	Public Property Let AttachmentList(val) 'As string
		m_AttachmentList = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	Public Property Get DateModified() 'As date
		DateModified = m_DateModified
	End Property

	Public Property Get DateSent() 'As date
		DateSent = m_DateSent
	End Property

	Public Property Let DateSent(val) 'As date
		m_DateSent = val
	End Property
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cEmail"
		DAYS_TO_SAVE_ATTACHMENTS = -6 * Application.Value("EMAIL_ATTACHMENT_MONTHS_TO_LIVE")
		HOURS_TO_SAVE_UNSENT_DRAFTS = -(Application.Value("EMAIL_UNSENT_DRAFTS_HOURS_TO_LIVE"))
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(m_rs) Then
			If m_rs.State = adStateOpen Then m_rs.Close
			Set m_rs = Nothing
		End If
		If IsObject(m_cnn) Then
			If m_cnn.State = adStateOpen Then m_cnn.Close
			Set m_cnn = Nothing
		End If
		If IsObject(m_fso) Then Set m_fso = Nothing
	End Sub
	
	Public Function FirstRecipient()
		If Len(m_EmailID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".GetFirstRecipient():", "Required parameter EmailID not provided.")
	
		Dim recipientList				: If Len(m_recipientAddressList) > 0 Then recipientList = Split(m_recipientAddressList, ",")
		Dim ccList						: If Len(m_ccAddressList) > 0 Then ccList = Split(m_ccAddressList, ",")
		Dim bccList						: If Len(m_bccAddressList) > 0 Then bccList = Split(m_bccAddressList, ",")
		
		Dim hasRecipient				: hasRecipient = False
		Dim hasCc						: hasCc = False
		Dim hasBcc						: hasBcc = False
		Dim hasMultiple					: hasMultiple = False
		Dim totalRecipients				: totalRecipients = 0
		
		Dim recipient					: recipient = ""
		Dim additional					: additional = ""
		
		If IsArray(recipientList) Then
			hasRecipient = True
			totalRecipients = totalRecipients + UBound(recipientList) + 1
			recipient = recipientList(0)
		End If
		If IsArray(ccList) Then
			hasRecipient = True
			totalRecipients = totalRecipients + UBound(ccList) + 1
			If Len(recipient) = 0 Then recipient = ccList(0)
		End If
		If IsArray(bccList) Then
			hasRecipient = True
			totalRecipients = totalRecipients + UBound(bccList) + 1
			If Len(recipient) = 0 Then recipient = bccList(0)
		End If
		
		If totalRecipients > 1 Then additional = "plus " & totalRecipients & " other"
		If totalRecipients > 2 Then additional = additional & "s"
		If Len(additional) > 0 Then additional = " (" & additional & ")"
		
		' extract user name ..
		If Len(recipient) = 0 Then
			FirstRecipient = "multiple recipients"
		Else
			FirstRecipient = Split(recipient, "@")(0) & additional
		End If
	End Function
	
	Public Sub CleanUpHistory(outError)
		Dim i
		Dim tempError	: tempError = 0
		outError = 0

		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".CleanUpHistory():", "Required parameter MemberID not provided.")
		Dim history		: history = List()

		If Not IsArray(list) Then Exit Sub
		
		For i = 0 To UBound(history,2)
			m_EmailID= history(0,i)
			tempError = 0
			
			If history(5,i) = 1 Then
				' sent
				If history(4,i) = 1 Then
					' marked for delete, delete message
					Call Delete(tempError)
				End If
				If history(10,i) < DateAdd("d", DAYS_TO_SAVE_ATTACHMENTS, Now()) Then
					' more than 6 months old, delete attachment store
					Call ClearAttachment()
				End If
			Else
				' unsent 
				If history(8,i) < DateAdd("h", HOURS_TO_SAVE_UNSENT_DRAFTS, Now()) Then
					' draft older than 12 hours, delete message
					Call Delete(tempError)
				End If
			End If
			outError = outError + tempError
		Next
		m_EmailID = ""
	End Sub
	
	Public Sub ClearAttachment()
		Dim path		: path = Application.Value("EMAIL_ATTACHMENTS_DIRECTORY") & "\" & m_EmailID
		
		If Len(m_EmailID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".ClearAttachment():", "Required parameter EmailID not provided.")

		If Not IsObject(m_fso) Then Set m_fso = Server.CreateObject("Scripting.FileSystemObject")
		If m_fso.FolderExists(path) Then
			m_fso.DeleteFolder(path)
		End If
	End Sub
	
	Public Function GetDetailsByIDList(str)
		' accept list of memberID, return list of email details
		
		If Len(str) = 0 Then Exit Function
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-MemberID 1-NameLast 2-NameFirst 3-Email
		m_cnn.up_emailGetMemberDetailsByIDList str, m_rs
		If Not m_rs.EOF Then GetDetailsByIDList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function MessageCount(countType)
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".MessageCount():", "Required parameter MemberID not provided.")

		MessageCount = 0

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		m_cnn.up_emailGetMessageCount CLng(m_MemberID), CInt(countType), m_rs
		If Not m_rs.EOF Then MessageCount = m_rs("MessageCount").Value
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function

	Public Function SentMessageList(pageNumber, pageSize)
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".SentMessageList():", "Required parameter MemberID not provided.")
		If Len(pageNumber) = 0 Then Call Err.Raise(vbObjectError + 4, CLASS_NAME & ".SentMessageList():", "Required parameter PageNumber not provided.")
		If Len(pageSize) = 0 Then Call Err.Raise(vbObjectError + 5, CLASS_NAME & ".SentMessageList():", "Required parameter PageSize not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-EmailID 1-ClientID 2-Subject 3-Text 4-IsMarkedForDelete 5-IsSent 6-RecipientIDList
		' 7-RecipientAddressList 8-BccAddressList 9-CcAddressList 10-DateCreated 11-DateModified 12-DateSent 
		' 13-GroupList 14-AttachmentList 15-RowID
		
		m_cnn.up_emailGetEmailListPaged CLng(m_MemberID), CInt(pageNumber), CInt(pageSize), m_rs
		If Not m_rs.EOF Then SentMessageList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function List()
		If Len(m_MemberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".List():", "Required parameter MemberID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_emailGetEmailList CLng(m_MemberID), m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()

		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Load() 'As Boolean
		If Len(m_EmailID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter EmailID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_emailGetEmail CLng(m_EmailID), m_rs
		If Not m_rs.EOF Then
			m_MemberID = m_rs("MemberID").Value
			m_ClientID = m_rs("ClientID").Value
			m_Subject = m_rs("Subject").Value
			m_Text = m_rs("Text").Value
			m_IsMarkedForDelete = m_rs("IsMarkedForDelete").Value
			m_IsSent = m_rs("IsSent").Value
			m_RecipientIDList = m_rs("RecipientIDList").Value
			m_RecipientAddressList = m_rs("RecipientAddressList").Value
			m_BccAddressList = m_rs("BccAddressList").Value
			m_CcAddressList = m_rs("CcAddressList").Value
			m_GroupList = m_rs("GroupList").Value
			m_AttachmentList = m_rs("AttachmentList").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateSent = m_rs("DateSent").Value
		
			Load = True
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_memberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Add():", "Required parameter MemberID not provided.")
		If Len(m_clientID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".Add():", "Required parameter ClientID not provided.")
		m_DateCreated = Now()
		m_DateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailInsert"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, m_MemberID)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, m_ClientID)
		If Len(m_Subject) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@Subject", adVarChar, adParamInput, 200, Null)
		Else 
			cmd.Parameters.Append cmd.CreateParameter("@Subject", adVarChar, adParamInput, 200, m_Subject)
		End If
		If Len(m_Text) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 4000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 4000, m_Text)
		End If
		If Len(m_IsMarkedForDelete) = 0 Then m_IsMarkedForDelete = 0
		cmd.Parameters.Append cmd.CreateParameter("@IsMarkedForDelete", adUnsignedTinyInt, adParamInput, 0, m_IsMarkedForDelete)
		If Len(m_IsSent) = 0 Then m_IsSent = 0
		cmd.Parameters.Append cmd.CreateParameter("@IsSent", adUnsignedTinyInt, adParamInput, 0, m_IsSent)
		
		cmd.Parameters.Append cmd.CreateParameter("@RecipientIDList", adVarChar, adParamInput, 4000, m_RecipientIDList)
		cmd.Parameters.Append cmd.CreateParameter("@RecipientAddressList", adVarChar, adParamInput, 4000, m_RecipientAddressList)
		cmd.Parameters.Append cmd.CreateParameter("@BccAddressList", adVarChar, adParamInput, 4000, m_BccAddressList)
		cmd.Parameters.Append cmd.CreateParameter("@BccAddressList", adVarChar, adParamInput, 4000, m_CcAddressList)
		
		If Len(m_GroupList) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@GroupList", adVarChar, adParamInput, 4000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@GroupList", adVarChar, adParamInput, 4000, m_GroupList)
		End If
		If Len(m_AttachmentList) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@AttachmentList", adVarChar, adParamInput, 4000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@AttachmentList", adVarChar, adParamInput, 4000, m_AttachmentList)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		If Len(m_DateSent) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@DateSent", adDate, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@DateSent", adDate, adParamInput, 0, m_DateSent)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@NewEmailID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_EmailID = cmd.Parameters("@NewEmailID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_EmailID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter EmailID not provided.")
		If Len(m_memberID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Add():", "Required parameter MemberID not provided.")
		If Len(m_clientID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".Add():", "Required parameter ClientID not provided.")
		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailUpdate"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EmailID", adBigInt, adParamInput, 0, m_EmailID)
		If Len(m_Subject) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@Subject", adVarChar, adParamInput, 200, Null)
		Else 
			cmd.Parameters.Append cmd.CreateParameter("@Subject", adVarChar, adParamInput, 200, m_Subject)
		End If
		If Len(m_Text) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 4000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 4000, m_Text)
		End If
		If Len(m_IsMarkedForDelete) = 0 Then m_IsMarkedForDelete = 0
		cmd.Parameters.Append cmd.CreateParameter("@IsMarkedForDelete", adUnsignedTinyInt, adParamInput, 0, m_IsMarkedForDelete)
		If Len(m_IsSent) = 0 Then m_IsSent = 0
		cmd.Parameters.Append cmd.CreateParameter("@IsSent", adUnsignedTinyInt, adParamInput, 0, m_IsSent)

		cmd.Parameters.Append cmd.CreateParameter("@RecipientIDList", adVarChar, adParamInput, 4000, m_RecipientIDList)
		cmd.Parameters.Append cmd.CreateParameter("@RecipientAddressList", adVarChar, adParamInput, 4000, m_RecipientAddressList)
		cmd.Parameters.Append cmd.CreateParameter("@BccAddressList", adVarChar, adParamInput, 4000, m_BccAddressList)
		cmd.Parameters.Append cmd.CreateParameter("@BccAddressList", adVarChar, adParamInput, 4000, m_CcAddressList)
		
		If Len(m_GroupList) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@GroupList", adVarChar, adParamInput, 4000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@GroupList", adVarChar, adParamInput, 4000, m_GroupList)
		End If
		If Len(m_AttachmentList) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@AttachmentList", adVarChar, adParamInput, 4000, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@AttachmentList", adVarChar, adParamInput, 4000, m_AttachmentList)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_DateModified)
		If Len(m_DateSent) = 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@DateSent", adDate, adParamInput, 0, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter("@DateSent", adDate, adParamInput, 0, m_DateSent)
		End If
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Save = True
		Else
			Save = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Sub Delete(ByRef outError) 'As Boolean
		Dim cmd
		Dim path		: path = Application.Value("EMAIL_ATTACHMENTS_DIRECTORY") & m_EmailID
		
		If Len(m_EmailID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter EmailID not provided.")
		outError = 0
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_emailDelete"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EmailID", adBigInt, adParamInput, 0, m_EmailID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		' delete any file attachments left on the server for this message ..
		Call ClearAttachment()
		
		Set cmd = Nothing
	End Sub

End Class
%>
