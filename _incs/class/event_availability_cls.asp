<script runat="server" language="vbscript" type="text/vbscript">
Class cEventAvailability

	Private m_iEventAvailabilityID		'as long int
	Private m_iMemberID		'as long int
	Private m_iEventID		'as long int
	Private m_sMemberNote		'as string
	Private m_iIsAvailable		'as small int
	Private m_dDateModified		'as date
	Private m_dDateCreated		'as date
	Private m_iIsViewedByMember		'as small int
	Private m_EventName		' as str
	Private m_EventDate		'as smalldatetime
	Private m_TimeStart		'as smalldatetime
	Private m_TimeEnd		'as smalldatetime
	
	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	Private m_iError	'as int
	Private m_bIsDirty  'as bool
	
	Private CLASS_NAME	'as string
	
	Public Property Get EventAvailabilityID() 'As long int
		EventAvailabilityID = m_iEventAvailabilityID
	End Property

	Public Property Let EventAvailabilityID(val) 'As long int
		m_iEventAvailabilityID = val
		m_bIsDirty = True
	End Property
	
	Public Property Get MemberID() 'As long int
		MemberID = m_iMemberID
	End Property

	Public Property Let MemberID(val) 'As long int
		m_iMemberID = val
		m_bIsDirty = True
	End Property
	
	Public Property Get EventID() 'As long int
		EventID = m_iEventID
	End Property

	Public Property Let EventID(val) 'As long int
		m_iEventID = val
		m_bIsDirty = True
	End Property
	
	Public Property Get MemberNote() 'As string
		MemberNote = m_sMemberNote
	End Property

	Public Property Let MemberNote(val) 'As string
		m_sMemberNote = val
		m_bIsDirty = True
	End Property
	
	Public Property Get IsAvailable() 'As small int
		IsAvailable = m_iIsAvailable
	End Property

	Public Property Let IsAvailable(val) 'As small int
		m_iIsAvailable = val
		m_bIsDirty = True
	End Property
	
	Public Property Get DateModified() 'As date
		DateModified = m_dDateModified
	End Property

	Public Property Get DateCreated() 'As date
		DateCreated = m_dDateCreated
	End Property

	Public Property Get IsViewedByMember() 'As small int
		IsViewedByMember = m_iIsViewedByMember
	End Property

	Public Property Let IsViewedByMember(val) 'As small int
		m_iIsViewedByMember = val
		m_bIsDirty = True
	End Property
	
	Public Property Let EventName(val) 'As smalldatetime
		m_EventName = val
	End Property
	
	Public Property Get EventName() 'As small int
		EventName = m_EventName
	End Property

	Public Property Let EventDate(val) 'As smalldatetime
		m_EventDate = val
	End Property
	
	Public Property Get EventDate() 'As small int
		EventDate = m_EventDate
	End Property

	Public Property Let TimeStart(val) 'As smalldatetime
		m_TimeStart = val
	End Property
	
	Public Property Get TimeStart() 'As small int
		TimeStart = m_TimeStart
	End Property

	Public Property Let TimeEnd(val) 'As smalldatetime
		m_TimeEnd = val
	End Property
	
	Public Property Get TimeEnd() 'As small int
		TimeEnd = m_TimeEnd
	End Property

	Public Property Get IsDirty()
		IsDirty = m_bIsDirty
	End Property
	
	Private Sub Class_Initialize()
		m_sMemberNote = ""
		m_iIsAvailable = 0
		m_dDateModified = ""
		m_dDateCreated = ""
		m_iIsViewedByMember = 0
	
		m_iError = 0
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cEventAvailability"
		m_bIsDirty = False
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
	End Sub
	
	Public Function AvailabilityList(programID, scheduleID, sortColumn)
		If Len(memberID) = 0 Then Call Err.Raise(1 + vbObjectError, CLASS_NAME & ".AvailabilityList();", "Missing required parameter MemberID.")
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		m_rs.CursorLocation = adUseClient	' makes sortable
		
		' 0-EventAvailabilityID 1-MemberId 2-MemberNote 3-IsAvailable 4-IsViewedByMember
		' 5-EventAvailabilityDateCreated 6-EventAvailabilityDateCreated 7-EventId 8-EventName
		' 9-EventDate 10-TimeStart 11-TimeEnd 12-EventDescription 13-ScheduleId 
		' 14-ScheduleName 15-ScheduleIsVisible 16-ProgramId 17-ProgramName
		' 18-ProgramIsEnabled 19-IsScheduled
		
		If Len(programID) = 0 Then
			m_cnn.up_memberGetAvailabilityList CLng(m_iMemberID), m_rs
		Else
			If Len(scheduleID) = 0 Then
				m_cnn.up_memberGetAvailabilityList CLng(m_iMemberID), CLng(programID), m_rs
			Else
				m_cnn.up_memberGetAvailabilityList CLng(m_iMemberID), CLng(programID), CLng(scheduleID), m_rs
			End If
		End If
		
		m_rs.Sort = sortColumn
		
		If Not m_rs.EOF Then AvailabilityList = m_rs.GetRows()
	
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function EventAvailabilityList()
		If Len(m_iEventId) = 0 Then Call Err.Raise(4 + vbObjectError, CLASS_NAME & ".EventAvailabilityList()", "Missing required paramenter m_iEventID")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")

		' 0-MemberId 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberIsActive 
		' 5-EventAvailabilityID 6-IsAvailable 7-IsViewedByMember 8-DateAvailabilityModified
		
		m_cnn.up_eventGetAvailabilityListByEventID CLng(m_iEventId), m_rs
		If Not m_rs.EOF Then EventAvailabilityList = m_rs.GetRows()
	
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Load() 'As Boolean
		If Len(m_iEventAvailabilityID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Load();", "Missing required parameter m_iEventAvailabilityID.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_eventGetAvailability CLng(m_iEventAvailabilityID), m_rs
		If Not m_rs.EOF Then
			m_iMemberID = m_rs("MemberID").Value
			m_iEventID = m_rs("EventID").Value
			m_sMemberNote = m_rs("MemberNote").Value
			m_iIsAvailable = m_rs("IsAvailable").Value
			m_dDateModified = m_rs("DateModified").Value
			m_dDateCreated = m_rs("DateCreated").Value
			m_iIsViewedByMember = m_rs("IsViewedByMember").Value
			m_EventName = m_rs("EventName").Value
			m_EventDate = m_rs("EventDate").Value
			m_TimeStart = m_rs("TimeStart").Value
			m_TimeEnd = m_rs("TimeEnd").Value
		
			Load = True
			m_bIsDirty = False
		Else
			Load = False
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function Add(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_iMemberID) = 0 Then Call Err.Raise(3 + vbObjectError, CLASS_NAME & ".Add();", "Missing required parameter m_iMemberID.")
		If Len(m_iEventID) = 0 Then Call Err.Raise(3 + vbObjectError, CLASS_NAME & ".Add();", "Missing required parameter m_iEventID.")
		
		m_dDateCreated = Now()
		m_dDateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_eventInsertAvailability"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(m_iMemberID))
		cmd.Parameters.Append cmd.CreateParameter("@EventID", adBigInt, adParamInput, 0, CLng(m_iEventID))
		If Len(m_sMemberNote) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@MemberNote", adVarChar, adParamInput, 1000, CStr(m_sMemberNote))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@MemberNote", adVarChar, adParamInput, 1000, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@IsAvailable", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsAvailable))
		cmd.Parameters.Append cmd.CreateParameter("@IsViewedByMember", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsViewedByMember))
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_dDateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@NewEventAvailabilityID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			m_iEventAvailabilityID = cmd.Parameters("@NewEventAvailabilityID").Value
			Add = True
		Else
			Add = False
		End If

		Set cmd = Nothing
	End Function
	
	Public Function Save(ByRef outError) 'As Boolean
		Dim cmd
		
		m_dDateModified = Now()
		
		If Len(m_iEventAvailabilityID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Load();", "Missing required parameter m_iEventAvailabilityID.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_eventUpdateAvailability"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventAvailabilityID", adBigInt, adParamInput, 0, CLng(m_iEventAvailabilityID))
		If Len(m_sMemberNote) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@MemberNote", adVarChar, adParamInput, 1000, CStr(m_sMemberNote))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@MemberNote", adVarChar, adParamInput, 1000, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@IsAvailable", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsAvailable))
		cmd.Parameters.Append cmd.CreateParameter("@IsViewedByMember", adUnsignedTinyInt, adParamInput, 0, CInt(m_iIsViewedByMember))
		cmd.Parameters.Append cmd.CreateParameter("@DateModified", adDate, adParamInput, 0, m_dDateModified)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Save = True
		Else
			Save = False
		End If
		
		Set cmd = Nothing
	End Function
	
	Public Function Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_iEventAvailabilityID) = 0 Then Call Err.Raise(2 + vbObjectError, CLASS_NAME & ".Load();", "Missing required parameter m_iEventAvailabilityID.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_eventDeleteAvailability"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@EventAvailabilityID", adBigInt, adParamInput, 0, CLng(m_iEventAvailabilityID))

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		If outError = 0 Then
			Delete = True
		Else
			Delete = False
		End If
		
		Set cmd = Nothing
	End Function

	Private Function RaiseError(ByVal id)
		Dim errorText
		
		errorText = errorText & "ERROR: Class " & CLASS_NAME & ":: Code:" & id & ":: Description:"
		
		Select Case id
			Case 1000
				'missing EventAvailabilityID
				errorText = errorText & "EventAvailabilityID not provided."
			Case 1001
				'missing MemberID
				errorText = errorText & "MemberID not provided."
			Case 1002
				'missing EventID
				errorText = errorText & "EventID not provided."
			Case 1003
				'missing ProgramID
				errorText = errorText & "ProgramID not provided."
			Case Else
				'unknown error
				errorText = errorText & "Unspecified error."
		End Select
		Response.Write "<div style=""font-size:1em;"">" & errorText & "</div>"
		Response.End
	End Function
End Class
</script>