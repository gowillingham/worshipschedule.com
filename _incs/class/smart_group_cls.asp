<script type="text/vbscript" runat="server" language="vbscript">
Class cSmartGroup
	Private m_smartGroupID

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Private IDX_MEMBER_ID 'as int
	Private IDX_LAST_NAME 'as int
	Private IDX_FIRST_NAME 'as int
	Private IDX_EMAIL 'as int
	Private RETURNED_FIELD_COUNT 'as it
	
	Private m_member_id
	Private m_last_name
	Private m_first_name
	Private m_email
	Private m_account_enabled
	Private m_program_member_enabled
	Private m_skill_enabled
	
	Public Property Let smartGroupID(val)
		m_smartGroupID = val
	End Property
	
	Public Property Get smartGroupID
		smartGroupID = m_smartGroupID
	End Property
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		
		CLASS_NAME = "cSmartGroup"
		
		IDX_MEMBER_ID		= 0
		IDX_LAST_NAME		= 1
		IDX_FIRST_NAME		= 2
		IDX_EMAIL			= 3
		
		' number of fields returned
		RETURNED_FIELD_COUNT = 4
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
	
	Private Function GroupType()
		If Len(m_smartGroupID) = 0 Then Exit Function
		GroupType = Split(m_smartGroupID, "|")(0)
	End Function
	
	Private Function GroupTypeID()
		If Len(m_smartGroupID) = 0 Then Exit Function
		GroupTypeID = Split(m_smartGroupID, "|")(1)
	End Function
	
	Public Function Name()
		If Len(m_smartGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Name();", "No SmartGroupID provided.")
		
		Dim str				: str = ""

		Dim skillGroup				
		Dim skill					
		Dim schedule				
		Dim evnt					
		Dim program					
		Dim emailGroup				
				
		Select Case GroupType
			Case "skillgroup"
				Set skillGroup = New cSkillGroup
				skillGroup.SkillGroupID = GroupTypeID()
				Call skillGroup.Load()
				str = "Skill Group:" & skillGroup.ProgramName & "/" & skillGroup.GroupName
				
			Case "skill"
				Set skill = New cSkill
				skill.SkillID = GroupTypeID()
				Call skill.Load()
				str = "Skill:" & skill.ProgramName & "/" & skill.SkillName				
				
			Case "availability"
				Set schedule = New cSchedule
				schedule.ScheduleID = GroupTypeID()
				Call schedule.Load()
				str = "Missing Availability:" & schedule.ScheduleName
				
			Case "event"
				Set evnt = New cEvent
				evnt.EventID = GroupTypeID()
				Call evnt.Load()
				str = "Event:" & evnt.EventName & "(" & evnt.EventDate & ")"
				
			Case "schedule"
				Set schedule = New cSchedule
				schedule.ScheduleID = GroupTypeID()
				Call schedule.Load()
				str = "Schedule:" & schedule.ProgramName & "/" & schedule.ScheduleName
				
			Case "program"
				Set program = New cProgram
				program.ProgramID = GroupTypeID()
				Call program.Load()
				str = "Program:" & program.ProgramName
				
			Case "emailgroup"
				Set emailGroup = New cEmailGroup
				emailGroup.EmailGroupID = GroupTypeID()
				Call emailGroup.Load()
				str = "My Groups:" & emailGroup.Name
				
			Case Else
				Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".Name();", "Unknown group type '" & GroupType() & "'.")
		End Select
		
		Name = str
	End Function
	
	Public Function GetAddressListAsString()
		Dim i
		
		If Len(m_smartGroupID) = 0 Then Exit Function
		
		Dim list			: list = MemberList()
		Dim addressList		: addressList = ""
		
		If IsArray(list) Then
			For i = 0 To UBound(list,2)
				addressList = addressList & list(IDX_EMAIL, i) & ","
			Next
			If Len(addressList) > 0 Then addressList = Left(addressList, Len(addressList) - 1)
		End If
		
		GetAddressListAsString = addressList
	End Function
	
	Public Function MemberList()
		Dim skillGroup
		Dim skill
		Dim availability
		Dim schedule
		Dim evnt
		Dim program
		Dim emailGroup
	
		Dim map											: map = ""
		
		If Len(m_smartGroupID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".MemberList();", "No SmartGroupID provided.")
		
		Select Case GroupType
			Case "skillgroup"
				' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 
				' 6-SkillIsEnabled 10-Email 11-ProgramMemberIsEnabled

				Set skillGroup = New cSkillGroup
				skillGroup.SkillGroupID = GroupTypeID

				map = map & "MemberID:0,"
				map = map & "LastName:1,"
				map = map & "FirstName:2,"
				map = map & "Email:10,"
				map = map & "MemberEnabled:3,"
				map = map & "ProgramMemberEnabled:11,"
				map = map & "ProgramMemberSkillEnabled:6"
				
				MemberList = GetListByMap(skillGroup.MemberList, map)
			
			Case "skill"
				' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-ProgramMemberSkillID
				' 5-ProgramMemberID 6-IsApproved 7-ProgramMemberIsActive 8-Email
			
				Set skill = New cSkill
				skill.SkillID = GroupTypeID			
				
				map = map & "MemberID:0,"
				map = map & "LastName:1,"
				map = map & "FirstName:2,"
				map = map & "Email:8,"
				map = map & "MemberEnabled:3,"
				map = map & "ProgramMemberEnabled:7"
				
				MemberList = GetListByMap(skill.MemberList(), map)
			
			Case "availability"
				' 0-MemberID 1-NameLast 2-NameFirst 3-LastLogin 4-Email 5-IsMemberAccountEnabled
				' 6-IsMissingAvailabilityInfo 7-IsProgramMemberEnabled
				
				Set schedule = New cSchedule
				schedule.ScheduleID = GroupTypeID()
				
				map = map & "MemberID:0,"
				map = map & "LastName:1,"
				map = map & "FirstName:2,"
				map = map & "Email:4,"
				map = map & "MemberEnabled:5,"
				map = map & "ProgramMemberEnabled:7,"
				map = map & "IsMissingAvailabilityInfo:6"
				
				MemberList = GetListByMap(schedule.AvailabilityList(), map)
				
			Case "event"
				' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-SkillName 5-IsAvailable 
				' 6-IsViewedByMember 7-email 8-ProgramMemberIsActive
			
				Set evnt = New cEvent
				evnt.EventID = GroupTypeID()
				
				map = map & "MemberID:0,"
				map = map & "LastName:1,"
				map = map & "FirstName:2,"
				map = map & "Email:7,"
				map = map & "MemberEnabled:3,"
				map = map & "ProgramMemberEnabled:8"
				
				MemberList = GetListByMap(evnt.MemberList(), map)
				
			Case "schedule"
				' 0-MemberID 1-NameLast 2-NameFirst 3-MemberActiveStatus 4-EventID 5-IsAvailable 6-IsViewedByMember
				' 7-SchedulePublishID 8-SkillName 9-email 10-ProgramMemberIsActive
				
				Set schedule = New cSchedule
				schedule.ScheduleID = GroupTypeID()
				
				map = map & "MemberID:0,"
				map = map & "LastName:1,"
				map = map & "FirstName:2,"
				map = map & "Email:8,"
				map = map & "MemberEnabled:3,"
				map = map & "ProgramMemberEnabled:9"
				
				MemberList = GetListByMap(schedule.MemberList(), map)
				
			Case "program"
				' 0-MemberID 1-NameLast 2-NameFirst 6-ProgramMemberIsEnabled 7-MemberActiveStatus 8-DateCreated 9-DateModified
				' 10-ProgramMemberID 11-IsApproved 12-HasMissingAvailability 13-Email
				
				Set program = New cProgram
				program.ProgramID = GroupTypeID()
				
				map = map & "MemberID:0,"
				map = map & "LastName:1,"
				map = map & "FirstName:2,"
				map = map & "Email:13,"
				map = map & "MemberEnabled:7,"
				map = map & "ProgramMemberEnabled:6"
				
				MemberList = GetListByMap(program.MemberList(), map)
				
			Case "emailgroup"
				' 0-EmailGroupMemberID 1-EmailGroupId 2-MemberID 3-Email 4-NameLast 
				' 5-NameFirst 6-DateCreated 7-MemberActiveStatus
			
				Set emailGroup = New cEmailGroup
				emailGroup.EmailGroupID = GroupTypeID()
			
				map = map & "MemberID:2,"
				map = map & "LastName:4,"
				map = map & "FirstName:5,"
				map = map & "Email:3,"
				map = map & "MemberEnabled:7"
				
				MemberList = GetListByMap(emailGroup.MemberList(), map)
				
			Case Else
				Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".MemberList();", "Unknown group type '" & GroupType() & "'.")
		End Select
	End Function
	
	Private Function GetListByMap(list, map)
		Dim outList(), i, idx
		Dim mapList
		Dim dict								: Set dict = Server.CreateObject("Scripting.Dictionary")
		Dim uniqueIdStringList					: uniqueIdStringList = ""
		Dim lastRow								: lastRow = 0
		Dim addThisRow							: addThisRow = True
		
		If IsArray(list) Then
			mapList = Split(map, ",")
			

			If IsArray(mapList) Then
			
				' hold the map/hash in dictionary object
				For i = 0 To UBound(mapList)
					dict.Add Split(mapList(i), ":")(0), Split(mapList(i), ":")(1)
				Next
				
				' initialize the array (
				ReDim Preserve outList(RETURNED_FIELD_COUNT - 1, 0)
			Else
				Err.Raise vbObjectError + 1, CLASS_NAME & ".GetListByMap()", "Required parameter 'map' is not a valid hash."
			End If
			
			For i = 0 To UBound(list,2)
				addThisRow = True
			
				' test for duplicate rows
				If IsDuplicate(list(dict(0), i), uniqueIdStringList) Then
					addThisRow = False
				End If
				
				' test for memberEnabled
				If Len(dict.Item("MemberEnabled")) > 0 Then
					If list(dict.Item("MemberEnabled"), i) = 0 Then
						addThisRow = False
					End If
				End If 
				
				' test for programMemberEnabled
				If Len(dict.Item("ProgramMemberEnabled")) > 0 Then
					If list(dict.Item("ProgramMemberEnabled"), i) = 0 Then
						addThisRow = False
					End If
				End If
				
				' test for programMemberSkillEnabled
				If Len(dict.Item("ProgramMemberSkillEnabled")) > 0 Then
					If list(dict.Item("ProgramMemberSkillEnabled"), i) = 0 Then
						addThisRow = False
					End If
				End If
				
				' test for isMissingAvailabilityInfo
				If Len(dict.Item("IsMissingAvailabilityInfo")) > 0 then
					If list(dict.Item("IsMissingAvailabilityInfo"), i) = 0 Then
						addThisRow = False
					End If
				End If
				
				If addThisRow Then

					' add this memberid to the string list of unique IDs ..
					If Len(uniqueIdStringList) > 0 Then uniqueIdStringList = uniqueIdStringList & ","
					uniqueIdStringList = uniqueIdStringList & list(dict.Item("MemberID"), i)
				
					' load outlist with new record ..
					lastRow = UBound(outList,2)
					
					outList(IDX_MEMBER_ID, lastRow) = list(dict.Item("MemberID"), i)
					outList(IDX_LAST_NAME, lastRow) = list(dict.Item("LastName"), i)
					outList(IDX_FIRST_NAME, lastRow) = list(dict.Item("FirstName"), i)
					outList(IDX_EMAIL, lastRow) = list(dict.Item("Email"), i)

					' incremement outlist array by one ..
					ReDim Preserve outList(UBound(outList,1), lastRow + 1)
				End If
			Next
			
			' remove last item from array if empty ..
			If Len(outList(0,UBound(outList,2))) = 0 Then
				ReDim Preserve outList(UBound(outList,1), UBound(outList,2) - 1)
			End If
			
			GetListByMap = outList
		Else
			' input var list was not an array so do nothing ..
		End If
	End Function

	Private Function IsDuplicate(id, stringList)
		Dim i
		
		IsDuplicate = False
		If Len(stringList) = 0 Then Exit Function
		
		Dim list			: list = Split(stringList, ",")
		
		For i = 0 To UBound(list)
			If CLng(id) = CLng(list(i)) Then
				IsDuplicate = True
				Exit For
			End If
		Next
	End Function
	
End Class
</script>
