<script runat="server" type="text/vbscript" language="vbscript">

Function ScheduleDropdownOptionsToString(program, scheduleId)
	Dim str, i
	
	Dim list			: list = program.ScheduleList()
	Dim selected		: selected = ""
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		selected = ""
		If CStr(scheduleId & "") = CStr(list(0,i) & "") Then selected = " selected=""selected"""
		
		str = str & "<option value=""" & list(0,i) & """" & selected & ">" & server.HTMLEncode(list(1,i)) & "</option>"
	Next
	
	ScheduleDropdownOptionsToString = str
End Function

</script>