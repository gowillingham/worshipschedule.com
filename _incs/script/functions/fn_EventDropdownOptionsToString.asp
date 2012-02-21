<script runat="server" type="text/vbscript" language="vbscript">

Function EventDropdownOptionsToString(schedule, eventId)
	Dim str, i
	Dim dateTime		: Set dateTime = New cFormatDate
	
	Dim list			: list = schedule.EventList("")
	Dim selected		: selected = ""
	
	If Not IsArray(list) Then Exit Function
	For i = 0 To UBound(list,2)
		selected = ""
		If CStr(eventId & "") = CStr(list(0,i) & "") Then selected = " selected=""selected"""
		
		str = str & "<option value=""" & list(0,i) & """" & selected & ">" & dateTime.Convert(list(2,i), "MM/dd/YYYY") & " - " & server.HTMLEncode(list(1,i)) & "</option>"
	Next
	
	EventDropdownOptionsToString = str
End Function

</script>