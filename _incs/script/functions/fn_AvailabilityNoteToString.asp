<script runat="server" type="text/vbscript" language="vbscript">

Function AvailabilityNoteToString(note, modifiedDate)
	Dim str

	If Len(note & "") > 0 Then
		str = str & "<div class=""availability-note"">"
		str = str & "<a href=""#"" class=""remove-link"">-remove-</a>"
		str = str & "<h4>You said on " & MonthName(Month(modifiedDate), True) & " " & Day(modifiedDate) & " " & Year(modifiedDate) & ": </h4>"
		str = str & "<p>" & Server.HTMLEncode(note & "") & "</p>"
		str = str & "</div>"
	End If
	
	AvailabilityNoteToString = str		
End Function

</script>