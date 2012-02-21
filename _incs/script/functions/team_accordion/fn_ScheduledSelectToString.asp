<script runat="server" type="text/vbscript" language="vbscript">

Function ScheduledSelectToString(optionList)
	Dim str 
	
	str = str & "<select name=""scheduled_id_list"" multiple=""multiple"" class=""scheduled-members"">"
	str = str & optionList
	str = str & "</select>"
	
	ScheduledSelectToString = str
End Function

</script>