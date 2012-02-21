<script runat="server" type="text/vbscript" language="vbscript">

Function AvailableSelectToString(optionList)
	Dim str 
	
	str = str & "<select name=""unscheduled_id_list"" multiple=""multiple"" class=""available-members"">"
	str = str & optionList
	str = str & "</select>"
	
	AvailableSelectToString = str
End Function

</script>