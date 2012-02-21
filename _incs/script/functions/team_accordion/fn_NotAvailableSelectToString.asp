<script runat="server" type="text/vbscript" language="vbscript">

Function NotAvailableSelectToString(optionList)
	Dim str 
	
	str = str & "<select name=""unscheduled_id_list"" multiple=""multiple"" class=""not-available-members"">"
	str = str & optionList
	str = str & "</select>"
	
	NotAvailableSelectToString = str
End Function

</script>