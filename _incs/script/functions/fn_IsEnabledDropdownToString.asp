<script language="vbscript" type="text/vbscript" runat="server">

Function IsEnabledDropdownToString(val)
	Dim str, arr
	
	ReDim arr(1,1)
	arr(0,0) = "0"
	arr(0,1) = "1"
	arr(1,0) = "Disabled"
	arr(1,1) = "Enabled"
	
	str = str & "<select name=""IsEnabled"">"
	str = str & SelectOption(arr, val)
	str = str & "</select>"
	
	IsEnabledDropdownToString = str
End Function

</script>