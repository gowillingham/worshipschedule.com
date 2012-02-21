<script language="vbscript" type="text/vbscript" runat="server">

Function YesNoDropdownToString(val, inputName)
	Dim str, i
	Dim arr
	Dim selected		: selected = ""
	
	ReDim arr(1,1)
	arr(0,0) = "1"
	arr(0,1) = "0"
	arr(1,0) = "Yes"
	arr(1,1) = "No"
	
	str = str & "<select class=""small"" name=""" & inputName & """>"
	For i = 0 To UBound(arr,2)
		selected = ""
		If CStr(arr(0,i)) = CStr(val) Then selected = " selected=""selected"""
		str = str & "<option value=""" & arr(0,i) & """" & selected & ">" & arr(1,i) & "</option>"
	Next
	str = str & "</select>"
	
	YesNoDropdownToString = str
End Function

</script>