<script runat="server" type="text/vbscript" language="vbscript">
	Function RemoveDupesFromStringList(ByRef str)
		Dim i
		Dim arr
		Dim list
		
		If Len(str) > 0 Then str = Replace(str, " ", "")
		
		arr = Split(str, ",")
		If Not IsArray(arr) Then Exit Function
		
		For i = 0 To UBound(arr)
			If InStr(list, arr(i) & ",") <= 0 Then
				list = list & arr(i) & ","
			End If
		Next
		If Len(list) > 0 Then list = Left(list, Len(list)-1)
		
		RemoveDupesFromStringList = list
	End Function
</script>