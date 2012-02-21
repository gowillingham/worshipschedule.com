<script runat="server" type="text/vbscript" language="vbscript">

	' accept two-dimensional array and an dimension index for that array, and return
	' a one-dimensional array of unique values for that dimension

	Function ArrayDimensionToList(ByVal masterList, idx)
		Dim list(), thisValue, isNewValue, i, j
		
		ReDim list(0)
		thisValue = masterList(idx,0)
		list(0) = thisValue
		For i = 0 To UBound(masterList,2)
		
			' check if this name is different than previous from masterList
			If CStr(masterList(idx,i)) <> CStr(thisValue) Then
			
				' check if it has already been stored in list
				isNewValue = True
				For j = 0 To UBound(list)
					If CStr(masterList(idx,i)) = CStr(list(j)) Or Len(masterList(idx,i)) = 0 Then
						isNewValue = False
						Exit For
					End If
				Next
				
				' if it is new then store it
				If isNewValue Then
					ReDim Preserve list(UBound(list) + 1)
					list(UBound(list)) = masterList(idx,i)
				End If
			End If
			thisValue = masterList(idx,i)
		Next
		
		ArrayDimensionToList = list
	End Function
</script>