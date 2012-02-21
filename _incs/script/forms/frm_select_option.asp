<%
'*******************************************************
'*	Function SelectOption(arr, sSelVal) - fills and writes option tags for select control
'*	Function MultiSelectOption(arr, arrSel)
'*	Function CheckboxArray(arr, arrSel, sName)
'*	Function RadioArray(arr, sSelVal, bNone, sName)
'*******************************************************

Function RadioArray(arr, sSelVal, bNone, sName)
	Dim str, sNone, i
	If IsArray(arr) Then
		If bNone Then str = str & "<div><input type=""radio"" name=""" & sName & """ value="""" style=""margin:0px;margin-right:5px;padding:0px;width:15px;vertical-align:middle;"" checked=""checked"" />None</div>" & vbCrLf
		For i = 0 To UBound(arr,2)
			str = str & "<div><input type=""radio"" name=""" & sName & """ value=""" & arr(0,i) & """" & IsChecked(arr(0,i), sSelVal) & " style=""margin:0px;margin-right:5px;padding:0px;width:15px;vertical-align:middle;"" />" & Server.HTMLEncode(arr(1,i)) & "</div>" & vbCrLf
		Next
	End If
	RadioArray = str
End Function

Function CheckboxArray(arr, arrSel, sName)
	'arrSel is 2-dimensional array
	Dim str, i
	If IsArray(arr) Then
		For i = 0 To UBound(arr,2)
			str = str & "<div><input type=""checkbox"" name=""" & sName & """ value=""" & arr(0,i) & """" & IsChecked(arr(0,i), arrSel) & " style=""margin:0px;margin-right:5px;padding:0px;width:15px;"" />" & Server.HTMLEncode(arr(1,i)) & "</div>" & vbCrLf
		Next
	End If
	CheckBoxArray = str	
End Function

Function IsChecked(val, arrSel)
	'helper function for CheckboxArray()
	Dim i
	
	IsChecked = ""
	If IsArray(arrSel) Then
		For i = 0 To UBound(arrSel,2)
			If CStr(val & "") = CStr(arrSel(0,i) & "") Then 
				IsChecked = " checked=""checked"""
				Exit For
			End If
		Next
	End If
End Function

Function SelectOption(arr, sSelVal)
	Dim sOptions, i
	sOptions = ""
	
	If IsArray(arr) Then
		For i = 0 To UBound(arr, 2)
			sOptions = sOptions & "<option value=""" & arr(0, i) & """" & IsSelected(arr(0, i), sSelVal) & ">" & Server.HTMLEncode(arr(1, i)) & "</option>"
		Next
	End If
	
	SelectOption = sOptions
End Function

Function IsSelected(val, sSelected)
	'helper function for SelectOption()
	Dim i
	
	IsSelected = ""
	If IsArray(sSelected) Then
		'assume 2-dim array 
		For i = 0 To UBound(sSelected,2) 
			If CStr(val & "") = CStr(sSelected(0,i) & "") Then 
				IsSelected = " selected=""selected"""
				Exit For
			End If
		Next	
	Else
		If CStr(val & "") = CStr(sSelected & "") Then IsSelected = " selected=""selected"""
	End If
End Function

Function MultiSelectOption(arr, vList)
	'vList can be either comma-delim string, one-dim array, or two-dim array
	Dim i, str
	If IsArray(arr) Then
		For i = 0 To UBound(arr, 2)
			str = str & "<option value=""" & arr(0, i) & """" & IsMultiSelected(arr(0, i), vList) & ">" & Server.HTMLEncode(arr(1, i)) & "</option>"
		Next
	End If
	MultiSelectOption = str
End Function

Function IsMultiSelected(val, vList)
	'helper function for MultiSelectOption()
	Dim arr, i, str, x, iDimension
	str = ""
	
	'determine if 
	On Error Resume Next
	iDimension = 0
	For i = 1 To 3
		x = UBound(vList, i)
		If err <> 0 Then	
			iDimension = i - 1
			Exit For
		End If
	Next
	On Error GoTo 0
	
	Select Case iDimension
		Case 1	'one-dimensional
			For i = 0 To UBound(vList, 1)
				If Trim(CStr(val & "")) = Trim(CStr(vList(i) & "")) Then 
					str = " selected=""selected"""
					Exit For
				End If
			Next
		Case 2	'two-dimensional
			'assume value to match is in first dimension
			For i = 0 To UBound(vList, 2)
				If Trim(CStr(val & "")) = Trim(CStr(vList(0,i) & "")) Then 
					str = " selected=""selected"""
					Exit For
				End If
			Next
		Case Else
			'check to see if it can be converted
			arr = Split(vList, ",")
			If IsArray(arr) Then
				For i = 0 To UBound(arr)
					If Trim(CStr(val & "")) = Trim(CStr(arr(i) & "")) Then
						str = " selected=""selected"""
						Exit For
					End If
				Next
			End If
	End Select 
	
	IsMultiSelected = str
End Function
%>