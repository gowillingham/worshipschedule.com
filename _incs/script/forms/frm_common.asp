<script runat="server" language="vbscript" type="text/vbscript">

Const SHOW_REQUIRED_FORMAT = "SHOW_REQUIRED_FORMAT"			

'***************************************************************************
'COMMON FUNCTIONS:
'-----------------------------------------
'	- Function FormatRequiredElement(element, label, settingList)
'	- Function IsRequiredElement(element, settingList)
'	- Function RequiredElementToString(isRequired, str)
'	- Function WasProvided(sValue)
'***************************************************************************

Function FormatRequiredElement(element, label, settingList)
	Dim str
	
	str = label

	' - use a default value to get this formatting to display in form instructions, etc
	If element = SHOW_REQUIRED_FORMAT Then
		str = RequiredElementToString(True, str)
	End If
	
	If IsRequiredElement(element, settingList) Then
		str = RequiredElementToString(True, str)
	End If
	
	FormatRequiredElement = str
End Function

Function IsRequiredElement(element, settingList)
	' new version of this to replace IsRequiredField
	Dim i
	
	IsRequiredElement = False
	If IsArray(settingList) Then
		For i = 0 To UBound(settingList,2)
			If element = settingList(0,i) Then
				If settingList(1,i) = 1 Then
					IsRequiredElement = True
				End If
			End If
		Next
	End If
End Function

Function RequiredElementToString(isRequired, str)
	If isRequired Then
		str = str & "<span class=""required"" style=""color:red;"">*</span>"
	End If
	
	RequiredElementToString = str
End Function

Function WasProvided(sValue)
	'generate placeholder for missing data
	'	- be sure to Server.HTMLEncode the output of this to 
	'		handle the angle brackets.
	WasProvided = sValue
	If sValue = "" Then WasProvided = "<Not Provided>"
End Function
</script>
