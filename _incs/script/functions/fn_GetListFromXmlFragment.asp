<script language="vbscript" type="text/vbscript" runat="server">

' todo: replace deprecated function in places where it is used ..

' legacy version to return list that should be replaced ..
Function GetListFromXMLFragment(sLeadIn, sLeadOut, sDelim, fragment)
	Dim str, delimiter, arr, i
	
	If Len(fragment & "") = 0 Then
		GetListFromXMLFragment = ""
		Exit Function
	End If
	
	delimiter = sLeadOut & sLeadIn
	arr = Split(fragment, delimiter)
	If IsArray(arr) Then
		For i = 0 To UBound(arr)
			arr(i) = Replace(arr(i), sLeadin, "")
			arr(i) = Replace(arr(i), sLeadOut, "")
			If Len(arr(i)) > 0 Then
				str = str & arr(i) & sDelim
			End If
		Next
		If Len(str) > 0 Then 
			'remove trailing delim
			str = Left(str, Len(str) - Len(sDelim))
		End If
	End If
	
	GetListFromXMLFragment = str
End Function

' new version to return list ..
Function XmlFragmentToList(xml, listDelim, ByRef xmlDoc) 'returns str
	Dim str, node, needsCleanup
	
	' check for no xml string
	If Len(xml & "") = 0 Then Exit Function
	
	needsCleanup = False
	If Not IsObject(xmlDoc) Then
		Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
		needsCleanup = True
	End If
	
	xmlDoc.LoadXml(xml)
	xmlDoc.Async = False
	
	For Each node In xmlDoc.DocumentElement.ChildNodes
		str = str & node.Text & listDelim
	Next
	If Len(str) > 0 Then
		' remove the last delim
		str = Left(str, Len(str) - Len(listDelim))
	End If
	
	If needsCleanup Then
		Set xmlDoc = Nothing
	End If
	
	XmlFragmentToList = str
End Function


'******************************************************************************
'	GetListFromXMLFragment(): 
'	----------------------------------------------------------------------------
'	Accept xml fragment as returned from SQL sproc that returns FOR XML AUTO. 
'	Removes XML and generates list using custom seperator. 

'	XmlFragmentToList(xml, listDelim, xmlDoc): 
'	----------------------------------------------------------------------------
'	Accept valid xml fragment from and return a list. Only one element is allowed 
'	per row. Accept xmlDoc as an optional object to preserve resources. Example 
'	fragment below ..
'
'	<?xml version="1.0"?>
'		<root>
'			<row>
'				<skillname>Bass</skillname>
'			</row>
'			<row>
'				<skillname>Violin</skillname>
'			</row>
'			<row>
'				<skillname>Flute</skillname>
'			</row>
'		</root>
'	
'******************************************************************************
</script>