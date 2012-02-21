<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>

<%

Dim str
Dim i
Dim encoded
Dim decoded
Dim length
Dim startIndex		: startIndex = 0
Dim endIndex		: endIndex = 100000
Dim startTime		: startTime = Now()
Dim errors			: errors = 0

	str = str & "<table>"
	For i = startIndex To endIndex

		encoded = Base64Encode(i)
		decoded = Base64Decode(encoded)
		length = len(decoded) - len(i)
		
		If (CStr(i) <> CStr(decoded)) or (length <> 0) then
			str = str & "<tr>"
			str = str & "<td>'" & i & "'</td>"
			str = str & "<td>'" & encoded & "'</td>"
			str = str & "<td>'" & decoded & "'</td>"
			str = str & "<td>" & length & "'</td>"
			str = str & "</tr>"
		End If
	Next

	str = str & "</table>"
	str = str & "<p>" & (endIndex - startIndex) & " integers tested .."
	str = str & "<br />" & errors & " errors detected .."
	str = str & "<br />calc time " & DateDiff("s", startTime, Now()) & " seconds for i = " & startIndex & " to " & endIndex & " .."


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>sandbox two</title>
</head>
	<body>
		<h3>Test Base64 Encode</h3>
		<%=str %>
	</body>
</html>

<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/rc4encrypt_cls.asp"-->





