<%
Option Explicit

Call Main()

Sub Main
	Dim str, i
	
	Dim sess				: Set sess = New cSession
	sess.SessionID = Request.QueryString("sid")
	Call sess.Load()
	
	Dim client				: Set client = New cClient
	client.ClientID = sess.ClientId
	Call client.Load()
	
	' 0-MemberID 1-NameLast 2-NameFirst 3-NameLogin 4-PWord 5-Email 6-DOB 7-Gender
	
	Dim list				: list = client.MemberList("", "")
	For i = 0 To UBound(list,2)
		str = str & list(5,i) & ","
	Next
	str = Left(str, Len(str) - 1)

	Response.Write str
End Sub
%>

<!--#INCLUDE VIRTUAL="/_incs/class/session_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/client_cls.asp"-->

