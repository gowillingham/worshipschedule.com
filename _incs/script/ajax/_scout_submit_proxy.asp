<%
Option Explicit

Call Main()

Sub Main()
	Dim xml								: Set xml = server.Createobject("MSXML2.ServerXMLHTTP")

	Dim scoutUserName					: scoutUserName = Request.Form("ScoutUserName")
	Dim scoutProject					: scoutProject = Request.Form("scoutProject")
	Dim scoutArea						: scoutArea = Request.Form("scoutArea")
	Dim description						: description = server.URLEncode(Request.Form("description"))
	Dim forceNewBug						: forceNewBug = Request.Form("forceNewBug")
	Dim scoutDefaultMessage				: scoutDefaultMessage = Server.URLEncode(Request.Form("scoutDefaultMessage"))
	Dim friendlyResponse				: friendlyResponse = Request.Form("friendlyResponse")
	Dim email							: email = Request.Form("email")
	Dim extra							: extra = Server.URLEncode(Request.Form("extra"))
	
	Dim data

	data = data & "ScoutUserName=" & scoutUserName
	data = data & "&scoutProject=" & scoutProject
	data = data & "&scoutArea=" & scoutArea
	data = data & "&description=" & description
	data = data & "&forceNewBug=" & forceNewBug
	data = data & "&scoutDefaultMessage=" & scoutDefaultMessage
	data = data & "&friendlyResponse=" & friendlyResponse
	data = data & "&email=" & email
	data = data & "&extra=" & extra

	Call xml.Open("POST", Application.Value("FOGBUGZ_SCOUT_SUBMIT_URL"), False)	
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	Call xml.Send(data)
	
	response.Write xml.ResponseText
End Sub
%>
