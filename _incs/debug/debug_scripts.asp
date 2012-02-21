<%
Function EnumQueryStringData()
	Dim var
	With Response
		.Write "<br>BEGIN ENUM QUERYSTRING DATA ********************************<br>"
		For Each var in Request.QueryString
			.Write var & "--" & Request.QueryString(var) & "<br>"
		Next
		.Write "<br>END DEBUG **************************************************<br>"			
	End With
End Function

Function EnumFormData()
	Dim var
	With Response
		.Write "<br>BEGIN ENUM FORM DATA ***************************************<br>"
		For Each var In Request.form
			.Write var & "--" & Request.Form(var) & "<br>"
		Next
		.Write "<br>END DEBUG **************************************************<br>"			
	End With
End Function

Function EnumServerVariables()
	Dim var
	With Response
		.Write "<div style=""font-size:10px;"">"
		.Write "<br>DEBUG SERVER VARIABLES *************************************<br>"
		For Each var In Request.ServerVariables
			.Write var & "--" & Request.ServerVariables(var) & "<br>"
		Next
		.Write "<br>END DEBUG **************************************************<br>"
		.Write "</div>"
	End With
End Function


'*********************************************************************************


    '**************************************
    ' for :Cookie Debugger
    '**************************************
    'Copyright (c) 2001, Lewis Edward Moten III. All rights reserved.
    '**************************************
    ' Name: Cookie Debugger
    ' Description:Creates a list of all cook
    '     ies and there crumbs along with the valu
    '     es assigned to each one.
    ' By: Lewis Moten
    ' Inputs:None
    ' Returns:Returns an orderd list of name
    '     s and values of cookies and crumbs.
    'Assumes:None
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    'Please see http://www.1JavaStreet.com/x
    '     q/ASP/txtCodeId.6724/lngWId.4/qx/vb/scri
    '     pts/ShowCode.htm
    'for details.
    '**************************************
    
    Function CookieData(bDebug)
    	Dim llngMaxCookieIndex
    	Dim llngCookieIndex
    	Dim llngMaxCrumbIndex
    	Dim llngCrumbIndex
    	Dim lstrDebug
    	
    	If Not bDebug Then Exit Function
    	
    	' Count Cookies
    	llngMaxCookieIndex = Request.Cookies.Count
    	
    	' Let user know if cookies do not exist
    	If llngMaxCookieIndex = 0 Then
    		CookieData = "<hr>cookie data is empty.<hr>"
    		Exit Function
    	End If
    	
    	' Begin building a list of all cookies
    	lstrDebug = "<hr>Cookie Data:<OL style=""font-size:0.7em;"">"
    	
    	' Loop through each cookie
    	For llngCookieIndex = 1 To llngMaxCookieIndex
    		lstrDebug = lstrDebug & "<LI>" & Server.HTMLEncode(Request.Cookies.Key(llngCookieIndex))
    		
    		' Count the crumbs
    		llngMaxCrumbIndex = Request.Cookies(llngCookieIndex).Count
    		
    		' If the cookie doesn't have crumbs ...
    		If llngMaxCrumbIndex = 0 Then
    			lstrDebug = lstrDebug & " = "
    			lstrDebug = lstrDebug & Server.HTMLEncode(Request.Cookies.Item(llngCookieIndex))
    		' Else loop through each crumb
    		Else
    			lstrDebug = lstrDebug & "<OL>"
    			For llngCrumbIndex = 1 to llngMaxCrumbIndex
    				lstrDebug = lstrDebug & "<LI>"
    				lstrDebug = lstrDebug & Server.HTMLEncode(Request.Cookies(llngCookieIndex).Key(llngCrumbIndex))
    				lstrDebug = lstrDebug & " = "
    				lstrDebug = lstrDebug & Server.HTMLEncode(Request.Cookies(llngCookieIndex)(llngCrumbIndex))
    				lstrDebug = lstrDebug & "</LI>"
    			Next
    			lstrDebug = lstrDebug & "</OL>"
    		End If
    		lstrDebug = lstrDebug & "</LI>"
    	Next
    	lstrDebug = lstrDebug & "</OL><hr>"
    	' Return the data
    	CookieData = lstrDebug
    	
    End Function

%>
