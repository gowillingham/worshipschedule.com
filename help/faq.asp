<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Const DISPLAY_FAQ_TOPIC	= "1"
Const DISPLAY_FAQ_CATEGORY = "2"
Dim m_bodyText
Dim m_breadcrumb

Sub OnPageLoad(ByRef page)
	page.Action = Decrypt(Request.QueryString("act"))
	page.MessageID = Request.QueryString("msgid")
	
	Set page.Faq = New cFaq
	page.Faq.FaqID = Request.Querystring("fqid")
	If Len(page.Faq.FaqID) > 0 Then page.Faq.Load()
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/outside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=Application.Value("APPLICATION_NAME") & " - Simple Web Scheduling for Worship Teams" %></title>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><%=m_breadcrumb %></h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<%=m_bodyText %>
			</div>
		</div>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_footer.asp"-->
	</body>
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)
	Call SetBreadcrumb(page)

	page.MessageID = ""
	
	Select Case page.Action
		Case DISPLAY_FAQ_TOPIC
			str = str & FaqToString(page)
		Case DISPLAY_FAQ_CATEGORY
			str = str & FaqCategoryListToString(page)
		Case Else
			str = str & NoFaqMessageToString()
			
			' todo: 
			' -----------------------------------
			' str = str & FaqListToString(page)
	End Select
	
	m_bodyText = str
	Set page = Nothing
End Sub

Function NoFaqMessageToString()
	Dim str
	
	str = str & "<h3>" & Application.Value("APPLICATION_NAME") & " FAQ currently not available</h3>"
	str = str & "<p>Sorry, but the <strong>Frequently asked questions</strong> portion of online help for " & Application.Value("APPLICATION_NAME") & " is being upgraded to version 2.0 and is not available at this time. "
	str = str & "We apologize for the inconvenience. </p>"
	str = str & "<p><strong>We still want to help! </strong>"
	str = str & "Please send any questions you have about " & Application.Value("APPLICATION_NAME") & " and/or your account to "
	str = str & "<a href=""/support.asp"">support</a> and we will get back to you with an answer as quickly as possible "
	str = str & "(usually in less than 24 hours). </p>"

	NoFaqMessageToString = str	
End Function

Function FaqListToString(page)
	Dim str, i, j
	
	Dim pg						: Set pg = page.Clone()
	Dim categoryList			: categoryList = page.Faq.CategoryList()
	Dim faqList					: faqList = page.Faq.List()
	
	Dim href					: href = ""
	Dim isColumn1
	Dim count								
	Dim faqTotalCount
		
	str = str & HelpMessageToString()
	str = str & "<div id=""help"">"
	str = str & "<h1>General Use</h1>"
	str = str & FaqListForSuperCategoryToString(page, "General Use", faqList, categoryList)
	
	str = str & "<h1>Administration</h1>"	
	str = str & FaqListForSuperCategoryToString(page, "Administration", faqList, categoryList)
	
	str = str & "</div>"
	
	FaqListToString = str
End Function

Function FaqListForSuperCategoryToString(page, superCategory, faqList, categoryList)
	Dim str, i, j
	Dim pg				: Set pg = page.Clone()
	Dim isColumn1		: isColumn1 = True
	Dim count			: count = 0								
	Dim faqTotalCount	: faqTotalCount = GetFaqCountForSuperCategory(superCategory, faqList)
	Dim href
	
	pg.Action = DISPLAY_FAQ_TOPIC
	
	' faqList
	' 0-faqID 1-Title 2-Text 3-Priorit 4-DateCreated 5-CategoryID 6-Category
	' 7-CategoryDescription 8-CategoryPriority 9-SuperCategory

	' categoryList
	' 0-FaqCategoryID 1-SuperCategory 2-Category 3-Description 4-Priority

	str = str & "<div class=""column1"">"
	For i = 0 To UBound(categoryList,2)
	
		' set super category
		If categoryList(1,i) = superCategory Then
			
			' category header
			str = str & "<h3>" & categoryList(2,i) & "</h3>"
			str = str & "<ol>"
			For j = 0 To UBound(faqList,2)
				If categoryList(0,i) = faqList(5,j) Then
					count = count + 1
					pg.Faq.FaqID = faqList(0,j)
					href = Request.ServerVariables("URL") & pg.UrlParamsToString(True)
					str = str & "<li><a href=""" & href & """>" & faqList(1,j) & "</a></li>"
				End If
			Next
			str = str & "</ol>"
			
			' here's where I change columns ..
			If isColumn1 Then
				If (2 * count) => faqTotalCount Then
					str = str & "</div><div class=""column2"">"
					isColumn1 = False
				End If
			End If
		End If
	Next 
	str = str & "</div>"
	
	FaqListForSuperCategoryToString = str
End Function

Function GetFaqCountForSuperCategory(superCategory, faqList)
	Dim i
	Dim count			: count = 0
	
	For i = 0 To UBound(faqList,2)
		If superCategory = faqList(9,i) Then
			count = count + 1
		End If
	Next
	
	GetFaqCountForSuperCategory = count	
End Function

Function FaqCategoryListToString(page)
	Dim str, i
	Dim pg				: Set pg = page.Clone()
	Dim list			: list = page.Faq.ListByCategoryID()
	Dim href
	
	pg.Action = DISPLAY_FAQ_TOPIC
	
	
	str = str & HelpMessageToString()
	str = str & "<div id=""help"" style=""margin-bottom:15px;"">"
	str = str & "<ol>"
	For i = 0 To UBound(list,2)
		pg.Faq.FaqID = list(0,i)
		href = Request.ServerVariables("URL") & pg.UrlParamsToString(True)
		str = str & "<li><a href=""" & href & """>" & list(1,i) & "</a></li>"
	Next
	str = str & "</ol>"
	str = str & "</div>"
	
	FaqCategoryListToString = str
End Function

Function HelpNavToString(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href

	Dim prevID, prevTitle, nextID, nextTitle
	
	Call GetPrevNextFaqID(page, prevID, prevTitle, nextID, nextTitle)
	
	str = str & "<div id=""helpnav"">"
	If Len(prevID) > 0 Then
		pg.Faq.FaqID = prevID
		href = Request.ServerVariables("URL") & pg.UrlParamsToString(True)
		str = str & "<div style=""float:left;"">"
		str = str & "<a href=""" & href & """ title=""" & prevTitle & """>&laquo; Prev</a></div>"
	End If
	If Len(nextID) > 0 Then
		pg.Faq.FaqID = nextID
		href = Request.ServerVariables("URL") & pg.UrlParamsToString(True)
		str = str & "<div style=""float:right;"">"
		str = str & "<a href=""" & href & """ title=""" & nextTitle & """>Next &raquo;</a></div>"
	End If 
	str = str & "</div>"
	
	HelpNavToString = str
End Function

Sub GetPrevNextFaqID(page, prevID, prevTitle, nextID, nextTitle)
	Dim i
	Dim list		: list = page.Faq.ListByCategoryID()
	
	For i = 0 To UBound(list,2)
		If CStr(list(0,i)) = CStr(page.Faq.FaqID) Then
			' get prev id ..
			If (i - 1) => 0 Then
				prevID = list(0, i - 1)
				prevTitle = list(1, i - 1)
			End If
			' get next id
			If (i + 1) <= UBound(list,2) Then
				nextID = list(0, i + 1)
				nextTitle = list(1, i + 1)
			End If
		End If
	Next
End Sub

Function FaqToString(page)
	Dim str

	str = str & page.Faq.Text
	str = str & HelpNavToString(page)
	
	FaqToString = str
End Function

Function HelpMessageToString()
	Dim str
	
	str = str & "<p style=""font-size:1.1em;"">"
	str = str & "Please check the list of frequently asked questions below to help address your problem. "
	str = str & "If your question is not answered, you might also like to try contacting "
	str = str & "" & Application.Value("APPLICATION_NAME") & " <a href=""/support.asp"">support</a></p>"
	
	HelpMessageToString = str
End Function

Function SetBreadcrumb(page)
	Dim str
	Dim pg			: Set pg = page.Clone()
	Dim href
	
	str = str & "<a href=""/help/help.asp"">Help</a> / "
	Select Case page.Action
		Case DISPLAY_FAQ_TOPIC
		
			pg.Action = DISPLAY_FAQ_CATEGORY
			href = Request.ServerVariables("URL") & pg.UrlParamsToString(True)
			
			str = str & "<a href=""/help/faq.asp"">FAQ</a> / <a href=""" & href & """>" & page.Faq.Category & "</a> / "
			str = str & page.Faq.Title
			
		Case DISPLAY_FAQ_CATEGORY
			str = str & "<a href=""/help/faq.asp"">FAQ</a> / "
			str = str & page.Faq.Category
			
		Case Else
			str = str & "FAQ"
			
	End Select
	
	m_breadcrumb = str
End Function

%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/security_common.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/faq_cls.asp"-->
<%
Class cPage
	Public Action
	Public MessageID
	
	' object
	Public Faq
	
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(Action) > 0 Then str = str & "act=" & Encrypt(Action) & amp
		If Len(Faq.FaqID) > 0 Then str = str & "fqid=" & Faq.FaqID & amp
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		
		If Len(str) > 0 Then 
			str = Left(str, Len(str) - Len(amp))
		Else
			' qstring needs at least one param in case more params are appended ..
			str = str & "noparm=true"
		End If
		str = "?" & str
		
		UrlParamsToString = str
	End Function
	
	Public Function Clone()
		Dim c
		Set c = New cPage
		
		c.Action = Action
		c.MessageID = MessageID
		Set c.Faq = Faq
		
		Set Clone = c
	End Function
End Class
%>

