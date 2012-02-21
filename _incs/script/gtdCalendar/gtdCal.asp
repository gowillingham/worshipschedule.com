<script runat="server" type="text/vbscript" language="vbscript">

Class gtdCalendar

'	---------------------------------------------------------------------	
	Private m_textItems	'as scripting.dictionary
	Private m_customNavUrlParmList ' as scripting.dictionary
	Private m_customDateUrlParmList ' as scripting.dictionary
	Private m_datePointer
	Private m_customNavUrlParms
	Private m_customDateUrlParms
	Private m_displayHeader
	Private m_displayTopNav
	Private m_displayDatePicker
	Private m_displayWeekdayRow 
	Private m_displayWeekdayRowStyle
	
	Private m_displayDayNumbersAsLinks
	
	Private m_textItemDelimiter
	Private m_pathToScript 'virtual path to script directory
	
'	---------------------------------------------------------------------	
	Public Function ToString()
		' return the caldendar grid as a string
		Dim str, i, dayIDX, startDate, cellStyle
		
		' save first day of month ..
		startDate = m_datePointer
		
		str = str & "<div class=""gtdCalContainer"">"
		str = str & HeaderToString()
		str = str & TopNavToString()
		str = str & "<table class=""gtdCalGrid"">"
		
		' weekday labels
		str = str & WeekdayRowToString()
		
		' track day of week ..
		dayIDX = 1

		str = str & "<tr>"
		
		' empty days before first of month
		For i = 1 To Weekday(m_datePointer) - 1
			str = str & "<td" & GetGridCellClass(dayIDX, Month(startDate)) & ">&nbsp;</td>"
			dayIDX = dayIDX + 1
		Next
		
		' loop through each day until reach next month
		Do While Month(m_datePointer) = Month(startDate)
		
			str = str & "<td" & GetGridCellClass(dayIDX, Month(startDate)) & ">"
			
			' day number
			str = str & "<div class=""gtdDayNumberContainer"">"
			If m_displayDayNumbersAsLinks Then
				str = str & "<a href=""" & Request.ServerVariables("URL") & GetDateParams() & """ title=""View This Date"">" & Day(m_datePointer) & "</a>"
			Else
				str = str & Day(m_datePointer)
			End If
			str = str & "</div>"
			
			' event item data
			str = str & ItemsToString(m_datePointer)
			str = str & "</td>"
			
			' start new week
			If Weekday(m_datePointer) = 7 Then
				str = str & "</tr><tr>"
				dayIDX = 0
			End If
			
			' move pointer and weekday index to next day
			m_datePointer = DateAdd("d", 1, m_datePointer)
			dayIDX = dayIDX + 1
		Loop
		
		' finish with any empty day cells at end of month
		If Weekday(m_datePointer) <> 1 Then
			For i = Weekday(m_datePointer) To 7
				str = str & "<td" & GetGridCellClass(dayIDX, Month(startDate)) & ">&nbsp;</td>"
				dayIDX = dayIDX + 1
			Next
		End If
		 
		' set datePointer back to start date
		m_datePointer = startDate
		
		str = str & "</tr>"
		str = str & "</table>"
		str = str & "</div>"
		
		ToString = str
	End Function

'	---------------------------------------------------------------------	
	Public Function Items(eventDate) 
		' returns array of event items
		Dim key
		key = Month(eventDate) & "/" & Day(eventDate) & "/" & Year(eventDate)
			
		Items = m_textItems(key)
	End Function
	
'	---------------------------------------------------------------------	
	Public Sub AddItem(eventDate, itemText, itemStyle)
		Dim key, arr, val
		
		key = Month(eventDate) & "/" & Day(eventDate) & "/" & Year(eventDate)
		
		' append itemStyle to itemText with private delimiter ..
		val = itemText & m_textItemDelimiter & itemStyle	
		
		' if date already has item, then get it out and append to end of item array
		If m_textItems.Exists(key) Then
			arr = m_textItems(key)
			ReDim Preserve arr(UBound(arr) + 1)
			arr(UBound(arr)) = val
			m_textItems(key) = arr
		
		' otherwise add new item ..
		Else
			ReDim arr(0)
			arr(0) = val
			Call m_textItems.Add(key, arr)
		End If	
	End Sub
	
'	---------------------------------------------------------------------	
	Public Sub SetDate(ByRef thisYear, ByRef thisMonth)
		' sets month/year for calendar to display, defaults to Now()
		If Len(thisYear) = 0 Then thisYear = Year(Now())
		If Len(thisMonth) = 0 Then thisMonth = Month(Now())
		
		m_datePointer = DateSerial(thisYear, thisMonth, 1)
	End Sub
	
'	---------------------------------------------------------------------
	Public Sub AddNavUrlParams(parmName, parmValue)
		' accepts value/pair and appends to urls exposed by the grid navigation
		Dim str
		
		If m_customNavUrlParmList.Exists(parmName) Then
			m_customNavUrlParmList(parmName) = parmValue
		Else
			Call m_customNavUrlParmList.Add(parmName, parmValue)
		End If
	End Sub	
	
'	---------------------------------------------------------------------
	Public Sub AddDateUrlParams(parmName, parmValue)
		' accepts value/pair and appends to urls exposed by the grid date link
		Dim str
		
		If m_customDateUrlParmList.Exists(parmName) Then
			m_customDateUrlParmList(parmName) = parmValue
		Else
			Call m_customDateUrlParmList.Add(parmName, parmValue)
		End If
	End Sub	
	
'	---------------------------------------------------------------------	
   Private Function NavUrlParamsToString(stringType)
      Dim str, keys, i
      
      If m_customNavUrlParmList.Count > 0 Then
      
         keys = m_customNavUrlParmList.Keys
         Select Case stringType
            Case "qstring"
               for i = 0 To UBound(keys)
                  If Len(m_customNavUrlParmList(keys(i))) > 0 Then
                     str = str & keys(i) & "=" & m_customNavUrlParmList(keys(i)) & "&amp;"
                  End If
               Next
            If Len(str) > 0 Then str = Left(str, Len(str) - 5)
            Case "form_elements"
               For i = 0 To UBound(keys)
                  str = str & "<input type=""hidden"" name=""" & keys(i) & """ value=""" & m_customNavUrlParmList(keys(i)) & """ />"
               Next
         End Select
      End If
      
      NavUrlParamsToString = str
   End Function
   
'	---------------------------------------------------------------------	
   Private Function DateUrlParamsToString()
      Dim str, keys, i
      
      If m_customDateUrlParmList.Count > 0 Then
      
         keys = m_customDateUrlParmList.Keys
         for i = 0 To UBound(keys)
            If Len(m_customDateUrlParmList(keys(i))) > 0 Then
               str = str & keys(i) & "=" & m_customDateUrlParmList(keys(i)) & "&amp;"
            End If
         Next
         If Len(str) > 0 Then str = Left(str, Len(str) - 5)
      End If
      
      DateUrlParamsToString = str
   End Function
   
'	---------------------------------------------------------------------	
	Private Function ItemsToString(eventDate)
		Dim key, arr, str, item, itemText, itemStyle, i
		
		key = Month(eventDate) & "/" & Day(eventDate) & "/" & Year(eventDate)	
		
		If Not m_textItems.Exists(key) Then Exit Function
		
		' iterate through item array and wrap each item in div
		arr = m_textItems(key)
		For i = 0 To UBound(arr)
		
			' item is stored as deliminated string of text, style
			item = Split(arr(i), m_textItemDelimiter)
			itemText = item(0)
			itemStyle = item(1)
			If Len(itemStyle) > 0 Then
				itemStyle = " style=""" & itemStyle & """"
			End If
		
			str = str & "<div class=""gtdCalCellItem""" & itemStyle & ">" & itemText & "</div>"
		Next
		
		ItemsToString = str
	End Function
	
'	---------------------------------------------------------------------	
	Private Function MonthToText(val, style)
		Dim monthList
		
		' use placeholder in first element so can return month for number without doing math ..
		If style = "SHORT_NAME" Then
			monthList = Split("placeholder,Jan,Feb,Mar,Apr,May,June,July,Aug,Sept,Oct,Nov,Dec", ",")
		Else
			monthList = Split("placeholder,January,February,March,April,May,June,July,August,September,October,November,December", ",")
		End If
		
		MonthToText = monthList(CInt(val))
	End Function
	
'	---------------------------------------------------------------------
	Private Function WeekdayToText(val, style)
		Dim weekdayList
	
		' use placeholder for first element so that can return weekday numbers without doing math
		If style = 1 Then
			weekdayList = Split("placeholder,S,M,T,W,T,F,S", ",")
		ElseIf style = 2 Then
			weekdayList = Split("placeholder,Su,Mo,Tu,We,Th,Fr,Sa", ",")
		ElseIf style = 3 Then 
			weekdayList = Split("placeholder,Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")
		Else
			weekdayList = Split("placeholder,Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", ",")
		End If
		
		WeekdayToText = weekdayList(val)
	End Function
	
'	---------------------------------------------------------------------	
	Private Function WeekdayRowToString()
		Dim str, i
		
		If Not m_displayWeekdayRow Then Exit Function
		
		str = str & "<tr class=""gtdCalWeekdayHeader"">"
		For i = 1 To 7 
			str = str & "<td>" & WeekdayToText(i, m_displayWeekdayRowStyle) & "</td>"
		Next
		str = str & "</tr>"
		
		WeekdayRowToString = str
	End Function

'	---------------------------------------------------------------------	
	Private Function TopNavToString()
		Dim str
		If Not m_displayTopNav Then Exit Function
		
		str = str & "<table class=""gtdCalTopNav"">"
		str = str & "<tr>"
		
		str = str & "<td class=""gtdCalTopNavPrevYear"">"
		str = str & "<a href=""" & Request.ServerVariables("URL") & GetNavParams("lastyear") & """ title=""Previous Year"">"
		str = str & "<img src=""" & m_pathToScript & "lastyear.gif"" alt=""Previous Year"" />"
		str = str & "</a></td>"
		
		str = str & "<td class=""gtdCalTopNavPrevMonth"">"
		str = str & "<a href=""" & Request.ServerVariables("URL") & GetNavParams("lastmonth") & """ title=""Previous Month"">"
		str = str & "<img src=""" & m_pathToScript & "lastmonth.gif"" alt=""Previous Month"" />"
		str = str & "</a></td>"
		
		str = str & "<td  class=""gtdCalTopNavToday"">"
		str = str & "<a href=""" & Request.ServerVariables("URL") & GetNavParams("today") & """ title=""Today"">"
'		str = str & "<img src=""" & m_pathToScript & "today.gif"" alt=""Today"" />"
'		str = str & "</a></td>"
		str = str & "Today&nbsp;!</a></td>"
		
		str = str & "<td class=""gtdCalTopNavMonthYearPicker"">"
		str = str & MonthYearPickerToString()
		str = str & "</td>"
				
		str = str & "<td  class=""gtdCalTopNavNextMonth"">"
		str = str & "<a href=""" & Request.ServerVariables("URL") & GetNavParams("nextmonth") & """ title=""Next Month"">"
		str = str & "<img src=""" & m_pathToScript & "nextmonth.gif"" alt=""Next Month"" />"
		str = str & "</a></td>"
		
		str = str & "<td  class=""gtdCalTopNavNextYear"">"
		str = str & "<a href=""" & Request.ServerVariables("URL") & GetNavParams("nextyear") & """ title=""Next Year"">"
		str = str & "<img src=""" & m_pathToScript & "nextyear.gif"" alt=""Next Year"" />"
		str = str & "</a></td>"
		
		str = str & "</tr>"
		str = str & "</table>"
		
		TopNavToString = str
	End Function
	
'	---------------------------------------------------------------------	
	Private Function MonthYearPickerToString()
		' returns month/year dropdown pickers
		Dim str, i, keys, url
		If Not m_displayDatePicker Then Exit Function

		str = str & "<form id=""newDate"" method=""get"" action=""" & url & """>"
		
		' month picker
		str = str & "<select name=""m"" onchange=""submit();"">"
		For i = 1 To 12
			str = str & "<option value=""" & i & """" & IsSelected(i, Month(m_datePointer)) & ">" & MonthToText(i, "SHORT_NAME") & "</option>"
		Next
		str = str & "</select>"
		
		' year picker
		str = str & "<select name=""y"" onchange=""submit();"">"
		For i = Year(DateAdd("yyyy", -5, Now())) To Year(DateAdd("yyyy", 5, Now()))
			str = str & "<option value=""" & i & """" & IsSelected(i, Year(m_datePointer)) & ">" & i & "</option>"
		Next
		str = str & "</select>"
		
		' custom url params
		str = str & NavUrlParamsToString("form_elements")
		
		str = str & "</form>"	
		
		MonthYearPickerToString = str
	End Function
	
'	---------------------------------------------------------------------	
	Private Function HeaderToString()
		Dim str 
		If Not m_displayHeader Then Exit Function

		str = str & "<div class=""gtdCalHeader"">" & MonthToText(Month(m_datePointer), "") & " " & Year(m_datePointer) & "</div>"
		
		HeaderToString = str
	End Function
	
'	---------------------------------------------------------------------	
	Private Function GetGridCellClass(dayIDX, thisMonth)
		' generate css class for cal.grid cells
		Dim str
		
		If dayIDX = 1 Then
			str = str & "gtdCalGridSun"
		ElseIf dayIDX = 7 Then
			str = str & "gtdCalGridSat"
		End If
		
		' format days before the first
		If dayIDX < Weekday(m_datePointer) Then
			If Len(str) > 0 Then str = str & " "
			str = str & "gtdCalGridNotPartOfMonth"
		End If
		' format days after the last of the month
		If Month(m_datePointer) > thisMonth Then
			If Len(str) > 0 Then str = str & " "
			str = str & "gtdCalGridNotPartOfMonth"
		End If
		
		' if this is today ..
		If FormatDateTime(m_datePointer, vbShortDate) = FormatDateTime(Now(), vbShortDate) Then
			' don't format non-month days before the first
			If dayIDX = Weekday(m_datePointer) Then
				str = str & " gtdCalGridToday"
			End If
		End If
		
		If Len(str) > 0 Then 
			str = " class=""" & str & """"
		End If
		
		GetGridCellClass = str
	End Function
	
'	---------------------------------------------------------------------	
	Private Function GetNavParams(direction)
		' return month/year url params for cal.grid nav date links in cells
		Dim newMonth, newYear, str
		
		If direction = "lastyear" Then
			newMonth = Month(m_datePointer)
			newYear = Year(m_datePointer) - 1
		ElseIf direction = "lastmonth" Then
			newMonth = Month(m_datePointer) - 1
			newYear = Year(m_datePointer)
			If newMonth < 1 Then
				newMonth = 12
				newYear = newYear - 1
			End If
		ElseIf direction = "nextmonth" Then
			newMonth = Month(m_datePointer) + 1
			newYear = Year(m_datePointer)
			If newMonth > 12 Then
				newMonth = 1
				newYear = newYear + 1
			End If
		ElseIf direction = "nextyear" Then
			newMonth = Month(m_datePointer)
			newYear = Year(m_datePointer) + 1
		ElseIf direction = "today" Then
			newMonth = Month(Now())
			newYear = Year(Now())
		ElseIf direction = "current" Then
			newMonth = Month(m_datePointer)
			newYear = Year(m_datePointer)
		End If
		
		' build the qstring and append any custom params to the end
		str = "?" & NavUrlParamsToString("qstring")
		If Len(str) > 1 Then 
			str = str & "&amp;"
		End If
		str = str & "m=" & newMonth & "&amp;y=" & newYear
		
		GetNavParams = str
	End Function
	
'	---------------------------------------------------------------------	
   Private Function GetDateParams()
      Dim str

      str = str & "?" & DateUrlParamsToString()
	  if Len(str) > 1 Then
		 str = str & "&amp;"
      end if
      str = str & "m=" & Month(m_datePointer) & "&amp;y=" & Year(m_datePointer) & "&amp;d=" & Day(m_DatePointer)

	  GetDateParams = str
   End Function
'	---------------------------------------------------------------------	
	Private Sub Class_Initialize()
		Set m_textItems = Server.CreateObject("Scripting.Dictionary")
		Set m_customNavUrlParmList = Server.CreateObject("Scripting.Dictionary")
		Set m_customDateUrlParmList = Server.CreateObject("Scripting.Dictionary")
		
		m_datePointer = DateSerial(Year(Now()), Month(Now()), 1)
		m_textItemDelimiter = "|%@!"
		m_displayHeader = True
		m_displayTopNav = True
		m_displayDatePicker = True
		m_displayWeekdayRow = True
		m_displayDayNumbersAsLinks = True
		m_displayWeekdayRowStyle = 0
		
		' set this for your own web directory
		m_pathToScript = "/_incs/script/gtdCalendar/"
	End Sub
	
'	---------------------------------------------------------------------	
	Private Sub Class_Terminate()
		Set m_textItems = Nothing
		Set m_customNavUrlParmList = Nothing
		Set m_customDateUrlParmList = Nothing
	End Sub
	
'	---------------------------------------------------------------------	
	Public Property Let DisplayTopNav(val) 'val = bool
		m_displayTopNav = val
	End Property
	
'	---------------------------------------------------------------------	
	Public Property Get DisplayTopNav()
		DisplayTopNav = m_displayTopNav
	End Property
	
'	---------------------------------------------------------------------	
	Public Property Let DisplayDatePicker(val) 'val = bool
		m_displayDatePicker = val
	End Property
	
'	---------------------------------------------------------------------	
	Public Property Get DisplayDatePicker()
		DisplayDatePicker = m_displayDatePicker
	End Property
	
'	---------------------------------------------------------------------	
	Public Property Let DisplayHeader(val) 'val as bool
		m_displayHeader = val
	End Property
	
'	---------------------------------------------------------------------	
	Public Property Get DisplayHeader()
		DisplayHeader = m_displayHeader
	End Property
	
'	---------------------------------------------------------------------	
	Public Property Let DisplayWeekdayRow(val)
		m_displayWeekdayRow = val
	End Property

'	---------------------------------------------------------------------
	Public Property Get DisplayWeekdayRow()
		DisplayWeekdayRow = m_displayWeekdayRow
	End Property
		
'	---------------------------------------------------------------------	
	Public Property Let DisplayWeekdayRowStyle(val)
		m_displayWeekdayRowStyle = val
	End Property

'	---------------------------------------------------------------------
	Public Property Get DisplayWeekdayRowStyle()
		DisplayWeekdayRowStyle = m_displayWeekdayRowStyle
	End Property
		
'	---------------------------------------------------------------------
	Public Property Let DisplayDayNumbersAsLinks(val)
		m_displayDayNumbersAsLinks = val
	End Property

'	---------------------------------------------------------------------	
	Public Property Get DisplayDayNumbersAsLinks()
		DisplayDayNumbersAsLinks = m_displayDayNumbersAsLinks
	End Property
'	---------------------------------------------------------------------	
	Private Function IsSelected(val, checkVal)
		If CStr(val) = CStr(checkVal) Then
			IsSelected = " selected=""selected"""
		End If
	End Function
	
End Class

</script>