<script language="vbscript" type="text/vbscript" runat="server">

' YY or YYYY - year
' MM - month number
' MMM - short month name
' MMMM - long month name
' D or DD - day
' DDD or DDDD - day name
' HH or HHH - 24 hour time
' hh or hhh - 12 hour time
' nn - minutes
' ss - seconds
' pp - pm/am
' PP - PM/AM
' px - p/a
Class cFormatDate

	Private m_dDate 'as date
	Private m_sMask 'as string
	Private SHORT_NAME 'as int
	Private LONG_NAME 'as int
	
	Public Property Let DateToConvert(var)	'as date
		m_dDate = var
	End Property
	
	Public Property Let Mask(var) 'as string
		m_sMask = var
	End Property
	
	Public Function Casual(ByVal dateIn)
		Dim dateOut
	
		Dim isToday			: isToday = False
		Dim isYesterday		: isYesterday = False
		
		' see if date is today
		If DateDiff("d", Now(), CDate(dateIn)) = 0 Then
			dateOut = "Today" 
			isToday = True
		End If
				
		' see if date is yesterday
		If DateDiff("d", Now(), CDate(dateIn)) = -1 Then
			dateOut = "Yesterday"
			isYesterday = True
		End If
		
		If Not(isToday Or isYesterday) Then
			dateOut = MonthToText(Month(dateIn), SHORT_NAME)
			dateOut = dateOut & " " & Day(dateIn)
			If Year(dateIn) <> Year(Now()) Then
				dateOut = dateOut & ", " & Year(dateIn)
			End If
		End If
		
		Casual = dateOut
	End Function
	
	Public Function Convert(ByVal theDate, ByVal formatMask) 'as string
		DateToConvert = theDate
		Mask = formatMask
		
		' drop in year
		m_sMask = Replace(m_sMask, "YYYY", GetYear("YYYY"))
		m_sMask = Replace(m_sMask, "YY", GetYear("YY"))
		
		' drop in month text
		m_sMask = Replace(m_sMask, "MMMM", GetMonth("MMMM"))
		m_sMask = Replace(m_sMask, "MMM", GetMonth("MMM"))
		m_sMask = Replace(m_sMask, "MM", GetMonth("MM"))
		m_sMask = Replace(m_sMask, "mm", GetMonth("mm"))
		
		' drop in day
		m_sMask = Replace(m_sMask, "DDDD", GetDay("DDDD"))
		m_sMask = Replace(m_sMask, "DDD", GetDay("DDD"))
		m_sMask = Replace(m_sMask, "DD", GetDay("DD"))
		m_sMask = Replace(m_sMask, "dd", GetDay("dd"))
		
		' drop in time
		m_sMask = Replace(m_sMask, "HHH", GetTime("HHH"))
		m_sMask = Replace(m_sMask, "hhh", GetTime("hhh"))
		m_sMask = Replace(m_sMask, "HH", GetTime("HH"))
		m_sMask = Replace(m_sMask, "hh", GetTime("hh"))
		
		' drop in minutes/seconds
		m_sMask = Replace(m_sMask, "nn", GetTime("nn"))
		m_sMask = Replace(m_sMask, "ss", GetTime("ss"))
		
		' drop in am/pm indicator
		m_sMask = Replace(m_sMask, "PP", GetIndicator("PP"))
		m_sMask = Replace(m_sMask, "pp", GetIndicator("pp"))
		m_sMask = Replace(m_sMask, "px", GetIndicator("px"))
		
		Convert = m_sMask
	End Function
	
	Public Function MonthToText(monthNumber, style)
		Dim str, list
		
		list = MonthList()
		If style = 1 Then
			str = list(SHORT_NAME, monthNumber)
		Else
			str = list(LONG_NAME, monthNumber)
		End If
		
		MonthToText = str
	End Function
	
	Public Function DayToText(dayNumber, style)
		Dim str, list
		
		list = DayList()
		If style = 1 Then
			str = list(SHORT_NAME, monthNumber)
		Else
			str = list(LONG_NAME, monthNumber)
		End If
		
		DayToText = str
	End Function
	
	Private Function GetIndicator(mask)
		Dim indicator
		
		Select Case mask
			Case "PP"
				indicator = Right(FormatDateTime(m_dDate), 2)
				indicator = UCase(indicator)
			Case "pp"
				indicator = Right(FormatDateTime(m_dDate), 2)
				indicator = LCase(indicator)
			Case "px"
				indicator = Left(Right(FormatDateTime(m_dDate), 2), 1)
				indicator = LCase(indicator)
			Case Else
		End Select
		
		GetIndicator = indicator
	End Function
	
	Private Function GetTime(mask)
		Dim timeStr
		Select Case mask
			Case "HHH"
				timeStr = Split(FormatDateTime(m_dDate, vbShortTime), ":")(0)
				timeStr = Right(("0" & timeStr), 2)
			Case "hhh"
				timeStr = Split(FormatDateTime(m_dDate, vbLongTime), ":")(0)
				timeStr = Right(("0" & timeStr), 2)
			Case "HH"
				timeStr = Split(FormatDateTime(m_dDate, vbShortTime), ":")(0)
				timeStr = CStr(CInt(timeStr))
			Case "hh"
				timeStr = Split(FormatDateTime(m_dDate, vbLongTime), ":")(0)
				timeStr = CStr(CInt(timeStr))
			Case "nn"
				timeStr = Minute(m_dDate)
				timeStr = Right(("0" & timeStr), 2)
			Case "ss"
				timeStr = Second(m_dDate)
				timeStr = Right(("0" & timeStr), 2)
			Case "ms"
			Case Else
		End Select
		
		GetTime = timeStr
	End Function
	
	Private Function GetDay(mask)
		Dim days
		
		Select Case mask
			Case "DDDD"
				days = DayList()
				GetDay = days(LONG_NAME, Weekday(m_dDate, vbSunday)) 
			Case "DDD"
				days = DayList()
				GetDay = days(SHORT_NAME, Weekday(m_dDate, vbSunday)) 
			Case "DD"
				GetDay = Right(("0" & Day(m_dDate)), 2)
			Case "dd"
				GetDay = Day(m_dDate)
			Case Else
		End Select
	End Function
	
	Private Function DayList()
		Dim dayTextList(2,8)
		
		dayTextList(LONG_NAME,1)="Sunday": dayTextList(SHORT_NAME,1)="Sun"
		dayTextList(LONG_NAME,2)="Monday": dayTextList(SHORT_NAME,2)="Mon"
		dayTextList(LONG_NAME,3)="Tuesday": dayTextList(SHORT_NAME,3)="Tues"
		dayTextList(LONG_NAME,4)="Wednesday": dayTextList(SHORT_NAME,4)="Wed"
		dayTextList(LONG_NAME,5)="Thursday": dayTextList(SHORT_NAME,5)="Thur"
		dayTextList(LONG_NAME,6)="Friday": dayTextList(SHORT_NAME,6)="Fri"
		dayTextList(LONG_NAME,7)="Saturday": dayTextList(SHORT_NAME,7)="Sat"
		
		DayList = dayTextList
	End Function
	
	Private Function GetMonth(mask)
		Dim months
		
		Select Case mask
			Case "MMMM"
				months = MonthList()
				GetMonth = months(LONG_NAME, Month(m_dDate))
			Case "MMM"
				months = MonthList()
				GetMonth = months(SHORT_NAME, Month(m_dDate))
			Case "MM"
				GetMonth = Right(("0" & Month(m_dDate)), 2)
			Case "mm"
				GetMonth = Month(m_dDate)
			Case Else
				RaiseError("Illegal placeholder for GetMonth() [" & mask & "]")
		End Select
	End Function
	
	Private Function MonthList()
		Dim monthTextList(2,13)

		monthTextList(LONG_NAME,1)="January": monthTextList(SHORT_NAME,1)="Jan"
		monthTextList(LONG_NAME,2)="February": monthTextList(SHORT_NAME,2)="Feb"
		monthTextList(LONG_NAME,3)="March": monthTextList(SHORT_NAME,3)="Mar"
		monthTextList(LONG_NAME,4)="April": monthTextList(SHORT_NAME,4)="Apr"
		monthTextList(LONG_NAME,5)="May": monthTextList(SHORT_NAME,5)="May"
		monthTextList(LONG_NAME,6)="June": monthTextList(SHORT_NAME,6)="June"
		monthTextList(LONG_NAME,7)="July": monthTextList(SHORT_NAME,7)="July"
		monthTextList(LONG_NAME,8)="August": monthTextList(SHORT_NAME,8)="Aug"
		monthTextList(LONG_NAME,9)="September": monthTextList(SHORT_NAME,9)="Sept"
		monthTextList(LONG_NAME,10)="October": monthTextList(SHORT_NAME,10)="Oct"
		monthTextList(LONG_NAME,11)="November": monthTextList(SHORT_NAME,11)="Nov"
		monthTextList(LONG_NAME,12)="December": monthTextList(SHORT_NAME,12)="Dec"
		
		MonthList = monthTextList
	End Function
	
	Private Function GetYear(mask)
		Select Case mask
			Case "YYYY"
				GetYear = Year(m_dDate)
			Case "YY"
				GetYear = Right(Year(m_dDate), 2)
			Case Else
				RaiseError("Illegal placeholder for GetYear()")
		End Select
	End Function
	
	Private Function RaiseError(errText)
		Dim str
		str = str & "<div>"
		str = str & "<span style=""font-weight:bold;color:red;"">Error::cFormatDate::</span>"
		str = str & errText
		str = str & "</div>"
		
		Response.Write str
		Response.End
	End Function	
	
	Private Sub Class_Initialize()
		SHORT_NAME = 1
		LONG_NAME = 0
	End Sub

	Private Sub Class_Terminate()
	
	End Sub

End Class

</script>