		<div id="topbar">
			<div id="appheader">
				<div id="topnav">
					<%=m_topBarText %>
					<div id="topdatetime" >
						<%=WeekdayName(Weekday(Now())) & ", " & MonthName(Month(Now())) & " " & Day(Now()) & ", " & Year(Now()) & "&nbsp;|&nbsp;" & Left(TimeValue(Now()), Len(TimeValue(Now())) - 6) & Right(TimeValue(Now()), 2) %>
					</div>
					
				</div>
				<div id="appheaderlogo">
					<a href="/">
						<img src="/_images/logo_bluebutton_help.png" alt="Home" />
					</a>
				</div>
			</div>
		</div>
