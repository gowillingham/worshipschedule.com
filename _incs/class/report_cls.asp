<script type="text/vbscript" runat="server" language="vbscript">

Class cReport
	Private m_clientId
	Private m_programId
	Private m_columnList
	Private m_sortColumns

	Private CLASS_NAME
	Private IS_TYPE_OF
	
	Public Property Let ClientID(val)
		m_clientID = val
	End Property
	
	Public Property Get ClientID()
		ClientID = val
	End Property
	
	Public Property Let ProgramID(val)
		m_programId = val
	End Property
	
	Public Property Get ProgramID()
		ProgramID = val
	End Property
	
	Public Property Let ColumnList(val)
		m_columnList = val
	End Property
	
	Public Property Get ColumnList()
		ColumnList = m_ColumnList
	End Property
	
	Public Property Let SortColumns(val)
		m_sortColumns = val
	End Property
	
	Public Property Get SortColumns()
		SortColumns = m_sortColumns
	End Property
	
	Function ToString(columnsToInclude)
		Dim str, i, j
		
		Dim header
		Dim body
		
		Dim rs				: Set rs = Report()
		Dim list			: If Not rs.EOF Then list = rs.GetRows()
		
		' clean '[' and ']' and '[space] ' from columnList() and  split
		Dim colList			: colList = Split(Replace(Replace(ColumnList(), "[", ""), "]", ""), ",")
		
		' idx() is an array list of columns to include ..
		Dim idx			
		If Len(columnsToInclude) > 0 Then
			idx = Split(Replace(columnsToInclude, " ", ""), ",")
		Else
			ReDim idx(UBound(colList))
			For i = 0 To UBound(idx)
				idx(i) = i
			Next
		End If

		header = header & "<thead><tr>"
		For i = 0 To UBound(idx)
		
			' only show the columns from columnsToInclude ..
			header = header & "<th>" & html(colList(idx(i))) & "</th>"
		Next
		header = header & "</tr></thead>"
		
		body = body & "<tbody>"
		If IsArray(list) Then
			For j = 0 To UBound(list,2)
				body = body & "<tr>"
				
				' only show the columns from columnsToInclude ..
				For i = 0 To UBound(idx)
					body = body & "<td>" & Server.HTMLEncode(list(idx(i), j) & "") & "</td>"
				Next
				body = body & "</tr>"
			Next
		Else
			body = body & "<tr><td colspan=""" & UBound(idx) + 1  & """></td></tr>"
		End If
		body = body & "</tbody>"
		
		str = str & "<div class=""report-container"">"
		str = str & "<table id=""report-table"" class=""tablesorter"">"
		str = str & header
		str = str & body
		str = str & "</table></div>"

		ToString = str
	End Function

	Public Function Report()
		' returns a sorted recordset ..
		
		Dim view
		Dim sql	
		
		If Len(m_programId) > 0 Then
			view = "dbo.vw_ProgramMemberListForReport vw"
			sql = "SELECT " & ColumnList() & " FROM " & view & " WHERE vw.ProgramId = " & m_programId & " ORDER BY [Last name],[First name]"
		Else
			view = "dbo.vw_MemberListForReport vw"
			sql = "SELECT " & ColumnList() & " FROM " & view & " WHERE vw.ClientID = " & m_clientID & " ORDER BY [Last name],[First name]"
		End If
		
		Dim cnn, rs
		Set cnn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")
		cnn.CursorLocation = adUseClient
		
		cnn.Open Application.Value("CNN_STR")
		Set rs = cnn.Execute(sql)
		rs.Sort = m_sortColumns
		
		Set Report = rs
	End Function
	
	Private Sub Class_Initialize()
		CLASS_NAME = "cReport"
		IS_TYPE_OF = "ws.Report"
	End Sub
	
	Private Sub Class_Terminate()
	
	End Sub
End Class

</script>
