<script type="text/vbscript" runat="server" language="vbscript">
Class cState

	Private m_stateID		' int
	Private m_stateCode		' varchar(2)
	Private m_longName		' varchar
	Private m_isActive		' tinyint
	Private m_countryID		' int
	Private m_countryName	' varchar
	
	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Let StateID(val)
		m_stateID = val
	End Property
	
	Public Property Get StateID()
		StateID = m_StateID
	End Property
	
	Public Property Let CountryID(val)
		m_countryID = val
	End Property
	
	Public Property Get CountryID()
		CountryID = m_CountryID
	End Property
	
	Public Property Get IsActive()
		IsActive = m_isActive
	End Property
	
	Public Property Get LongName()
		LongName = m_longName
	End Property
	
	Public Property Get CountryName()
		CountryName = m_countryName
	End Property
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cState"
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(m_rs) Then
			If m_rs.State = adStateOpen Then m_rs.Close
			Set m_rs = Nothing
		End If
		If IsObject(m_cnn) Then
			If m_cnn.State = adStateOpen Then m_cnn.Close
			Set m_cnn = Nothing
		End If
	End Sub
	
	Public Function OptionListToString(val)
		Dim str, i
		
		If Len(m_countryID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".List():", "Required parameter CountryID not provided.")

		Dim states				: states = List()
		Dim selected			: selected = ""
		
		For i = 0 To UBound(states,2)
			selected = ""
			If CStr(states(0,i) & "") = CStr(val & "") Then selected = " selected=""selected"""
			
			str = str & "<option value=""" & states(0,i) & """" & selected & ">" & Server.HTMLEncode(states(1,i) & " - " & states(2,i)) & "</option>"
		Next		
		
		OptionListToString = str
	End Function
	
	Public Sub Load()
		If Len(m_stateID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter StateID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_adminGetState CLng(m_stateID), m_rs
		If Not m_rs.EOF Then
			m_stateCode = m_rs("StateCode").Value
			m_longName = m_rs("LongName").Value
			m_isActive = m_rs("IsActive").Value
			m_countryID = m_rs("CountryID").Value
			m_countryName = m_rs("CountryName").Value
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Sub
	
	Public Function List()
		If Len(m_countryID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".List():", "Required parameter CountryID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-StateID 1-StateCode 2-LongName 3-IsActive 4-CountryID 5-CountryName

		m_cnn.up_adminGetStateList CLng(m_countryID), m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()

		m_rs.Close()
	End Function
End Class
</script>
