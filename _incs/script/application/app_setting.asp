<script runat="server" language="vbscript" type="text/vbscript">
Function GetSettingValue(sSettingName, aSettings)
	Dim i
	GetSettingValue = 0
	
	If IsArray(aSettings) Then
		For i = 0 To UBound(aSettings, 2)
			If aSettings(0,i) = sSettingName Then
				GetSettingValue = aSettings(1,i)
				Exit For
			End If
		Next
	End If
End Function

Function GetApplicationSetting(iClientID, sSettingType)
	'return an array of application settings from tbl UserSettings
	Dim cnn, rs
	
	'get default settings if no ClientID is passed
	If Len(iClientID) = 0 Then iClientID = 0
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	With cnn
		.ConnectionString = Application.Value("CNN_STR")
		.Open
		Set rs = Server.CreateObject("ADODB.Recordset")
		.up_adminGetApplicationSetting CInt(iClientID), sSettingType, rs
	End With
	If Not rs.EOF Then GetApplicationSetting = rs.GetRows
	rs.Close: Set rs = Nothing
	cnn.Close: Set cnn = Nothing
End Function
</script>
