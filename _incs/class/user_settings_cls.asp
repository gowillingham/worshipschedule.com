<script type="text/vbscript" runat="server" language="vbscript">

Class cUserSettings
	Private m_SettingsList	'dictionary obj
	Private m_cnn
	Private m_CnnString
	
	Public Sub ChangeSetting(settingName, val)
	
		If m_SettingsList.Exists(settingName) Then
			' update setting if it exists
			m_SettingsList(settingName) = val
		Else
			' add setting if new
			m_SettingsList.Add settingName, val
		End If
	End Sub
	
	Public Function GetSetting(settingName)
		If m_SettingsList.Exists(settingName) Then
			GetSetting = m_SettingsList(settingName)
		Else
			GetSetting = 0
		End If
	End Function
	
	Public Function IsChecked(settingName)
		Dim str
		
		If m_SettingsList.Exists(settingName) Then
			If CStr(m_SettingsList(settingName)) = "1" Then
				str = " checked=""checked"""
			End If
		Else
			str = ""
		End If
		
		IsChecked = str
	End Function
	
	Public Sub Save(clientID, outError)
	
		' clear existing from db for this clientID
		Call DeleteSettings(clientID, outError)
		
		' insert new settings for this clientID
		Call SaveSettings(clientID, outError)
	End Sub
	
	Public Sub Load(clientID)
		Dim arr, i
		
		' clear any settings already loaded
		m_SettingsList.RemoveAll
		
		' load from db
		arr = RetrieveSettings(clientID)
		
		For i = 0 To UBound(arr,2)
			m_SettingsList.Add arr(0,i), arr(1,i)
		Next
		
	End Sub
	
	Private Function RetrieveSettings(clientID)
		Dim rs
		
		If Not IsObject(m_cnn) Then Set m_cnn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")
		m_cnn.Open m_CnnString
		
		m_cnn.up_adminGetApplicationSetting CLng(clientID), "MemberRequiredField", rs
		If Not rs.EOF Then RetrieveSettings = rs.GetRows
		
		rs.Close: Set rs = Nothing
		m_cnn.Close
	End Function
	
	Private Function DeleteSettings(clientID, outError)
		Dim cmd

		If Not IsObject(m_cnn) Then Set m_cnn = Server.CreateObject("ADODB.Connection")
		Set cmd = Server.CreateObject("ADODB.Command")
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_adminDeleteApplicationSettings"
			m_cnn.Open m_CnnString
			.ActiveConnection = m_cnn
		End With

		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, CLng(clientID))
		cmd.Parameters.Append cmd.CreateParameter("@SettingType", adVarChar, adParamInput, 50, "MemberRequiredField")

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
		m_cnn.Close
	End Function
	
	Private Sub SaveSettings(clientID, outError)
		Dim cmd, i, keys
		
		If Not IsObject(m_cnn) Then Set m_cnn = Server.CreateObject("ADODB.Connection")
		Set cmd = Server.CreateObject("ADODB.Command")
		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "up_adminInsertApplicationSetting"
			m_cnn.Open m_CnnString
			.ActiveConnection = m_cnn
		End With

		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, CLng(clientID))
		cmd.Parameters.Append cmd.CreateParameter("@SettingType", adVarChar, adParamInput, 50, "MemberRequiredField")
		cmd.Parameters.Append cmd.CreateParameter("@SettingName", adVarChar, adParamInput, 25)
		cmd.Parameters.Append cmd.CreateParameter("@SettingValue", adUnsignedTinyInt, adParamInput, 0)
		
		keys = m_SettingsList.Keys
		For i = 0 To UBound(keys)
			cmd.Parameters("@SettingName").Value = keys(i)
			cmd.Parameters("@SettingValue").Value = m_SettingsList(keys(i))
			cmd.Execute ,,adExecuteNoRecords
			If cmd.Parameters("@RETURN_VALUE").Value <> 0 Then
				outError = cmd.Parameters("@RETURN_VALUE").Value
			End If
		Next
		
		Set cmd = Nothing
		m_cnn.Close
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(m_SettingsList) Then Set m_SettingsList = Nothing
		If IsObject(m_cnn) Then Set m_cnn = Nothing
	End Sub
	
	Private Sub Class_Initialize()
		m_CnnString = Application.Value("CNN_STR")
		Set m_SettingsList = Server.CreateObject("Scripting.Dictionary")
	End Sub

End Class

</script>


