<script runat="server" language="vbscript" type="text/vbscript">
Class cFaq
	Private m_FaqID		'as int
	Private m_Title		'as string
	Private m_Text		'as string
	Private m_Priority		'as small int
	Private m_DateCreated		'as date
	Private m_CategoryID		'as int
	Private m_Category			' as string

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	
	Public Property Get FaqID() 'As int
		FaqID = m_FaqID
	End Property

	Public Property Let FaqID(val) 'As int
		m_FaqID = val
	End Property
	
	Public Property Get Title() 'As string
		Title = m_Title
	End Property

	Public Property Let Title(val) 'As string
		m_Title = val
	End Property
	
	Public Property Get Text() 'As string
		Text = m_Text
	End Property

	Public Property Let Text(val) 'As string
		m_Text = val
	End Property
	
	Public Property Get Priority() 'As small int
		Priority = m_Priority
	End Property

	Public Property Let Priority(val) 'As small int
		m_Priority = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	Public Property Get CategoryID() 'As int
		CategoryID = m_CategoryID
	End Property

	Public Property Let CategoryID(val) 'As int
		m_CategoryID = val
	End Property
	
	Public Property Get Category()
		Category = m_Category
	End Property
	
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		CLASS_NAME = "cFaq"
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
	
	Public Function List() ' as array
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_faqGetFaqList m_rs
		If Not m_rs.EOF Then List = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function ListByCategoryID() ' as array
		If Len(m_CategoryID) = 0 Then Call Err.Raise(vbObjectError + 3, CLASS_NAME & ".ListByCategoryID():", "Required parameter CategoryID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-FaqCategoryID 1-SuperCategory 2-Category 3-Description 4-Priority
		m_cnn.up_faqGetFaqListByCategoryID CLng(m_CategoryID), m_rs
		If Not m_rs.EOF Then ListByCategoryID = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Function CategoryList() ' as array
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-FaqCategoryID 1-SuperCategory 2-Category 3-Description 4-Priority
		m_cnn.up_faqGetFaqCategoryList m_rs
		If Not m_rs.EOF Then CategoryList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Sub Load() 'As Boolean
		If Len(m_FaqID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter FaqID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_faqGetFaq CLng(m_FaqID), m_rs
		If Not m_rs.EOF Then
			m_FaqID = m_rs("FaqID").Value
			m_Title = m_rs("Title").Value
			m_Text = m_rs("Text").Value
			m_Priority = m_rs("Priority").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_CategoryID = m_rs("CategoryID").Value
			m_Category = m_rs("Category").Value
		End If
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Sub
	
	Public Sub Add(ByRef outError) 'As Boolean
		Dim cmd
		
		m_DateCreated = Now()
		m_DateModified = Now()

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_faqInsertFaq"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@Title", adVarChar, adParamInput, 256, m_Title)
		cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 8000, m_Text)
		cmd.Parameters.Append cmd.CreateParameter("@Priority", adUnsignedTinyInt, adParamInput, 0, m_Priority)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@CategoryID", adInteger, adParamInput, 0, m_CategoryID)
		cmd.Parameters.Append cmd.CreateParameter("@NewFaqID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		m_FaqID = cmd.Parameters("@NewFaqID").Value

		Set cmd = Nothing
	End Sub
	
	Public Sub Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_FaqID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter FaqID not provided.")
		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_faqUpdateFaq"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@FaqID", adInteger, adParamInput, 0, m_FaqID)
		cmd.Parameters.Append cmd.CreateParameter("@Title", adVarChar, adParamInput, 256, m_Title)
		cmd.Parameters.Append cmd.CreateParameter("@Text", adVarChar, adParamInput, 8000, m_Text)
		cmd.Parameters.Append cmd.CreateParameter("@Priority", adUnsignedTinyInt, adParamInput, 0, m_Priority)
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, m_DateCreated)
		cmd.Parameters.Append cmd.CreateParameter("@CategoryID", adInteger, adParamInput, 0, m_CategoryID)
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Sub
	
	Public Sub Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_FaqID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter FaqID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_faqDeleteFaq"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@FaqID", adInteger, adParamInput, 0, m_FaqID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
End Class
</script>

