<script runat="server" language="vbscript" type="text/vbscript">

Class cFile

	Private m_FileID		'as long int
	Private m_FileName		'as string
	Private m_FriendlyName		'as string
	Private m_Description		'as string
	Private m_ClientID		'as long int
	Private m_ProgramID		'as long int
	Private m_FileOwnerID		'as long int
	Private m_DateCreated		'as date
	Private m_DateModified		'as date
	Private m_FileExtension		'as string
	Private m_FileSize		'as long int
	Private m_MIMEType		'as string
	Private m_MIMESubType		'as string
	Private m_IsPublic			' tinyint
	Private m_DownloadCount		' int
	Private m_EventCount		' int
	Private m_ProgramName	' str

	Private m_sSQL		'as string
	Private m_cnn		'as ADODB.Connection
	Private m_rs		'as ADODB.Recordset
	
	Private CLASS_NAME	'as string
	Private FILESTORE_PATH	' as string
	
	Public Property Get FileID() 'As long int
		FileID = m_FileID
	End Property

	Public Property Let FileID(val) 'As long int
		m_FileID = val
	End Property
	
	Public Property Get FileName() 'As string
		FileName = m_FileName
	End Property

	Public Property Let FileName(val) 'As string
		m_FileName = val
	End Property
	
	Public Property Get FriendlyName() 'As string
		FriendlyName = m_FriendlyName
	End Property

	Public Property Let FriendlyName(val) 'As string
		m_FriendlyName = val
	End Property
	
	Public Property Get Description() 'As string
		Description = m_Description
	End Property

	Public Property Let Description(val) 'As string
		m_Description = val
	End Property
	
	Public Property Get ClientID() 'As long int
		ClientID = m_ClientID
	End Property

	Public Property Let ClientID(val) 'As long int
		m_ClientID = val
	End Property
	
	Public Property Get ProgramID() 'As long int
		ProgramID = m_ProgramID
	End Property

	Public Property Let ProgramID(val) 'As long int
		m_ProgramID = val
	End Property
	
	Public Property Get FileOwnerID() 'As long int
		FileOwnerID = m_FileOwnerID
	End Property

	Public Property Let FileOwnerID(val) 'As long int
		m_FileOwnerID = val
	End Property
	
	Public Property Get DateCreated() 'As date
		DateCreated = m_DateCreated
	End Property

	Public Property Get DateModified() 'As date
		DateModified = m_DateModified
	End Property

	Public Property Get FileExtension() 'As string
		FileExtension = m_FileExtension
	End Property

	Public Property Let FileExtension(val) 'As string
		m_FileExtension = val
	End Property
	
	Public Property Get FileSize() 'As long int
		FileSize = m_FileSize
	End Property

	Public Property Let FileSize(val) 'As long int
		m_FileSize = val
	End Property
	
	Public Property Get MIMEType() 'As string
		MIMEType = m_MIMEType
	End Property

	Public Property Let MIMEType(val) 'As string
		m_MIMEType = val
	End Property
	
	Public Property Get MIMESubType() 'As string
		MIMESubType = m_MIMESubType
	End Property

	Public Property Let MIMESubType(val) 'As string
		m_MIMESubType = val
	End Property
	
	Public Property Get IsPublic() 'As tinyint
		IsPublic = m_IsPublic
	End Property

	Public Property Let IsPublic(val) 'As tinyint
		m_IsPublic = val
	End Property
	
	Public Property Get DownloadCount() 'As int
		DownloadCount = m_DownloadCount
	End Property

	Public Property Get EventCount() 'As int
		EventCount = m_EventCount
	End Property

	Public Property Let EventCount(val) 'As int
		m_EventCount = val
	End Property
	
	Public Property Get ProgramName()
		ProgramName = m_ProgramName
	End Property
	
	Public Property Get Path()
		Dim str
		
		str = FILESTORE_PATH & m_ClientID
		If Len(m_ProgramID) > 0 Then
			str = str & "\" & m_ProgramID 
		End If
		str = str & "\" & m_fileName
		
		Path = str
	End Property
	
	Private Sub Class_Initialize()
		m_sSQL = Application.Value("CNN_STR")
		FILESTORE_PATH = Application.Value("FILE_MANAGER_FILESTORE")
		CLASS_NAME = "cFile"
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
	
	Public Sub ToFilestore(upload, outError)
		' accept ASPSmartUpload.File object and save to filestore
		
		Dim fso			: Set fso = Server.CreateObject("Scripting.FileSystemObject")
		Dim path		: path = Application.Value("FILE_MANAGER_FILESTORE") & clientId
	
		Dim extension	: extension = upload.FileExt
		Dim rootName	: rootName = Replace(upload.FileName, "." & extension, "")
		Dim fileName	: fileName = ""
		Dim suffix		: suffix = ""
		
		Dim counter		: counter = 1
		
		'create client dir if necessary
		If Not fso.FolderExists(path) Then fso.CreateFolder(path)
		If Len(programId) > 0 Then
			path = path & "\" & programId
		End If
		If Not fso.FolderExists(path) Then fso.CreateFolder(path)
		
		' clean the file name eof illegal characters and spaces
		rootName = CleanFileName(rootName, "_")
		
		' rename the file if a file with that name already exists ..
		fileName = rootName & "." & extension
		Do While fso.FileExists(path & "\" & fileName)
			fileName = rootName & "(" & counter & ")" & "." & extension
			counter = counter + 1
		Loop
		
		' save the file to the filestore
		Call upload.SaveAs(path & "\" & fileName)
		
		m_fileName = fileName
		m_friendlyName = Replace(fileName, "." & extension, "")
		m_FileExtension = extension
		m_fileSize = upload.Size
		m_mimeType = upload.TypeMIME
		m_mimeSubType = upload.SubTypeMIME
		
		Call Add(outError)	
		Call Load()	
	End Sub
	
	Public Function GetFileStoreInfo()
		If Len(m_ClientID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".GetFileStoreInfo():", "Required parameter ClientID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-NameClient 1-FileCount 2-EventCount 3-Used 4-Available
		m_cnn.up_filesGetFilestoreInfo CLng(m_ClientID), m_rs
		If Not m_rs.EOF Then GetFileStoreInfo = m_rs.GetRows()		
		
		If m_rs.State = adStateOpen Then m_rs.Close	
	End Function
	
	Public Sub StreamFile(memberID, outError)
		If Len(m_FileID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".StreamFile():", "Required parameter FileID not provided.")
		If Len(memberID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".StreamFile():", "Required parameter MemberID not provided.")
		
		Dim fso			: Set fso = Server.CreateObject("Scripting.FileSystemObject")
		Dim stream		: Set stream = Server.CreateObject("ADODB.Stream")
		outError = 0
		
		Call Load()
		If fso.FileExists(Path) Then
			Response.Clear
			Response.Buffer = False
			Response.AddHeader "Content-Disposition", "attachment; filename=" & m_fileName
			Response.AddHeader "Content-Length", m_fileSize
			Response.ContentType = "application/octet-stream"	
			stream.Open
			stream.Type = 1
			Response.CharSet = "UTF-8"
			stream.LoadFromFile(Path)
			Response.BinaryWrite(stream.Read)
			stream.Close
		Else
			outError = -1
		End If 
		
		' add row to dbo.FileDownload
		Call InsertFileDownload(memberID, "")
		
		Set fso = Nothing
		Set stream = Nothing
	End Sub
	
	Public Function List(sortColumn)
		If Len(m_ClientID) = 0 Then Call Err.Raise(vbObjectError + 2, CLASS_NAME & ".List():", "Required parameter ClientID not provided.")
		
		Dim rs		: Set rs = Server.CreateObject("ADODB.Recordset")
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		rs.CursorLocation = adUseClient	' makes sortable
		
		' 0-FileID 1-FileName 2-FriendlyName 3-Description 4-ClientID 5-ProgramID 6-FileOwnerID
		' 7-DateCreated 8-DateModified 9-FileExtension 10-FileSize 11-MIMEType 12-MIMESubType 13-EventFileCount
		' 14-IsPublic 15-DownloadCount 16-ProgramName
		
		If Len(m_ProgramID) > 0 Then
			m_cnn.up_filesGetFileList CLng(m_ClientiD), CLng(m_ProgramID), rs
		Else
			m_cnn.up_filesGetFileList CLng(m_ClientiD), rs
		End If
		rs.Sort = sortColumn
		
		If Not rs.EOF Then List = rs.GetRows()
	End Function
	
	Public Function EventList()
		If Len(m_FileID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".EventList():", "Required parameter FileID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_filesGetEventsForFileList CLng(m_FileID), m_rs
		If Not m_rs.EOF Then EventList = m_rs.GetRows()
		
		If m_rs.State = adStateOpen Then m_rs.Close
	End Function
	
	Public Sub Load() 'As Boolean
		If Len(m_FileID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Load():", "Required parameter FileID not provided.")
	
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		If Not IsObject(m_rs) Then Set m_rs = Server.CreateObject("ADODB.Recordset")
		
		m_cnn.up_fileGetfile CLng(m_FileID), m_rs
		If Not m_rs.EOF Then
			m_FileName = m_rs("FileName").Value
			m_FriendlyName = m_rs("FriendlyName").Value
			m_Description = m_rs("Description").Value
			m_ClientID = m_rs("ClientID").Value
			m_ProgramID = m_rs("ProgramID").Value
			m_FileOwnerID = m_rs("FileOwnerID").Value
			m_DateCreated = m_rs("DateCreated").Value
			m_DateModified = m_rs("DateModified").Value
			m_FileExtension = m_rs("FileExtension").Value
			m_FileSize = m_rs("FileSize").Value
			m_MIMEType = m_rs("MIMEType").Value
			m_MIMESubType = m_rs("MIMESubType").Value
			m_IsPublic = m_rs("IsPublic").Value
			m_DownloadCount = m_rs("DownloadCount").Value
			m_EventCount = m_rs("EventCount").Value
			m_ProgramName = m_rs("ProgramName").Value
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
			.CommandText = "dbo.up_fileInsertFile"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@FileName", adVarChar, adParamInput, 256, m_FileName)
		cmd.Parameters.Append cmd.CreateParameter("@FriendlyName", adVarChar, adParamInput, 256, m_FriendlyName)
		cmd.Parameters.Append cmd.CreateParameter("@Description", adVarChar, adParamInput, 1000, m_Description)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, m_ClientID)
		If Len(m_ProgramID) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, CLng(m_ProgramID))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@FileOwnerID", adBigInt, adParamInput, 0, m_FileOwnerID)
		cmd.Parameters.Append cmd.CreateParameter("@FileExtension", adVarChar, adParamInput, 10, m_FileExtension)
		cmd.Parameters.Append cmd.CreateParameter("@FileSize", adBigInt, adParamInput, 0, m_FileSize)
		cmd.Parameters.Append cmd.CreateParameter("@MIMEType", adVarChar, adParamInput, 100, m_MIMEType)
		cmd.Parameters.Append cmd.CreateParameter("@MIMESubType", adVarChar, adParamInput, 100, m_MIMESubType)
		cmd.Parameters.Append cmd.CreateParameter("@IsPublic", adTinyInt, adParamInput, 0, CInt(m_IsPublic))
		cmd.Parameters.Append cmd.CreateParameter("@NewFilesID", adBigInt, adParamOutput)
	
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		m_FileID = cmd.Parameters("@NewFilesID").Value

		Set cmd = Nothing
	End Sub
	
	Public Sub Save(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_FileID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Save():", "Required parameter FileID not provided.")
		m_DateModified = Now()
		
		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_filesUpdateFile"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@FileID", adBigInt, adParamInput, 0, m_FileID)
		cmd.Parameters.Append cmd.CreateParameter("@FileName", adVarChar, adParamInput, 256, m_FileName)
		cmd.Parameters.Append cmd.CreateParameter("@FriendlyName", adVarChar, adParamInput, 256, m_FriendlyName)
		cmd.Parameters.Append cmd.CreateParameter("@Description", adVarChar, adParamInput, 1000, m_Description)
		cmd.Parameters.Append cmd.CreateParameter("@ClientID", adBigInt, adParamInput, 0, m_ClientID)
		If Len(m_ProgramID) > 0 Then
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, CLng(m_ProgramID))
		Else
			cmd.Parameters.Append cmd.CreateParameter("@ProgramID", adBigInt, adParamInput, 0, Null)
		End If
		cmd.Parameters.Append cmd.CreateParameter("@FileOwnerID", adBigInt, adParamInput, 0, m_FileOwnerID)
		cmd.Parameters.Append cmd.CreateParameter("@FileExtension", adVarChar, adParamInput, 10, m_FileExtension)
		cmd.Parameters.Append cmd.CreateParameter("@FileSize", adBigInt, adParamInput, 0, m_FileSize)
		cmd.Parameters.Append cmd.CreateParameter("@MIMEType", adVarChar, adParamInput, 100, m_MIMEType)
		cmd.Parameters.Append cmd.CreateParameter("@MIMESubType", adVarChar, adParamInput, 100, m_MIMESubType)
		cmd.Parameters.Append cmd.CreateParameter("@IsPublic", adTinyInt, adParamInput, 0, CInt(m_IsPublic))
		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value
		
		Set cmd = Nothing
	End Sub
	
	Public Sub Delete(ByRef outError) 'As Boolean
		Dim cmd
		
		If Len(m_FileID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".Delete():", "Required parameter FileID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_filesDeleteFile"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@FileID", adBigInt, adParamInput, 0, m_FileID)

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub
	
	Private Sub InsertFileDownload(memberID, outError)
		Dim cmd
		
		If Len(m_FileID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".InsertFileDownload():", "Required parameter FileID not provided.")
		If Len(memberID) = 0 Then Call Err.Raise(vbObjectError + 1, CLASS_NAME & ".InsertFileDownload():", "Required parameter MemberID not provided.")

		If Not IsObject(m_cnn) Then 
			Set m_cnn = Server.CreateObject("ADODB.Connection")
			m_cnn.Open m_sSQL
		End If
		Set cmd = Server.CreateObject("ADODB.Command")

		With cmd
			.CommandType = adCmdStoredProc
			.CommandText = "dbo.up_filesInsertFileDownload"
			.ActiveConnection = m_cnn
		End With
		
		cmd.Parameters.Append cmd.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, 0)
		cmd.Parameters.Append cmd.CreateParameter("@FileID", adBigInt, adParamInput, 0, CLng(m_FileID))
		cmd.Parameters.Append cmd.CreateParameter("@MemberID", adBigInt, adParamInput, 0, CLng(memberID))
		cmd.Parameters.Append cmd.CreateParameter("@DateCreated", adDate, adParamInput, 0, Now())

		cmd.Execute ,,adExecuteNoRecords
		outError = cmd.Parameters("@RETURN_VALUE").Value

		Set cmd = Nothing
	End Sub

	Function CleanFileName(str, token)
		' remove illegal characters from filename and replace with token
		' illegal characters ..  < > : " / \ |
	
		str = Replace(str, " ", token)
		str = Replace(str, """", token)
		str = Replace(str, "<", token)
		str = Replace(str, ">", token)
		str = Replace(str, ":", token)
		str = Replace(str, "\", token)
		str = Replace(str, "/", token)
		str = Replace(str, "|", token)
		
		CleanFileName = str
	End Function
End Class



</script>