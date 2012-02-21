<script runat="server" language="vbscript" type="text/vbscript">

Class cFileDisplay
	Private m_sCssImgClass
	Private m_sIconBGColor
	Private m_sFileName
	Private m_sFileExtension
	
	Property Let CssImgClass(val)
		m_sCssImgClass = val
	End Property
	
	Public Function DisplayFile(ByVal sFile, sIconBGColor)
		'return icon+filename
		Dim sExt, sName, sIconPath, str
		
		m_sIconBGColor = sIconBGColor
		If Len(sFile) = 0 Then Exit Function
		
		str = str & "<img src=""" & GetIconPath(GetFileExtension(sFile)) & """ style=""display:inline;padding:0 0 0 0;margin:0 0 -2px 0;"" alt=""" & sFile & """ />&nbsp;" & sFile
		DisplayFile = "<span style=""white-space:nowrap;"">" & str & "</span>"
	End Function
	
	Public Function DisplayFileLink(ByVal sFile, sLinkText, sIconBGColor)
		'return icon+linktext
		Dim str
		
		m_sIconBGColor = sIconBGColor
		If Len(sFile) = 0 Then Exit Function
		If Len(sLinkText) = 0 Then Exit Function
		
		str = str & "<img src=""" & GetIconPath(GetFileExtension(sFile)) & """ style=""display:inline;padding:0 0 0 0;margin:0 0 -2px 0;"" alt=""" & sFile & """ />&nbsp;" & sLinkText
		DisplayFileLink = "<span style=""white-space:nowrap;"">" & str & "</span>"
	End Function
	
	Private Function GetFileExtension(ByVal sFileName)
		Dim arr
		GetFileExtension = ""
		arr = Split(sFileName, ".")
		If IsArray(arr) Then
			GetFileExtension = Trim(UCase(CStr(arr(UBound(arr)))))
		End If
	End Function
	
	Public Function GetIconPath(sExtension)
		Dim str, sIconDir
		sIconDir = "/_images/icons/"
		
		Select Case UCase(sExtension)
			Case "M4A"
				str = sIconDir & "page_white_cd.png"
			Case "WMV"
				str = sIconDir & "page_white_cd.png"
			Case "AVI"
				str = sIconDir & "page_white_cd.png"
			Case "MP3"
				str = sIconDir & "page_white_cd.png"
			Case "WAV"
				str = sIconDir & "page_white_cd.png"
			Case "MPEG"
				str = sIconDir & "page_white_cd.png"
			Case "DOC"
				str = sIconDir & "page_white_word.png"
			Case "RTF"
				str = sIconDir & "page_white_word.png"
			Case "XLS"
				str = sIconDir & "page_white_excel.png"
			Case "PDF"
				str = sIconDir & "page_white_acrobat.png"
			Case "HTM"
				str = sIconDir & "page_white_world.png"
			Case "HTML"
				str = sIconDir & "page_white_world.png"
			Case "ZIP"
				str = sIconDir & "page_white_compressed.png"
			Case "TXT"
				str = sIconDir & "page_white_text.png"
			Case "BMP"
				str = sIconDir & "page_white_picture.png"
			Case "JPG"
				str = sIconDir & "page_white_picture.png"
			Case "JPEG"
				str = sIconDir & "page_white_picture.png"
			Case "PNG"
				str = sIconDir & "page_white_picture.png"
			Case "GIF"
				str = sIconDir & "page_white_picture.png"
			Case "ICO"
				str = sIconDir & "page_white_picture.png"
			Case "TIF"
				str = sIconDir & "page_white_picture.png"
			Case "TIFF"
				str = sIconDir & "page_white_picture.png"
			Case Else
				str = sIconDir & "page_white.png"
		End Select
		GetIconPath = str
	End Function
	
	Private Sub Class_Initialize()
	
	End Sub
	
	Private Sub Class_Terminate()
	
	End Sub
End Class

</script>