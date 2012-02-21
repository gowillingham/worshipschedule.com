<script language="vbscript" runat="server" type="text/vbscript">

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
	
</script>
