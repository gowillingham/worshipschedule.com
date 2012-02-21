<script runat="server" type="text/vbscript" language="vbscript">
	
	' copies schedule to toEventID from fromEventID
	' dependencies: /_incs/class/schedule_build_cls.asp
	
	Sub CopyEventSchedule(toEventID, fromEventID, outError) 
		Dim scheduleBuild			: Set scheduleBuild = New cScheduleBuild
		
		If Len(toEventID) = 0 Or Len(fromEventID) = 0 Then
			outError = -100 
			Exit Sub
		End If

		outError = 0
		scheduleBuild.EventID = toEventID
		Call scheduleBuild.CopyFromEvent(fromEventID, outError)
		
		Set scheduleBuild = Nothing
	End Sub
</script>


