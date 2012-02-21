<script language="vbscript" type="text/vbscript" runat="server">

Function Wait(interval)
		dim t
		
		If Application.Value("IsLiveServer") Then Exit Function
		
         ' Wait the specified number of seconds.  But if the
         ' end-of-wait time is beyond midnight then don't wait.
         t = timer() + interval
         if t < 86399 then       ' 86400 = 24 hrs * 60 minutes * 60 seconds
                 while timer() < t
                 wend
         end if
End Function

</script>