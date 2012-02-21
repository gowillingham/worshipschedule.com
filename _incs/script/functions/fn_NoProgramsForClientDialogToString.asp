<script runat="server" type="text/vbscript" language="vbscript">

Function NoProgramsForClientDialogToString(page)
	Dim dialog				: Set dialog = New cDialog
	Dim pg					: Set pg = page.Clone()
	
	Dim helpLink
	
	dialog.Headline	= "Whoa, something is missing here .."
	
	dialog.Text = dialog.Text & "<p>It looks like the " & Application.Value("APPLICATION_NAME") & " account you are logged into (" & html(page.Client.NameClient) & ") doesn't have any programs set up. "
	dialog.Text = dialog.Text & "Perhaps this account is brand new and no programs have been created yet. </p>"
	dialog.Text = dialog.Text & "<p>Before you will can use this account to manage your " & html(page.Client.NameClient) & " schedule and events, "
	dialog.Text = dialog.Text & "you'll need an account administrator to set up at least one program. "
	If page.Member.IsAdmin Then
		dialog.Text = dialog.Text & "Click <strong>Create your first program</strong> to get started. "
	End If
	dialog.Text = dialog.Text & "</p>"
	
	dialog.SubText = dialog.SubText & "<p>When this is fixed, you can use this page to add, remove, or change the programs that belong to your " & html(page.Client.NameClient) & " account. </p>"
	dialog.SubText = dialog.SubText & "<p>" & Application.Value("APPLICATION_NAME") & " uses programs to organize your church's schedule. "
	dialog.SubText = dialog.SubText & "Programs keep track of your church's events, schedules, and members. </p>"

	If page.Member.IsAdmin Then
		pg.Action = ADDNEW_RECORD
		dialog.LinkList = dialog.LinkList & "<li><a href=""/admin/programs.asp" & pg.UrlParamsToString(True) & """>Create your first program</a></li>"
	End If
	dialog.LinkList = dialog.LinkList & "<li><a href=""/member/contacts.asp"">Email an account administrator</a></li>"
	
	helpLink = "/help/topic.asp?hid=6"
	If (page.Member.IsAdmin = 1) Or (page.Member.IsLeader = 1) Then helpLink = "/help/topic.asp?hid=14"
	
	dialog.LinkList = dialog.LinkList & "<li><a href=""" & helpLink & """ target=""_blank"">Learn more about programs</a></li>"
	
	NoProgramsForClientDialogToString = dialog.ToString()
End Function

</script>