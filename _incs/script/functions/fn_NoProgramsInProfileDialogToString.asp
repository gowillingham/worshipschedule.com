<script runat="server" type="text/vbscript" language="vbscript">

Function NoProgramsInProfileDialogToString(page)
	Dim dialog				: Set dialog = New cDialog
	Dim pg					: Set pg = page.Clone()
	
	dialog.Headline = "Ok, let's get started here!"
	
	dialog.Text = dialog.Text & "<p>It looks like you haven't added any " & html(page.Client.NameClient) & " programs to your account. "
	dialog.Text = dialog.Text & "Click <strong>Add my first program</strong> to see a list of programs that are available for you. </p>"

	dialog.SubText = dialog.SubText & "<p>" & Application.Value("APPLICATION_NAME") & " uses programs to organize events, schedules and members for the " & html(page.Client.NameClient) & " account. </p>"
	dialog.SubText = dialog.SubText & "<p>Once you've added a program to your account, this page will show you a list of the programs you belong to. "
	dialog.SubText = dialog.SubText & "You use this page to add, remove, or change the programs that belong to your account. </p>"

	pg.Action = SHOW_AVAILABLE_PROGRAMS
	dialog.LinkList = dialog.LinkList & "<li><a href=""/member/programs.asp" & pg.UrlParamsToString(True) & """>Add my first program</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/member/contacts.asp"">Email an account administrator</a></li>"
	dialog.LinkList = dialog.LinkList & "<li><a href=""/help/topic.asp?hid=6"" target=""_blank"">Learn more about programs</a></li>"
	
	NoProgramsInProfileDialogToString = dialog.ToString()
End Function

</script>