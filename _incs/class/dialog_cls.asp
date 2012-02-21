<script type="text/vbscript" runat="server" language="vbscript">

Class cDialog
	Public headLine
	Public text
	Public subText
	Public linkList
	Public id
	
	Public Function ToString()
		Dim str
		
		str = str & "<div class=""dialog""" & getId() & ">"
		str = str & "<h2>" & headLine & "</h2>"
		str = str & "<div class=""text"">" & text & "</div>"
		str = str & "<p class=""gray-line""></p>"
		str = str & "<ul class=""links"">" & linkList & "</ul>"
		str = str & "<div class=""sub-text"">" & subText & "</div>"
		str = str & "</div>"
		
		ToString = str
	End Function
	
	Private Function getId()
		Dim str
		
		If Len(id) = 0 Then Exit Function
		
		str = str & " id=""" & id & """"
	End Function
End Class

</script>
