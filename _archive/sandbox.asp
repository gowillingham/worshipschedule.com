<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
	<title>SandBox</title>
		<script src="http://www.google.com/jsapi" type="text/javascript" language="javascript"></script>
		<script language="javascript" type="text/javascript">
			google.load("jquery", "1.2.6");
		</script>
		<script language="javascript" type="text/javascript">
		
			$(document).ready(function(){
				$("#master").click(function(){
					var master = this
					$("[name=slave]").each(function(){
						this.checked = master.checked;
					})
				})			
			})
			
		</script>
</head>
	<body>
		<%Call Main() %>
		<ul>
			<li><input type="checkbox" id="master" value="theValue" />Master</li>
		</ul>
		<ul>
			<li><input type="checkbox" name="slave" value="" />slave</li>
			<li><input type="checkbox" name="slave" value="" />slave</li>
			<li><input type="checkbox" name="slave" value="" />slave</li>
			<li><input type="checkbox" name="slave" value="" />slave</li>
			<li><input type="checkbox" name="slave" value="" />slave</li>
		</ul>
	</body>
</html>
<%
Sub Main()


End Sub
%>


