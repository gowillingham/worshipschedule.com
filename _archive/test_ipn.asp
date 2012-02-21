<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Call Main()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>Sandbox</title>
	</head>
	<body style="text-align:left;margin:5px;">
		<h1>Sandbox.asp</h1>
		<h2>test ipn post</h2>
			<form method="post" action="/client/ipn.asp">
				<table>
					<tr><td>guid</td>
					<td><input type="text" name="invoice" style="width:300px;"/></td></tr>
					<tr><td>mc_gross</td>
					<td><input type="text" name="mc_gross" value="285" /></td></tr>
					<tr><td>test_ipn</td>
					<td><input type="text" name="test_ipn" value="1" /></td></tr>
					<tr><td>txn_id</td>
					<td><input type="text" name="txn_id" value="paypal transaction id" /></td></tr>
					<tr><td>payment_status</td>
					<td><input type="text" name="payment_status" value="Completed" /></td></tr>
					<tr><td>business</td>
					<td><input type="text" name="business" value="willin_1248207716_biz@lakevillejuniors.com" /></td></tr>
					<tr><td>mc_currency</td>
					<td><input type="text" name="mc_currency" value="USD" /></td></tr>
					<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td>
					<td><input type="submit" name="submit" value="Post to IPN" /></td></tr>
				</table>
			</form>
<%Function Main()


End Function
%>
