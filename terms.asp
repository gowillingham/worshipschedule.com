<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Sub OnPageLoad(ByRef page)
	page.MessageID = Request.QueryString("msgid")
End Sub

Call Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<link rel="stylesheet" type="text/css" href="/_incs/style/outside.css" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title><%=Application.Value("APPLICATION_NAME") & " - Simple Web Scheduling for Worship Teams" %></title>
	</head>
	<body>
		<%=GetServerIndicator("DIV") %>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topbar.asp"-->
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_topnav.asp"-->
		<div id="pageheader">
			<div class="content">
				<h1><a href="/">Home</a> / <a href="/policies.asp">Policies Overview</a> / Terms of Use</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="column">
					<div class="feature-box">
						<h3>More Info ..</h3>
						<ul>
							<li><a href="/policies.asp"><strong>Policy Overview</strong></a></li>
							<li><a href="/privacy.asp"><strong>Privacy</strong></a></li>
						</ul>
					</div>
				</div>
				<h1><%=Application.Value("APPLICATION_NAME")%> Terms of Use</h1>
				<h3>Terms of Use</h3>
				<p>By selecting "I Agree" or accessing and using the <%=Application.Value("APPLICATION_NAME")%> 
					web site, you agree that you have read, understand and accept the terms and conditions below for all use 
					of the <%=Application.Value("APPLICATION_NAME")%> services and website (collectively, the "Services"). If you do not agree to these 
					terms and conditions, do not use any of the <%=Application.Value("APPLICATION_NAME")%> Services. <%=Application.Value("APPLICATION_NAME")%> is 
					operated by GTD Solutions LLC.
				</p>
				<h3>These Terms are Subject to Change</h3>
				<p><%=Application.Value("APPLICATION_NAME")%> reserves the right to update and change the Terms of Service from time to time without 
					notice. Any new features that augment or enhance the current Service, including the release of new 
					tools and resources, shall be subject to the Terms of Service. Continued use of the Service after 
					any such changes shall constitute your consent to such changes. You can review the most current 
					version of the Terms of Service at any time at: http://<%=Request.ServerVariables("SERVER_NAME") %>/terms.asp.
				</p>
				<h3>Proprietary Rights</h3>
				<p>All right, title, and interest in and to the Services are and will remain the exclusive property of 
					<%=Application.Value("APPLICATION_NAME")%> and its licensors. The Services are protected by copyright, trademark, and other laws 
					of both the United States and foreign countries. Except as expressly permitted in these terms of use, 
					you may not reproduce, modify, or prepare derivative works based upon, distribute, sell, transfer, 
					publicly display, publicly perform, transmit, or otherwise use the Services.
				</p>
				<h3>Trademarks</h3>
				<p><%=Application.Value("APPLICATION_NAME")%>, the <%=Application.Value("APPLICATION_NAME")%> logo, and all other <%=Application.Value("APPLICATION_NAME")%> trademarks, service marks, 
					product names, and trade names of <%=Application.Value("APPLICATION_NAME")%> appearing on the Services are owned by <%=Application.Value("APPLICATION_NAME")%>. 
					All other trademarks, service marks, product names, and logos appearing on the Services are the 
					property of their respective owners. You may not use or display any trademark, service mark, product name, 
					trade name, or logo appearing on the Services without the owner's prior written consent.
				</p>
				<h3>Account Terms</h3>
				<ol>
					<li>You must be at least 13 years old to be eligible to use the Services. However, if you are younger 
						than 13 years old, you may use the Services if and only if you have your parents' or guardians' 
						prior permission. By selecting "I Agree" you are representing that you are at least 
						13 years old or are under 13 years old and have parent or guardian permission to register for the Services.
					</li>
					<li>You must be human. Accounts registered by "bots" or other automated methods are not permitted.</li>
					<li>You must provide a valid email address, and any other information requested in order to complete the 
						sign-up process.
					</li>
					<li>You are responsible for maintaining the security of your account and password. <%=Application.Value("APPLICATION_NAME")%>
						cannot and will not be liable for any loss or damage from your failure to comply with this security 
						obligation.
					</li>
					<li>You are responsible for all content posted and activity that occurs under your account (even when 
						content is posted by others who are sharing your pages).
					</li>
					<li>You may not use the Service for any illegal or unauthorized purpose. You must not, in the use of 
						the Service, violate any laws in your jurisdiction (including but not limited to copyright laws).
					</li>
					<li>You have sole responsibility for all user files that you store on <%=Application.Value("APPLICATION_NAME")%> servers through use of the 
						Services. You acknowledge and agree that <%=Application.Value("APPLICATION_NAME")%> will not be responsible for any failure of the 
						Services to store a user file, for the deletion of a user file stored on the Services, or for the 
						corruption of or loss or any data, information or content contained in a user file.
					</li>
					<li>You agree not to upload or transmit as part of a user file or otherwise any data, text, graphics, 
						content, or material that: (i) is false or misleading; (ii) is defamatory; (iii) invades another's 
						privacy; (iv) is obscene, pornographic, or offensive; (v) promotes bigotry, racism, hatred, or 
						harm against any individual or group; (vi) infringes another's rights, including any intellectual 
						property rights; or (vii) violates, or encourages any conduct that would violate, 
						any applicable law or regulation or would give rise to civil liability.</li>
					<li>You agree not to access, tamper with, or use any non-public areas of the Services or <%=Application.Value("APPLICATION_NAME")%>'s 
						computer systems or the technical delivery systems of <%=Application.Value("APPLICATION_NAME")%>'s providers.
					</li>
					<li>You agree not to attempt to probe, scan, or test the vulnerability of the Services or any 
						related system or network or breach any security or authentication measures.
					</li>
					<li>You agree not to attempt to decipher, decompile, disassemble, or reverse engineer any of the 
						software used to provide the Services.
					</li>
					<li>You agree not to interfere with, or attempt to interfere with, the access of any user, host or 
						network, including, without limitation, sending a virus, overloading, flooding, spamming, 
						or mail-bombing the Services.
					</li>
					<li>You agree not to impersonate or misrepresent your affiliation with any person or entity.</li>
				</ol>
				<p>Violation of any of these agreements will result in the termination of your account. While <%=Application.Value("APPLICATION_NAME")%> 
					prohibits such conduct and content on the Service, you understand and agree that <%=Application.Value("APPLICATION_NAME")%> 
					cannot be responsible for the content posted on the Service and you nonetheless may be exposed to such 
					materials. You agree to use the Service at your own risk.
				</p>
				<h3>Termination</h3>
				<p><%=Application.Value("APPLICATION_NAME")%>, in its sole discretion, has the right to suspend or terminate your account and 
					refuse any and all current or future use of the Service, or any other GTD Solutions LLC service, for 
					any reason at any time. Such termination of the Service will result in the deactivation or 
					deletion of your account or your access to your account, and the forfeiture and relinquishment of 
					all content in your account. <%=Application.Value("APPLICATION_NAME")%> reserves the right to refuse service to anyone for 
					any reason at any time. All of your content will be immediately deleted from the Service upon cancellation.
				</p>
				<h3>Modifications to the Service</h3>
				<ol>

					<li><%=Application.Value("APPLICATION_NAME")%> reserves the right at any time and from time to time to modify or discontinue, 
						temporarily or permanently, the Service (or any part thereof) with or without notice.
					</li>
					<li>Prices of all Services, including but not limited to monthly subscription plan fees to the Service, 
						are subject to change upon 30 days notice from us. Such notice may be provided at any time by 
						posting the changes to the <%=Application.Value("APPLICATION_NAME")%> site (<%=Request.ServerVariables("SERVER_NAME") %>) or the Service itself.
					</li>
					<li><%=Application.Value("APPLICATION_NAME")%> shall not be liable to you or to any third party for any modification, price 
						change, suspension or discontinuance of the Service.
					</li>
				</ol>
				<h3>Copyright and Content Ownership</h3>
				<p><%=Application.Value("APPLICATION_NAME")%> claims no intellectual property rights over the material you provide to the Service. 
					Your profile and materials uploaded remain yours. <%=Application.Value("APPLICATION_NAME")%> does not pre-screen content, but 
					<%=Application.Value("APPLICATION_NAME")%> and its designee have the right (but not the obligation) in their sole discretion 
					to refuse or remove any content that is available via the Service.
				</p>
				<h3>General Conditions</h3>
				<ol>
					<li>Your use of the Service is at your sole risk. The service is provided on an "as is" and 
						"as available" basis.
					</li>
					<li>Technical support is only provided via email (we try to respond within 12 hours).</li>
					<li>You must not modify, adapt or hack the Service or modify another website so as to falsely 
						imply that it is associated with the Service, <%=Application.Value("APPLICATION_NAME")%>, or any other <%=Application.Value("APPLICATION_NAME")%> service.
					</li>
					<li>You agree not to reproduce, duplicate, copy, sell, resell or exploit any portion of the 
						Service, use of the Service, or access to the Service without the express written permission 
						by <%=Application.Value("APPLICATION_NAME")%>.
					</li>
					<li>We may, but have no obligation to, remove content and accounts containing content that we 
						determine in our sole discretion are unlawful, offensive, threatening, libelous, defamatory, 
						pornographic, obscene or otherwise objectionable or violates any party’s intellectual 
						property or these terms.
					</li>
					<li>Verbal, physical, written or other abuse (including threats of abuse or retribution) of any 
						<%=Application.Value("APPLICATION_NAME")%> customer, employee, member, or officer will result in immediate account termination.
					</li>
					<li>You understand that the technical processing and transmission of the Service, including your content, 
						may be transferred unencrypted and involve (i) transmissions over various networks; and (ii) changes 
						to conform and adapt to technical requirements of connecting networks or devices.
					</li>
					<li>You must not upload, post, host, or transmit unsolicited email or “spam” messages.</li>
					<li>You must not transmit any worms or viruses or any code of a destructive nature.</li>
					<li>If your bandwidth usage exceeds significantly exceeds the average bandwidth usage (as determined 
						solely by <%=Application.Value("APPLICATION_NAME")%>) of other <%=Application.Value("APPLICATION_NAME")%> customers, we reserve the right to immediately 
						disable your account or throttle your file hosting until you can reduce your bandwidth consumption.
					</li>
					<li><%=Application.Value("APPLICATION_NAME")%> does not warrant that (i) the service will meet your specific requirements, 
						(ii) the service will be uninterrupted, timely, secure, or error-free, (iii) the results 
						that may be obtained from the use of the service will be accurate or reliable, (iv) the 
						quality of any products, services, information, or other material purchased or obtained by you 
						through the service will meet your expectations, and (v) any errors in the Service will be 
						corrected.
					</li>
					<li>You expressly understand and agree that <%=Application.Value("APPLICATION_NAME")%> shall not be liable for any direct, indirect, 
						incidental, special, consequential or exemplary damages, including but not limited to, damages 
						for loss of profits, goodwill, use, data or other intangible losses (even if <%=Application.Value("APPLICATION_NAME")%> 
						has been advised of the possibility of such damages), resulting from: (i) the use or the inability 
						to use the service; (ii) the cost of procurement of substitute goods and services resulting from 
						any goods, data, information or services purchased or obtained or messages received or 
						transactions entered into through or from the service; (iii) unauthorized access to or alteration 
						of your transmissions or data; (iv) statements or conduct of any third party on the service; 
						(v) termination of your account; or (vi) any other matter relating to the service.
					</li>
					<li>The failure of <%=Application.Value("APPLICATION_NAME")%> to exercise or enforce any right or provision of the Terms of Use 
						shall not constitute a waiver of such right or provision. The Terms of Use constitutes the entire 
						agreement between you and <%=Application.Value("APPLICATION_NAME")%> and govern your use of the Service, superseding any 
						prior agreements between you and <%=Application.Value("APPLICATION_NAME")%> (including, but not limited to, any prior 
						versions of the Terms of Use).
					</li>
				</ol>
				<h3>Contact Us</h3>
				<p>If you have questions or suggestions regarding our Terms of Use please contact us:</p>
				<ul>
					<li>
						<a href="mailto:<%=Application.Value("TERMS_OF_SERVICE_EMAIL_ADDRESS") %>" title="Terms of Service Inquiry"><%=Application.Value("TERMS_OF_SERVICE_EMAIL_ADDRESS") %></a>
					</li>
					<li><%=Application.Value("APPLICATION_NAME")%> Terms of Use, GTD Solutions LLC, 16827 Interlachen Blvd, Lakeville, MN 55044</li>
					<li>
						Send us <a href="mailto:<%=Application.Value("SUPPORT_EMAIL_ADDRESS") %>" title="Support Email">feedback</a> about <%=Application.Value("APPLICATION_NAME")%>
					</li>
					<li>These terms last updated on February 7, 2008</li>
				</ul>
				</div>
		</div>
		<!--#INCLUDE VIRTUAL="/_incs/navigation/outside_footer.asp"-->
	</body>
</html>
<%
Sub Main()
	Dim str, rv
	Dim page		: Set page = New cPage
	
	Call OnPageLoad(page)

	page.MessageID = ""
	
	Set page = Nothing
End Sub

%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<%
Class cPage
	Public MessageID
	
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(MessageID) > 0 Then str = str & "msgid=" & MessageID & amp
		
		If Len(str) > 0 Then 
			str = Left(str, Len(str) - Len(amp))
		Else
			' qstring needs at least one param in case more params are appended ..
			str = str & "noparm=true"
		End If
		str = "?" & str
		
		UrlParamsToString = str
	End Function
	
	Public Function Clone()
		Dim c
		Set c = New cPage
		
		c.MessageID = MessageID
		
		Set Clone = c
	End Function
End Class
%>

