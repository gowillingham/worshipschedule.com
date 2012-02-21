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
				<h1><a href="/">Home</a> / <a href="/policies.asp">Policies Overview</a> / Privacy</h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<div class="column">
					<div class="feature-box">
						<h3>More Info ..</h3>
						<ul>
							<li><a href="/policies.asp"><strong>Policy Overview</strong></a></li>
							<li><a href="/terms.asp"><strong>Terms</strong></a></li>
						</ul>
					</div>
				</div>
				<h1>Website Privacy and Security</h1>
				<h3>Privacy and Security</h3>
				<p><%=Application.Value("APPLICATION_NAME")%> is a group management and scheduling service from GTD Solutions LLC. 
					<%=Application.Value("APPLICATION_NAME")%> makes your event calendar easily accessible to your volunteer teams. We designed this 
					innovative service with both its usefulness and your privacy in mind. <%=Application.Value("APPLICATION_NAME")%> recognizes the 
					potential sensitivity of confidential information that may be placed in group web pages or 
					transmitted via group messaging and is firmly committed to protecting your privacy. 
					The following discloses our information gathering and dissemination practices for this site.
					If you have any questions about our privacy policy or the terms and conditions, feel free to contact us at
					<a href="mailto:<%=Application.Value("PRIVACY_EMAIL_ADDRESS") %>"><%=Application.Value("PRIVACY_EMAIL_ADDRESS") %></a>.
				</p>
				
				<h3>When Does <%=Application.Value("APPLICATION_NAME")%> Collect Personal Information?</h3>
				<p><strong>Registration</strong>. When you register with <%=Application.Value("APPLICATION_NAME")%>, we will request some personal information 
					including your first name, your last name, and an email address to create your account. We use your 
					registration information to manage your account and to provide you with the <%=Application.Value("APPLICATION_NAME")%> Services. 
					Therefore, when you register as a member of <%=Application.Value("APPLICATION_NAME")%>, we ask for your honest responses to this information, 
					and you represent that your responses are correct and accurate, so that we can serve you more 
					efficiently and effectively.
				</p>
				<p><strong>Member Account Information</strong>. Your <%=Application.Value("APPLICATION_NAME")%> member account includes group membership information, 
					personal and group calendar data, personal and group message history, and personal and group address book 
					information. Your member account data is stored and maintained on <%=Application.Value("APPLICATION_NAME")%> servers in order to provide the 
					service. <%=Application.Value("APPLICATION_NAME")%>'s computers process the information in your account for various purposes, including 
					formatting and displaying the information to you, preventing unsolicited bulk email (spam), 
					backing up your calendar and message history, and other purposes relating to offering you <%=Application.Value("APPLICATION_NAME")%>. 
					Because we keep back-up copies of data for the purposes of recovery from errors or system failure, 
					residual copies of email, short messages, calendar events, and address book information may remain 
					on our systems for some time, even after you have deleted a group, specific messages or events, 
					or after the termination of your account. 
				</p>
				<p><%=Application.Value("APPLICATION_NAME")%> employees do not access the content of any 
					account unless you specifically request them to do so (for example, if you are having technical 
					difficulties accessing your account) or to maintain our system. We may disclose your information
					if necessary to protect our legal rights or if the information relates to actual or threatened 
					harmful conduct or potential threats to the physical safety of any person. Disclosure may be 
					required by law or if we receive legal processes.
				</p>
				<p><strong>Site Usage Information</strong>. We also may collect information about the use of <%=Application.Value("APPLICATION_NAME")%>, 
					such as how much storage you are using, how often you log in and other information related to your 
					registration. Information displayed or clicked on in your <%=Application.Value("APPLICATION_NAME")%> account (including 
					UI elements and other information) is also recorded. We use this information internally to 
					deliver the best possible service to you, such as improving the <%=Application.Value("APPLICATION_NAME")%> user interface. 
					<%=Application.Value("APPLICATION_NAME")%> will not sell, rent or share your personal information, including your 
					member account data, with any third parties for marketing purposes. For each visitor to our web page, 
					our web server does not recognize any information regarding the domain or email address. 
					Your IP address is logged and used by <%=Application.Value("APPLICATION_NAME")%> for measuring usage or eliminating abuse.
				</p>
				<p><strong>Cookies</strong>. In order to provide you with secure and personalized service, <%=Application.Value("APPLICATION_NAME")%> 
					uses cookies to keep and occasionally track information. A cookie is a small text file that is 
					stored on a user's computer for record-keeping purposes. The cookies that are gathered by <%=Application.Value("APPLICATION_NAME")%> 
					are only used by <%=Application.Value("APPLICATION_NAME")%>. Cookies are never shared with third parties. We use both session ID 
					cookies and persistent cookies. We use session cookies to make navigation of our site easy and secure. 
					A session ID cookie expires when you close your browser. A persistent cookie remains on your hard drive 
					for an extended period of time. We set a persistent cookie to store your password if you have enabled the 
					auto-login feature, so you don't have to enter it more than once. You can remove persistent cookies 
					by following the directions provided in your browser's help file.
				</p>
				<p><strong>Links to Other Sites</strong>. Please note that you could be directed to another site while using 
					<%=Application.Value("APPLICATION_NAME")%>. When you link to another site, you should review their privacy policy, as it may 
					be different from ours. <%=Application.Value("APPLICATION_NAME")%> is not responsible, and shall not be liable, for the privacy 
					practices of linked sites or any use such sites may make of any information collected from you. 
					This privacy policy applies solely to information collected by this web site.
				</p>
				<p><strong>Children</strong>. Children should not submit any personal information without the permission
					of their parents or guardians. <%=Application.Value("APPLICATION_NAME")%> encourages all parents or guardians to instruct 
					their children in the safe and responsible use of their personal information while using the Internet. 
				</p>
				
				<h3>When Does <%=Application.Value("APPLICATION_NAME")%> Share My Personal Information?</h3>
				<p><strong>Member Initiated Communication</strong>. When you send an email using <%=Application.Value("APPLICATION_NAME")%> services, 
					<%=Application.Value("APPLICATION_NAME")%> includes your email address and name in the email header. In order to facilitate 
					group communication, <%=Application.Value("APPLICATION_NAME")%> may include mailto links with your email address in the 
					footer of group email messages.
				</p>
				<p><strong>Group Member Information</strong>. <%=Application.Value("APPLICATION_NAME")%> publishes roster and event information for 
					each <%=Application.Value("APPLICATION_NAME")%> group that is only accessible to other members of the specific 
					<%=Application.Value("APPLICATION_NAME")%> account. Your first name, last name, and group attributes may be published on these pages.
				</p>
				<p><strong>Third Parties</strong>. We do not disclose your personally identifying information to 
					third parties unless we believe we are required to do so by law or have a good faith belief 
					that such access, preservation or disclosure is reasonably necessary to (a) satisfy any applicable law, 
					regulation, legal process or governmental request, (b) enforce the <%=Application.Value("APPLICATION_NAME")%> Terms of Use, 
					including investigation of potential violations thereof, (c) detect, prevent, or otherwise address 
					fraud, security or technical issues (including, without limitation, the filtering of spam), 
					(d) respond to user support requests, or (e) protect the rights, property or safety of 
					<%=Application.Value("APPLICATION_NAME")%>, its members and the public.
				</p>
				<p><strong>Business Transfers</strong>. <%=Application.Value("APPLICATION_NAME")%> does not sell, rent or disclose any of your 
					member account information with third parties. In the event <%=Application.Value("APPLICATION_NAME")%> goes through a 
					business transition, such as a merger, acquisition by another company, or sale of all or a 
					portion of its assets, your member account information will likely be among the assets transferred. 
					You will be notified via email and notice on our website for 30 days of any such change in 
					ownership or control of your member account information.

				</p>
				<p><strong>Aggregate Information</strong>. When you access <%=Application.Value("APPLICATION_NAME")%>, our server automatically 
					collects certain information that is not personally identifiable, such as your IP address, 
					pages viewed, and length of time spent on the web site. We collect and use this information, 
					in the aggregate only, to diagnose problems with our servers and to improve <%=Application.Value("APPLICATION_NAME")%>.
				</p>
				
				<h3>Can I Update My Personal Information?</h3>
				<p>You have the ability and the right to update your member account information at
					any time. To do so, log in to <%=Application.Value("APPLICATION_NAME")%> and go to the Member Control Panel. From there you may 
					access or edit your member account information, group membership information, and your group or <%=Application.Value("APPLICATION_NAME")%> 
					membership preferences. You may also remove yourself from any or all <%=Application.Value("APPLICATION_NAME")%> groups.
				</p>
				
				<h3>Security</h3>
				<p><%=Application.Value("APPLICATION_NAME")%> takes precautions to ensure the security of your member account information and 
					strives to keep it accurate. We follow generally accepted industry standards to protect your member 
					information once you have entrusted it to us, both during transmission and once we receive it. 
					We have appropriate security measures in place in our physical facilities to protect against the loss, 
					misuse, or alteration of information that we have collected from you at our site. Files stored on our 
					servers are only accessible by <%=Application.Value("APPLICATION_NAME")%> and through the clickable link displayed within the 
					password protected areas of our website. All files stored are deleted immediately from the <%=Application.Value("APPLICATION_NAME")%>
					servers when you delete them from your account or your account is terminated.
				</p>
				<p>
					However, no method of transmission over the internet or method of electronic storage is 100% secure. 
					Therefore, while we strive to use commercially acceptable means to protect your member account information,
					we cannot guarantee its absolute security.
				</p>
				<h3>Your Consent to Our Privacy Policy</h3>
				<p>By using this website you consent to the collection and use of information about you in 
					the ways described in this privacy policy. <%=Application.Value("APPLICATION_NAME")%> reserves the right to update and 
					amend this policy from time to time to remain consistent with new products, services, 
					processes and Internet privacy legislation. We will post changes to the privacy policy 
					on this page. If we materially change how we use your personal information we will notify 
					you by email or by means of a notice on our home page. The amended policy shall automatically 
					be effective five (5) days after they are initially posted on the web site. Your continued 
					use of <%=Application.Value("APPLICATION_NAME")%> services after the effective date of any posted change constitutes 
					your acceptance of the amended policy as modified by the posted changes. For this reason, 
					we encourage you to review this privacy policy whenever you use the <%=Application.Value("APPLICATION_NAME")%> services. 
					The last date this privacy policy was revised is set forth below.
				</p>
				<h3>Contact Us</h3>
				<p>If you have questions or suggestions regarding our privacy policy, please contact us:</p>
				<ul>
					<li>
						<a href="mailto:<%=Application.Value("PRIVACY_EMAIL_ADDRESS") %>" title="Privacy Email"><%=Application.Value("PRIVACY_EMAIL_ADDRESS") %></a>
					</li>
					<li><%=Application.Value("APPLICATION_NAME")%> Privacy Support, GTD Solutions LLC, 16827 Interlachen Blvd, Lakeville, MN 55044</li>
					<li>
						Send us <a href="mailto:<%=Application.Value("SUPPORT_EMAIL_ADDRESS") %>" title="Support Email">feedback</a> about <%=Application.Value("APPLICATION_NAME")%>
					</li>
					<li>This privacy policy last updated on February 7, 2008</li>
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

