		<div id="footer">
			<div id="footer-content">
				<ul id="footer-nav">
					<li class="footer-nav-accounts">
						<a href="/tryit.asp">Accounts</a>
						<ul>
							<li><a href="/tryit.asp">Sign-Up (Free!)</a></li>
							<li><a href="/member/login.asp">Login</a></li>
							<li><a href="/overview.asp">Learn More</a></li>
							<li><a href="/client/account.asp">Upgrade</a></li>
						</ul>
					</li>
					<li class="footer-nav-about">
						<a href="/about.asp">About</a>
						<ul>
							<li><a href="/howitworks.asp">How It Works</a></li>
							<li><a href="/screenshots.asp">Screenshots</a></li>
							<li><a href="/support.asp">Contact Us</a></li>
						</ul>
					</li>
					<li class="footer-nav-forums">
						<a href="<%=Application.Value("SUPPORT_FORUM_URL")%>">Community</a>
						<ul>
							<li><a href="<%=Application.Value("SUPPORT_FORUM_URL")%>">Support Forum</a></li>
							<li><a href="/releasenotes.asp">Release Notes</a></li>
						</ul>
					</li>
					<li class="footer-nav-help">
						<a href="/help/help.asp">Help</a>
						<ul>
							<li><a href="/help/topic.asp?hid=14">Getting Started</a></li>
							<li><a href="/help/faq.asp">FAQ</a></li>
							<li><a href="/support.asp">Help by Email</a></li>
						</ul>
					</li>
				</ul>
				<div id="footer-legal">
					<span style="float:right;">
						<a href="/terms.asp">Terms of Use</a>
						|
						<a href="/privacy.asp">Privacy</a>
					</span>
					&copy; <%=Year(Now())%>&nbsp;<%=Application.Value("APPLICATION_NAME")%>
				</div>
			</div>
		</div>
