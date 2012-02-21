<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Dim m_bodyText
Dim m_helpTitle

Sub OnPageLoad(ByRef page)
	page.MessageID = Request.QueryString("msgid")
	
	Set page.Help = New cHelp
	page.Help.HelpID = Request.QueryString("hid")
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
				<h1><a href="/help/help.asp">Help</a> / <%=m_helpTitle %></h1>
			</div>
		</div>
		<div id="container">
			<div class="content">
				<%=m_bodyText %>
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
	
	Call SetHelpTopic(page.Help)
	m_bodyText = page.Help.Text
	m_helpTitle = page.Help.Title
	
	Set page = Nothing
End Sub

Sub SetHelpTopic(help)
	Dim str
	Dim title
	
	Select Case help.HelpID
		Case 1
			help.title = "Add members to my account"
			
			' todo: set up anchors for the three ways to add members ..
			
			str = str & "<div id=""topic"">"
			str = str & "<p class=""dotline"">Welcome! This guide is intended to help you (as an account administrator) get your members into your " & Application.Value("APPLICATION_NAME") & " account. "
			str = str & "You can add members by typing into a <a href=""#anchor-add"">form</a>, importing from a <a href=""#anchor-import"">file</a>, or sending an <a href=""#anchor-invite"">invite</a> by email. "
			str = str & "For additional information on working with your members, please also see the <a href=""/help/faq.asp"">FAQ</a>. </p>"
			
			str = str & "<h1>Adding members to my account</h1>"
			str = str & "<h3 id=""anchor-add"">Adding one member</h3>"
			str = str & "<p>Login to " & Application.Value("APPLICATION_NAME") & " with an administrative account. "
			str = str & "Click the link for <strong>administration</strong> at the top right of any page. </p>"
			str = str & "<p><img src=""/help/_img/globals/click_administration_link.png"" alt=""Click administration"" /></p>"
			
			str = str & "<p>Click the <strong>members</strong> tab from any administrative page. </p>"
			str = str & "<p><img src=""/help/_img/admin_add_members/click_members_tab.png"" alt=""Members tab"" /></p>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " will open the listing of all of your account's members. "
			str = str & "Click the <strong>new member</strong> button from the toolbar in the upper right of the page. </p>"
			str = str & "<p><img src=""/help/_img/admin_add_members/click_new_member_button.png"" alt=""New member"" /></p>"
			
			str = str & "<p>Complete and save the new member form. "
			str = str & "You'll need to provide a first name, last name, and email address for each member you add. </p>"
			str = str & "<p><img src=""/help/_img/admin_add_members/new_member_form.png"" alt=""New member form"" /></p>"
			
			str = str & "<p>To save you time, " & Application.Value("APPLICATION_NAME") & " automatically creates a new account for that member and "
			str = str & "sends a welcome email message with instructions for using " & Application.Value("APPLICATION_NAME") & " and their login credentials. </p>"
			
			str = str & "<h3 id=""anchor-invite"">Inviting members</h3>"
			str = str & "<p>If you know the email address of someone you'd like to add to your " & Application.Value("APPLICATION_NAME") & " account, "
			str = str & "you can send them an email invitation to create their own account. "
			str = str & "From your <strong>members</strong> page, click the <strong>invite</strong> button in the toolbar on the upper right. </p>"
			str = str & "<p><img src=""/help/_img/admin_add_members/click_invite_button.png"" alt=""Invite member"" /></p>"
			
			str = str & "<p>Complete the form and click <strong>send</strong>. "
			str = str & "You can invite multiple members at once by separating each address with a comma. </p>"
			str = str & "<p><img src=""/help/_img/admin_add_members/invite_member_form.png"" alt=""Invite form"" /></p>"
			
			str = str & "<p>After you click send, " & Application.Value("APPLICATION_NAME") & " will send an email message to each invitee's address with a link to click to set up their own accounts. </p>"
			
			str = str & "<h3 id=""anchor-import"">Importing members from a file </h3>"

			str = str & "<div class=""tip-box""><h3>Tip!</h3>"
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " will automatically send a welcome message to your imported members with their new username and password. </p></div>" 
			
			str = str & "<p>Prepare for your import by converting your list of members to a text file. "
			str = str & "Your list needs to be in comma-delimited (sometimes known as CSV - Comma Separated Value) format, and be named 'members.txt'. "
			str = str & "Make sure the first line in your file is <strong>FirstName</strong>, <strong>LastName</strong>, and <strong>Email</strong> separated by commas (the column names), "
			str = str & "and that there are no blank lines at the end of your list of names. </p>"
			str = str & "<p>Each member that should be imported gets their own line that contains their first name, last name, and email address separated by commas. "
			str = str & "Here is a picture of a sample file. </p>"
			str = str & "<p><img src=""/help/_img/admin_add_members/import_file_notepad.png"" alt=""Notepad"" /></p>"
			
			str = str & "<p>From your members page, click <strong>import</strong> in the toolbar on the upper right of the page. </p>"
			str = str & "<p><img src=""/help/_img/admin_add_members/click_import_members.png"" alt=""Import button"" /></p>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " will display the import members form. "
			str = str & "Click <strong>browse</strong> and navigate to the members.txt file on your computer that you created in the step above and click <strong>open</strong>. "
			str = str & "If you would like the members that you are importing to be added to a particular program as they are imported, make sure there is a program selected in the program dropdown. </p>"
			str = str & "<p><img src=""/help/_img/admin_add_members/import_members_form.png"" alt=""Import form"" /></p>"
			
			str = str & "<p>Click <strong>import</strong> and " & Application.Value("APPLICATION_NAME") & " will add the members from the file to your account. </p>"

			str = str & "</div>"
			
		Case 6
			help.title = "Working with programs"
			
			str = str & "<div id=""topic"">"
			str = str & "<p class=""dotline"">Thid guide is intended to help you add a program to your " & Application.Value("APPLICATION_NAME") & " member profile. "
			str = str & "For additional information, please also see the <a href=""/help/faq.asp"">FAQ</a>. </p>"
			
			str = str & "<h1>" & help.Title & "</h1>"
			str = str & "<h3>Add a program to my profile</h3>"
			str = str & "<p>Login to your " & Application.Value("APPLICATION_NAME") & " account. " 
			str = str & "From your member home page, click the <strong>programs</strong> tab. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_programs/click_programs_tab.png"" alt=""Click programs"" /></p>"
			
			str = str & "<p>Your member programs page will open with a list of all programs that already belong to your profile. "
			str = str & "Click the <strong>Add program</strong> button in the toolbar on the upper right. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_programs/click_add_program_button.png"" alt=""Click programs"" /></p>"
			
			str = str & "<p>In the list of available programs, click <strong>Add program</strong> for the program you want to add to your profile. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_programs/click_add_to_profile_button.png"" alt=""Click programs"" /></p>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " will add the program to your profile and show you a list of skills that belong to the program you added. "
			str = str & "Select the skill or skills that should belong to your profile for this program and click <strong>Save</strong>. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_programs/select_skills_and_save.png"" alt=""Save skills"" /></p>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " will return you to your programs page, "
			str = str & "and you will see the program you just added is now in your program listing. "
			str = str & "If there are events already scheduled for the program you just added, "
			str = str & "then " & Application.Value("APPLICATION_NAME") & " will show a reminder for you to update your availablity for those events. </p>"
			
			str = str & "<h3 id=""availability-anchor"">Updating my availability</h3>"

			str = str & "<div class=""tip-box""><h3>Tip!</h3>"
			str = str & "<p>Use your availability to tell your account administrator and " & Application.Value("APPLICATION_NAME") & " when you can be scheduled. </p></div>"
			
			str = str & "<p>From your member home page, click the <strong>availability</strong> tab. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_programs/click_availability_tab.png"" alt=""Click availability"" /></p>"
			
			str = str & "<div class=""tip-box""><h3>Tip!</h3>"
			str = str & "<p>If events have been added to a schedule since your last login, " & Application.Value("APPLICATION_NAME") & " will remind you to update your availability. </p>"
			str = str & "</div>"
			
			str = str & "<p>In the event listing, select the events for which you are available, and then click <strong>save</strong>. "
			str = str & "You may optionally leave a note for with any event for the person who will be scheduling your team. "
			str = str & "Your note will be displayed to them while they are working on the schedule. </p>"
			str = str & "<p>If there are new events that you haven't yet saved, " & Application.Value("APPLICATION_NAME") & " shows them highlighted in red. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_programs/select_events_and_save.png"" alt=""Save events"" /></p>"

			str = str & "</div>"
					
		Case 12
			help.title = "Before You Begin with " & Application.Value("APPLICATION_NAME")
			
			str = str & "<div id=""topic"">"
			str = str & "<p class=""dotline"">Welcome! This guide is intended as a brief introduction to using " & Application.Value("APPLICATION_NAME") & ". "
			str = str & "For additional information, please also see the <a href=""/help/topic.asp?hid=14""><strong>Getting Started Guide</strong></a> or the <a href=""/help/faq.asp""><strong>FAQ</strong></a></p>"
			str = str & "<h1>" & help.Title & "</h1>"
			
			str = str & "<p>Before trying to schedule with " & Application.Value("APPLICATION_NAME") & ", it is important for you to understand how " & Application.Value("APPLICATION_NAME") & "  handles scheduling for you. "
			str = str & "If you are reading this, you may have already created your " & Application.Value("APPLICATION_NAME") & "  account. "
			str = str & "If you haven't, go get a free <a href=""/tryit.asp"" title=""Get a Free Trial Account"">trial account</a> (there is no obligation). "
			str = str & "A short overview of how " & Application.Value("APPLICATION_NAME") & " works follows. </p>"
			
			str = str & "<div class=""subtopic dotline"">"
			str = str & "<img class=""float-image"" src=""/help/_img/admin_before_you_begin/my_account.png"" alt=""My Account"" />"
			str = str & "<h3>Who Are My Members? </h3>"
			str = str & "<p>Your " & Application.Value("APPLICATION_NAME") & "  account can have any number of <strong>Members</strong>. "
			str = str & "Your members are the people in your church or organization that you would like to schedule and manage with " & Application.Value("APPLICATION_NAME") & " . "
			str = str & "You'll set up each of your members with their own login to " & Application.Value("APPLICATION_NAME") & " where they can access the site, view their calendars, and let you know when they are available. "
			str = str & "</p>"
			str = str & "<p>It's easy to get your members into your " & Application.Value("APPLICATION_NAME") & " account (more on the three different ways you can do that later!) "
			str = str & "but whichever way you choose, " & Application.Value("APPLICATION_NAME") & " will take care of sending their new login information (login name and password) to them by email automatically. "
			str = str & "</p>"
			str = str & "</div>"
			
			str = str & "<div class=""subtopic dotline"">"
			str = str & "<img class=""float-image"" src=""/help/_img/admin_before_you_begin/my_account_programs.png"" alt=""My Programs"" />"
			str = str & "<h3>It Starts With Your Programs</h3>"
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " organizes your members inside your account into <strong>Programs</strong>. "
			str = str & "A program is just a group of your members for which you would like to provide scheduling. "
			str = str & "You probably have a number of programs in your organization that you are already scheduling (like a worship team or a prayer team). "
			str = str & "You'll set up a program for any team that you want to manage with " & Application.Value("APPLICATION_NAME") & " . </p>"
			str = str & "<p>Adding or removing members from a program is easy and just takes a few clicks. "
			str = str & "It's not a problem for members to belong to more than one program because " & Application.Value("APPLICATION_NAME") & " can keep track of all of their activities for you. </p>"
			str = str & "</div>"
			
			str = str & "<div class=""subtopic dotline"">"
			str = str & "<img class=""float-image"" src=""/help/_img/admin_before_you_begin/program_skills.png"" alt=""Program Skills"" />"
			str = str & "<h3>Programs Know What You Do</h3>"
			str = str & "<p>When you set up a program for your members in " & Application.Value("APPLICATION_NAME") & " , you will also set up a number of program <strong>Skills</strong> at the same time. "
			str = str & "Skills are just the roles or jobs that the members of your program will be doing when they are scheduled. "
			str = str & "For example, the skills for a worship team might be vocalist, instrumentalist, and worship leader. "
			str = str & "Programs can have as many skills as you like, but each program will need to have at least one. </p>"
			str = str & "<p style=""padding-bottom:20px;"">You can add or remove skills at any time with just a few clicks. </p>"
			str = str & "</div>"
			
			str = str & "<div class=""subtopic dotline"">"
			str = str & "<h3>Programs Know Your Members Too</h3>"
			str = str & "<p>Once a member belongs to a program, that program's skills become part of their profile. "
			str = str & "Since you define a program's skills (and you can add or remove them at any time) you now have a powerful way to organize your events. "
			str = str & "When it comes time for you to start scheduling, " & Application.Value("APPLICATION_NAME")
			str = str & " uses those member skills to intelligently build your event schedules based on what your members can do. </p>"
			str = str & "</div>"
			
			str = str & "<div class=""subtopic"">"
			str = str & "<img class=""float-image"" src=""/help/_img/admin_before_you_begin/my_schedule_events.png"" alt=""My Schedules"" />"
			str = str & "<h3>Where Are My Schedules?</h3>"
			str = str & "<p>Here is where you can really save time and effort. " 
			str = str & Application.Value("APPLICATION_NAME") & " organizes your program's <strong>Events</strong> into a <strong>Schedule</strong>. "
			str = str & "Your programs can have as many schedules or events as needed. "
			str = str & "When you are ready, " & Application.Value("APPLICATION_NAME") & " helps you <strong>Build</strong> your schedule by creating a <strong>Team</strong> of your members for each event. "
			str = str & "</p><p>You build your event team by selecting the members for each skill that will be needed for that event. "
			str = str & "If you have events where you need identical or similar teams, you can even copy event teams from one event to another. </p>"
			str = str & "</div><div class=""subtopic dotline"">"
			str = str & "<img class=""float-image"" src=""/help/_img/admin_before_you_begin/my_event_team.png"" alt=""My Event Team"" />"
			str = str & "<p>After you have set up a team for each event in the schedule, " 
			str = str & Application.Value("APPLICATION_NAME") & "  will <strong>Publish</strong> that schedule to your member's calendar page. "
			str = str & "When your schedule is published, your members login to " & Application.Value("APPLICATION_NAME") & " and view the completed schedule "
			str = str & "(including the details and team for each event) on their calendar. "
			str = str & "The events for which they are scheduled will be highlighted. </p>"
			str = str & "<p>Their calendar is always available to them from any computer with internet access. "
			str = str & "Your members can even download or print a hard copy of any schedule you've created as a PDF document. </p>"
			str = str & "</div>"

			str = str & "<div><h3>Who's Available?</h3>"
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " keeps track of who is available and who has conflicts for the events on your calendar. "
			str = str & "Right from their " & Application.Value("APPLICATION_NAME") & "   home page, "
			str = str & "your members can indicate their <strong>Availability</strong> for your events before you even start assigning members to your event teams. "
			str = str & "That means no more phone tag and digging through stacks of paper or email to figure out who can be scheduled. " 
			str = str & Application.Value("APPLICATION_NAME") & " will show you who is available for your events while you are scheduling! </p>"
			str = str & "</div>"
			str = str & "<div>"
			str = str & "<h3>Email is Easy</h3>"
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " manages your member communication via built-in intelligent <strong> Email Groups</strong>. "
			str = str & "As you work on your schedules, " & Application.Value("APPLICATION_NAME") & " automatically organizes your teams into smart email lists by program, "
			str = str & "schedule, event, skill or who's available and who's not. "
			str = str & "Send an email anytime to any of your members from any computer with internet access. "
			str = str & "You won't need to manage your multiple contact lists by hand any longer as " & Application.Value("APPLICATION_NAME") & "  does it for you with no extra effort on your part.</p>"
			str = str & "</div>"
			str = str & "<div>"
			str = str & "<h3>Admin In Charge</h3>"
			str = str & "<p>Your " & Application.Value("APPLICATION_NAME") & " login is a special type of account called the <strong>Administrator</strong> account. "
			str = str & "An administrator is a member who is able to use the powerful management and scheduling features of " & Application.Value("APPLICATION_NAME") & ". "
			str = str & "Administrators can add or remove members, create and change schedules, or use the email features of your account. "
			str = str & "When any other member logs in to " & Application.Value("APPLICATION_NAME") & " , they only have access to their own information and any schedule information that you have published to their member calendar. " 
			str = str & "As administrator, you can designate any additional members as administrators to also allow them access to the management features of your account. </p>"
			str = str & "</div></div>"
			
		Case 14
		
			help.title = "Getting Started Guide"

			str = str & "<div id=""topic"">"
			str = str & "<p>Welcome! This guide is intended as a brief introduction to using " & Application.Value("APPLICATION_NAME") & " for administrators. "
			str = str & "For additional information, please also see <a href=""/help/topic.asp?hid=12""><strong>Before You Begin</strong></a> or the <a href=""/help/faq.asp""><strong>FAQ</strong></a></p>"
			str = str & "<hr class=""dotted"" />"
			str = str & "<h2>Getting started with " & Application.Value("APPLICATION_NAME") & "</h2>"
			str = str & "<h3>Account administrators</h3>"
			str = str & "<div class=""tip-box""><h3>Tip</h3>"
			str = str & "<p>If you are an administrator you set any of your members as administrators also. </p></div>"
			
			str = str & "<p>Log in to your " & Application.Value("APPLICATION_NAME") & " account. "
			str = str & "Click the link for <strong>Administration</strong> in the upper right of any page "
			str = str & "(only members designated as administrators or program leaders will have this link available). </p>"
			str = str & "<p><img src=""/help/_img/globals/click_administration_link.png"" alt=""Administration link"" /></p>"
			
			str = str & "<p>Your administration home page will open "
			str = str & "(<strong>Admin home</strong> will be showing in the black location bar near the top of the page, "
			str = str & "and the <strong>Overview</strong> tab will be highlighted). "
			str = str & "From here you can add or change events, members and settings for your " & Application.Value("APPLICATION_NAME") & " account. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/admin_overview.png"" alt=""Administration overview"" /></p>"

			str = str & "<hr class=""dotted"" />"
			str = str & "<h2>Working with your members</h2>"
			str = str & "<h3>Your " & Application.Value("APPLICATION_NAME") & " account members</h3>"
			str = str & "<div class=""tip-box""><h3>Tip</h3>"
			str = str & "<p>todo:</p></div>"
			
			str = str & "<p>To open your member page, click on the <strong>Members</strong> tab near the top of any admin page. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/admin_members.png"" alt=""Members"" /></p>"
			str = str & "<p>This is where you can add, remove, or change the members that belong to your " & Application.Value("APPLICATION_NAME") & " account. "
			str = str & "Click on the <strong>Edit</strong> button in the toolbar for a member to open that member's profile. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_edit_member.png"" alt=""Edit member"" /></p>"
			
			str = str & "<p>From here you can access all the account and program settings (programs, skills, availability, etc.) for the member you selected. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/admin_member_account_profile.png"" alt=""Member profile"" /></p>"
			
			str = str & "<hr class=""dotted"" />"
			str = str & "<h2>Working with programs</h2>"
			str = str & "<h3 id=""anchor-add-program"">Create a program</h3>"
			
			str = str & "<p>To create a new program for your account, click on the <strong>Program</strong> tab near the top of any admin page. </p>"
			str = str & "<p><img src=""/help/_img/globals/click_admin_programs_tab.png"" alt=""Programs tab"" /></p>"
			
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " will display your programs listing. </p>"
			str = str & "<p><img src=""/help/_img/globals/admin_programs_list.png"" alt=""Programs"" /></p>"
			
			str = str & "<p>Select <strong>New program</strong> from the toolbar on the upper right of the programs tab. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_new_program.png"" alt=""New program"" /></p>"
			str = str & "<p>Complete the new program form and click <strong>Save</strong> and you'll be returned to your program listing. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/new_program_form.png"" alt=""Form"" /></p>"
			
			str = str & "<h3 id=""anchor-add-skills"">Your program's skills</h3>"
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " keeps track of what your program's members can do (their skills) to help you schedule. "
			str = str & "For that reason, you need to set up at least one skill for each program. "
			str = str & "On your program admin page, click <strong>Skills</strong> in the toolbar for the program you would like to change. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_skills_button.png"" alt=""Skills"" /></p>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " shows you the skill listing for that program. "
			str = str & "Click <strong>Add skill</strong> above the skill listing to add a new skill. "
			str = str & "Repeat the process for all the different skills you will need for the program. "
			str = str & "You will set which members have which skills later. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_add_skill.png"" alt=""Skills"" /></p>"
			
			str = str & "<p>Note: If there are no skills for this program yet, then you will see this message instead of a skill listing. "
			str = str & "In that case, you should click <strong>Create the first skill</strong> from the link list to get started. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/no_skills_for_program_dialog.png"" alt=""Skills"" /></p>"
			
			str = str & "<h3 id=""anchor-add-members"">Your program's members</h3>"
			str = str & "<p>So you have a program, and you know what the members of that program can do (their skills). "
			str = str & "It's time to add some of your account's member to the program. "
			str = str & "On your program admin page, click <strong>Members</strong> in the program's toolbar to tell " & Application.Value("APPLICATION_NAME") & " which of your account members will belong to this program. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_program_members.png"" alt=""Program members"" /></p>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " shows you a listing of all the members you have previously assigned to this program "
			str = str & "(you'll see a message if you haven't assigned any members yet). </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/admin_program_members.png"" alt=""Program members"" /></p>"
			
			str = str & "<p>Click <strong>Select members</strong> from the toolbar in the upper right of the page. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_select_members.png"" alt=""Select members"" /></p>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " shows you a list of your account members next to a list of members that already belong to the program. "
			str = str & "Highlight the members you would like to add or remove from the program and click the arrow buttons to commit your choice. "
			str = str & "You can add or remove members like this at any time. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/admin_select_program_members.png"" alt=""Select members"" /></p>"

			str = str & "<h3>Assigning skills to your members</h3>"
			str = str & "<p>You have a program with skills and members. Now you just need to let " & Application.Value("APPLICATION_NAME") & " know which members have which skills before you begin scheduling. "
			str = str & "On your program admin page, click <strong>Skills</strong> for the program you would like to change. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_skills_button.png"" alt=""Skills"" /></p>"
			
			str = str & "<p>In the upper right of the page, click <strong>Member skills</strong>. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_member_skills.png"" alt=""Member skills"" /></p>"
			str = str & "<p>If no members are shown, be sure a skill is selected in the dropdown on the upper right. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/no_skill_selected.png"" alt=""Member skills"" /></p>"
			str = str & "<p>Highlight the members from the program that should have the skill you selected and click the arrow buttons to save your choices. "
			str = str & "Then repeat the process for each of the program's skills. "
			str = str & "You are free to add or change your member's skills like this at any time. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/admin_member_skills.png"" alt=""Member skills"" /></p>"
			
			str = str & "<hr class=""dotted"" />"
			str = str & "<h2 id=""anchor-add-schedule"">Working with schedules</h2>"
			str = str & "<p>Scheduling with " & Application.Value("APPLICATION_NAME") & " is a three step process. "
			str = str & "First you create a schedule. "
			str = str & "Then you create any number of events that belong to that schedule. "
			str = str & "Finally, you create an event team (what your members will see on their calendar) by assigning program members to those events. </p>"

			str = str & "<h3>Create a schedule</h3>"
			str = str & "<p>Click the <strong>Schedules</strong> tab near the top of any administrator page. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_schedules_tab.png"" alt=""Schedules"" /></p>"
			
			str = str & "<p>This takes you to your admin calendar page, a handy way for you to access the events or event teams for any of the programs you are managing with your " & Application.Value("APPLICATION_NAME") & " account. "
			str = str & "Click on <strong>New schedule</strong> in the toolbar at the upper right of the window. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_new_schedule.png"" alt=""New Schedule"" /></p>"			
			
			str = str & "<p>Provide a name for your schedule and set which program it will belong to. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/new_schedule_form.png"" alt=""New Schedule"" /></p>"
			str = str & "<p>Click <strong>Save</strong> to add the new schedule to your calendar. </p>"
						
			str = str & "<h3>Create an event</h3>"
			str = str & "<p>From your admin schedule page, find the schedule listing to the right of the calendar "
			str = str & "(all the schedules you have created will be listed there). "
			str = str & "Click the <strong>New event</strong> button in the toolbar for the schedule you would like to change. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_new_event.png"" alt=""New Event"" /></p>"
		
			str = str & "<p>Provide a name and date for your event (start/end times are optional) and click <strong>Save</strong>. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/new_event_form.png"" alt=""New Event"" /></p>"

			str = str & "<h3 id=""anchor-add-team"">Schedule an event team</h3>"
			str = str & "<p>From your admin schedule page, navigate to the event you would like to change on your calendar. "
			str = str & "Click <strong>Event team</strong> in the toolbar for that event to move to the <strong>Event team editor</strong>. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_event_team.png"" alt=""Event team"" /></p>"
			
			str = str & "<p>In the team editor, click a skill from the list that you will need on your event team. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/event_team_editor.png"" alt=""Event editor"" /></p>"
			
			str = str & "<p>That skill will open to display the members (available and unavailable) who have the selected skill in their profile. "
			str = str & "Highlight one or more members and click the arrow buttons to add them to the team. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/add_team_member.png"" alt=""Add member"" /></p>"
			str = str & "<p>Repeat for each skill that you need on your event team. </p>"
			
			str = str & "<h3>Publish</h3>"
			
			str = str & "<p>Whenever you make a change to an event team, your changes are not immediately reflected on your member's calendars until you publish that event or schedule. "
			str = str & "You can easily see what members you've added to your team as " & Application.Value("APPLICATION_NAME") & " will display them in green. "
			str = str & "Any team members that have been removed will be displayed in gray. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/unpublished_team_members.png"" alt=""Not published"" /></p>"
			
			str = str & "<p>When you are done assigning members to your event team, you should publish your changes to your member's calendar. "
			str = str & "Click <strong>Publish</strong> in the toolbar for the event (left of the team editor). </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_publish_team.png"" alt=""Publish team"" /></p>"
					
			str = str & "<hr class=""dotted"" />"
			str = str & "<h2 id=""anchor-email-team"">Sending email</h2>"
			
			str = str & "<h3>Composing a message</h3>"
			str = str & "<p>Click on the <strong>Email</strong> tab at the top of any administration page. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_email_tab.png"" alt=""Email"" /></p>"
			
			str = str & "<p>Provide a subject and/or message (you won't be able to send a blank message) "
			str = str & "and fill in the <strong>To</strong> field with any recipients who should receive your message. "
			str = str & "Multiple addresses should be separated by a comma. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/compose_email_message.png"" alt=""Compose email"" /></p>"
			
			str = str & "<p>If you would like " & Application.Value("APPLICATION_NAME") & " to fill in member addresses for you, "
			str = str & "click <strong>Address book</strong> in the toolbar on the upper right. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_address_book.png"" alt=""Address book"" /></p>"
			
			str = str &	"<p>When you are finished, click <strong>Send</strong>. "
			str = str & Application.Value("APPLICATION_NAME") & " will save a copy of your message in your <strong>Sent mail</strong> folder. </p>"
			
			str = str & "<h3>Working with email groups</h3>"
			str = str & "<p>Sometimes you'll want to send email to only a certain group of your account's members. "
			str = str & "You can save your frequently used email addresses into a group to make sending mail to those members faster. </p>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " also includes a number of built-in smart email groups to make that even easier. "
			str = str & "Select <strong>Email groups</strong> from the toolbar on the upper right. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_email_groups.png"" alt=""Email groups"" /></p>"
			
			str = str & "<p>A list of built-in groups (programs, events, etc) is displayed. "
			str = str & "If the built-in group you would like belongs to a specific program, you'll need to set a program in the <strong>Select a program</strong> dropdown list to make them selectable. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/select_a_program_for_groups_dropdown.png"" alt=""Select a program"" /></p>"
			
			str = str & "<p>Check the group or groups you would like to add to your message and click <strong>Save</strong>. </p>"
			str = str & "<p>To create your own custom group, click <strong>My groups</strong> in the toolbar. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_my_groups.png"" alt=""My groups"" /></p>"
			
			
			str = str & "<h3>Email shortcuts</h3>"
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " provides a number of shortcuts throughout your account to quickly email groups of your members. "
			str = str & "Select the <strong>Email</strong> button from any toolbar to go directly to your email page with that group of your members loaded into your message's recipients. </p>"
			str = str & "<p><img src=""/help/_img/admin_getting_started_guide/click_email_shortcut.png"" alt=""Shortcuts"" /></p>"
			
			str = str & "</div>"
			
		Case 16
			help.Title = "Upgrade to a paid " & Application.Value("APPLICATION_NAME") & " subscription"
			
			str = str & "<div id=""topic"">"
			str = str & "<p class=""dotline"">Welcome! This guide is intended to walk your through the steps for upgrading your " & Application.Value("APPLICATION_NAME") & " account to a paid subscription. "
			str = str & "We would be happy to extend your trial account with additional time to evaluate " & Application.Value("APPLICATION_NAME") & ". "
			str = str & "Please request an extension from <a href=""/support.asp"">support</a>. "
			str = str & "Additional information regarding account subscriptions may be found in the <a href=""/help/faq.asp"">FAQ</a>. </p>"
			
			str = str & "<h1>Upgrade your " & Application.Value("APPLICATION_NAME") & " account</h1>"
			str = str & "<h3>View your account information</h3>"
			str = str & "<p>Login to your account as an administrator. "
			str = str & "Click the link for <strong>administration</strong> in the upper right of any page "
			str = str & "(only members designated as administrators or program leaders will have this link available). </p>"
			str = str & "<p><img src=""/help/_img/globals/click_administration_link.png"" alt=""Administration link"" /></p>"

			str = str & "<p>Click the <strong>account</strong> tab at the top of any administration page. </p>"
			str = str & "<p><img src=""/help/_img/admin_upgrade_account/click_account_tab.png"" alt=""Account tab"" /></p>"
			
			str = str & "<div class=""tip-box""><h3>Tip!</h3>"
			str = str & "<p>Your account billing contact person does not need to have a " & Application.Value("APPLICATION_NAME") & " account. </p></div>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " shows you your account history and billing contact information. "
			str = str & "Click the <strong>edit info</strong> button to update your account contact info. "
			str = str & "This should be the person at your church that will be responsible for paying for your account. </p>"
			str = str & "<p><img src=""/help/_img/admin_upgrade_account/account_info_view.png"" alt=""Account"" /></p>"
			
			str = str & "<h3>Upgrade or extend your account</h3>"
			str = str & "<p>To extend your account (or convert your trial account to a full account), click the <strong>extend account</strong> button under your account info. "
			str = str & "Check that your contact information is current. Select a subscription package (6 or 12 months) from the dropdown list and click <strong>save</strong>. </p>"
			str = str & "<p><img src=""/help/_img/admin_upgrade_account/upgrade_account_form.png"" alt=""Account"" /></p>"
			
			str = str & "<div class=""tip-box""><h3>Tip!</h3>"
			str = str & "<p>You may use any major credit card to make payment. "
			str = str & Application.Value("APPLICATION_NAME") & " uses Paypal to manage online payments. </p></div>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " displays a confirmation page. "
			str = str & "If all of your information is correct, click the <strong>pay now</strong> button at the bottom right of the form. </p>"
			str = str & "<p><img src=""/help/_img/admin_upgrade_account/click_pay_now_button.png"" alt=""Pay now"" /></p>"
			
			str = str & "<p>If you wish to use a credit card to make payment, click <strong>continue</strong> on the form where it says <strong>Don't have a PayPal account?</strong> "
			str = str & "Otherwise you may login and use a PayPal account to save time. </p>"
			str = str & "<p><img src=""/help/_img/admin_upgrade_account/paypal_gateway_form.png"" alt=""Paypal"" /></p>"
			
			str = str & "<p>After providing the requested payment information, you will be returned to your " & Application.Value("APPLICATION_NAME") & " accounts page. </p>"
			
			str = str & "<h3>Canceling your account</h3>"
			str = str & "<p>You may cancel your " & Application.Value("APPLICATION_NAME") & " account at any time. "
			str = str & "Click the <strong>cancel account</strong> button on your accounts page. "
			str = str & "We'll disable your account immediately, and refund any pro-rated portion of your subscription that is remaining. </p>"
			 
		
			
			str = str & "</div>"
			
		Case 17
			help.Title = "Work with my calendar"
			
			
			str = str & "<div id=""topic"">"
			str = str & "<p class=""dotline"">Welcome! This guide will help you work with your member calendar. "
			str = str & "For additional help, see the <a href=""/help/faq.asp"">FAQ</a>. </p>"
			
			str = str & "<h1>Working with calendars</h1>"
			str = str & "<h3>The main calendar</h3>"
			str = str & "<p>Login to your " & Application.Value("APPLICATION_NAME") & " account. "
			str = str & "Click the <strong>calendar</strong> tab from any page. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_calendar_tab.png"" alt=""Calendar tab"" /></p>"
			
			str = str & "<div class=""tip-box""><h3>Tip!</h3>"
			str = str & "<p>Events on your calendar for which you are scheduled are highlighted in blue. </p></div>"
			
			str = str & "<p>Your member calendar is displayed for the current month. "
			str = str & "You can use the month dropdown list or the previous/next arrow buttons to the above right of the calendar to navigate from month to month. </p>"
			str = str & "<p>From here you can view all the events for your programs, download an event to your desktop, or receive the events by email. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/full_calendar_view.png"" alt=""Member calendar"" /></p>"
			
			str = str & "<h3>Working with events</h3>"
			str = str & "<p>In your calendar, click the <strong>event details</strong> button in the toolbar for any event. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_event_details_button.png"" alt=""Details button"" /></p>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " will show you the details for that event. "
			str = str & "From here, you'll see the event start and end times (if the event has them). "
			str = str & "If there is an event team scheduled for this event, the team will be listed in the summary. "
			str = str & "Any files that have been saved for the event are available from this view also. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/event_detail_view.png"" alt=""Event details"" /></p>"
			
			str = str & "<p>To set your availability for this event without visiting your availability page, "
			str = str & "click the <strong>availability</strong> button in the event toolbar. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_availability_button.png"" alt=""Availability"" /></p>"
			
			str = str & "<p>If you have Outlook or another calendaring application installed on your computer, "
			str = str & Application.Value("APPLICATION_NAME") & " can convert the event to an iCal file for you. "
			str = str & "Clicking <strong>iCal</strong> in the toolbar saves the event to your computer. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_ical_button.png"" alt=""iCal"" /></p>"
			
			str = str & "<p>Click <strong>email</strong> to have the event details sent to your email address. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_email_event_button.png"" alt=""Email event"" /></p>"
			
			str = str & "<h3>The event list</h3>"
			str = str & "<p>To view all of your events in single list, click <strong>event list</strong> in the toolbar at the upper right of the calendar page. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_event_list_button.png"" alt=""Event list"" /></p>"
			
			str = str & "<div class=""tip-box""><h3>Tip!</h3>"
			str = str & "<p>Events in the list for which you are scheduled will be highlighted in blue. </p></div>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " shows you a sortable, filterable listing of all the events for the programs that are active in your profile. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/event_list_view.png"" alt=""Event list"" /></p>"
			
			str = str & "<p>Click the <strong>my events</strong> checkbox in the toolbar for a listing of the events for which you are scheduled. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_my_events_checkbox.png"" alt=""My events"" /></p>"
			
			str = str & "<p>Sort your events by date, program, etc. by picking an option in the <strong>sort by</strong> list in the toolbar. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_sort_by_dropdown.png"" alt=""Sort events"" /></p>"
			
			str = str & "<p>Choose a program from the <strong>program</strong> dropdown list in the toolbar to filter the event list by that program. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_program_dropdown.png"" alt=""Program filter"" /></p>"
			
			str = str & "<h3>Working with your team view</h3>"
			str = str & "<p>Sometimes you might wish to see the event teams for all of a programs events at one time. "
			str = str & Application.Value("APPLICATION_NAME") & " calls this your <strong>team view</strong>. </p>"
			str = str & "<p>To access this view, make sure a program is selected in the <strong>select a program</strong> dropdown list in the toolbar. "
			str = str & "Then click the <strong>team view</strong> button in the toolbar. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/click_team_view_button.png"" alt=""Team button"" /></p>"
			
			str = str & "<div class=""tip-box""><h3>Tip!</h3>"
			str = str & "<p>If you belong to an event team, " & Application.Value("APPLICATION_NAME") & " will highlight your name. </p></div>"
			
			str = str & "<p>" & Application.Value("APPLICATION_NAME") & " shows you a grid view of the event team members and what they will be doing for each event. </p>"
			str = str & "<p><img src=""/help/_img/member_work_with_calendar/event_team_view.png"" alt=""Event teams"" /></p>"
			
			str = str & "</div>"
		Case Else
			help.Title = "Help unavailable"
			
			str = str & "<h3>No help topic available at this time. </h3>"
			str = str & "<p>Sorry, but the online help for " & Application.Value("APPLICATION_NAME") & " is currently being upgraded to version 2.0. </p>"
			str = str & "<p>We'd still like to help! Send your questions about " & Application.Value("APPLICATION_NAME") & " to <a href=""/support.asp"">support</a> "
			str = str & "and we'll get back to you as quickly as we can (usually within 24 hours). </p>"

	End Select
	
	help.Text = str
End Sub

%>
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_message.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/script/application/app_checkaccess.asp"-->
<%
Class cPage
	Public MessageID

	' obj
	Public Help
		
	Public Function UrlParamsToString(isEncoded)
		Dim str, amp
		
		amp = "&"
		If isEncoded Then amp = "&amp;"
		
		If Len(page.HelpID) > 0 Then str = str & "hid=" & page.HelpID & amp
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
		Set c.Help = Help
		
		Set Clone = c
	End Function
End Class

Class cHelp
	Public HelpID
	Public Text
	Public Title
End Class
%>

