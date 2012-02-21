<%
Option Explicit

Call Main()

Sub Main
	Dim str
	
	Call wait(.5)
	
	Dim eventAvailabilityId							: eventAvailabilityId = Request.QueryString("eaid")
	Dim skillId										: skillId = Request.QueryString("skid")
	
	Dim eventAvailability							: Set eventAvailability = New cEventAvailability
	Dim programMemberSkill							: Set programMemberSkill = New cProgramMemberSkill
	Dim skill										: Set skill = New cSkill
	Dim scheduleBuild								: Set scheduleBuild = New cScheduleBuild
	
	eventAvailability.eventAvailabilityID = eventAvailabilityID
	Call eventAvailability.Load()
	
	skill.SkillId = skillId
	Call skill.Load()
	
	programMemberSkill.SkillID = skill.SkillID
	Call programMemberSkill.LoadByMemberIdSkillId(eventAvailability.MemberID)
	
	scheduleBuild.EventID = eventAvailability.EventId
	scheduleBuild.ProgramMemberSkillId = programMemberSkill.ProgramMemberSkillId
	Call scheduleBuild.Load()
	
	If Len(scheduleBuild.PublishStatus & "") = 0 Then
		' insert row ..
		scheduleBuild.EventID = eventAvailability.EventId
		scheduleBuild.ProgramMemberSkillId = programMemberSkill.ProgramMemberSkillID
		scheduleBuild.PublishStatus = IS_MARKED_FOR_PUBLISH
		Call scheduleBuild.Add("")
		
		' return marked_for_publish image
		str = "<img src=""/_images/icons/user.png"" title=""Add to event team"" alt=""Published"" />"
		
	ElseIf scheduleBuild.PublishStatus = IS_PUBLISHED Then
		' set to MARK_FOR_UNPUBLISH ..
		scheduleBuild.EventID = eventAvailability.EventId
		scheduleBuild.ProgramMemberSkillId = programMemberSkill.ProgramMemberSkillID
		scheduleBuild.PublishStatus = IS_MARKED_FOR_UNPUBLISH
		Call scheduleBuild.Save("")
		
		' return marked_for_unpublish image
		str = "&nbsp;"
		
		
	ElseIf scheduleBuild.PublishStatus = IS_MARKED_FOR_PUBLISH Then
		' delete row ..
		scheduleBuild.EventID = eventAvailability.EventId
		scheduleBuild.ProgramMemberSkillId = programMemberSkill.ProgramMemberSkillID
		Call scheduleBuild.Delete("")
		
	ElseIf scheduleBuild.PublishStatus = IS_MARKED_FOR_UNPUBLISH Then
		' set to IS_PUBLISHED
		scheduleBuild.EventID = eventAvailability.EventId
		scheduleBuild.ProgramMemberSkillId = programMemberSkill.ProgramMemberSkillID
		scheduleBuild.PublishStatus = IS_PUBLISHED
		Call scheduleBuild.Save("")
		
		' return published image 
		str = "<img src=""/_images/icons/user.png"" title=""Event team"" alt=""Published"" />"
		
	End If

	Response.Write str
End Sub

%>
<!--#INCLUDE VIRTUAL="/_incs/constant/constants.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/script/functions/fn_Wait.asp"-->

<!--#INCLUDE VIRTUAL="/_incs/class/skill_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/event_availability_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/schedule_build_cls.asp"-->
<!--#INCLUDE VIRTUAL="/_incs/class/program_member_skill_cls.asp"-->
