$(document).ready(function(){
	
	// wire up save buttons ..
	$("#save-availability-button a").click(function(){
		$("#form-availability").submit();
	});
	$("#save-program-button a").click(function(){
		$("#form-select-program").submit();
	});
	$("#save-skills-button a").click(function(){
		$("#form-skill").submit();
	});
	
	// jquery code to check/uncheck all
	$("#master").click(function(){
		var master = this
		
		$("[name=SkillIDList]").each(function(){
			this.checked = master.checked;
		});
		
		$("[name=ProgramIDList]").each(function(){
			this.checked = master.checked;
		});
		
		$(".is-available-checkbox").each(function(){
			this.checked = master.checked;
		});
	});
});	
