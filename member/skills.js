$(document).ready(function(){
	// check for unsaved changes and bind modal dialog to all links on page ..
	var hasChanges = false
	$("[name=SkillIDList]").click(function(){
		hasChanges = true
	});
	
	$("a").not($("#logoff a, #help a")).click(function(){
		var link = this
		var message = "You have unsaved changes to your skill list. Are you sure you would like to leave this page without saving them?"
		var title = "Your changes haven't been saved!"
		if (hasChanges) {
			$("#confirm-unsaved-changes-dialog").remove();
			$("body").append('<div id="confirm-unsaved-changes-dialog"><\/div>');
		
			$("#confirm-unsaved-changes-dialog").dialog({
				bgiframe: true,
				autoOpen: false,
				modal: true,
				width: 400,
				height: 135,
				overlay: {"background-color": "#000", opacity : 0.5},
				title: title,
				buttons: {
					"Ok (leave page)": function(){
						$(this).dialog("close")
						window.location = $(link).attr("href")
					},
					"Cancel (stay)": function(){$(this).dialog("close")}
				}
			});
			$("#confirm-unsaved-changes-dialog").html('<p>' + message + "<\/p>");
			$("#confirm-unsaved-changes-dialog").dialog("open");
			
			return false
		};
	});
	
	// jquery code to check/uncheck all
	$("#master").click(function(){
		var master = this
		$("[name=SkillIDList]").each(function(){
			this.checked = master.checked;
		})
	})
});
