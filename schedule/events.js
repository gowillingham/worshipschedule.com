$(document).ready(function(){
	
	// stripe tables on page
	$(".grid table tr:nth-child(even)").addClass("alt");
	
	// file selector
	$("#file-list").asmSelect();

	// datepicker
	$("#event-date").datepicker({
		showOn: "button", 
		buttonImage: "/_images/icons/calendar_edit.png", 
		buttonText: 'Choose a date ..',
		buttonImageOnly: true 				
	});
	
	// this gives the file inputs in multi-file plugin unique names ..
	$('#form-event').submit(function(){
		var files = $('#form-event input:file');
		var count=0;
		files.attr('name',function(){return this.name+''+(count++);});
	});
	
	// hide/show file input controls
	$("#event-file").hide();
	$("#upload-trigger").click(function(){
		$(this).hide();
		$("#event-file").show();
		return false;
	});
	
	// jquery code to check/uncheck all in delete multiple form ..
	$("#master").click(function(){
		var master = this
		$("[name=remove_event_id_list]").each(function(){
			this.checked = master.checked;
		})
	})
	
	// wire up delete multiple event button ..
	$(".delete-multiple-events").click(function(){
		var hasChecked = false
		var message = "You did not select any events to delete. Select at least one event to use the bulk delete feature. "
		var title = "No events were selected!"
		
		$("[name=remove_event_id_list]").each(function(){
			if (this.checked) {hasChecked = true};
		});
		
		if (hasChecked) {
			$("#form-delete-multiple-events").submit();
			return false;
		}
		else {
			$("#no-events-selected-dialog").remove();
			$("body").append('<div id="no-events-selected-dialog"><\/div>');
			
			$("#no-events-selected-dialog").dialog({
				bgiframe: true,
				autoOpen: false,
				modal: true,
				width: 400,
				height: 135,
				overlay: {"background-color": "#000", opacity : 0.5},
				title: title,
				buttons: {"Ok": function(){$(this).dialog("close")}}
			});
			$("#no-events-selected-dialog").html('<p>' + message + "<\/p>");
			$("#no-events-selected-dialog").dialog("open");
			
			return false
		};				
	});
});

