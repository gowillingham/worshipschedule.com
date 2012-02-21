$(document).ready(function(){

	// wire up master check/uncheck checkbox ..
	$("#master").click(function(){
		var master = this
		$(".file-checkbox").each(function(){
			this.checked = master.checked;
		})
	});
	
	// wire up bulk delete links ..
	$(".bulk-delete-link").click(function(){
		var hasChecked = false
		var message = "You did not select any files to delete. Select at least one file to use the bulk delete feature. "
		var title = "No files were selected!"
		
		$(".file-checkbox").each(function(){
			if (this.checked) {hasChecked = true};
		});					
	
		if (hasChecked) {
			$("#form-file-grid").submit();
			return false;
		}
		else {
		
			$("#no-files-selected-dialog").remove();
			$("body").append('<div id="no-files-selected-dialog"><\/div>');
		
			$("#no-files-selected-dialog").dialog({
				bgiframe: true,
				autoOpen: false,
				modal: true,
				width: 400,
				height: 135,
				overlay: {"background-color": "#000", opacity : 0.5},
				title: title,
				buttons: {"Ok": function(){$(this).dialog("close")}}
			});
			$("#no-files-selected-dialog").html('<p>' + message + "<\/p>");
			$("#no-files-selected-dialog").dialog("open");
		};				
	});
});


