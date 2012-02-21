$(document).ready(function(){
	// wire up bulk delete button
	$("#bulk-delete-button a").click(function(){
		var hasChecked = false
		var message = "You did not select any messages to delete. Select at least one message to use the bulk delete feature. "
		var title = "No messages were selected!"
	
		$("[name=EmailIdList]").each(function(){
			if (this.checked) {hasChecked = true};
		});
		
		if (hasChecked) {
			$("#form-bulk-delete").submit();
			return false;
		}
		else {
			$("#no-messages-selected-dialog").remove();
			$("body").append('<div id="no-messages-selected-dialog"><\/div>');
		
			$("#no-messages-selected-dialog").dialog({
				bgiframe: true,
				autoOpen: false,
				modal: true,
				width: 400,
				height: 135,
				overlay: {"background-color": "#000", opacity : 0.5},
				title: title,
				buttons: {"Ok": function(){$(this).dialog("close")}}
			});
			$("#no-messages-selected-dialog").html('<p>' + message + "<\/p>");
			$("#no-messages-selected-dialog").dialog("open");
		};				
	});

	// jquery code to check/uncheck all
	$("#master").click(function(){
		var master = this
		$("[name=EmailIdList]").each(function(){
			this.checked = master.checked;
		})
	})
	
	// set initial state for email details ..
	$("#email-message-collapsed").show()
	$("#email-message-expanded").hide()
	
	// attach click event to showdetails link ..
	$("#expand-details").click(function(){
		$("#email-message-expanded").show()
		$("#email-message-collapsed").hide()
		return false;
	})	
	
	// attach click event to hidedetails link ..	
	$("#collapse-details").click(function(){
		$("#email-message-collapsed").show()
		$("#email-message-expanded").hide()
		return false;
	})		
})
