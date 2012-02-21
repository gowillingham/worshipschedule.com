function doAvailableView(){
	$.cookie("available-view", "on");
	$.cookie("publish-view", null);

	$("#master-team-view table li.unpublish").removeClass("style-for-unpublish")
	$("#master-team-view table li.publish").removeClass("style-for-publish")
	
	$("#master-team-view table li.unknown-available").addClass("style-for-unknown-available")
	$("#master-team-view table li.not-available").addClass("style-for-not-available")
	$("#master-team-view table li.available").addClass("style-for-available")
	
	$(".event-item li.unknown-available").addClass("style-for-unknown-available")
	$(".event-item li.not-available").addClass("style-for-not-available")
	$(".event-item li.available").addClass("style-for-available")
}

function doPublishView(){
	$.cookie("publish-view", "on");
	$.cookie("available-view", null);
	
	$("#master-team-view table li.unknown-available").removeClass("style-for-unknown-available")
	$("#master-team-view table li.not-available").removeClass("style-for-not-available")
	$("#master-team-view table li.available").removeClass("style-for-available")
	
	$("#master-team-view table li.unpublish").addClass("style-for-unpublish")
	$("#master-team-view table li.publish").addClass("style-for-publish")
}

function doDefaultView(){
	$.cookie("publish-view", null);
	$.cookie("available-view", null);
	
	$("#master-team-view table li.unknown-available").removeClass("style-for-unknown-available")
	$("#master-team-view table li.not-available").removeClass("style-for-not-available")
	$("#master-team-view table li.available").removeClass("style-for-available")
	$("#master-team-view table li.unpublish").removeClass("style-for-unpublish")
	$("#master-team-view table li.publish").removeClass("style-for-publish")
	
	$(".event-item li.unknown-available").removeClass("style-for-unknown-available")
	$(".event-item li.not-available").removeClass("style-for-not-available")
	$(".event-item li.available").removeClass("style-for-available")
}

function refreshAvailableStyles(){
	if ($("#see-available-button").hasClass("selected")){
		doAvailableView();
	}
	else {
		doDefaultView();
	}
};

function wireUpAccordion(){
	// wire up accordian control ..
	$("#team-editor").accordion({
		clearStyle: true,
		active: false,
		header: ".head",
		alwaysOpen: false						
	});
	
	$("#team-editor h5").each(function(){
		var scheduleItem 
		var height
		
		// wire up click event to h5
		$(this).click(function(){
			// get a reference to this accordion header
			var header = this
			var parent = $(header).parent();
		
			// reset all indicators to right ..
			$("#team-editor h5").css("background-image", "url('/_images/icons/delta_right.gif')");
		
			// change the clicked bar indicator to point down ..
			$(this).css("background-image", "url('/_images/icons/delta_down.gif')");
			
			var form = $(header).next("form");
			
			var options = {
				dataType: "json",
				beforeSubmit: function(){
					scheduleItem = $(".event-item");
					height = scheduleItem.height();
					
					$(scheduleItem).css("height", height);
					$(scheduleItem).children("h5, h6, .toolbar, ul, .bottom, .alert").hide()
					$(scheduleItem).css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat center center");
					
					$(form).children("table").css("height", $(form).children("table").height())
					$(form).children("table").children().hide();
					$(form).children("table").css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat center center")
				},
				success: function(responseText){
					
					// remove and replace the select options ..
					$("select", form).children().remove();
					$("select.scheduled-members", form).append(responseText.scheduled);
					$("select.available-members", form).append(responseText.available);
					$("select.not-available-members", form).append(responseText.notAvailable);

					// refresh the event team listing ..
					$(".event-item").replaceWith(responseText.eventItem)
					
					// replace progress indicator ..
					$(header).next("div").replaceWith($(form).ajaxForm(options));
					$(form).children("table").css("background-image", "none")
					$(form).children("table").children().show();
					$(form).children("table").css("height", "auto")
					
					refreshAvailableStyles();
				},
				error: function(responseText){
					console.log(responseText)
					$(header).next("div").replaceWith('<div id="ajax-error">There was a problem. Your changes were not saved. <\/div>');
				}
			}
			$(form).ajaxForm(options)
		});
	});
}

$(document).ready(function(){
	// set initial state for view buttons ..
	if ($.cookie("publish-view") == "on") {
		$("#see-published-button").addClass("selected")
		doPublishView();
	}
	else if ($.cookie("available-view") == "on") {
		$("#see-available-button").addClass("selected")
		doAvailableView();
	}

	// wire event-item links ..
	$(".team-list-for-skill a").livequery("click", function(){
		var mid = $(this).attr("class").replace("mid-", "")
		var parentItem = $(this).parent().parent().parent().parent().parent();
		var scid = $(".schedule-dropdown", parentItem).val();
		
		// get scheduleId from parent event-item and update widget ..
		$("#schedule-id", "#availability-widget").val(scid);

		$("#member-dropdown").val(mid);
		updateMemberEventWidget();
		
		return false
	});
	
	// set initial show/hide state for skill rows, skill row checkboxes
	$(".empty,.not-empty", "#master-team-view").each(function(){
		var id = $(this).attr("class").split(" ").slice(-1)[0].replace("skill-row-id-", "")
	
		if ($(this).hasClass("empty")){
			$(this).hide()
			$("#skill-checkbox-id-" + id).attr("checked", false)
		}
		else if ($(this).hasClass("not-empty")){
			$("#skill-checkbox-id-" + id).attr("checked", true)
		}
	});
	
	// over-ride initial state if there are zero not-empty rows ..
	if ($(".not-empty").size() == 0){
		$("tr.empty").show()
		$("#hide-skills input").attr("checked", true)
		$("#hide-empty-skills").attr("checked", true)
	}
	
	// wire up hide-empty-skills checkbox ..
	$("#hide-empty-skills").change(function(){
		if ($(this).attr("checked") == true) {
			$(".empty").each(function(){
				var id = $(this).attr("class").split(" ").slice(-1)[0].replace("skill-row-id-", "")
				
				$(this).show();
				$("#skill-checkbox-id-" + id).attr("checked", true)
			});
		}
		else {
			$(".empty").each(function(){
				var id = $(this).attr("class").split(" ").slice(-1)[0].replace("skill-row-id-", "")
				
				$(this).hide();
				$("#skill-checkbox-id-" + id).attr("checked", false)
			});
		}
	});
	
	// wire up show past events link in no events dialog ..
	$("#show-past-events-link").click(function(){
		$("#form-show-past-events").submit();
		return false;
	});
	
	// wire up click event for skillName filter checkboxes
	$("#hide-skills input").each(function(){
		$(this).change(function(){
			var id = $(this).attr("id").replace("skill-checkbox-id-", "")
			
			if ($(this).attr("checked") == true){
				$(".skill-row-id-" + id).show();
			}
			else {
				$(".skill-row-id-" + id).hide();
			};
		});
	});

	// wire view buttons ..
	$("#view-button-container a").click(function(){
	
		// set button state ..
		if($(this).hasClass("selected")){
			$(this).removeClass("selected");
		}
		else {
			$("#view-button-container a").removeClass("selected");
			$(this).addClass("selected");
		}
		
		if ($("#see-published-button").hasClass("selected")){
			doPublishView();
		}
		else if ($("#see-available-button").hasClass("selected")){
			doAvailableView();
		}
		else {
			doDefaultView();
		}
		
		return false
	});

	// wire up availability-widget dropdown ..
	$("#availability-widget #member-dropdown").change(function(){
		updateMemberEventWidget();
	});
	
	// wire past events button ..
	$("#past-events-button").click(function(){
		$("#form-show-past-events").submit();
		return false
	});
	
	// wire up copy button click event ..
	$("#copy-button a").livequery("click", function(){
		var scheduleItem = $("#copy-to-item .event-item")
		var height = $(scheduleItem).height();
		var qs

		qs = $("#form-copy-event").serialize() 
		qs = qs + "&to_event_id=" + $("#copy-to-item .event-dropdown").val();
		qs = qs + "&from_event_id=" + $("#copy-from-item .event-dropdown").val();
		qs = qs + "&schedule_id=" + $("#copy-to-item .schedule-dropdown").val();
		
		$.ajax({
			type: "GET",
			dataType: "json",
			url: "/_incs/script/ajax/_event_team.asp",
			data: qs, 
			beforeSend: function(){
				$(scheduleItem).css("height", height);
				$(scheduleItem).children("h5, h6, .toolbar, ul, .bottom, .alert").hide()
				$(scheduleItem).css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat center center");
			},
			success: function(response){
				scheduleItem.replaceWith(response.scheduleViewItem);
			},
			error: function(response){console.log(response)}
		});
		
		return false
	});
	
	$(".form-set-schedule-item .schedule-dropdown").livequery("change", function(){
		var qs
		var scheduleItem = $(this).parent().parent().parent();
		var height = $(scheduleItem).height();
		
		// blank out before generating qs ..
		$(this).parent().children(".event-dropdown").val("");
		qs = $(this).parent().serialize();

		$.ajax({
			dataType: "json",
			url: "/_incs/script/ajax/_event_team.asp?" + qs,
			beforeSend: function(){
				$(scheduleItem).css("height", height);
				$(scheduleItem).children("h5, h6, .toolbar, ul, .bottom, .alert").hide()
				$(scheduleItem).css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat center center");
			},
			success: function(data){
				$(scheduleItem).replaceWith(data.scheduleViewItem)
			},
			error: function(response){console.log(response)}
		});
	});
	$(".form-set-schedule-item .event-dropdown").livequery("change", function(){
		var qs = $(this).parent().serialize();
		var scheduleItem = $(this).parent().parent().parent();
		
		$.ajax({
			dataType: "json",
			url: "/_incs/script/ajax/_event_team.asp?" + qs,
			beforeSend: function(){
				$(scheduleItem).css("height", $(scheduleItem).height());
				
				$(scheduleItem).children("h5, h6, .toolbar, ul, .bottom, .alert").hide()
				$(scheduleItem).css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat center center");
			},
			success: function(data){
				$(scheduleItem).replaceWith(data.scheduleViewItem)
				refreshAvailableStyles();
			},
			error: function(response){console.log(response)}
		});
	});

	// wire up publish event button 
	$("a[class^='ajax-publish']").livequery("click", function(){
		var className = $(this).attr("class");
		var scheduleItem = $(this).parent().parent();
		var parentCell = $(scheduleItem).parent();
		var height = $(scheduleItem).height();
		var qs
		
		qs = "action=" + PUBLISH_EVENT + "&event_id=" + className.replace("ajax-publish-", "");
		if ($(parentCell).attr("id") == "copy-from-item"){
			qs = qs + "&item_type=" + SCHEDULE_ITEM_TYPE_COPY_FROM
		}
		else if ($(parentCell).attr("id") == "copy-to-item"){
			qs = qs + "&item_type=" + SCHEDULE_ITEM_TYPE_COPY_TO
		}
		
		$.ajax({
			dataType: "json",
			type: "GET", 
			url: "/_incs/script/ajax/_event_team.asp", 
			data: qs,
			beforeSend: function(){
				$(scheduleItem).css("height", height);
				$(scheduleItem).children("h5, h6, .toolbar, ul, .bottom, .alert").hide()
				$(scheduleItem).css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat center center");
			},
			success: function(response){
				$(scheduleItem).replaceWith(response.scheduleViewItem);
				refreshAvailableStyles();
			},
			error: function(response){console.log(response)}
		});
		return false
	});	

	// wire up remove team event button
	$("a[class^='ajax-remove-team']").livequery("click", function(){
		var className = $(this).attr("class")
		var scheduleItem = $(this).parent().parent();
		var parentCell = $(scheduleItem).parent();
		var height = $(scheduleItem).height();
		var qs
		
		var qs = "action=" + CLEAR_EVENT_TEAM_FROM_EVENT + "&event_id=" + className.replace("ajax-remove-team-", "");
		if ($(parentCell).attr("id") == "copy-from-item"){
			qs = qs + "&item_type=" + SCHEDULE_ITEM_TYPE_COPY_FROM
		}
		else if ($(parentCell).attr("id") == "copy-to-item"){
			qs = qs + "&item_type=" + SCHEDULE_ITEM_TYPE_COPY_TO
		}

		$.ajax({
			type: "GET",
			dataType: "json",
			url: "/_incs/script/ajax/_event_team.asp",
			data: qs,
			beforeSend: function(){
				$(scheduleItem).css("height", height);
				$(scheduleItem).children("h5, h6, .toolbar, ul, .bottom, .alert").hide()
				$(scheduleItem).css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat center center");
			},
			success: function(response){
				$(scheduleItem).replaceWith(response.scheduleViewItem);
				refreshAvailableStyles();
			},
			error: function(response){console.log(response);}
		});		
		return false
	});
	
	$(".form-goto-event-dropdown .event-dropdown").livequery("change", function(){
		$(this).parent().submit();
		return false
	});
	
	$("#go-to-schedule-dropdown").change(function(){
		$(this).parent().submit();
	});
	
	// wire up accordion
	wireUpAccordion()
});
