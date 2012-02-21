// gets index of a color from global list ..
function getCurrentIndex(color){
	var idx = 0 
	var list = SCHEDULE_HTML_BACKGROUND_COLOR.split(",");
	
	// handles case where form is not being displayed ..
	if (typeof(color) == "undefined"){return idx};
	
	$.each(list, function(i, val){
		if(val.toUpperCase() == color.toUpperCase()){
			idx = i
		};
	});
	return idx
};

$(document).ready(function(){

	// capture publish click in toolbar ..
	$("a.publish-schedule-link").click(function(){
		var href = $(this).attr("href");
		var message = "Select <strong>Publish changes<\/strong> to refresh your member calendars with any event team information that has changed. "
		var title = "Publish event team .."

		$("#publish-schedule-dialog").remove();
		$("body").append('<div id="publish-schedule-dialog"><\/div>');
		
		$("#publish-schedule-dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			modal: true,
			width: 400,
			height: 135,
			overlay: {"background-color": "#000", opacity : 0.5},
			title: title,
			buttons: {
				"Publish changes": function(){
					$(this).dialog("close")
					href = href + "&act=" + PUBLISH_SCHEDULE_ENCRYPTED
					window.location = href
				},
				"Unpublish all": function(){
					$(this).dialog("close")
					href = href + "&act=" + UNPUBLISH_SCHEDULE_ENCRYPTED
					window.location = href
				},
				"Cancel": function(){$(this).dialog("close")}
			}
		});
		$("#publish-schedule-dialog").html('<p>' + message + "<\/p>");
		$("#publish-schedule-dialog").dialog("open");

		return false;
	});

	// wire click event to program dropdown ..
	$("#program-select").change(function(){
		$("#form-program-dropdown").submit();
	});
	
	// set up the color picker ..
	$("#color-picker").colorPicker({
		color: SCHEDULE_HTML_BACKGROUND_COLOR.split(","),
		defaultColor: getCurrentIndex($("#html-color").val()),
		click: function(color){$("#html-color").attr("value", color);}
	});
});
