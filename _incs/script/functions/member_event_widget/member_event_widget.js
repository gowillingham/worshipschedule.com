function updateMemberEventWidget(){
	$.ajax({
		type: "post",
		url: "/_incs/script/ajax/_availability.asp",
		data: $("#form-select-member").serialize(),
		beforeSend: function(){
			var height = $("#availability-widget .event-list").height()
			var width = $("#availability-widget .event-list").width()
			
			$("#availability-widget .event-list").css("height", height + "px");
			$("#availability-widget .event-list").css("width", width + "px");
			$("#availability-widget .event-list").children().hide();
			
			$("#availability-widget .event-list").css("background", "#fff url(/_images/icons/loader.gif) no-repeat center center")
		},
		success: function(response){
			$("#availability-widget .event-list").css("background", "none")
			$("#availability-widget .event-list").css("height", "auto")
			$("#availability-widget .event-list").css("width", "auto")
			
			$("#availability-widget .event-list").children().remove();
			$("#availability-widget .event-list").html(response);
			
			//
			$(".event-list .note").jTruncate({length: 200, ellipsisText: " .."});
		},
		error: function(response){console.log(response)}
	});
};

$(document).ready(function(){
	// set widget to default on refresh ..
	$("#availability-widget #member-dropdown").val("")
});

