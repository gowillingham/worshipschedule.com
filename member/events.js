$(document).ready(function(){

	function wireRemoveNoteLinks(){
		// unbind existing click events ..
		$(".availability-note a.remove-link").unbind("click");
		
		// bind click event to remove link ..
		$(".availability-note a.remove-link").click(function(){
			var eaid = $(this).parent().parent().children(".button-container").attr("id").replace("eaid-", "")
			var note = $(this).parent();
			var qs
			
			qs = "eaid=" + eaid + "&act=" + DELETE_RECORD
			
			$.ajax({
				url: "/_incs/script/ajax/_member_events.asp",
				type: "post",
				data: qs,
				beforeSend: function(){
					$(note).css("width", $(note).width());
					$(note).css("height", $(note).height());
					$(note).css("background-image", "url(/_images/icons/loader.gif)")
					$(note).css("background-repeat", "no-repeat")
					$(note).css("background-position", "center center")				
				},
				success: function(response){
					$(note).remove();
				},
				error: function(){console.log("error: wireRemoveNoteLinks()")}
			});
					
			return false
		});
	};
	
	function hideSavedEvents(){
		$(".event-item").not($("a.unknown").parent().parent().parent()).hide();
	}
	
	function showAllEvents(){
		$(".event-item").show();
	}
	
	// wire up unviewed events checkbox ..
	$("#unviewed-event-switch-checkbox").attr("checked", "");
	$("#unviewed-event-switch-checkbox").click(function(){
		if (this.checked) {
			hideSavedEvents()
		}
		else {
			showAllEvents()
		};
	});	
	
	// wire up hover toolbar ..
	$(".event-item").hover(
		function() {
			$(this).children(".hover-toolbar").show();
		},
		function() {
			$(this).children(".hover-toolbar").hide();
		}
	);
	
	
	
	// wire up dropdown filters ..
	$("#program-dropdown").change(function(){
		$(this).parent().submit();
	});
	$("#schedule-dropdown").change(function(){
		$(this).parent().submit();
	});

	// wire delete note links ..
	wireRemoveNoteLinks()

	// handle availability click ..
	$("span.available a, span.not-available a").click(function(){
		var eaid = $(this).parent().parent().attr("id").replace("eaid-", "")
		var eventItem = $("#eaid-" + eaid).parent()
		var act
		
		var qs = "eaid=" + eaid
		
		// eat click if this is selected button already ..
		if ($(this).hasClass("selected")) {return false}
		
		if ($(this).parent().hasClass("available")) {
			act = SET_MEMBER_TO_AVAILABLE
		}
		else if ($(this).parent().hasClass("not-available")) {
			act = SET_MEMBER_TO_NOT_AVAILABLE
		}
		qs = qs + "&act=" + act

		$.ajax({
			url: "/_incs/script/ajax/_member_events.asp",
			type: "post",
			data: qs,
			beforeSend: function (){
				$(eventItem).css("width", $(eventItem).width());
				$(eventItem).css("height", $(eventItem).height());
				$(eventItem).children().hide()
				$(eventItem).css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat center center")
			},
			success: function(){
				// removing highlighting styles ..
				$("#eaid-" + eaid + " span a").removeClass("unknown")
				$("#eaid-" + eaid + " span a").removeClass("selected")
				
				//remove question marks ..
				$("#eaid-" + eaid + " span.available a").html("Available")
				$("#eaid-" + eaid + " span.not-available a").html("Not available")
				
				$("#eaid-" + eaid).parent().children(".left").children(".date").removeClass("unknown")
				$("#eaid-" + eaid).parent().children(".left").children(".date").removeClass("available")
				$("#eaid-" + eaid).parent().children(".left").children(".date").removeClass("not-available")
				
				if (act == SET_MEMBER_TO_AVAILABLE) {
					$("#eaid-" + eaid + " span.available a").addClass("selected")
					$("#eaid-" + eaid).parent().children(".left").children(".date").addClass("available")
				}
				else if (act == SET_MEMBER_TO_NOT_AVAILABLE) {
					$("#eaid-" + eaid + " span.not-available a").addClass("selected")
					$("#eaid-" + eaid).parent().children(".left").children(".date").addClass("not-available")
				}
				
				// remove loader.gif and show contents ..
				eventItem.css("background", "none")
				eventItem.children().show();
			},
			error: function (){console.log("error: $.ajax() call for updating availability")}
		});
		
		return false
	});
	
	// check for unsaved changes and bind modal dialog to all links on page ..
	$("a").not($(".available a, .not-available a, .unknown a, .tablist .availability a, .note a, .availability-note a.remove-link, .hover-toolbar a, #logoff a, #help a")).click(function(){
		var link = this
		
		var title = "Some events need your attention!"
		var message = "<p>You have't indicated if you are available for one or more events on this page. "
		message = message + "Is it ok to leave this page? <\/p>"
		
		if ($(".unknown").length > 0) {
			$("#confirm-nav-dialog").remove();
			$("body").append('<div id="confirm-nav-dialog" title="' + title + '">' + message + '<\/div>')
			
			$("#confirm-nav-dialog").dialog({
				modal: true,
				width: "475px",
				height: "135px",
				overlay: {"background-color": "#000", opacity : 0.5},
				bgiframe: true,
				autoOpen: false,
				buttons: {
					"Ok (leave page)": function(){
						$("#confirm-nav-dialog").dialog("close");
						window.location = $(link).attr("href")
 					},  
					"Show those events": function(){
						hideSavedEvents()
						$("#unviewed-event-switch-checkbox").attr("checked", "checked");
						$("#confirm-nav-dialog").dialog("close");
 					},  
					"Cancel (stay)": function(){
						$("#confirm-nav-dialog").dialog("close");
					}
				}
			});
			$("#confirm-nav-dialog").dialog("open");
			return false
		}
	});
	
	// handle add/update note link ..
	$(".note a").click(function(){
		var eaid = $(this).parent().parent().attr("id").replace("eaid-", "")
		var qs = "eaid=" + eaid + "&act=" + GET_AVAILABILITY_FORM
		var loaderStyle = "#fff url(/_images/icons/loader_lg.gif) no-repeat 50% 30%"

		$("#edit-note-dialog").remove();
		$("body").append('<div id="edit-note-dialog" title="Edit availability note"><\/div>')
		
		$.ajax({
			url: "/_incs/script/ajax/_member_events.asp",
			type: "post",
			data: qs,
			beforeSend: function(){
				$("#edit-note-dialog").css("background", loaderStyle)
			},
			success: function(response){
				$("#edit-note-dialog").css("background", "none")
				$("#edit-note-dialog").html(response)
				$("#edit-note-dialog textarea").focus();
			},
			error: function(){console.log("error: bind click for add/update note")}
		});
		
		$("#edit-note-dialog").dialog({
			modal: true,
			overlay: {"background-color": "#000", opacity : 0.5},
			width: "400px",
			height: "260px",
			bgiframe: true,
			autoOpen: false,
			buttons: {
				"Save": function(){
					$.ajax({
						url: "/_incs/script/ajax/_member_events.asp",
						type: "post",
						data: $("#form-availability-note").serialize(),
						dataType: "html",
						beforeSend: function(){
							$("#edit-note-dialog").children().hide();
							$("#edit-note-dialog").css("background", loaderStyle)
						},
						success: function(response){
							// remove existing note from event-item ..
							if ($("#eaid-" + eaid + " + .availability-note").length > 0){
								$("#eaid-" + eaid + " + .availability-note").remove();
							}
							
							// insert note returned from server ..
							$("#eaid-" + eaid).after(response)
							
							// unbind/rebind remove note click event ..
							wireRemoveNoteLinks();
							
							// close the dialog ..
							$("#edit-note-dialog").dialog("close");
						},
						error: function(){console.log("error: update note dialog save button $.ajax call")}
					});
				},  
				"Cancel": function(){
					$("#edit-note-dialog").dialog("close");
				}
			}
		})
		$("#edit-note-dialog").dialog("open")

		return false
	});
});
