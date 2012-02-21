var scheduleHasUnpublishedChanges
var programHasUnpublishedChanges

$(document).ready(function(){

	$("td.member-name").click(function(){
		var row = $(this).parent("tr");
		
		// hide checkboxes for all rows not clicked ..
		$("#availability-grid table tr").not(row).each(function(){
			$(this).children("td").children(".checkbox").hide();
			$(this).children("td").children("div").show();
		});
		
		// toggle checkbox visibility for selected row ..
		$(row).children("td").each(function(){
			$(this).children("div").toggle();
			$(this).children(".checkbox").toggle();
		});
		
	});
	
	// ajax post to update event team info ..
	$("#availability-grid .checkbox").click(function(){
		var skillId = $("#availability-grid").attr("class").replace("skid-", "")
		var eventAvailabilityId = $(this).val()
		
		var qs = "eaid=" + eventAvailabilityId + "&skid=" + skillId;
		
		$.ajax({
			dataType: "text",
			type: "GET", 
			cache: false,
			url: "/_incs/script/ajax/_update_availability_view.asp",
			data: qs,
			beforeSend: function(){
			
				// hide the checkbox and show loader.gif ..
				$("#eaid-" + eventAvailabilityId + " .checkbox").hide()
				$("#eaid-" + eventAvailabilityId + " .checkbox").before('<img class="loader" src="/_images/icons/loader.gif" alt="" />');
			},
			success: function(responseText){
			
				// refresh cell contents ..
				$("#eaid-" + eventAvailabilityId + " div").html(responseText)
				
				// remove loader.gif and unhide the checkbox ..
				$("#eaid-" + eventAvailabilityId + " .loader").remove();
				$("#eaid-" + eventAvailabilityId + " .checkbox").show();
				
				// make sure publish button is showing 
				$("#publish-button").show();
			},
			error: function(){
				// error ..
			}					
		});
	});
		
	// wire up owned program dropdown list ..
	$("#owned-program-dropdown").change(function(){
		$("#form-owned-program-dropdown").submit();
	});
	
	// wire up schedule dropdown list ..
	$("#schedule-dropdown").change(function(){
		$("#form-schedule-dropdown").submit();
	});
	
	// wire up goto skill dropdown list ..
	$("#goto-skill-dropdown").change(function(){
		$("#form-goto-skill-dropdown").submit();
	});
});


