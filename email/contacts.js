
$(document).ready(function(){
	function refreshQuickSearch() {
		$(".qs_input").val("Search members ..");
	}

	function updateActionPane() {
		var selectedCount = $(".listing input:checked", "#contact-pane").length
		var contactCount = $(".listing .member-item:visible", "#contact-pane").length
		
		var selectedCountText = selectedCount + " contact"
		if (selectedCount != 1) {selectedCountText = selectedCountText + "s"}
		selectedCountText = "<strong>" + selectedCountText + " selected<\/strong>"
		
		var contactCountText = contactCount + " member"
		if (contactCount != 1) {contactCountText = contactCountText + "s"};
		if (contactCount == 0) {contactCountText = "No members"};
		
		var nodeTitle = $("a.highlight", "#tree-view").parent("span").parent("li").attr("title")
		if (nodeTitle == undefined) {nodeTitle = $("#root-node").attr("title")};
		
		if (selectedCount > 0) {
			$("#notifier .header").html(selectedCountText);
			$("#notifier .details").html(" ");
			$("#recipient-buttons").show();
			$("#notifier .help").hide();
		}
		else {
			// no members selected ..
			$("#notifier .header").html(nodeTitle);
			$("#notifier .details").html(contactCountText);
			$("#recipient-buttons").hide();
			
			$("#notifier .help").hide();
			if ($("#root-node a").hasClass("highlight")) {$("#notifier .help").show()};
		}
	};			

	function clearSelected() {
		$(".listing input[type=checkbox]", "#contact-pane").each(function(){
			this.checked = false
			$(this).parent("li").removeClass("highlight")
		});
	};
	
	function checkAll() {
		$(".listing input[type=checkbox]", "#contact-pane").each(function(){
			if ($(this).parent("li").is(":visible")) {
				this.checked = true
				$(this).parent("li").addClass("highlight")
			}
		});
	};
	
	function disableGroupToolbar(){
		$("#action-pane .header .button").each(function(){
			$(this).attr("disabled", true);
		});
		$("#group-member-dropdown").attr("disabled", true);
	};
	
	function enableGroupToolbarEditDelete(){
		$("#action-pane .header .button").each(function(){
			$(this).attr("disabled", false);
		});
	};
	
	function enableGroupToolbarMemberDropdown(){
		$("#action-pane .header select").each(function(){
			$(this).attr("disabled", false);
		});
	};
	
	function disableGroupToolbarMemberDropdown(){
		$("#action-pane .header select").each(function(){
			$(this).attr("disabled", true);
		});
	}
	
	function setGroupToolbarMemberDropdown(){
		var hasSelected = false
		
		$("#contact-pane .listing input[type=checkbox]").each(function(){
			if (this.checked) {
				hasSelected = true

				// break ..
				return false
			};
		});
		
		if (hasSelected) {
			enableGroupToolbarMemberDropdown();
		}
		else {
			disableGroupToolbarMemberDropdown();
		}
	};
	
	function setRemoveFromOptions() {
		var emgid = ""
		
		// first disable all the remove from options ..
		$("#remove-from-optgroup option").each(function(){
			$(this).attr("disabled", "disabled")
		});
		
		// enable all the add to options ..
		$("#add-to-optgroup option").each(function(){
			$(this).attr("disabled", "")
		});
		
		if ($("#tree-view a.highlight").hasClass("email-group-node")) {
			emgid = $("#tree-view a.highlight").attr("id").replace("emgid-", "")

			// enable the remove from option ..
			$("#remove-from-optgroup option").each(function(){
				if ($(this).val() == "emgid" + "-" + emgid + "-" + DELETE_EMAIL_GROUP_MEMBERS){
					$(this).attr("disabled", "");
					return false;
				}
			});
				
			// disable the addto option ..
			$("#add-to-optgroup option").each(function(){
				if ($(this).val() == "emgid" + "-" + emgid + "-" + INSERT_EMAIL_GROUP_MEMBERS){
					$(this).attr("disabled", "disabled");
					return false;
				}
			});
		};
	};
	
	function unsetTreeNodeClickBehavior(){
		$("#tree-view li a").unbind("click");
	};
	
	function setTreeNodeClickBehavior(){
		$(".program-missing-availability-node a, .missing-availability-for-skill-node a, .missing-availability-for-event-node a, .not-available-for-event-node a, .available-for-event-node a, .event-node a, .schedule-missing-availability-node a, .program-node a, .email-group-node a, .skill-group-node a, .skill-node a, .schedule-node a").click(function(){
			var val = $(this).attr("id").split("-")[1]
			var key = $(this).attr("id").split("-")[0]
			var emgid = ""
			var action
			var qs
			
			clearSelected();
			refreshQuickSearch();
			
			// highlight clicked node ..
			$("#tree-view a").removeClass("highlight")
			$(this).addClass("highlight")
			
			if (key == "emgid") {
				action = SMART_GROUP_CUSTOM_GROUP
				emgid = val
			}
			else if (key == "pid") {
				action = SMART_GROUP_PROGRAM
			}
			else if (key == "ungroupedskillpid") {
				action = SMART_GROUP_SKILL_UNGROUPED
				key = "pid"
			}
			else if (key == "skgid") {
				action = SMART_GROUP_SKILL_GROUP
			}
			else if (key == "skid") {
				action = SMART_GROUP_SKILL
			}
			else if (key == "scid") {
				action = SMART_GROUP_SCHEDULE_TEAM
			}
			else if (key == "schedulemissingavailabilityscid") {
				action = SMART_GROUP_SCHEDULE_AVAILABILITY_MISSING
				key = "scid"
			}
			else if (key == "eid") {
				action = SMART_GROUP_EVENT_TEAM
			}
			else if (key == "availableforeventeid") {
				action = SMART_GROUP_EVENT_AVAILABLE
				key = "eid"
			}
			else if (key == "notavailableforeventeid") {
				action = SMART_GROUP_EVENT_NOT_AVAILABLE
				key = "eid"
			}
			else if (key == "missingavailabilityinfoforeventeid") {
				action = SMART_GROUP_EVENT_AVAILABILITY_MISSING
				key = "eid"
			}
			else if (key == "missingavailabilityinfoforskillskid") {
				action = SMART_GROUP_SKILL_AVAILABILITY_MISSING
				key = "skid"
			}
			else if (key == "missingavailabilityinfoforprogrampid") {
				action = SMART_GROUP_PROGRAM_AVAILABILITY_MISSING
				key = "pid"
			};
			
			qs = key + "=" + val + "&act=" + action			
			refreshContactListing(qs);
			
			// enable edit/delete group ..
			disableGroupToolbar();
			
			if ($(this).hasClass("email-group-node")){
				enableGroupToolbarEditDelete();
			};
			
			setRemoveFromOptions()

			return false	
		});
	};
	
	function refreshContactListing(qs) {
		$.ajax({
			type: "GET",
			dataType: "text",
			url: "/_incs/script/ajax/_contacts.asp",
			cache: false,
			data: qs,
			beforeSend: function(){
				// hide no members returned message if it's there ..
				$(".no-members").remove();
				
				// hide the member items
				$(".member-item").hide();
				
				// show the loading gif ..
				$(".listing").css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat 44% 20%")
			},
			success: function(responseText){
				var list = responseText.split(",")
				
				// check length of response text for no members returned ..
				if (responseText.length == 0) { 
					$(".listing", "#contact-grid").append('<div class="no-members">You selected a smart contact group that does not have any members. <\/div>')
				}
				else {
					list = responseText.split(",");
					count = list.length;
					$.each(list, function(i, n){
						$(".listing").css("background", "none")
						$("li#mid-" + n).show();
					});
				};
				$(".listing").css("background", "none")
				$(".listing input[type=checkbox]").attr("checked", "");
				$(".listing input[type=checkbox]").parent("li").removeClass("highlight");
				
				updateActionPane();
			},
			error: function(response){
				console.log("error: refreshContactListing()")
				$(".listing").css("background", "none")
			}
		});
	};
			
	function refreshEmailGroupTreeNodes(nodes){
		// replace nodes with html returned from the server ..
		$("li.email-group-node").remove();
		$("#root-node").after(nodes)
		
		// add click events back in ..
		unsetTreeNodeClickBehavior();
		setTreeNodeClickBehavior();
	};
	
	function setTreeToRootNode(){
		clearSelected();
		refreshQuickSearch();
		disableGroupToolbar();
		
		$(".no-members").remove();
		$(".member-item").show();
		$("#tree-view a").removeClass("highlight");
		$("#root-node a").addClass("highlight");

		updateActionPane();
	};
	
	function deleteEmailGroup(){
		$("li a.highlight").each(function(){
			var emgid = $(this).attr("id").replace("emgid-", "")
			var sid = $.cookie("sid")
			var qs = "act=" + DELETE_RECORD + "&emgid=" + emgid + "&sid=" + sid

			$.ajax({
				type: "POST",
				dataType: "json",
				data: qs,
				url: "/_incs/script/ajax/_email_groups.asp",
				success: function(response){
					refreshEmailGroupTreeNodes(response.nodes);
					$("#group-member-dropdown").html(response.optionList);
					
					$("#root-node a").addClass("highlight");
					setTreeToRootNode();
					setRemoveFromOptions();
				},
				error: function(){console.log("error: wire up delete group button")}
			});
		});
	};
	
	// group-member-dropdown change event ..
	$("#group-member-dropdown").change(function(){
		var emgid = $(this).val().replace("emgid-", "").split("-")[0]
		var act = $(this).val().replace("emgid-", "").split("-")[1]
		
		var qs = "emgid=" + emgid + "&act=" + act
		
		var member_id_list = ""
					
		// parse selected checkbox IDs to get a list of memberId for the form post ..
		$("#form-add-recipients input:checked").parent("li").each(function(){
			member_id_list = member_id_list + $(this).attr("id").replace("mid-", "") + ","
		});
		member_id_list = member_id_list.slice(0, member_id_list.length - 1);
		qs = qs + "&member_id_list=" +  member_id_list
		
		$.ajax({
			type: "POST",
			url: "/_incs/script/ajax/_email_groups.asp",
			data: qs,
			success: function(response){
				var data = "emgid=" + emgid + "&act=" + SMART_GROUP_CUSTOM_GROUP
				
				// highlight the relevant node ..
				$("#tree-view a.highlight").removeClass("highlight");
				$("#tree-view a#emgid-" + emgid).addClass("highlight");
				
				// refresh the contact listing ..
				refreshContactListing(data);
				setRemoveFromOptions();
			},
			error: function(response){console.log("error: #group-member-dropdown change event ..")}
		});
		
		
		// return to unselected state 
		$(this).val("default");
	});
	
	// wire up delete group button ..
	$("#delete-group-button").click(function(){
		// get the group name ..
		var group = $("#tree-view a.highlight").html()
		var message = "<p>Are you sure you want to delete the group " + group + "? This action cannot be undone. <\/p>"
	
		// show confirm delete dialog ..
		$("#delete-group-confirm-dialog").remove();
		$("body").append('<div id="delete-group-confirm-dialog" title="Delete this group?">' + message + '</div>')
		$("#delete-group-confirm-dialog").dialog({
			modal: true,
			overlay: {"background-color": "#000", opacity : 0.5},
			width: "400px",
			height: "150px",
			bgiframe: true,
			autoOpen: false,
			buttons: {
				"Delete the group": function(){
					deleteEmailGroup();
					$("#delete-group-confirm-dialog").dialog("close");
				},  
				"Cancel": function(){
					$("#delete-group-confirm-dialog").dialog("close");
				}
			}
		});
		$("#delete-group-confirm-dialog").dialog("open");
	});
	
	// wire new group dialog to button click ..
	$("#add-group-button, #edit-group-button").click(function(){
		var emgid = $("#tree-view a.highlight").attr("id").replace("emgid-", "")
		if ($(this).attr("id") == "add-group-button") {
			emgid = ""
		}; 

		$("#add-group-dialog").dialog({
			width: "400px",
			height: "220px",
			bgiframe: true,
			autoOpen: false,
			modal: true,
			overlay: {"background-color": "#000", opacity : 0.5},
			buttons: {
				"Save": function(){
					var sid = $.cookie("sid")
					var qs = ""
					var emgid
					
					// pass the session id in the form ..
					$("#session-id", "#add-group-dialog form").val(sid);
				
					$("#add-group-dialog form").ajaxSubmit({
						dataType: "json",
						success: function(response){
						
							if (response.errorMessage.length > 0) {
							
								// remove old messages and display new ..
								$("#add-group-dialog .dialog-notify-message").remove();
								$("#form-email-group").before(response.errorMessage);
							}
							else {
								// remove highlight from original node ..
								$("#tree-view .highlight").each(function(){
									$(this).removeClass("highlight");
								});
								
								refreshEmailGroupTreeNodes(response.nodes);
								$("#group-member-dropdown").html(response.optionList);
								enableGroupToolbarEditDelete();
								
								// generate a qs for contact pane refresh ..
								emgid = $("#tree-view a.highlight").attr("id").replace("emgid-", "")
								qs = "emgid=" + emgid + "&act=" + SMART_GROUP_CUSTOM_GROUP
								refreshContactListing(qs)
								
								setRemoveFromOptions();
								disableGroupToolbarMemberDropdown()
								
								// close dialog ..
								$("#add-group-dialog").dialog("close");
							}
	
							return false
						},
						error: function(response){
							console.log("error: ajaxSubmit add-group-dialog-form")
							console.log(response)	
						}
					});
				},
				"Cancel": function(){$("#add-group-dialog").dialog("close")}
			}
		});
		
		$("#add-group-dialog").dialog("open");
		
		$.ajax({
			type: "POST",
			dataType: "html",
			data: "emgid=" + emgid,
			url: "/_incs/script/ajax/_email_groups.asp",
			beforeSend: function(){
				$("#add-group-dialog").html(" ");
				$("#add-group-dialog").css("background", "#fff url(/_images/icons/loader_lg.gif) no-repeat 50% 25%")
			},
			success: function(response){
				$("#add-group-dialog").css("background-image", "none");
				$("#add-group-dialog").html(response);
				
				// disable enter key for this form ..
				$("#form-email-group").bind("keypress", function(e) {
				  if (e.keyCode == 13) return false;
				});
				
				// focus textbox for this form ..
				$("#add-group-dialog form input:visible:enabled:first").focus();
			},
			error: function(response){
				console.log("error: $.ajax post #add-group-dialog")
				console.log(response)
				$("#add-group-dialog").css("background-image", "none");
			}
		});
	});
	
	// wire click for recipient links ..
	$("#recipient-buttons a").click(function(){

		if ($(this).attr("id") == "email-recipient-button") {
			$("#recipient-type").val(TYPE_EMAIL_RECIPIENTS)
		}
		else if ($(this).attr("id") == "cc-recipient-button") {
			$("#recipient-type").val(TYPE_CC_RECIPIENTS)
		}
		else if ($(this).attr("id") == "bcc-recipient-button") {
			$("#recipient-type").val(TYPE_BCC_RECIPIENTS)
		};
		$("#form-add-recipients").submit();
		
		return false
	});
	
	// highlight root node on refresh ..
	$("#root-node a").addClass("highlight");
	
	// clear checkboxes and refresh action pane on refresh ..
	clearSelected();
	updateActionPane();

	// quicksearch plugin ..
	$("ul.listing li").quicksearch({
		formId: "listing-search-form",
		attached: "td.toolbar",
		position: "append",
		delay: 100,
		loaderText: " ",
		labelText: " ",
		inputText: "Search members ..",
		onAfter: function(){
			clearSelected();
			
			// un-highlight tree nodes ..
			$(".highlight", "#tree-view").removeClass("highlight");

			updateActionPane();
			
			// todo: make these next two lines their own function that can be called ..
			$("#notifier .header").prepend("Searching ")
			$("#notifier .header").append(" for '" + $.trim($("#listing-search-form input.qs_input").val()) + "'")
		}
	});
	
	$(".listing input[type=checkbox]").click(function(){
		if (this.checked) {
			$(this).parent("li").addClass("highlight");
		}
		else {
			$(this).parent("li").removeClass("highlight");
		}
		
		setGroupToolbarMemberDropdown()
	});
	
	// wire up checkbox click ..
	$(".listing input[type=checkbox]").click(function(){
		updateActionPane()
		setGroupToolbarMemberDropdown()
	});
	
	// wire-up check/clear all ..
	$("#check-all-button").click(function(){
		checkAll();
		updateActionPane()
		setGroupToolbarMemberDropdown()
		
		return false;
	});
	$("#clear-all-button").click(function(){
		clearSelected();
		updateActionPane()
		setGroupToolbarMemberDropdown()

		return false;
	});
	
	// wire-up nodes for tree ..
	$("#root-node").click(function(){
		setTreeToRootNode()
		setRemoveFromOptions()
		
		return false;
	});
	
	unsetTreeNodeClickBehavior();
	setTreeNodeClickBehavior();
	disableGroupToolbar();
	setRemoveFromOptions();
	
	// set up treeview ..
	$("#tree-view").treeview({
		animated: "fast",
		collapsed: true,
		control: "#tree-control"
	});
});	
