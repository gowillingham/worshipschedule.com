$(document).ready(function(){

	var backgroundLoaderStyle = "#fff url(/_images/icons/loader_lg.gif) no-repeat 50% 50%";
	var programMemberModalHasChanges = false;
	
	function wireAddRemoveClick(pid, sid){
		
		$("#add-members-button").click(function(){
			var	qs 
			var list = $("#members-to-add-dropdown").val()
			if (list == null){return false};
			
			qs = "pid=" + pid
			qs = qs + "&sid=" + sid
			qs = qs + "&act=" + ADD_PROGRAM_MEMBER
			qs = qs + "&id_list=" + list
			
			$.ajax({
				type: "get",
				dataType: "json",
				url: "/_incs/script/ajax/_program_members.asp",
				data: qs,
				beforeSend: function(){
					$("#form-set-program-members").hide();
					$("#program-member-widget").css("background", backgroundLoaderStyle);
				},
				success: function(response){
					$("#members-to-remove-dropdown").html(response.programMemberOptions);
					$("#members-to-add-dropdown").html(response.clientMemberOptions);
					$("#program-member-widget").css("background-image", "none");			
					$("#form-set-program-members").show();
				},
				error: function(response){
					console.log(response)
				}			
			});
			
			programMemberModalHasChanges = true;
			return false;
		});
		$("#remove-members-button").click(function(){
			var	qs 
			var list = $("#members-to-remove-dropdown").val()
			if (list == null){return false};
			
			qs = "pid=" + pid
			qs = qs + "&sid=" + sid
			qs = qs + "&act=" + REMOVE_PROGRAM_MEMBER
			qs = qs + "&id_list=" + list
			
			$.ajax({
				type: "get",
				dataType: "json",
				url: "/_incs/script/ajax/_program_members.asp",
				data: qs,
				beforeSend: function(){
					$("#form-set-program-members").hide();
					$("#program-member-widget").css("background", backgroundLoaderStyle);
				},
				success: function(response){
					$("#members-to-remove-dropdown").html(response.programMemberOptions);
					$("#members-to-add-dropdown").html(response.clientMemberOptions);
					$("#program-member-widget").css("background-image", "none");			
					$("#form-set-program-members").show();
				},
				error: function(response){
					console.log(response)
				}			
			});
			
			programMemberModalHasChanges = true;
			return false;
		});
	}
	
	// wire up program member widget to button ..
	$(".program-member-button, .program-member-tip-link, .program-member-link").click(function(){
		var pid = $(".program-member-button").attr("id").replace("pid-", "").toString();
		var sid = $.cookie("sid")
		var qs = "pid=" + pid + "&sid=" + sid
		
		$("#program-member-widget").dialog({
			bgiframe: true,
			autoOpen: false,
			modal: true,
			width: "655px",
			height: "375px",
			overlay: {"background-color": "#000", opacity : 0.5},
			beforeclose: function(){
				if (programMemberModalHasChanges) {
					window.location = "/admin/members.asp?pid=" + PROGRAM_ID_ENCRYPTED;	
				};
			},
			buttons: {"Done": function(){$(this).dialog("close")}}
		});
	
		// display the program-member dialog ..
		$("#program-member-widget").html("");
		$("#program-member-widget").dialog("open");
		$.ajax({
			url: "/_incs/script/ajax/_program_members.asp",
			data: qs, 
			beforeSend: function(){
				$("#program-member-widget").css("background", backgroundLoaderStyle);
			},
			success: function(response){
				$("#program-member-widget").css("background-image", "none");			
				$("#program-member-widget").html(response);
				wireAddRemoveClick(pid, sid);
			},
			error: function(response){
				console.log(response)
			}
		});
		return false;
	});
	
	// wire up bulk delete link ..
	$(".bulk-delete-member").click(function(){
		var hasChecked = false
		var message = "You did not select any members to delete. Select at least one member to use the bulk delete feature. "
		var title = "No members were selected!"
		
		$("[name=MemberIDList]").each(function(){
			if (this.checked) {hasChecked = true};
		});
		
		if (hasChecked) {
			$("#form-bulk-delete-members").submit();
			return false;
		}
		else {
			$("#bulk-delete-modal-dialog").dialog({
				bgiframe: true,
				autoOpen: false,
				modal: true,
				width: 400,
				height: 135,
				overlay: {"background-color": "#000", opacity : 0.5},
				title: title,
				buttons: {"Ok": function(){$(this).dialog("close")}}
			});
			$("#bulk-delete-modal-dialog").html('<p>' + message + "<\/p>");
			$("#bulk-delete-modal-dialog").dialog("open");
		};				
	});
	
	// focus to first element in form ..
	$(".gets-focus").focus();

	// wire up master checkbox ..
	$("#master").click(function(){
		var master = this;
		$("[name=MemberIDList]").each(function(){
			this.checked = master.checked;
		});
	});
	
	// wire up go to member dropdown ..
	$("#go-to-member-dropdown").change(function(){
		$("#form-go-to-member").submit();
	});
	
	// wire up go to program dropdown ..
	$("#go-to-program-dropdown").change(function(){
		$("#form-go-to-program").submit();
	});
	
});


