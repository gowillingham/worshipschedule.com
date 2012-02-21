$(document).ready(function(){
	// save button ..
	$("#save-button a").click(function(){
		$("#form-assign-groups").submit();
	});

	// focus to first element in form ..
	$(".gets-focus").focus();
});
