$(function() {
	$('button').click(function() {
		var date = $('#searchCondition').val();
		$.ajax({
			url: '/results',
			type: 'POST',
			success: function(response) {
				console.log(response);
			},
			error: function(error) {
				console.log(error);
			}
		})
	})
})