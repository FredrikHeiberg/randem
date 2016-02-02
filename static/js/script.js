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

$(function() {
  $( "#dateOfOrder" ).datepicker( {dateFormat: 'dd.mm.yy'} );
});

$(function() {
  $( "#searchCondition" ).datepicker( {dateFormat: 'dd.mm.yy'} );
});

$(function() {
  $( "#searchCondition2" ).datepicker( {dateFormat: 'dd.mm.yy'} );
});

var $rows = $('#table1 tr');
$('#searchBox').keyup(function() {
    
    var val = '^(?=.*\\b' + $.trim($(this).val()).split(/\s+/).join('\\b)(?=.*\\b') + ').*$',
        reg = RegExp(val, 'i'),
        text;
    
    $rows.show().filter(function() {
        text = $(this).text().replace(/\s+/g, ' ');
        return !reg.test(text);
    }).hide();
});