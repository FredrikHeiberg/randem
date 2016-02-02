// Implement logic for search of projects

var projectId = "";

$(document).ready(function () {
	$("#searchButton").click(function () {
		projectId = $("#searchCondition").val();
		console.log(projectId);

		// if string is on correct format
		window.location.href = "results.html";

		var tempHeader = "<h1>"+projectId+"</h1>";
		document.getElementById('date').innerHTML += tempHeader;
		console.log(tempHeader);
	})
})

