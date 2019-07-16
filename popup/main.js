function onload() {
	console.log(dates)
	showHTML(dates)
}

function send(i) {
	var date = dates[i].date
	var time = dates[i].time
	var room = dates[i].room

	py_function(date, time, room)
	window.close()
}

function showHTML(list) {
	var html = "";

	var counter = 0;

	for (i = 0; i < list.length; i++)
	{
		if( counter == 0)
		{
			var container = "<div class=\'container\'>";
			html = html.concat(container);
		}
		var string = "<div class=\"card-obj\"><span>Date: " + list[i].date +  "<div id=\"date\"></div> Time: " + list[i].time + " <div id=\"time\"></div> Location: " + list[i].room + " <div id=\"location\"></div> Webex: Yes <div id=\"webex\"></div> <button type=\"button\" onclick = \"send(" + i + ")\">Schedule</button></span></div>";
		html = html.concat(string);
		if(counter == 2)
		{
			var end = "</div>";
			html = html.concat(end);
			counter = 0;
		}
		else
		{
			counter++;
		}
	}

	document.getElementById("data").innerHTML = html;
}