/*
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 04/26/1999
*/

var ie4 = document.all;
var ns4 = document.layers;
var x = 0;
var y = 0;

/*
Netscape doesn't track mouse position over form
buttons so keep running track at all times.
My thanks to Sam Kirchmeier (skirch@uidaho.edu)
for helping me figure this out.
*/

function init() {
	if (ns4) { document.captureEvents(Event.MOUSEMOVE); }
	document.onmousemove = mouseMove;
}

function mouseMove(e) {
	if (ie4) {
		x = event.screenX
		y = event.screenY
	}
	if (ns4) {
		x = e.screenX;
		y = e.screenY;
	}
}

/*
Now we can take those coordinates and popup a mini
calendar in the right spot.  Clicking on a calendar
date updates the form element el with the selected
date.
*/

function calpopup(el) {
	x = x + 15; y = y - 10;
	var date = document.forms[0].elements[el].value;
	var url = "webCal4_mini.asp?element=" + el + "&date=" + date
	popWin=window.open(url,"calendar","height=140,width=140,scrollbars=no,titlebar=no,resizable,screenX="+x+",left="+x+",screenY="+y+",top="+y);
}