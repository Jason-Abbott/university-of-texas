//preload images and text for faster operation

if (document.images) {
// back icon
	var iconPrev = new Image();
	iconPrev.src = "images/icon_calprev_grey.gif";
	var iconPrevOn = new Image();
	iconPrevOn.src = "images/icon_calprev.gif";

// forward icon
	var iconNext = new Image();
	iconNext.src = "images/icon_calnext_grey.gif";
	var iconNextOn = new Image();
	iconNextOn.src = "images/icon_calnext.gif";

// print icon
	var iconPrint = new Image();
	iconPrint.src = "images/icon_print_grey.gif";
	var iconPrintOn = new Image();
	iconPrintOn.src = "images/icon_print.gif"
	
// week view icon
	var iconWeek = new Image();
	iconWeek.src = "images/week_grey.gif";
	var iconWeekOn = new Image();
	iconWeekOn.src = "images/week.gif";

// rules icon
	var iconRules = new Image();
	iconRules.src = "images/icon_users_grey.gif";
	var iconRulesOn = new Image();
	iconRulesOn.src = "images/icon_users.gif";
	
// day view icon
	var iconDay = new Image();
	iconDay.src = "images/day_grey.gif";
	var iconDayOn = new Image();
	iconDayOn.src = "images/day.gif";

// goto icon
	var iconGoto = new Image();
	iconGoto.src = "images/icon_goto_grey.gif";
	var iconGotoOn = new Image();
	iconGotoOn.src = "images/icon_goto.gif";
}