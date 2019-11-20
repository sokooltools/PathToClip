// =======================================================================================================
// Help.js
// =======================================================================================================
$(document).ready(function() {

	// Add 'Table of Contents' to the page.
	(function createTableOfContents() {
		var sel = $("LI.li2 > a");
		sel.attr('title', "Click to go to the top of this page...");
		sel.attr('href', "#");
		var mid = Math.round(sel.length / 2);
		var str = "<table id='help_toc_table'>";
		str += "<caption id='help_toc_caption'>Table of Contents</caption>";
		str += "<tr>";
		str += "<td class='td1'>";
		str += "<ol start='1'>";
		for (var i = 0; i < mid; i++) {
			sel[i].id = "A" + (i + 1);
			if (sel[i].innerHTML != "")
				str += "<li class='li1'><a href='#" + sel[i].id + "'>" + sel[i].innerHTML + "</a></li>";
			else {
				var sec = "Step " + (i + 1);
				str += "<li class='li1'><a href='#" + sel[i].id + "'>" + sec + "</a></li>";
				sel[i].innerHTML = sec;
			}
		}
		str += "</ol>";
		str += "</td>";
		str += "<td class='td1'>";
		str += "<ol start='" + (mid + 1) + "'>";
		for (i = mid; i < sel.length; i++) {
			sel[i].id = "A" + (i + 1);
			if (sel[i].innerHTML != "")
				str += "<li class='li1'><a href='#" + sel[i].id + "'>" + sel[i].innerHTML + "</a></li>";
			else {
				sec = "Step " + (i + 1);
				str += "<li class='li1'><a href='#" + sel[i].id + "'>" + sec + "</a></li>";
				sel[i].innerHTML = sec;
			}
		}
		str += "</ol>";
		str += "</td>";
		str += "</tr>";
		str += "</table>";
		$('#help_toc_table').replaceWith(str);
		sel = $("LI.li1 > a");
		sel.attr('title', "Click to jump directly to this topic on the page...");
		sel.click(function() {
			scrollToElement(this.hash.replace('#',''));
		});
	})();

	// Add a tooltip to the 'Go to Top' button.
	$("#help_btnScrollToTop").attr('title', "Click to go to the top of this page...");

	// Add a tooltip to the 'Close' button.
	$("#help_btnClose").attr('title', "Click to close this online help...");

	// Add event to close the window when the 'Close' button is clicked.
	$("#help_btnClose").click(function() {
		window.close();
	});

	// Add event to scroll to top of page when any caption is clicked.
	$("#help_btnScrollToTop, .li2 > a").click(function() {
		window.scrollTo(0, 0);
	});

	// Scroll the selected topic into view adjusting for the header bar.
	function scrollToElement(elId) {
		document.getElementById(elId).scrollIntoView(true);
		if  ($(window).height() + $(window).scrollTop() !== $(document).height())
		{
			$('html, body').animate({
				scrollTop: "-=40"
			}, "fast");
		}
//		if (window.navigator.userAgent.indexOf('MSIE') != -1) {
//			var url = window.unescape(window.self.location.hash);
//			if (url != "") {
//				window.scrollBy(0, -40);
//			}
//		}
	};


	//$('#titleee a').trigger('click');

	//	(function updatePage() {
	//		// Create a dictionary from the table of contents.
	//		var sel = $("LI.li1 > a");
	//		var dict = [];
	//		for (var i = 0; i < sel.length; i++) {
	//			dict.push({
	//				key: sel[i].hash,
	//				value: sel[i].outerText
	//			});
	//			$(sel[i]).attr('title', "Click to jump directly to this topic on the page...");
	//		}
	//		// Update the page using the dictionary.
	//		for (var n = 0; n < dict.length; n++) {
	//			$(dict[n].key).attr('title', "Click to return to top of this page...");
	//			$(dict[n].key).attr('href', "#top");
	//			$(dict[n].key).text(dict[n].value);
	//		}
	//		$("#btnClose").attr('title', "Click to close this online help...");
	//	} ());

});
