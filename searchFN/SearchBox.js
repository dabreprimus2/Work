/**
 * This JavaScript file contains all the functions regarding the search function.
 *
 * @author pdabre@hi-techhealth.com
 */

/**
 * Set up the search box to show.
 */
var global=[];
function SetSearchBox() {
	searchBox = "<input id=\"searchBox\" name=\"search\" type=\"search\" onkeypress=\"CheckKeyEntered(event)\">";
	searchIco = "<span id=\"searchIco\" title=\"Search\" class=\"clsSearchButtonOn\" onclick=submitdata() onmouseover=SwapViewStyle(this,\'clsSearchButtonHover\') onmouseout=SwapViewStyle(this,\'clsSearchButtonOn\')></span>"
	document.getElementById("HeaderDiv").innerHTML = searchBox + searchIco;
}


function submitdata() {
	var searchTxt = document.getElementById("searchBox").value;
	var searchUCase = searchTxt.toUpperCase();
	var searchlCase = searchTxt.toLowerCase();
	var searchFCase = toTitleCase(searchTxt);
	var sql = "select * from cblib.testpgm where sdesc like '%" + searchTxt + "%' or sdesc like '%" + searchUCase + "%' or sdesc like '%" + searchlCase + "%' or sdesc like '%" + searchFCase + "%'ORDER BY SMENU1, SMENU2, SMENU3, SMENU4, SMENU5 ASC";
	if (searchTxt == "") {
		popuop("null");
	} else {
		$.support.cors = true;
		$.ajax({
			xhr: function () {
				if (window.XMLHttpRequest) {
					return new XMLHttpRequest();
				} else {
					return new ActiveXObject("Microsoft.XMLHTTP");
				}
			},
			type: 'get',
			url: URL,
			data: {
				sql: sql
			},
			dataType: 'json',
			cache: false,
			async: false,
			error: function (request, error, data) {
				//console.log(arguments);
				// console.log("data: "+data+" request :"+request);
				alert(" Can't do because: " + error);
			},
			success: function (data) {
				popuop(data);
			}

		});

		document.getElementById("searchBox").value = "";
	}
}

function popuop(data) {
	HideRuntimeDiv();
	global = data;
	var searchTxt = document.getElementById("searchBox").value;
	if (data == "" || data == "null") {
		noticeHtml = "";
		headerHtml = "<div class=\"clsSearchHeader\">No Search Results Found</div>";
		btnHtml = "<div class=\"clsEnterBtnOn\" onclick=ResetRuntimeDiv() onmouseover=SwapViewStyle(this,\"clsEnterBtnHover\") onmouseout=SwapViewStyle(this,\"clsEnterBtnOn\")>Exit</div>";
		if (data == "") {
			bodyHtml = "<div class=\"clsNoticeBody\"><br><br>No result found For: " + searchTxt + "<br><br></div>";
		} else {
			bodyHtml = "<div class=\"clsNoticeBody\"><br><br>Search feild cannot be empty<br><br></div>";
		}
		panelHtml = "<div class=\"clsNoticePanel\">" + btnHtml + "</div>";

		noticeHtml += headerHtml + bodyHtml + panelHtml + "<div class=\"clsNoticeBottom\"></div>";
		document.getElementById("SearchDiv").innerHTML = noticeHtml;
	} else {
		

		noticeHtml = "";
		headerHtml = "<div class=\"clsSearchHeader\">Search Results</div>";
		btnHtml = "<div class=\"clsEnterBtnOn\" onclick=ResetRuntimeDiv() onmouseover=SwapViewStyle(this,\"clsEnterBtnHover\") onmouseout=SwapViewStyle(this,\"clsEnterBtnOn\")>Exit</div>";

		bodyHtml = "<div class=\"clsNoticeBody\">";
		bodyHtml += "<!--[if lte IE 9]><div class=\"old_ie_wrapper\"><!--<![endif]--><br><table class =\"hoverTable\">";
		bodyHtml += "<tr><th class=\"hoverTableth\">Menu Description</th><th class=\"hoverTableth\">Menu Path</th></tr><tbody>";
		for (idx = 0; idx < data.length; idx++) {
			desc = TrimStr(data[idx].SDESC);
			op1 = TrimStr(data[idx].SMENU1);
			op2 = TrimStr(data[idx].SMENU2);
			op3 = TrimStr(data[idx].SMENU3);
			op4 = TrimStr(data[idx].SMENU4);
			op5 = TrimStr(data[idx].SMENU5);

			if (op2 == "") {
				bodyHtml += "<tr class=\"hoverTabletr\" onmouseover=SwapViewStyle(this,\'clsViewHover2\') onmouseout=SwapViewStyle(this,\'hoverTabletr\') onclick=openOption(\"" + op1 + "\",\"" + op2 + "\",\"" + op3 + "\",\"" + op4 + "\",\"" + op5 + "\")><td>" + data[idx].SDESC + "</td><td>" + data[idx].SMENU1;
				bodyHtml += "</td></tr>";

			} else if (op3 == "") {
				bodyHtml += "<tr class=\"hoverTabletr\" onmouseover=SwapViewStyle(this,\'clsViewHover2\') onmouseout=SwapViewStyle(this,\'hoverTabletr\') onclick=openOption(\"" + op1 + "\",\"" + op2 + "\",\"" + op3 + "\",\"" + op4 + "\",\"" + op5 + "\")><td>" + data[idx].SDESC + "</td><td>" + data[idx].SMENU1 + "-" + data[idx].SMENU2;
				bodyHtml += "</td></tr>";

			} else if (op4 == "") {
				bodyHtml += "<tr class=\"hoverTabletr\" onmouseover=SwapViewStyle(this,\'clsViewHover2\') onmouseout=SwapViewStyle(this,\'hoverTabletr\') onclick=openOption(\"" + op1 + "\",\"" + op2 + "\",\"" + op3 + "\",\"" + op4 + "\",\"" + op5 + "\")><td>" + data[idx].SDESC + "</td><td>" + data[idx].SMENU1 + "-" + data[idx].SMENU2 + "-" + data[idx].SMENU3;
				bodyHtml += "</td></tr>";

			} else if (op5 == "") {
				bodyHtml += "<tr class=\"hoverTabletr\" onmouseover=SwapViewStyle(this,\'clsViewHover2\') onmouseout=SwapViewStyle(this,\'hoverTabletr\') onclick=openOption(\"" + op1 + "\",\"" + op2 + "\",\"" + op3 + "\",\"" + op4 + "\",\"" + op5 + "\")><td>" + data[idx].SDESC + "</td><td>" + data[idx].SMENU1 + "-" + data[idx].SMENU2 + "-" + data[idx].SMENU3 + "-" + data[idx].SMENU4;
				bodyHtml += "</td></tr>";

			} else {
				bodyHtml += "<tr class=\"hoverTabletr\" onmouseover=SwapViewStyle(this,\'clsViewHover2\') onmouseout=SwapViewStyle(this,\'hoverTabletr\') onclick=openOption(\"" + op1 + "\",\"" + op2 + "\",\"" + op3 + "\",\"" + op4 + "\",\"" + op5 + "\")><td>" + data[idx].SDESC + "</td><td>" + data[idx].SMENU1 + "-" + data[idx].SMENU2 + "-" + data[idx].SMENU3 + "-" + data[idx].SMENU4 + "-" + data[idx].SMENU5;
				bodyHtml += "</td></tr>";

			}
		}

		bodyHtml += "</tbody></table><br><!--[if lte IE 9]></div><!--<![endif]-->";

		bodyHtml +="</div>";
		panelHtml = "<div class=\"clsNoticePanel\">" + btnHtml + "</div>";

		noticeHtml += headerHtml + bodyHtml + panelHtml + "<div class=\"clsNoticeBottom\"></div>";
		document.getElementById("SearchDiv").innerHTML = noticeHtml;
	}

}

function ResetRuntimeDiv() {
    document.getElementById("SearchDiv").innerHTML = "";
    ShowRuntimeDiv();
}


/**
 * Remove the search box from show.
 */
function RemoveSearchBox() {
	document.getElementById("HeaderDiv").innerHTML = "";
}

/**
 * Check for Enter key event.
 * @param event
 * 		Event of key pressed.
 */
function CheckKeyEntered(event) {
	if (event.keyCode == 13) {
		submitdata();
	}
}

function toTitleCase(str) {
	return str.replace(/\w\S*/g, function (txt) {
		return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
	});
}

function openOption(SMENU1, SMENU2, SMENU3, SMENU4, SMENU5) {
	var pathList = [];
	pathList.push(SMENU1, SMENU2, SMENU3, SMENU4, SMENU5);
	ShowRuntimeDiv();
	PassPath(pathList);
}
