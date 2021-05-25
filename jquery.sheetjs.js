
/*!
 *
 * The 'sheetjs' is a simple spreadsheet UI for your table tag as a jQuery-plugin.
 *
 *
 *
 * https://github.com/hidehitomobile/sheetjs.git
 * Released under the MIT license
 *
 * @version 1.1
 * @date 2021-05-11
 * @author Hidehito Tanaka
 * @see https://github.com/hidehitomobile
 * 
 */
(function(factory) {
	// UMD module
	// https://github.com/umdjs/umd/blob/master/templates/jqueryPlugin.js
	if (typeof define === 'function' && define.amd) {
		define(['jquery'], factory);
	}
	else if(typeof module === 'object' && module.exports) {
		module.exports = function(root, jQuery) {
			if(jQuery === undefined) {
				if(typeof window !== 'undefined') {
					jQuery = require('jquery');
				}
				else {
					jQuery = require('jquery')(root);
				}
			}
			factory(jQuery);
			return jQuery;
		};
	}
	else {
		factory(jQuery);
	}
}(function($) {

	'use strict';

	$.fn.sheetjs = function() {
	
		var context = this;

		//
		// Here is shared functions or values
		//

		var isIE = (/.*(Trident|MSIE).*/.test(navigator.userAgent));
		var isIE7 = (navigator.userAgent.indexOf('MSIE 7.0') > 0);
		var isJQuery3 = /^3/.test($.fn.jQuery);
	
		/*
		 * Check the string is valid number or not, exclude logarithm number.
		 * native function isNaN is include logarithm number.
		 */
		function isNum(str) {
			if(str) {
				return str.match(/^-?(\d+|\d+\.?\d*|\d*\.?\d+)$/);
			}
			return true;
		}
	
		/*
		 * Format num to comma separated string.
		 */
		function numFormat(num) {
			return (num + '').replace(/-?(\d)(?=(\d{3})+(?!\d))/g, '$1,');
		}
		
		/*
		 * Extract table cells to a two-dimensional array.
		 * If a cell has colspan attribute or rowspan attribute, push two or more cells into the array
		 * to prevent shifting.
		 */
		function getTableCells(tblElm) {
			var tHead = tblElm.tHead;
			var tBody = tblElm.tBodies[0];
			var tFoot = tblElm.tFoot;
			
			// This logic is reused 3 times, this is defined an internal function.
			var extract = function(rowSet) {
					var data = [];
					var rows = [];
					for(var rr in rowSet) {
						for(var i = 0, r = rowSet[rr].length; i < r; i++) {
							if(rowSet[rr][i].style.display !== 'none') { // ignore invisible row
								rows.push(rowSet[rr][i]);
								data.push([]);// row's cell array
							}
						}
					}
					// Expand to a matrix considers rowspan and colspan
					for(var r = 0; r < rows.length; r++) {
						var cells = rows[r].cells;
						for(var c = 0; c < cells.length; c++) {
							var x = c;
							var cell = cells[c];
							while(data[r][x] !== undefined) x++; // Check the cell is setted because of rowspan or colspan.
							for(var col = 0; col < cell.colSpan; col++) { // if it is marged cell, push the same value.
								for(var row = 0; row < cell.rowSpan; row++) { // if it is marged cell, push the same value.
									data[r+row][x+col] = cell;
								}
							}
						}
					}
					return data;
				};
			
			var headCells = extract([tHead ? tHead.rows : []]);
			var bodyCells = extract([tBody ? tBody.rows : []]);
			var footCells = extract([tFoot ? tFoot.rows : []]);
			return {
				headCells : headCells,
				bodyCells : bodyCells,
				footCells : footCells,
				allCells : headCells.concat(bodyCells).concat(footCells)
			};
		}
	
		/*
		 * Get specified cells
		 */
		function getRangeCells(data, beginCell, endCell) {
			var range = [];
			var row1 = Infinity, col1 = Infinity, row2 = 0, col2 = 0;
			// Find the start cell and the end cell
			// Sometime the range start and end is reversed,
			// to find the maximum value and the minimum value.
			for(var row = 0; row < data.length; row++) {
				for(var col = 0; col < data[row].length; col++) {
					if(data[row][col] === beginCell) {
						if(row1 > row) row1 = row;
						if(col1 > col) col1 = col;
						if(row2 < row) row2 = row;
						if(col2 < col) col2 = col;
					}
					if(data[row][col] === endCell) {
						if(row1 > row) row1 = row;
						if(col1 > col) col1 = col;
						if(row2 < row) row2 = row;
						if(col2 < col) col2 = col;
					}
				}
			}
			// After find the maximum and minimum value, extract the range cells.
			for(var row = row1; row <= row2; row++) {
				var subdata = [];
				for(var col = col1; col <= col2; col++) {
					subdata.push(data[row][col]);
				}
				range.push(subdata);
			}
			return {
				cells: range,
				row1: row1,
				row2: row2,
				col1: col1,
				col2: col2
			};
		}
		
		/*
		 * Convert cells text into TSV(Tab Separeted Values)
		 */
		function getCellsTsv(data) {
			var tsv = '';
			for(var row = 0; row < data.length; row++) {
				// Exclude invisible cells
				if(!$(data[row][0].parentNode).is(':visible')) {
					continue;
				}
				
				if(row != 0) tsv += '\n';
				for(var col = 0; col < data[row].length; col++) {
					if(col != 0) tsv += '\t';
					var text = getStaticText(data[row][col]);
					if(/[\t\n"]/.test(text)) {
						// If it contains line breaks and tabs, enclose it in a double coat.
						// And if there is " in the character, it is converted to "".
						// Because of IE7, I don't use replaceAll('"', '""') function.
						tsv += '"' + text.split('"').join('""') + '"';
					}
					else {
						tsv += text;
					}
				}
			}
			return tsv;
		}
		
		/*
		 * Convert the TSV text into tow-demensional array.
		 */
		function parseTsv(text) {
			var m = text.match(/(\t|\r?\n|[^\t"\r\n]+|"(?:[^"]|"")*")/g);
			var tsv = [];
			var line = [];
			tsv.push(line);
			for(var i = 0; i < m.length; i++) {
				var str = m[i];
				if(/^(\r\n|\r|\n)$/.test(str)) {
					line = [];
					tsv.push(line);
				}
				else if(/^\t$/.test(str)) {
					//noop
				}
				else {
					if(/^"[^"]*"$/.test(str)) {
						str = str.replace(/^"|"$/g, '');
					}
					str = str.replace(/""/g, '"');
					line.push(str);
				}
			}
			return tsv;
		}
	
		/*
		 * Set clipboard, works both of IE and Chrome
		 */
		function setClipboard(str) {
			// for IE
			if(window.clipboardData) {
				clipboardData.setData('Text', str);
			}
			// for Chrome
			else if(navigator.clipboard) {
				navigator.clipboard.writeText(str);
			}
		}
	
		/*
		 * Get element text recursively
		 * Include input, textarea and button value
		 * Trim the start whitespace
		 */
		function getStaticText(element) {
			if(element.className === 'markbox') return ''; // Ignore the triangle mark. 
	
			var str = (/checkbox|radio/i.test(element.type)
				? (element.checked ? element.value : "") // Get the value when it is checked.
				: (element.nodeValue || element.value || '')).replace(/^\s+|\s+$|\n\s+|\n/g, '');
			if(!/select/i.test(element.type)) { // Select tag has some chile not. but only get the selected option value.
				for(var i = 0; i < element.childNodes.length; i++) {
					str += getStaticText(element.childNodes[i]);
				}
			}
			
			if(/img/i.test(element.tagName)) {
				str = element.alt;
			}
			else if(str === '' && /br/i.test(element.tagName)) {
				str = '\n';
			}
			return str;
		}

		/*
		 * This is for developping.
		 */
		var preTime = new Date().getTime();
		function printTime(str) {
			var now = new Date().getTime();
			if(window.console) console.log(str, (now-preTime) + 'ms');
			preTime = now;
		}
		
		
		/*
		 * Main
		 */
		$('table.sheetjs', context).each(function() {
			var table = $(this);
			var isSelecting = false;
			var isMouseDown = false;
			var isCalculatingSum = false;
			var beginCell = null;
			var endCell = null;
			var clickedCell = null;
			var rectLine1 = $('<div class="rectline"><!--empty--></div>');
			var rectLine2 = $('<div class="rectline"><!--empty--></div>');
			var rectLine3 = $('<div class="rectline"><!--empty--></div>');
			var rectLine4 = $('<div class="rectline"><!--empty--></div>');
			var copyLine1 = $('<div class="rectline dashed"><!--empty--></div>');
			var copyLine2 = $('<div class="rectline dashed"><!--empty--></div>');
			var copyLine3 = $('<div class="rectline dashed"><!--empty--></div>');
			var copyLine4 = $('<div class="rectline dashed"><!--empty--></div>');
			var caption = table.find('caption');
			var fullcopyBox = $('<span class="captionmenu" style="margin-left:10px;"><span class="material-icons">content_copy</span></span>');
			var rowCountBox = $('<span class="captionbox">Row<span style="margin-left:3px;"></span></span>');
			var sumBox = $('<span class="captionbox">Sum<span style="margin-left:3px;"></span></span>');
			var ctrlBoxHeader = $('<div class="contextbox" style="display:none;"></div>');
			var ctrlAZHeader = $('<div class="contextmenu"><span class="material-icons">south</span> Sort　A-Z</div>');
			var ctrlZAHeader = $('<div class="contextmenu"><span class="material-icons">north</span> Sort　Z-A</div>');
			var ctrlCopyHeader = $('<div class="contextmenu"><span class="material-icons">content_copy</span> Copy　<span style="color:#aaa;font-size:95%;">Ctrl+C</span></div>');
			var ctrlSumHeader = $('<div class="contextmenu"><span class="material-icons">functions</span> Sum　<span class="sum"></span></div>');
			var ctrlFilterHeader = $('<div class="contextmenu"><span class="material-icons">search</span> Filter　<input class="searchfilter" autocomplete="off"></div>');
			var searchFilter = ctrlFilterHeader.children(".searchfilter");
			var ctrlClearFilterHeader = $('<div class="contextmenu"><span class="material-icons" style="color:#aaa;text-decoration:line-through">filter_alt</span> Clear Filter <div class="filtercondition"></div></div>');
			var ctrlBox = $('<div class="contextbox" style="display:none;"></div>');
			var ctrlCopy = $('<div class="contextmenu"><span class="material-icons">content_copy</span> Copy　<span style="color:#aaa;font-size:95%;">Ctrl+C</span></div>');
			var ctrlSum = $('<div class="contextmenu"><span class="material-icons">functions</span> Sum　<span class="sum"></span></div>');
			var ctrlFilter = $('<div class="contextmenu"><span class="material-icons">filter_alt</span> Filter　"<span class="flt"></span>"</div>');
			var ctrlNotFilter = $('<div class="contextmenu"><span class="material-icons">filter_alt</span> Not Filter　"<span class="flt" style="text-decoration:line-through"></span>"</div>');
			var ctrlClearFilter = $('<div class="contextmenu"><span class="material-icons" style="color:#aaa;text-decoration:line-through">filter_alt</span> Clear Filter <div class="filtercondition"></div></div>');
			var cells = getTableCells(table[0]);
			var headCells = cells.headCells;
			var bodyCells = cells.bodyCells;
			var footCells = cells.footCells;
			var allCells = cells.allCells;
			var filterCondition = '';
			var sumNum = 0;
			var hoverHead = null;
			var isHeadless = (headCells.length === 0);
	
			//
			// Initialize components
			//
			
			// append row count box and full copy box if there is caption tag
			if(caption.length > 0 && !isHeadless) {
				caption.append([fullcopyBox, rowCountBox, sumBox]);
			}
			else {
				// append caption tag as the parent element of rectangle lines.
				// the caption position is set 'relative'.
				caption = $('<caption></caption>');
				table.append(caption);
			}
			ctrlBoxHeader.append([ctrlAZHeader, ctrlZAHeader,ctrlCopyHeader, ctrlSumHeader, ctrlFilterHeader, ctrlClearFilterHeader]);
			ctrlBox.append([ctrlCopy, ctrlSum]);
			if(!isHeadless) ctrlBox.append([ctrlFilter, ctrlNotFilter, ctrlClearFilter]);
			caption.append([rectLine1, rectLine2, rectLine3, rectLine4, copyLine1, copyLine2, copyLine3, copyLine4, ctrlBoxHeader, ctrlBox]);
			
			// Fine-tune the border style
			if(isHeadless) table.addClass('headless');
			
			// Support IE7
			if(isIE7) {
				ctrlBoxHeader.css({width:'200px'}); // instedof min-width
				ctrlBox.css({width:'200px'}); // instedof min-width
			}
	
			// Append combo buttons to the header
			if(isIE) {
				table.find('>thead>tr>th').each(function() {
					// IE's th and td are become borderless when the position style be set 'relative'.
					// therefore leyout with 'float'.
					$(this).append('<div class="markbox" style="position:relative;float:right;"><div class="sortmark"></div><div class="menumark">&#9660;</div></div>');
				});
			}
			else {
				//layout with 'relative'
				table.find('>thead>tr>th').each(function() {
					$(this).css('position', 'relative').append('<div class="markbox"><div class="sortmark"></div><div class="menumark">&#9660;</div></div>');
				});
			}
	
			// Create hover header
			// None support IE
			if(!isIE && !isHeadless) {
				hoverHead = $('<table class="hoverhead"><thead><tr></tr></thead></table>');
				var preCell = null;
				for(var c = 0; c < headCells[headCells.length-1].length; c++) {
					var cell = headCells[headCells.length-1][c];
					if(preCell !== cell) { // Considering colspan
						var th = $('<th></th>').html(getStaticText(cell)).css('width', $(cell).innerWidth() + 'px');
						th.refCell = cell; // Use this refernce when table resized.
						hoverHead.find('tr').append(th);
						preCell = cell;
					}
				}
				hoverHead.css('width', table.innerWidth() + 'px');
				hoverHead.css('left', '1px');
				caption.append(hoverHead);
	
				// Scroll hover header within the window.
				$(window).scroll(scrollHoverHeadHundler);
			}
	
	
			//
			// Setup event actions
			//
	
			// Fully copy the cell's text
			fullcopyBox.on('click', function() {
				var tsv = getCellsTsv(allCells);
				setClipboard(tsv);
			});
	
			// Copy event
			$([ctrlCopyHeader[0], ctrlCopy[0]]).on('click', function() { 
				if(isIE) {
					copyHandler();
				}
				else {
					document.execCommand("copy");
				}
			});
	
			// Copy the sum number
			ctrlSumHeader.on('click', function() {
				var str = ctrlSumHeader.children('span').html();
				setClipboard(str);
			});

			// Copy the sum number
			ctrlSum.on('click', function() {
				var str = ctrlSum.children('span').html();
				setClipboard(str);
			});
	
			// SearchFilter
			ctrlFilterHeader.on('click', function(event) {
				// Cancel original event not to hide the menu
				if(event.preventDefault) event.preventDefault(); //support for IE7
				if(event.stopPropagation) event.stopPropagation(); //support for IE7
				event.returnValue = false;
				return false;
			});
	
			// Fuzzy text search
			searchFilter.on('change keypress', function(event) {
				if(event.type === 'change' || (event.type === 'keypress' && event.key === 'Enter')) {
					var str = this.value;
					if(str !== '') {
						applyFilter(function(text) { return text.indexOf(str) != -1; });
						filterCondition += str + '<br>';
					}
					this.value = '';
					event.stopPropagation();
					return false;
				}
			});

			// Apply the selection filter event
			ctrlFilter.on('click', function() {
				var str = ctrlFilter.children('span.flt').html();
				applyFilter(function(text) { return text === str; });
					filterCondition += str + '<br>';
			});

			// Apply the none selection filter event
			ctrlNotFilter.on('click', function() {
				var str = ctrlNotFilter.children('span.flt').html();
				applyFilter(function(text) { return text !== str; });
					filterCondition += '<s>' + str + '</s><br>';
			});

			// Apply the clear filter event
			ctrlClearFilterHeader.on('click', clearFilter)
			ctrlClearFilter.on('click', clearFilter);

			// Apply the sort event
			ctrlAZHeader.on('click', function() { sortBody(true); });
			ctrlZAHeader.on('click', function() { sortBody(false); });

			// Sort event from double click event
			table.find('>thead>tr>th').on('dblclick', function(event) {
				if(/th/i.test(event.target.tagName)) {
					clickedCell = event.target;
					var isAsc = $(clickedCell).find('div.sortmark').html() !== '↓';
					sortBody(isAsc);
				}
			});
	
			// Begin drawing the select rectangle
			//
			// Normaly below selecter is perfect. but this is slow at event inisializing.
			// table.find(">thead>tr>th, >thead>tr>td, >tbody>tr>th, >tbody>tr>td, >tfoot>tr>th, >tfoot>tr>td").on('mousedown'
			//
			// Note: When mousedown on the nested table's cell, rectangle draws on wrong cell.
			table.on('mousedown', function(event) {
				if(/th|td/i.test(event.target.tagName) && (event.button === 0 || event.button === 1)) { // Support IE7 : Right click is '1'
					
					if(!event.shiftKey) {
						beginCell = event.target;
					}
					endCell = event.target;
					isMouseDown = true;
					selectRectangle(event.ctrlKey || event.metaKey);
					
					// Support IE
					// Stop event to avoide text selection
					// Note: Eventhrough IE7 select text.
					if(isIE && event.shiftKey) {
						event.stopPropagation();
						return false;
					}
				}
			});

			// Stretch rectangle
			//
			// Normaly below selecter is perfect. but this is slow at event inisializing.
			// table.find(">thead>tr>th, >thead>tr>td, >tbody>tr>th, >tbody>tr>td, >tfoot>tr>th, >tfoot>tr>td").on('mouseup'
			table.on("mouseover", function(event){
				if(isMouseDown && /th|td/i.test(event.target.tagName)) {
					endCell = event.target;
					selectRectangle(event.ctrlKey || event.metaKey); // Multiple selection with the controle key. Note: macOS is metaKey
				}
			});

			// End drawing the select rectangle
			//
			// Normaly below selecter is perfect. but this is slow at event inisializing.
			// table.find(">thead>tr>th, >thead>tr>td, >tbody>tr>th, >tbody>tr>td, >tfoot>tr>th, >tfoot>tr>td").on('mouseup'
			table.on('mouseup', function(event) {
				if(/th|td/i.test(event.target.tagName) && (event.button === 0 || event.button === 1)) {
					endCell = event.target;
					isMouseDown = false;
					selectRectangle(event.ctrlKey || event.metaKey);
				}
			});


			// Move select rectangle or hide copy rectangle
			$(document).on('keydown', function(event) {

				// Simply hide copy lines with escape key
				// Eventhough clipboard is not cleared. (same as google spreadsheets)
				if(isSelecting && /Escape/.test(event.key)) {
					rectLine1.removeClass('dashed');
					rectLine2.removeClass('dashed');
					rectLine3.removeClass('dashed');
					rectLine4.removeClass('dashed');
					copyLine1.hide();
					copyLine2.hide();
					copyLine3.hide();
					copyLine4.hide();
				}
				// Move select rectangle with arrow keys
				// Node: Shift + Enter motion is different from google spreadsheets. (google move selection upward)
				else if(isSelecting && /ArrowLeft|ArrowUp|ArrowRight|ArrowDown|Enter/.test(event.key)) {
					var currentCell = (event.shiftKey ? endCell : beginCell);
					var movedCell = currentCell;
					var range = getRangeCells(allCells, currentCell, currentCell);
					var row = range.row1;
					var col = range.col1;

					// Jump to the last
					if(event.ctrlKey || event.metaKey) {
						if(event.key === 'ArrowLeft') col = 0;
						else if(event.key === 'ArrowUp') row = 0;
						else if(event.key === 'ArrowRight') col = allCells[0].length-1;
						else if(event.key === 'ArrowDown' || event.key === 'Enter') row = allCells.length-1;
						movedCell = allCells[row][col];
					}
					// Normal
					else {
						// Loop within colspan or rowspan ranges
						while(currentCell === movedCell) {
							if(event.key === 'ArrowLeft') {
								if(allCells[row][--col] === undefined) return;
							}
							else if(event.key === 'ArrowUp') {
								if(allCells[--row] === undefined) return;
							}
							else if(event.key === 'ArrowRight') {
								if(allCells[row][++col] === undefined) return;
							}
							else if(event.key === 'ArrowDown' || event.key === 'Enter') {
								if(allCells[++row] === undefined) return;
							}
							movedCell = allCells[row][col];
						}
					}

					if(event.shiftKey) {
						endCell = movedCell;
					}
					else {
						beginCell = movedCell;
						endCell = movedCell;
					}
					selectRectangle(false);
					
					// Cancel orignale event to cancel scrolling
					event.stopPropagation();
					return false;
				}
			});

			// Append a context event on thead
			table.find(">thead>tr>th").on('contextmenu', headerCtrlHandler);
			table.find(">thead>tr>th>div.markbox>div.menumark").on('click', headerCtrlHandler);
	
			// Append a context event on tbody
			table.find(">tbody, >tfoot").on('contextmenu', bodyCtrlHandler);
	
			// Draw the select rectangle again when the table size is changed
			// This is only for modern browsers.
			// Note: copy lines don't follow the size.
			if(window.ResizeObserver) {
				new ResizeObserver(function(entries) {
					selectRectangle();
				}).observe(table[0]);
			}

			// Clear context menu
			// when outside of the table is clicked
			$("body").on('click', function(event) {
				// hide menu
				ctrlBoxHeader.hide();
				ctrlBox.hide();
	
				// hide button
				table.find('>thead>tr>th>div.markbox>div.menumark').removeClass('active');
	
				// Check the event from own table or not recursively
				var tblElm = table[0];
				var element = event.target;
				while(element) {
					if(element === tblElm) return;
					element = element.parentNode;
				}
				clearRectangle();
				updateStat();
			});
	
			// Append copy event handlers
			if(isIE) {
				$(document).on('keydown', copyHandler);
				$(document).on('keydown', pasteHandler);
			}
			else {
				$(document).on('copy', copyHandler);
				$(document).on('paste', pasteHandler);
			}
			
			
			//
			// Define hundlers
			//
			
			/*
			 * Copy handler
			 * Chrome call this on copyevent
			 * IE call this on keyevent. (IE don't have copy event)
			 */
			function copyHandler(event) {
				event = event || window.event; //Support IE
				var target = event.target || event.srcElement; //Support IE
				
				if(!isSelecting) {
					return;
				}

				// Normal input method
				if(/input|textarea/i.test(target.tagName)) {
					return;
				}

				// Support IE. Observe 'Ctrl + C' key event
				if(event.type === 'keydown' && (event.ctrlKey === false || event.key !== 'c')) {
						return;
				}
				
				// Generate TSV text, then set TSV into clipboard
				var range = getRangeCells(allCells, beginCell, endCell).cells;
				var tsv = getCellsTsv(range);
				setClipboard(tsv);

				// Change the rectangle border style
				rectLine1.addClass('dashed');
				rectLine2.addClass('dashed');
				rectLine3.addClass('dashed');
				rectLine4.addClass('dashed');
				
				// Draw copylines
				copyLine1.css(rectLine1.css(['top', 'left', 'height', 'width'])).show();
				copyLine2.css(rectLine2.css(['top', 'left', 'height', 'width'])).show();
				copyLine3.css(rectLine3.css(['top', 'left', 'height', 'width'])).show();
				copyLine4.css(rectLine4.css(['top', 'left', 'height', 'width'])).show();

				// Cancel original event
				if(event.preventDefault) event.preventDefault(); // Support IE7
				if(event.stopPropagation) event.stopPropagation(); // Support IE7
				event.returnValue = false;
				return false;
			}
	
			/*
			 * Paste handler
			 * set TSV values into editable inputs, editable textarea or contenteditable element
			 */
			function pasteHandler(event) {
				event = event || window.event;
				var target = event.target || event.srcElement;

				// If the event from the input on the cell, the original text copy processing is performed.
				if(/input|textarea/i.test(target.tagName)) {
					return;
				}
				// Nothing selected
				if(!isSelecting) {
					return;
				}

				// Support IE. Observe 'Ctrl + V' key event
				if(event.type === 'keydown' && (event.ctrlKey === false || event.key !== 'v')) {
						return;
				}
	
				// Functionalized for Promis
				var paste = function(text) {
						var tsv = parseTsv(text);
						if(tsv.length == 0) return;
						
						var range = getRangeCells(allCells, beginCell, endCell).cells;
						for(var r = 0; r < range.length; r++) {
							for(var c = 0; c < range[r].length; c++) {
								var r2 = r % tsv.length; // Repeat values as much as rows length. like this 1,2,3,1,2,3,...
								var c2 = c % tsv[r2].length; // Repeat values as much as cols length. like this a,b,a,b,...
								var val = tsv[r2][c2];
								var cell = range[r][c];
								if(cell.contentEditable === 'true') {
									if(cell.innerHTML !== val) {
										$(cell).html(val).trigger('change'); // TODO: This is not sure that 'change' event is proper or not.
									}
								}
								else { // If includeing input elements, change the first input value.
									var elm =$('input:text:visible:enabled:not([readonly]),textarea:visible:enabled:not([readonly]),select:visible:enabled', cell).first();
									if(elm.length > 0 && elm.val() !== val) {
										elm.val(val).trigger('change');
									}
								}
							}
						}
					};
	
				// For IE
				if (window.clipboardData && window.clipboardData.getData) { 
					paste(window.clipboardData.getData('Text'));
				}
				// If you touch the clipboard
				else if (event.clipboardData && event.clipboardData.getData) {
					paste(event.clipboardData.getData('text/plain'));
				}
				// For Chrome
				else {
					navigator.clipboard.readText().then(paste); // Not handle a error, because of IE7 catch identifier error
				}
	
				// Cancel the original event
				if(event.preventDefault) event.preventDefault(); // For IE7
				if(event.stopPropagation) event.stopPropagation(); // For IE7
				event.returnValue = false;
				return false;
			}
	
			/*
			 * @param filter - filter function
			 */
			function applyFilter(filter) {
				// Get column index
				var col = getRangeCells(allCells, clickedCell, clickedCell).col1;
	
				// Mark the column head 'filtering'
				if(!isHeadless) {
					$('div.menumark', headCells[headCells.length-1][col]).addClass('filtering');
				}
				
				// Create only display settings array before remove from DOM to avoid reflow
				// (In IE6,7, if you disconnect, the contents of rows and cells will be undefined,
				// so the conditions will be calculated before remove.)）
				var rowsVisible = [];
				for(var row = 0; row < bodyCells.length; row++) {
					if(bodyCells[row][col].parentNode.style.display !== 'none') { //tr
						var text = getStaticText(bodyCells[row][col]);
						rowsVisible[row] = filter(text);
					}
				}
				
				var rowsRef = []; // Support IE7. Keep refernce before remove from DOM.
				for(var r = 0; r < bodyCells.length; r++) {
					rowsRef[r] = bodyCells[r][col].parentNode;
				}
				
				var dummy = document.createElement('tbody');
				var tblElm = table[0];
				var tbody = tblElm.tBodies[0];
				tblElm.replaceChild(dummy, tbody);
				for(var r = 0; r < bodyCells.length; r++) {
					rowsRef[r].style.display = (rowsVisible[r] ? '' : 'none');
				}
				tblElm.replaceChild(tbody, dummy);
	
				clearRectangle();
				updateSum();
				updateRowCount();
				updateStat();
			}
	
			/*
			 * Clear filters
			 */
			function clearFilter() {
				table.find('>thead>tr>th>div.markbox>div.menumark.filtering').removeClass('filtering');
				filterCondition = '';
	
				var rowsRef = [];  // Support IE7. Keep refernce before remove from DOM.
				for(var r = 0; r < bodyCells.length; r++) {
					rowsRef[r] = bodyCells[r][0].parentNode;
				}
				var dummy = document.createElement('tbody');
				var tblElm = table[0];
				var tbody = tblElm.tBodies[0];
				tblElm.replaceChild(dummy, tbody);
				for(var r = 0; r < bodyCells.length; r++) {
					rowsRef[r].style.display = '';
				}
				tblElm.replaceChild(tbody, dummy);
	
				clearRectangle();
				updateSum();
				updateRowCount();
				updateStat();
			}
	
			/*
			 * Draw select rectabgle
			 */
			function selectRectangle(withCrtlKey) {
				
				// this case is occur on loaded timming
				// called from resize observer
				if(beginCell == null) {
					return;
				}

				isSelecting = true;
				beginCell = beginCell || endCell;
	
				// Column selection mode when the target is 'thead>tr>th'
				if(!isHeadless) {
					if(/th/i.test(beginCell.tagName) && /thead/i.test(beginCell.parentNode.parentNode.tagName)) {
						var range = getRangeCells(allCells, endCell, endCell);
						for(var r = allCells.length; r > 0; r--) { // search from the last
							if(allCells[r-1][range.col2].parentNode.style.display !== 'none') { //tr
								endCell = allCells[r-1][range.col2];
								break;
							}
						}
					}
				}

				// Row selection mode the the target is 'tbody>tr>th:first-child'
				if(!isHeadless) {
					if(/th/i.test(beginCell.tagName) && /tbody/i.test(beginCell.parentNode.parentNode.tagName)) {
						if(beginCell.parentNode.cells[0] === beginCell) {
							var row = getRangeCells(bodyCells, endCell, endCell).row1;
							if(row != Infinity) { //Out of range. Ex.thead
								endCell = bodyCells[row][bodyCells[row].length-1];
							}
						}
					}
				}

				var range = getRangeCells(allCells, beginCell, endCell);
				
				// Continuous selection with Ctrl key
				if(withCrtlKey == false) {
					table.find('th.selectedcell, td.selectedcell').removeClass('selectedcell');
				}
				var rangeCells = range.cells;
				for(var row = 0; row < rangeCells.length; row++) {
					for(var col = 0; col < rangeCells[row].length; col++) {
						if(rangeCells[row][col].parentNode.style.display !== 'none') { //tr
							$(rangeCells[row][col]).addClass('selectedcell');
						}
					}
				}
				
				
				// Consider upside down case beginCell and endCell
				var a = $(beginCell).position();
				var b = $(endCell).position();
				var top = a.top < b.top ? a.top : b.top;
				var left = a.left < b.left ? a.left : b.left;
				var h = (a.top < b.top ? b.top - a.top + $(endCell).innerHeight() : a.top - b.top + $(beginCell).innerHeight());
				var w = (a.left < b.left ? b.left - a.left + $(endCell).innerWidth() :  a.left - b.left + $(beginCell).innerWidth());
				// jQuery Version3 has changed 'width' and 'height' calculation
				if(!isJQuery3) {
					h = h + 1; // +1 is cells border width
					w = w + 1; // +1 is cells border width
				}
				// Consider the cell has a rowspan or a colspan attribute.
				// Ex: If the endCell is 'colspan=2', The endCell's width is wider than 'w'.
				if(h < $(beginCell).innerHeight()) h = $(beginCell).innerHeight();
				if(h < $(endCell).innerHeight()) h = $(endCell).innerHeight();
				if(w < $(beginCell).innerWidth()) w = $(beginCell).innerWidth();
				if(w < $(endCell).innerWidth()) w = $(endCell).innerWidth();
				// Draw lines
				rectLine1.removeClass('dashed').css({top: top+'px', left: left+'px', height: h+'px', width: 0+'px'}).show();
				rectLine2.removeClass('dashed').css({top: top+'px', left: left+'px', height: 0+'px', width: w+'px'}).show();
				rectLine3.removeClass('dashed').css({top: top+h+'px', left: left+'px', height: 0+'px', width: w+'px'}).show();
				rectLine4.removeClass('dashed').css({top: top+'px', left: left+w+'px', height: h+'px', width: 0+'px'}).show();
				
				// Select the column head
				table.find('>thead>tr>th.selectedcol').removeClass('selectedcol');
				var col1 = range.col1;
				var col2 = range.col2;
				for(var row = 0; row < headCells.length; row++) {
					for(var col = col1; col <= col2; col++) {
						$(headCells[row][col]).addClass('selectedcol');
					}
				}

				updateSum();
				updateStat();
			}
			
			/*
			 * Clear selecting rectangle
			 */
			function clearRectangle() {
				isMouseDown = false;
				isSelecting = false;
				beginCell = null;
				endCell = null;
				rectLine1.hide();
				rectLine2.hide();
				rectLine3.hide();
				rectLine4.hide();
				copyLine1.hide();
				copyLine2.hide();
				copyLine3.hide();
				copyLine4.hide();
			}

			/*
			 * Sort rows in tbody. only tbody.
			 * @param isAsc true - ascending order
			 */
			function sortBody(isAsc) {
				var range = getRangeCells(allCells, clickedCell, clickedCell);
				var col = range.col1;
				var colLen = range.col2 - range.col1; // colspan

				// Set a mark
				table.find('>thead>tr>th>div.markbox>div.sortmark.active').removeClass('active').html(''); // Reset sorted mark first
				$(clickedCell).find('div.sortmark').html(isAsc ? '↓' : '↑').addClass('active'); // &darr; &uarr; 

				bodyCells.sort(function(rowA, rowB){
							var a = '';
							for(var i = 0; i <= colLen; i++) { a += getStaticText(rowA[col+i]); }
							var b = '';
							for(var i = 0; i <= colLen; i++) { b += getStaticText(rowB[col+i]); }
							var z = 0;
							// Compare number string as number.
							// Normaly: '2'> '12'
							// This function: '2'->2, '12'->12 then 12 > 2
							// In the 'isNaN' function, the logarithm '2e3' is also judged as a numerical value,
							// so I used my function 'isNum'.
							var a2 = a.replace(/,/g, ''); //-12,229 is treated as a number
							var b2 = b.replace(/,/g, '');
							if (a2 && b2 && isNum(a2) && isNum(b2)) {
								z = a2 - b2; // numerical comparison
							}
							else {
								z = (a > b ? 1 : (a < b ? -1 : 0)); // Character comparison
							}
							return (isAsc ? z : z * -1);
						});
				
				// replace allCells with sorted bodayCells
				allCells = headCells.concat(bodyCells).concat(footCells);
				

				var dummy = document.createElement('tbody'); //Avoide from Reflow, tbody unchain from DOM. This wey is very quick for IE.
				var tblElm = table[0];
				var tbody = tblElm.tBodies[0];
				tbody.style.display = 'none';

				// Once hide tbody to avoid from reflow.
				// This wey is effective for IE.
				tbody.style.display = 'none';
				for (var r = bodyCells.length - 1; r > 0; r--) {
					tbody.insertBefore(bodyCells[r-1][0].parentNode, bodyCells[r][0].parentNode);
				}
				tbody.style.display = '';

				// Hide copyLines. Note: not hide rectLines.
				copyLine1.hide();
				copyLine2.hide();
				copyLine3.hide();
				copyLine4.hide();
				updateStat();
			}
	
			/*
			 * Context menu of thead
			 */
			function headerCtrlHandler(event) {
				var target = event.target;
				
				// Hide menumark mark once
				// Then show menumark mark
				table.find('>thead>tr>th>div.markbox>div.menumark.active').removeClass('active');
				if($(target).hasClass('menumark')) {
					$(target).addClass('active');
					target = target.parentNode.parentNode;
				}
	
				if(/th/i.test(target.tagName)) {
					clickedCell = target;
					beginCell = clickedCell;
					endCell = clickedCell;
					isSelecting = true;
					selectRectangle(false);
					ctrlClearFilterHeader.children('div.filtercondition').html(filterCondition);
					ctrlSumHeader.children('span.sum').html(sumNum);
	
					var x = event.clientX, y = event.clientY;
					if(window.innerWidth < x + ctrlBoxHeader.outerWidth()) {
						x -= ctrlBoxHeader.outerWidth();
					}
					if(window.innerHeight < y + ctrlBoxHeader.outerHeight()) {
						y -= ctrlBoxHeader.outerHeight();
					}
					ctrlBoxHeader.css({'top': y + 'px', 'left': x + 'px'}).show();
	
					event.stopPropagation();
					event.preventDefault();
					return false;
				}
				else {
					clickedCell = null;
					updateStat();
				}
			}
	
			/*
			 * Context menu of tbody
			 */
			function bodyCtrlHandler(event) {
				var target = event.target;
				if(/th|td/i.test(target.tagName)) {
					clickedCell = event.target;
					// Keep cell selection if the cell is aleady selected
					// Otherwise select new.
					if(!$(target).hasClass('selectedcell')) {
						beginCell = clickedCell;
						endCell = clickedCell;
						
						// Continuous selection with Ctrl key
						var withCrtlKey = event.ctrlKey || event.metaKey
						if(withCrtlKey == false) {
							table.find('th.selectedcell, td.selectedcell').removeClass('selectedcell');
						}
					}
					isSelecting = true;
					selectRectangle(true); //Continuous selection
					var str = getStaticText(target);
					ctrlFilter.children('span.flt').html(str);
					ctrlNotFilter.children('span.flt').html(str);
					ctrlClearFilter.children('div.filtercondition').html(filterCondition);
					ctrlSum.children('span.sum').html(sumNum);

					var x = event.clientX, y = event.clientY;
					if(window.innerWidth < x + ctrlBox.outerWidth()) {
						x -= ctrlBox.outerWidth();
					}
					if(window.innerHeight < y + ctrlBox.outerHeight()) {
						y -= ctrlBox.outerHeight();
					}
					ctrlBox.css({'top': y + 'px', 'left': x + 'px'}).show();
	
					event.stopPropagation();
					event.preventDefault();
					return false;
				}
				else {
					clickedCell = null;
					ctrlBoxHeader.hide();
					ctrlBox.hide();
					table.find('>thead>tr>th>div.markbox>div.menumark').removeClass('active');
				}
			}
	
			/*
			 * Keep the hovering header in window
			 * TODO: internal scroll with 'overflow' attribute.
			 */
			function scrollHoverHeadHundler() {
				var tableTop = table.offset().top;
				var tbodyTop = table.find('>tbody').offset().top;
				var cloneH = hoverHead.outerHeight();
				if(window.scrollY > tbodyTop) {
					var y = window.scrollY - tableTop;
					y = (y > table.height()-cloneH ? table.height() - cloneH : y);
					hoverHead.css('top', y + 'px').fadeIn();
				}
				else {
						hoverHead.fadeOut();
				}
			}
			
			/*
			 * Update sum number
			 */
			function updateSum() {
				// Calculate every time is so slow.
				// Therefore this function execute after 500ms delay.
				if(isCalculatingSum === false) {
					isCalculatingSum = true;
					var func = function() {
						sumNum = 0;
						table.find('.selectedcell:visible').each(function() {
							var str = getStaticText(this);
							var m = str.match(/([,.\-0-9]+)/); //Numeric ducktype. If it looks like a number, calculate the sum.
							if(m) {
								str = m[0].replace(/,/g,'');
								if(isNum(str)) sumNum += parseFloat(str);
							}
						});
						sumBox.children('span').html(numFormat(sumNum));
						isCalculatingSum = false;
					};
					setTimeout(func, 500);
				}
			}

			/*
			 * Update row count
			 */
			function updateRowCount() {
				if(!isHeadless) {
					rowCountBox.children('span').html(table.find('>tbody>tr:visible').length);
				}
			}

			/*
			 * Update status
			 */
			function updateStat() {
				// Hide context menu
				ctrlBoxHeader.hide();
				ctrlBox.hide();

				// Ajust width of the hover header
				if(hoverHead) {
					hoverHead.css('width', table.innerWidth() + 'px');
					hoverHead.find('th').each(function(i) {
						$(this).css('width', $(this.refCell).innerWidth() + 'px');
					});
				}
			}

			// Apply initiale value
			updateSum();
			updateRowCount();
			updateStat();
		});
	};
}));

jQuery(function() { $(document).sheetjs(); });
