var xlsjs = require('node-xlsx'),
	path = require('path');

function decode_col(c) {
	var d = 0,
		i = 0;
	for (; i !== c.length; ++i) d = 26 * d + c.charCodeAt(i) - 64;
	return d - 1;
}

function decode_row(rowstr) {
	return Number(rowstr) - 1;
}

function split_cell(cstr) {
	return cstr.replace(/(\$?[A-Z]*)(\$?[0-9]*)/, "$1,$2").split(",");
}

function decode_cell(cstr) {
	var splt = split_cell(cstr);
	return {
		c: decode_col(splt[0]),
		r: decode_row(splt[1])
	};
}

function pad2(n) {
    return (n < 10 ? '0' : '') + n.toString();
}

function decode_range(range) {
	var x = range.split(":").map(decode_cell);
	return {
		s: x[0],
		e: x[x.length - 1]
	};
}

function ext(data) {
	if (data && data.t) {
		if (data.t == 's') {
			return data.v.replace(/"/g, '');
		} else if (data.t == 'n') {
			var date = data.w || '';
			var tab = date.split('/');
			if (tab.length > 2) {
				return pad2(tab[1]) + '/' + pad2(tab[0]) + '/' + tab[2];
			}
			return data.v;
		} else if (data.t == 'str') {
			if (data.rawt === 'n') {
				var date = data.v || '';
				if (date.length > 0) {
					var tab = date.split('/');
					return tab[1] + '/' + tab[0] + '/' + tab[2];
				} else {
					return date;
				}
			} else {
				return data.v || '';
			}
		} else {
			return data;
		}
	} else {
		return data;
	}
}

function getIndexFromCol(col) {
	return col.charCodeAt() - 65;
}

exports.load = function(cfg) {

	if (!cfg.filename) {
		console.warn('no filename ?');
		return [];
	}

	var data = xlsjs.parse(path.resolve(__dirname, cfg.filename)),
		sheet = data.worksheets[cfg.sheetNumber || 0],
		minRow = cfg.minRow || 0,
		max = sheet.maxRow,
		array = [];

	if (cfg.max && max>cfg.max) {
		max = cfg.max;
	}

	for (var i=minRow; i<max; ++i) {

		var obj = {};

		for (var prop in cfg.data) {
			if (cfg.data.hasOwnProperty(prop)) {
				var value = cfg.data[prop];
				var index = getIndexFromCol(value);
				var row = sheet.data[i];
				obj[prop] = row[index].value;/*ext(sheet.data[value + i]);*/
			}
		}

		array.push(obj);
	}

	return array;
};