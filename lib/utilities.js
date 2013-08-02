var letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

exports.cellToRowCol = function (cell) {

	var match;

	if (!cell) {
		return [0, 0];
	} else if (cell.indexOf(',') > 0) {
		// "row,col" notation
		match = /(\d+),(\d+)/.exec(cell);
		if (match) {
			return [parseInt(match[1], 10), parseInt(match[2], 10)];
		} else {
			throw new Error('invalid cell reference: ' + cell);
		}
	} else {
		// "A1" notation
		match = /(\$?)([A-Z]{1,3})(\$?)(\d+)/.exec(cell);
		if (match) {
			return [parseInt(match[4], 10) - 1, exports.colToInt(match[2]) - 1];
		} else {
			throw new Error('invalid cell reference: ' + cell);
		}
	}
};

exports.colToInt = function (col) {

	var n = 0;
	col = col.trim().split('');

	for (var i = 0; i < col.length; i++) {
		n *= 26;
		n += letters.indexOf(col[i]);
	}

	return n;
};
