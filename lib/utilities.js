var letters = ["",
	"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
	"N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
];

// utility functions to use as a mixin
module.exports = {

	/**
	 * Converts a cell string into an indexed row and column ("A1" -> [0,0] or "0,0" -> [0,0])
	 * @param {String} cell The cell address (A1) or cell position (0,0)
	 * @returns {Array} A two element array, containing the row, then the column
	 */
	cellToRowCol: function (cell) {

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
				return [parseInt(match[4], 10) - 1, this.colToInt(match[2]) - 1];
			} else {
				throw new Error('invalid cell reference: ' + cell);
			}
		}
	},

	/**
	 * Converts a column letter into an indexed column number (A -> 0)
	 * @param col The column letter
	 * @returns {number} The indexed column number
	 */
	colToInt: function (col) {

		var n = 0;
		col = col.trim().split('');

		for (var i = 0; i < col.length; i++) {
			n *= 26;
			n += letters.indexOf(col[i]);
		}

		return n;
	}

};
