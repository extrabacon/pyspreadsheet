var _ = require('underscore'),
	EventEmitter = require('events').EventEmitter,
	util = require('util');

var letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

var SpreadsheetParser = function (data) {

	var self = this;
	EventEmitter.call(this);

	function cellToRowCol(cell) {

		var match;

		if (!cell) {
			return [0, 0];
		} else if (cell.indexOf(',') > 0) {
			// "row,col" notation
			match = /(\d+),(\d+)/.exec(cell);
			if (match) {
				return [parseInt(match[1], 10), parseInt(match[2], 10)];
			} else {
				self.emit('error', new Error('invalid cell reference: ' + cell));
				return null;
			}
		} else {
			// "A1" notation
			match = /(\$?)([A-Z]{1,3})(\$?)(\d+)/.exec(cell);
			if (match) {
				return [parseInt(match[4], 10) - 1, colToInt(match[2]) - 1];
			} else {
				self.emit('error', new Error('invalid cell reference: ' + cell));
				return [0, 0];
			}
		}
	}

	function colToInt(col) {

		var n = 0;
		col = col.trim().split('');

		for (var i = 0; i < col.length; i++) {
			n *= 26;
			n += letters.indexOf(col[i]);
		}

		return n;
	}

	function parseFormats(data) {
		if (_.isArray(data)) {
			data.forEach(parseFormats);
		} else if (_.isObject(data)) {
			_.each(data, function (format, id) {
				self.emit('command', 'format', id, format);
			});
		} else {
			self.emit('error', new Error('invalid formats, expected an array or an object'));
		}
	}

	function parseSheets(data) {

		if (_.isArray(data)) {
			data.forEach(parseSheets);
		} else if (_.isObject(data)) {

			_.each(data, function (sheet, id) {
				if (_.isArray(sheet)) {
					addSheet(_.isString(id) ? id : null);
					parseSheetData(sheet);
				} else if (_.isObject(sheet)) {
					addSheet(sheet.name || _.isString(id) ? id : null, sheet.options);
					parseSheetData(sheet.data || sheet);
				}
			});

		} else if (_.isString(data)) {
			data.split(',').forEach(function (sheetName) {
				addSheet(sheetName);
			});
		}

	}

	function parseSheetData(data) {

		if (_.isArray(data)) {
			write(null, 0, data);
		} else if (_.isObject(data)) {

			if (data.v || data.values || data.value) {
				// parse as a single range
				parseRangeData(data);
			} else {
				_.each(data, function (value, cell) {
					parseCellData(cell, value);
				});
			}

		} else {
			self.emit('error', new Error('invalid sheet data'));
		}

	}

	function parseCellData(cell, value) {

		var offset = [0, 0];

		if (_.isString(cell)) {
			offset = cellToRowCol(cell);
		} else {
			offset = [parseInt(cell, 10), 0];
		}

		if (_.isArray(value)) {
			write(offset[0], offset[1], value);
		} else if (_.isObject(value)) {
			parseRangeData(value, offset);
		}
	}

	function parseRangeData(data, offset) {

		var row = 0, col = 0,
			value = data.values || data.value || data.v,
			format = data.format || data.f;

		offset = offset || data.offset;

		if (offset) {
			if (_.isArray(offset)) {
				row = offset[0];
				col = offset[1];
			} else {
				row = parseInt(offset, 10);
			}
		}

		write(row, col, value, format)
	}

	function addSheet(name, options) {
		self.emit('command', 'add_sheet', name);
		if (options) {
			self.emit('command', 'set_sheet', null, options);
		}
	}

	function write(row, col, value, format) {
		self.emit('command', 'write', row, col, value, format);
	}

	function parse(data) {

		if (!data) return;

		if (_.isArray(data)) {
			write(null, 0, data);
		} else if (data.sheets || data.formats) {
			if (data.formats) {
				parseFormats(data.formats);
			}
			if (data.sheets) {
				parseSheets(data.sheets);
			}
		} else {
			_.each(data, function (value, key) {
				switch (key) {
					case '$new-sheet':
						if (_.isString(value)) {
							addSheet(value);
						} else {
							addSheet(value.name, value.options);
						}
						break;
					case '$change-sheet':
						if (/\d+/.test(value)) {
							value = parseInt(value, 10);
						}
						self.emit('command', 'set_sheet', value);
						break;
					default:
						parseCellData(key, value);
						break;
				}
			});
		}
	}

	SpreadsheetParser.prototype.parse = function (data) {
		if (_.isArray(data)) {
			data.forEach(parse);
		} else {
			parse(data);
		}
	};

	if (data) {
		this.parse(data);
	}
};

util.inherits(SpreadsheetParser, EventEmitter);
module.exports = SpreadsheetParser;
