var _ = require('underscore'),
	EventEmitter = require('events').EventEmitter,
	util = require('util'),
	SpreadsheetWriter = require('./spreadsheet-writer');

var JSONSpreadsheetParser = function (data) {

	var self = this;
	EventEmitter.call(this);

	this.writer = new SpreadsheetWriter(data);

	if (data.sheets) {
		_.each(data.sheets, function (sheet, id) {
			self.writer.addSheet(sheet.name || id, sheet.options);
			self.parse(sheet.data || sheet);
		});
	}

	if (data.formats) {
		_.each(data.formats, function (format, id) {
			self.writer.addFormat(id, format);
		});
	}

};

// [{ "A1": [$values] }]
// [{ "0,0": [$values] }]
// { format: "", values: [] }
// [0, 1, 2]
// [[0, 1, 2], [3, 4, 5]]
// [{ format: "", value: 0 }, {}]

JSONSpreadsheetParser.prototype.parse = function (data, cell) {

	var self = this;

	function write(cell, data, format) {
		if (!cell) {
			self.writer.write(-1, 0, data, format);
		} else {
			self.writer.write(cell, data, format);
		}
	}

	function valueOf() {
		return data.values || data.value || data.v;
	}

	if (_.isArray(data) && data.length > 0) {
		if (!_.isArray(data[0]) && _.isObject(data[0])) {
			// if an array of objects, parse each element individually
			return _.each(data, function (el) { self.parse(el, cell) });
		} else {
			// if an array of values, write the data
			write(cell, data);
		}
	}

	if (valueOf()) {
		write(cell, valueOf(), data.format || data.f);
	} else {
		_.each(data, function (value, key) {
			self.parse(value, ''+key);
		});
	}

};

util.inherits(JSONSpreadsheetParser, EventEmitter);
module.exports = JSONSpreadsheetParser;
