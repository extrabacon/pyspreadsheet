var _ = require('underscore'),
	EventEmitter = require('events').EventEmitter,
	util = require('util'),
	spawn = require('child_process').spawn;

var SpreadsheetReader = function () {

};

/**
 * Parses a workbook, returning a workbook/sheet/row/cell structure
 *
 * @param {String|Array} path The path of the file(s) to load
 * @param {Object} options The options object (optional)
 * @param {Function} callback The callback method to invoke with the results
 */
exports.parse = function (path, options, callback) {

	if (_.isFunction(options)) {
		callback = options;
		options = null;
	}

	var reader = new XlrdParser(path, options),
		currentWorkbook,
		workbooks = [],
		errors;

	reader.on('open', function (workbook) {
		currentWorkbook = workbook;
		workbooks.push(workbook);
	});

	reader.on('data', function (data) {

		if (!currentWorkbook.sheets) {
			currentWorkbook.sheets = [];
		}

		var sheet = _.find(currentWorkbook.sheets, function (s) {
			return s.index === data.sheet.index;
		});

		if (!sheet) {
			sheet = _.extend({ rows: [] }, data.sheet);
			currentWorkbook.sheets.push(sheet);
			currentWorkbook.sheets[sheet.name] = sheet;
		}

		data.rows.forEach(function (row) {
			sheet.rows.push(row);
		});

	});

	reader.on('error', function (err) {
		if (!errors) {
			errors = err;
		} else {
			errors = [errors];
			errors.push(err);
		}
	});

	reader.on('close', function () {

		callback = callback || function () {};

		if (!errors && !workbooks.length) {
			errors = _.extend(new Error('file not found: ' + path), { id: 'file_not_found' });
		}

		if (workbooks.length === 0) {
			callback(errors, null);
		} else if (workbooks.length === 1) {
			callback(errors, workbooks[0])
		} else {
			callback(errors, workbooks);
		}
	});

};
