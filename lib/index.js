var _ = require('underscore'),
	SpreadsheetReader = require('./spreadsheet-reader'),
	SpreadsheetWriter = require('./spreadsheet-writer');

/**
 * Reads a spreadsheet file
 * @param path The path of the spreadsheet file to read
 * @param {Object} [options] The reading options
 * @param {Function} [callback] The callback function to invoke with the resulting workbook object - ignore to stream
 * the file manually
 * @returns {SpreadsheetReader} The SpreadsheetReader instance to use to read the file
 */
exports.read = function (path, options, callback) {

	if (_.isFunction(options)) {
		callback = options;
		options = null;
	}

	var reader = new SpreadsheetReader(path, options),
		currentWorkbook,
		workbooks = [],
		errors;

	// if a callback is specified, stream the results into a workbook/sheet/row structure
	if (callback) {

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
				// TODO: should use an accessor
				//currentWorkbook.sheets[sheet.name] = sheet;
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
	}

	return reader;
};

/**
 * Writes a spreadsheet file
 * @param {Object} workbook The workbook object to write
 * @param {Function} [callback] The callback function to invoke with the resulting stream - ignore to write and save
 * data manually
 * @returns {SpreadsheetWriter} The SpreadsheetWriter instance to use to write the file
 */
exports.write = function (workbook, callback) {

	var writer = new SpreadsheetWriter(workbook);

	// if a callback is specified, save the workbook right away
	if (callback) {
		writer.save(callback);
	}

	return writer;
};
