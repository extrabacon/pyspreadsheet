var _ = require('underscore'),
	fs = require('fs'),
	PythonExcelWriter = require('./python-writer');

/**
 * @this SpreadsheetWriter
 * @param workbook The workbook object to write
 * @constructor
 */
var SpreadsheetWriter = function (workbook) {

	var self = this;

	if (workbook && workbook.format) {

		//noinspection FallthroughInSwitchStatementJS
		switch (workbook.format.toLowerCase()) {
			case 'xls':
			case 'xlsx':
				this.writer = new PythonExcelWriter(workbook);
				break;
			default:
				throw new Error('unknown workbook format: ' + workbook.format);
		}

	} else {
		// use PythonExcelWriter as the default
		this.writer = new PythonExcelWriter(workbook);
	}

	/**
	 * Writes data to the spreadsheet
	 * @param data The data to write
	 */
	SpreadsheetWriter.prototype.write = function (data) {
		self.writer.write(data);
	};

	/**
	 * Saves the spreadsheet into a file, returning a stream
	 * @param callback
	 */
	SpreadsheetWriter.prototype.save = function (callback) {
		self.writer.save(callback);
	};

	/**
	 * Destroys the instance, releasing resources
	 * @param callback
	 */
	SpreadsheetWriter.prototype.destroy = function (callback) {
		self.writer.destroy(callback);
	};
};

module.exports = SpreadsheetWriter;
