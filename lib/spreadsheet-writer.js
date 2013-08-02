var _ = require('underscore'),
	async = require('async'),
	fs = require('fs'),
	PythonShell = require('./python-shell');

/**
 * The spreadsheet writer API
 * @param {Object} [options] The workbook options
 * @this SpreadsheetWriter
 * @constructor
 */
var SpreadsheetWriter = function (options) {

	var self = this;

	// mixin with utilities
	_.extend(this, require('./utilities'));

	this.state = {
		status: 'new',
		sheets: [],
		anonymousFormatCount: 0,
		currentSheetIndex: -1
	};
	this.filePath = null;
	this.errors = [];

	// by default, xlsx format is assumed and xlsxwriter will be used without a 'setup' command
	if (options && options.format) {
		switch (options.format.toLowerCase()) {
			case 'xls':
				// use the xlwt module for legacy XLS files
				this.pyshell.send('setup', 'xlwt');
				break;
			case 'xlsx':
				// use the xlsxwriter module for OpenOffice files
				this.pyshell.send('setup', 'xlsxwriter');
				break;
			default:
				throw new Error('unknown format "' + options.format + '", supported formats are "xlsx" and "xls"');
				break;
		}
	}

	// command events are fired when invoking "write"
	/*this.parser.on('command', function (command) {

		var args = _.rest(arguments);

		if (command == 'write') {
			// writing without a sheet - create a sheet now
			if (sheetCount == 0) {
				self.send('add_sheet');
				sheetCount++;
			}
			// if the row is not specified, use the next row from the last write
			if (args[0] == null) {
				args[0] = currentRow + 1;
			}
			// keep track of where data is being written
			if (_.isArray(args[2]) && _.isArray(args[2][0])) {
				// writing a jagged array, count rows in array
				currentRow = args[0] + args[2].length - 1;
			} else {
				currentRow = args[0];
			}
		} else if (command == 'add_sheet') {
			sheetCount++;
			currentRow = -1;
		} else if (command == 'set_sheet') {
			if (args[0] == null) {
				args[0] = sheetCount - 1;
			}
		}

		self.pyshell.send.apply(self, _.chain(args).unshift(command).value());

	}).on('error', function (err) {
		self.errors.push(err);
	});*/

	// the Python shell process we use to render Excel files
	this.pyshell = new PythonShell('python/excel_writer.py', null, { stdio: 'pipe' });

	this.pyshell.on('message', function (message, value) {
		switch (message) {
			case 'open':
				self.state = 'open';
				self.filePath = value;
				break;
			case 'close':
				self.state = 'closed';
				self.filePath = value;
				break;
			case 'err':
				var msg = ['an error occurred in the Python shell while executing the command "',
						value.command,
						'" >> ',
						value.error
					].join('');
				self.errors.push(_.extend(new Error(msg), value));
				break;
		}
	}).on('error', function (err) {
		self.errors.push(err);
	});

	this.pyshell.send('create_workbook', options && options.meta ? { properties: options.meta } : null);
	return this;
};

/**
 * Adds a sheet to the workbook
 * @param {String} [name] The name of the sheet to add
 * @param {Object} [options] The sheet options
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.addSheet = function (name, options) {

	this.state.sheets.push({
		name: name,
		index: sheets.length,
		currentRow: -1
	});

	this.state.currentSheetIndex = this.state.sheets.length;

	this.pyshell.send('add_sheet', name);

	if (options) {
		this.pyshell.send('set_sheet_settings', null, options);
	}

	return this;
};

/**
 * Activates a previously added sheet
 * @param {Number|String} sheet The name or the index of the sheet to activate
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.activateSheet = function (sheet) {

	var byindex = _.isNumber(sheet);

	var sheetToActivate = _.find(this.state.sheets, function (s) {
		if (byindex) {
			return s.index === sheet;
		} else {
			return s.name == sheet;
		}
	});

	if (sheetToActivate) {
		this.state.currentSheetIndex = sheetToActivate.index;
		this.pyshell.send('activate_sheet', sheet);
	} else {
		throw new Error('sheet not found: ' + sheet);
	}

	return this;
};

/**
 * Registers a reusable format which can be used with "write"
 * @param {String} id The format ID to use as the reference
 * @param {Object} format The format object
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.addFormat = function (id, format) {
	this.pyshell.send('format', format);
	return this;
};

/**
 * Writes data to the current sheet
 * @param {Number|String} row The row index to write to - or a cell address such as "A1"
 * @param {Number} [col] The column index to write to
 * @param {Array} data The data to write - use an array to write an entire row, or a 2-D array for multiple rows
 * @param format
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.write = function (row, col, data, format) {

	// writing without a sheet - create a sheet now
	if (this.state.currentSheetIndex == -1) {
		this.addSheet();
	}

	var sheet = this.state.sheets[this.state.currentSheetIndex];

	// look for a cell notation instead of row index (A1)
	if (_.isString(row)) {
		var r = this.cellToRowCol(row);
		row = r[0];
		col = r[1];
		data = arguments.length > 1 ? arguments[1] : null;
		format = arguments.length > 2 ? arguments[2] : null;
	}

	// if the row is not specified, use the next row from the last write
	if (row === -1) {
		row = sheet.currentRow + 1;
	}
	if (!col) {
		col = 0;
	}

	// keep track of where data is being written
	if (_.isArray(data) && data.length > 0 && _.isArray(data[0])) {
		// writing a jagged array, count rows in array
		sheet.currentRow = row + data.length - 1;
	} else {
		sheet.currentRow = row;
	}

	// look for an anonymous format
	if (_.isObject(format)) {
		var formatName = 'untitled_format_' + this.state.anonymousFormatCount;
		this.addFormat(formatName, format);
		this.state.anonymousFormatCount++;
		format = formatName;
	}

	// convert values into transport-friendly JSON in order to avoid loss of fidelity
	function translate(v) {
		if (_.isArray(v)) {
			return _.map(v, translate);
		} else if (_.isDate(v)) {
			return { $date: v };
		} else if (v === true) {
			return "=TRUE";
		} else if (v === false) {
			return "=FALSE";
		}
		return v;
	}
	data = translate(data);

	this.pyshell.send('write', row, col, data, format);
	return this;
};

/**
 * Appends data to the current sheet
 * @param {Array} data The data to append - a 2-D array will append multiple rows at once
 * @param {String|Object} [format] The format to apply, using its name or an anonymous format object
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.append = function (data, format) {
	return this.write(-1, 0, data, format);
};

/**
 * Saves this workbook, committing all changes
 * @param {Function|String} callback The callback function receiving the resulting stream - use a string to write
 * the stream to a file automatically instead
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.save = function (callback) {

	var self = this;
	callback = callback || function () {};

	// if callback is a string, use a callback to pipe the stream to the path instead
	if (_.isString(callback)) {
		var path = callback;
		callback = function (err, stream) {
			stream.pipe(fs.createWriteStream(path));
		}
	}

	this.pyshell.on('close', function () {

		if (self.errors.length > 0) {
			return callback(self.errors.length === 1 ? self.errors[0] : self.errors);
		}

		// create a readable stream with the output file
		var stream = fs.createReadStream(self.filePath);

		stream.on('close', function () {
			// destroy the file automatically upon closing the stream
			self.destroy();
		});

		return callback(null, stream);
	});

	// nextTick ensures that all data has been parsed
	process.nextTick(function () {
		self.pyshell.end();
	});

	return this;
};

/**
 * Destroys this instance, releasing all resources
 * Called automatically when the stream returned from "save" is closed
 * @param {Function} [callback] The function to invoke when the operation is complete
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.destroy = function (callback) {

	var self = this;

	async.series([
		function (callback) {
			if (!self.pyshell.terminated) {
				self.pyshell.on('close', callback);
				self.pyshell.childProcess.kill();
			} else {
				callback();
			}
		},
		function (callback) {
			if (self.filePath) {
				fs.exists(self.filePath, function (exists) {
					if (exists) {
						fs.unlink(self.filePath, callback);
					} else {
						callback();
					}
				});
			} else {
				callback();
			}
		}
	], callback);

	return this;
};

module.exports = SpreadsheetWriter;
