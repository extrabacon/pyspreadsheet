var _ = require('underscore'),
	async = require('async'),
	fs = require('fs'),
	PythonShell = require('./python-shell'),
	utils = require('./utilities'),
	JSONSpreadsheetParser = require('./json-spreadsheet-parser');

/**
 *
 * @param workbook
 * @constructor
 */
var PythonExcelWriter = function (workbook) {

	var self = this,
		sheets = [],
		anonymousFormat = 0,
		currentSheet = -1;

	/**
	 * Adds a sheet to the workbook
	 * @param name
	 * @param options
	 * @returns {*}
	 */
	PythonExcelWriter.prototype.addSheet = function (name, options) {

		sheets.push({
			name: name,
			index: sheets.length,
			currentRow: -1
		});
		currentSheet = sheets.length;

		self.pyshell.send('add_sheet', name);

		if (options) {
			self.pyshell.send('set_sheet_settings', null, options);
		}

		return self;
	};

	/**
	 * Activates a previously added sheet
	 * @param sheet
	 * @returns {*}
	 */
	PythonExcelWriter.prototype.activateSheet = function (sheet) {

		var byindex = _.isNumber(sheet);
		var sheetToActivate = _.find(sheets, function (s) {
			if (byindex) {
				return s.index === sheet;
			} else {
				return s.name == sheet;
			}
		});

		if (sheetToActivate) {
			currentSheet = sheetToActivate.index;
			self.pyshell.send('activate_sheet', sheet);
		} else {
			throw new Error('sheet not found: ' + sheet);
		}

		return self;
	};

	/**
	 * Registers a reusable format which can be used with "write"
	 * @param id
	 * @param format
	 * @returns {*}
	 */
	PythonExcelWriter.prototype.addFormat = function (id, format) {
		self.pyshell.send('format', format);
		return self;
	};

	/**
	 * Writes data to the current sheet
	 * @param row
	 * @param col
	 * @param data
	 * @param format
	 * @returns {*}
	 */
	PythonExcelWriter.prototype.write = function (row, col, data, format) {

		// writing without a sheet - create a sheet now
		if (currentSheet == -1) {
			self.addSheet();
		}

		var sheet = sheets[currentSheet];

		// look for a cell notation instead of row index (A1)
		if (_.isString(row)) {
			var r = utils.cellToRowCol(row);
			row = r[0];
			col = r[1];
			data = arguments.length > 1 ? arguments[1] : null;
			format = arguments.length > 2 ? arguments[2] : null;
		}

		// if the row is not specified, use the next row from the last write
		if (row === -1) {
			row = sheet.currentRow + 1;
		}
		if (col == null) {
			col = 0;
		}

		// keep track of where data is being written
		if (_.isArray(data) && _.isArray(data[0])) {
			// writing a jagged array, count rows in array
			sheet.currentRow = row + data.length - 1;
		} else {
			sheet.currentRow = row;
		}

		// look for an inline format
		if (_.isObject(format)) {
			var formatName = 'untitled_format_' + anonymousFormat;
			self.addFormat(formatName, format);
			anonymousFormat++;
			format = formatName;
		}

		self.pyshell.send('write', row, col, data, format);
		return self;
	};

	PythonExcelWriter.prototype.append = function (data, format) {
		self.write(-1, 0, data, format);
	};

	PythonExcelWriter.prototype.save = function (callback) {

		callback = callback || function () {};

		self.pyshell.on('close', function () {

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

		return self;
	};

	PythonExcelWriter.prototype.destroy = function (callback) {

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

		return self;
	};

	this.filePath = null;
	this.state = 'initial';
	this.errors = [];
	this.parser = new JSONSpreadsheetParser();

	// by default, xlsx format is assumed and xlsxwriter will be used without a 'setup' command
	if (workbook && workbook.format) {
		switch (workbook.format.toLowerCase()) {
			case 'xls':
				// use the xlwt module for legacy XLS files
				this.pyshell.send('setup', 'xlwt');
				break;
			case 'xlsx':
				// use the xlsxwriter module for OpenOffice files
				this.pyshell.send('setup', 'xlsxwriter');
				break;
			default:
				throw new Error('unknown format "' + workbook.format + '", supported formats are "xlsx" and "xls"');
				break;
		}
	}

	// command events are fired when invoking "write"
	this.parser.on('command', function (command) {

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
	});

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

	this.pyshell.send('create_workbook', workbook && workbook.meta ? { properties: workbook.meta } : null);

	if (workbook) {
		this.write(workbook);
	}

};

module.exports = PythonExcelWriter;
