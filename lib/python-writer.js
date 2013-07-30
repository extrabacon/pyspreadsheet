var _ = require('underscore'),
	async = require('async'),
	fs = require('fs'),
	PythonShell = require('./python-shell'),
	SpreadsheetParser = require('./spreadsheet-parser');

var PythonExcelWriter = function (workbook) {

	var self = this,
		sheetCount = 0,
		currentRow = -1;

	PythonExcelWriter.prototype.write = function (data) {
		self.parser.parse(data);
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
	};

	this.filePath = null;
	this.state = 'initial';
	this.errors = [];
	this.parser = new SpreadsheetParser();

	// by default, xlsx format is assumed and xlsxwriter will be used without a 'setup' command
	if (workbook && workbook.format) {
		switch (workbook.format) {
			case 'xls':
				// use the xlwt module
				this.pyshell.send('setup', 'xlwt');
				break;
			default:
				// use the xlsxwriter module as default
				this.pyshell.send('setup', 'xlsxwriter');
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

};

module.exports = PythonExcelWriter;
