var _ = require('./utilities');
var EventEmitter = require('events').EventEmitter;
var os = require('os');
var path = require('path');
var util = require('util');
var PythonShell = require('./python-shell');

/**
 * The spreadsheet writer API
 * @param {String} [filePath] The path of the file to write
 * @param {Object} [options]  The workbook options
 * @this SpreadsheetWriter
 * @constructor
 */
var SpreadsheetWriter = function (filePath, options) {
    var self = this;
    if (filePath && filePath.constructor === Object) {
        options = filePath;
        filePath = null;
    }

    EventEmitter.call(this);

    this.state = {
        status: 'new',
        sheets: [],
        anonymousFormatCount: 0,
        currentSheetIndex: -1,
        lastError: null
    };
    this.filePath = filePath || path.join(os.tmpdir(), _.randomString(8));

    // the Python shell process we use to render Excel files
    this.pyshell = new PythonShell('excel_writer.py');

    this.pyshell.on('message', function (message, value) {
        switch (message) {
            case 'open':
                self.state = 'open';
                self.filePath = value;
                self.emit('open');
                break;
            case 'close':
                self.state = 'closed';
                self.emit('close');
                break;
            case 'err':
                var error = _.extend(new Error(
                    'failed to execute command "' + value.command + '": ' + value.error
                ), value);
                self.emit('error', error);
                self.state.lastError = error;
                break;
        }
    }).on('error', function (err) {
        self.emit('error', err);
        self.state.lastError = err;
    });

    // by default, xlsx format is assumed and xlsxwriter will be used
    if (options && options.format) {
        switch (options.format.toLowerCase()) {
            case 'xls':
                // use the xlwt module for legacy XLS files
                this.pyshell.send('open', 'xlwt', this.filePath);
                break;
            case 'xlsx':
                // use the xlsxwriter module for OpenOffice files
                this.pyshell.send('open', 'xlsxwriter', this.filePath);
                break;
            default:
                throw new Error('unsupported format: ' + options.format);
        }
    } else {
        // default behavior
        this.pyshell.send('open', 'xlsxwriter', this.filePath);
    }

    this.pyshell.send('create_workbook', options);
};
util.inherits(SpreadsheetWriter, EventEmitter);

/**
 * Adds a sheet to the workbook
 * @param {String} [name] The name of the sheet to add
 * @param {Object} [options] The sheet options
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.addSheet = function (name, options) {
    this.state.sheets.push({
        name: name,
        index: this.state.sheets.length,
        currentRow: -1
    });
    this.state.currentSheetIndex = this.state.sheets.length - 1;
    this.pyshell.send('add_sheet', name);
    options && this.pyshell.send('set_sheet_settings', null, options);
    return this;
};

/**
 * Activates a previously added sheet
 * @param {Number|String} sheet The name or the index of the sheet to activate
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.activateSheet = function (sheet) {
    var sheetToActivate = _.find(this.state.sheets, function (s) {
        if (typeof sheet === 'number') {
            return s.index === sheet;
        } else {
            return s.name === sheet;
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
 * @param {Object} properties The formatting properties
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.addFormat = function (id, properties) {
    this.pyshell.send('format', id, properties);
    return this;
};

/**
 * Writes data to the current sheet
 * @param {Number|String} row The row index to write to - or a cell address such as "A1"
 * @param {Number} [col] The column index to write to - ignore if using a cell address
 * @param {*} data The data to write - use an array to write a row, or a 2-D array for multiple rows
 * @param {String|Object} [format] The format to apply to written cells - can be a format ID or a format object
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.write = function (row, col, data, format) {

    // look for a cell notation instead of row index (A1)
    if (typeof row === 'string') {
        var address = this.cellToRowCol(row);
        return this.write(address[0], address[1], col, data);
    }

    if (data === null || typeof data === 'undefined') return this; // nothing to write

    if (this.state.currentSheetIndex == -1) {
        // writing without a sheet - create a sheet now
        this.addSheet();
    }

    var sheet = this.state.sheets[this.state.currentSheetIndex];

    // if the row is not specified, use the next row from the last write
    if (row === -1) {
        row = sheet.currentRow + 1;
    }
    if (!col) {
        col = 0;
    }

    // keep track of where data is being written
    if (Array.isArray(data) && data.length > 0 && Array.isArray(data[0])) {
        // writing a jagged array, count rows in array
        sheet.currentRow = row + data.length - 1;
    } else {
        sheet.currentRow = row;
    }

    // look for an anonymous format
    if (format && format.constructor === Object) {
        var formatName = 'untitled_format_' + (++this.state.anonymousFormatCount);
        this.addFormat(formatName, format);
        format = formatName;
    }

    // convert values into transport-friendly JSON in order to avoid loss of fidelity
    function translate(v) {
        if (Array.isArray(v)) {
            return v.map(translate);
        } else if (v instanceof Date) {
            return { $date: v.getTime() };
        } else if (v === true) {
            return '=TRUE';
        } else if (v === false) {
            return '=FALSE';
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
 * @param {Function} callback The callback function
 * @returns {SpreadsheetWriter} The same writer instance for chaining calls
 */
SpreadsheetWriter.prototype.save = function (callback) {

    var self = this;
    callback = callback || function () {};

    this.pyshell.on('close', function () {
        if (self.state.lastError) {
            return callback(self.state.lastError);
        }
        return callback();
    });

    process.nextTick(function () {
        self.pyshell.end();
    });

    return this;
};

SpreadsheetWriter.prototype.cellToRowCol = _.cellToRowCol;
SpreadsheetWriter.prototype.colToInt = _.colToInt;
module.exports = SpreadsheetWriter;
