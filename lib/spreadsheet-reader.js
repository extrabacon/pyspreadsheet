var _ = require('./utilities');
var EventEmitter = require('events').EventEmitter;
var util = require('util');
var PythonShell = require('python-shell');

/**
 * The spreadsheet reader API
 * @param path The path of the file to read
 * @param [options] The reading options
 * @this SpreadsheetReader
 * @constructor
 */
var SpreadsheetReader = function (path, options) {
    var self = this;
    var workbook;
    var sheet;
    var currentRow;
    var rowIndex = -1;
    var rowBuffer = [];
    var inputEnded = false;
    var consoleEndedOrError = false;
    var isClosed = false;
    options = _.extend({}, SpreadsheetReader.defaultOptions, options);

    EventEmitter.call(this);

    function flushRows() {
        // if no accumulated rows, skip
        if (!rowBuffer.length) return;

        // emit an event with the accumulated rows
        self.emit('data', {
            workbook: workbook,
            sheet: sheet,
            rows: rowBuffer
        });

        // reset buffer for the next iteration
        rowBuffer = [];
    }

    this.pyshell = new PythonShell('excel_reader.py', {
        args: formatArgs(path, options),
        mode: 'json'
    });

    this.pyshell.on('message', function (message) {
        var value = message[1];
        switch (message[0]) {
            case 'w':
                workbook = {
                    file: value.file,
                    meta: {
                        user: value.user,
                        sheets: value.sheets
                    }
                };
                self.emit('open', workbook);
                break;
            case 's':
                // if we are changing sheets, flush rows immediately
                rowIndex = -1;
                sheet && flushRows();
                sheet = {
                    index: value.index,
                    name: value.name,
                    bounds: {
                        rows: value.rows,
                        columns: value.columns
                    },
                    visibility: (function () {
                        switch (value.visibility) {
                            case 1: return 'hidden';
                            case 2: return 'very hidden';
                            default: return 'visible';
                        }
                    })()
                };
                break;
            case 'c':
                // switching rows
                if (value[0] !== rowIndex) {
                    // flush accumulated rows if the buffer limit has been reached
                    rowBuffer.length >= options.bufferSize && flushRows();
                    // append a new row to the buffer
                    currentRow = [];
                    rowIndex = value[0];
                    rowBuffer.push(currentRow);
                }
                // append cell to current row
                currentRow.push({
                    row: value[0],
                    column: value[1],
                    address: value[2],
                    value: parseValue(value[3])
                });
                break;
            case 'x':
                inputEnded = true;
                if (consoleEndedOrError && !isClosed) {
                    isClosed = true;
                    // flush remaining rows, then close
                    flushRows();
                    self.emit('close');
                }
                break;
            case 'error':
                self.emit('error', _.extend(self.pyshell.parseError(value.traceback), value));
                break;
        }
    }).on('error', function (err) {
        // Safety disable to avoid stucks 
        consoleEndedOrError = true;
        self.emit('error', err);
    }).on('close', function () {
        consoleEndedOrError = true; 
        if (inputEnded && !isClosed) {
            isClosed = true;
            // flush remaining rows, then close
            flushRows();
            self.emit('close');        
        }
    });
};
util.inherits(SpreadsheetReader, EventEmitter);

SpreadsheetReader.defaultOptions = {
    bufferSize: 20
};

function formatArgs(path, options) {
    var args = [];

    if (options.meta) {
        args.push('-m');
    }
    if (options.hasOwnProperty('sheets') && Array.isArray(options.sheets)) {
        args.push(options.sheets.map(function (s) {
            return ['-s', s];
        }));
    } else if (options.hasOwnProperty('sheet')) {
        args.push(['-s', options.sheet]);
    }
    if (options.hasOwnProperty('maxRows')) {
        args.push(['-r', options.maxRows]);
    }

    args.push(path);
    return _.flatten(args);
}

function parseValue(value) {
    if (Array.isArray(value)) {
        // parse non-native data types
        if (value[0] === 'date') {
            return new Date(
                value[1], value[2] - 1, value[3],
                value[4], value[5], value[6]
            );
        } else if (value[0] === 'error') {
            return _.extend(new Error(value[1]), { id: 'cell_error', errorCode: value[1] });
        } else if (value[0] === 'empty') {
            return null;
        }
    }
    return value;
}

/**
 * Reads an entire spreadsheet file into a workbook object
 * @param path The path of the spreadsheet file to read
 * @param {Object} [options] The reading options
 * @param {Function} callback The callback function to invoke with the resulting workbook object
 * @returns {SpreadsheetReader} The SpreadsheetReader instance to use to read the file
 */
SpreadsheetReader.read = function (path, options, callback) {
    if (typeof options === 'function') {
        callback = options;
        options = null;
    }

    var reader = new SpreadsheetReader(path, options);
    var currentWorkbook;
    var workbooks = [];
    var errors;

    function getCell(address) {
        var pos = _.cellToRowCol(address);
        return this.rows[pos[0]][pos[1]];
    }

    reader.on('open', function (workbook) {
        currentWorkbook = workbook;
        workbooks.push(workbook);
    }).on('data', function (data) {
        currentWorkbook.sheets = currentWorkbook.sheets || [];

        var sheet = _.find(currentWorkbook.sheets, function (s) {
            return s.index === data.sheet.index;
        });

        if (!sheet) {
            sheet = _.extend({ rows: [] }, data.sheet);
            sheet.cell = getCell.bind(sheet);
            currentWorkbook.sheets.push(sheet);
            currentWorkbook.sheets[sheet.name] = sheet;
        }

        data.rows.forEach(function (row) {
            sheet.rows.push(row);
        });

    }).on('error', function (err) {
        if (!errors) {
            errors = err;
        } else {
            errors = [errors];
            errors.push(err);
        }
    }).on('close', function () {
        callback = callback || function () {};

        if (!errors && !workbooks.length) {
            errors = _.extend(new Error('file not found: ' + path), { id: 'file_not_found' });
        }

        if (!workbooks.length) {
            return callback(errors, null);
        } else if (workbooks.length === 1) {
            return callback(errors, workbooks[0]);
        }
        return callback(errors, workbooks);
    });

    return reader;
};

module.exports = SpreadsheetReader;
