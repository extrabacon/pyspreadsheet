var _ = require('./utilities');
var EventEmitter = require('events').EventEmitter;
var util = require('util');
var PythonShell = require('./python-shell');

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
    var bufferSize = (options || {}).bufferSize || 50;

    EventEmitter.call(this);

    function formatArgs() {
        var args = [];
        options = options || {};

        if (options.meta) {
            args.push('-m');
        }
        if (options.hasOwnProperty('sheet') && !options.hasOwnProperty('sheets')) {
            options.sheets = options.sheet;
            delete options.sheet;
        }
        if (options.hasOwnProperty('sheets')) {
            if (Array.isArray(options.sheets)) {
                args.push(options.sheets.map(function (s) {
                    return ['-s', s];
                }));
            } else {
                args.push(['-s', options.sheets]);
            }
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

    function flushRows() {
        // if no accumulated rows, skip
        if (!rowBuffer.length) return;

        // emit an event with the accumulated rows
        self.emit('data', {
            workbook: workbook,
            sheet: sheet,
            rows: rowBuffer.slice(0)
        });

        // reset buffer for the next iteration
        rowBuffer = [];
    }

    this.pyshell = new PythonShell('excel_reader.py', formatArgs());

    this.pyshell.on('message', function (message, value) {
        switch (message) {
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
                if (value.r !== rowIndex) {
                    // flush accumulated rows if the buffer limit has been reached
                    rowBuffer.length >= bufferSize && flushRows();
                    // append a new row to the buffer
                    currentRow = [];
                    rowIndex = value.r;
                    rowBuffer.push(currentRow);
                }
                // append cell to current row
                currentRow.push({
                    row: value.r,
                    column: value.c,
                    address: value.a,
                    value: parseValue(value.v)
                });
                break;
            case 'err':
                self.emit('error', _.extend(
                    new Error('[' + value.exception + '] ' + value.type + ': ' + value.details),
                    value)
                );
                break;
        }
    }).on('error', function (err) {
        self.emit('error', err);
    }).on('close', function () {
        // flush remaining rows, then close
        flushRows();
        self.emit('close');
    });
};
util.inherits(SpreadsheetReader, EventEmitter);

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
        var pos = reader.cellToRowCol(address);
        return this.rows[pos[0]][pos[1]];
    }

    reader.on('open', function (workbook) {
        currentWorkbook = workbook;
        workbooks.push(workbook);
    }).on('data', function (data) {

        if (!currentWorkbook.sheets) {
            currentWorkbook.sheets = [];
        }

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

SpreadsheetReader.prototype.cellToRowCol = _.cellToRowCol;
SpreadsheetReader.prototype.colToInt = _.colToInt;
module.exports = SpreadsheetReader;
