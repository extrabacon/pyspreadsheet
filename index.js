var path = require('path');
var PythonShell = require('python-shell');
PythonShell.defaultOptions = {
    scriptPath: path.join(__dirname, 'python')
};

exports.SpreadsheetReader = require('./lib/spreadsheet-reader');
exports.SpreadsheetWriter = require('./lib/spreadsheet-writer');
