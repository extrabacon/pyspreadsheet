/**
 * This example demonstrates how to read a simple file from memory
 */

var SpreadsheetReader = require('../lib').SpreadsheetReader;
var util = require('util');

SpreadsheetReader.read('examples/sample.xlsx', function (err, workbook) {
    if (err) throw err;
    console.log(util.inspect(workbook, { depth: null, colors: true }));
});
