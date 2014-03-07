/**
 * This example demonstrates how to read a large file as a stream, without loading everything in memory.
 * The reader triggers events as data is being received.
 */

var SpreadsheetReader = require('../lib').SpreadsheetReader;

new SpreadsheetReader('examples/sample.xlsx').on('open', function (workbook) {
    // file is open
    console.log('opened ' + workbook.file);
}).on('data', function (data) {
    // data is being received
    console.log('buffer contains %d rows from sheet "%s"', data.rows.length, data.sheet.name);
}).on('close', function () {
    // file is now closed
    console.log('file closed');
}).on('error', function (err) {
    throw err;
});
