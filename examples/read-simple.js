/**
 * This example demonstrates how to read a simple file
 * @type {*}
 */

var util = require('util'),
	pyspreadsheet = require('../lib');

pyspreadsheet.read('./examples/sample.xlsx', function (err, workbook) {
	if (err) throw err;
	console.log(util.inspect(workbook, { depth: null, colors: true }));
});
