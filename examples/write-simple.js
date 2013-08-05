/**
 * This example writes a simple spreadsheet using a JSON object
 */

var fs = require('fs'),
	pyspreadsheet = require('../lib');

var jsonWorkbook = {
	"meta": {
		"title": "title",
		"subject": "subject",
		"author": "me",
		"comments": "written with pyspreadsheet"
	},
	"sheets": {
		"employees": [
			["ID", "Name", "Age", "Birthdate", "Points", "Active?"],
			["1", "John Doe", 31, new Date('2012-01-01T01:35:33Z'), 55.34, true],
			["2", "John Dow", 34, new Date('2013-02-03T01:35:33Z'), 12.0002, false],
			["3", "John Doh", 33, new Date('2014-03-06T01:35:33Z'), 78.901, true],
			[null, null, null, "Total:", "=SUM(E2:E4)"]
		]
	}
};

pyspreadsheet.write(jsonWorkbook, function (err, stream) {
	if (err) throw err;
	stream.pipe(fs.createWriteStream('./examples/output.xlsx'));
});
