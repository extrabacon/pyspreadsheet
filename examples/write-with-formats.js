/**
 * This example writes a spreadsheet with styles
 */

var fs = require('fs')
	pyspreadsheet = require('../lib');

var jsonWorkbook = {
	"sheets": {
		"employees": [
			{
				format: "header",
				values: ["ID", "Name", "Age", "Birthdate", "Points", "Active?"]
			},
			["1", "John Doe", 31, new Date('2012-01-01T01:35:33Z'), 55.34, true],
			["2", "John Dow", 34, new Date('2013-02-03T01:35:33Z'), 12.0002, false],
			["3", "John Doh", 33, new Date('2014-03-06T01:35:33Z'), 78.901, true],
			{
				format: "total",
				values: [null, null, null, "Total:", "=SUM(E2:E4)"]
			}
		]
	},
	"formats": {
		"header": {
			"fill": "black",
			"color": "white",
			"font": { "bold": true }
		},
		"total": {
			"font": { "bold": true },
			"alignment": "right"
		}
	}
};

pyspreadsheet.write(jsonWorkbook, function (err, stream) {
	if (err) throw err;
	stream.pipe(fs.createWriteStream('./examples/output.xlsx'));
});
