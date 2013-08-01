var fs = require('fs'),
	pyspreadsheet = require('./lib');

/*var writer = new SpreadsheetWriter();

writer.write(require('./sample-workbook.json'));
writer.write(require('./sample-stream.json'));

writer.save(function (err, stream) {
	if (err) throw err;
	var output = fs.createWriteStream('output.xlsx');
	stream.pipe(output);
});

SpreadsheetReader('test.xlsx').parse(function (args) {
	
})*/


pyspreadsheet.read('output.xlsx', function (err, workbook) {
	console.log(workbook);
});
