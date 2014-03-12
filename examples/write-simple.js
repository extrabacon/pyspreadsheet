/**
 * This example writes a simple spreadsheet
 */

var SpreadsheetWriter = require('../lib').SpreadsheetWriter;
var writer = new SpreadsheetWriter('examples/output.xlsx');

writer.addFormat('heading', { font: { bold: true } });
writer.write(0, 0, ['ID', 'Name', 'Age', 'Birthdate', 'Balance', 'Active?'], 'heading');

writer.append([
    ['1', 'John Doe', 31, new Date('2012-01-01T01:35:33Z'), 55.34, true],
    ['2', 'John Dow', 34, new Date('2013-02-03T01:35:33Z'), 12.0002, false],
    ['3', 'John Doh', 33, new Date('2014-03-06T01:35:33Z'), 78.901, true]
]);

writer.write('D5', ['Total:', '=SUM(E2:E4)'], { font: { bold: true }, alignment: 'right' });

writer.save(function (err) {
    if (err) throw err;
    console.log('file saved');
});
