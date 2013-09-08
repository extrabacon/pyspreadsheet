var async = require('async'),
	expect = require('chai').expect,
	SpreadsheetReader = require('../lib/spreadsheet-reader'),
	SpreadsheetWriter = require('../lib/spreadsheet-writer');

SpreadsheetWriter.prototype.saveAndRead = function (path, callback) {
	this.save(path, function (err) {
		if (err) return callback(err);
		SpreadsheetReader.read(path, callback);
	});
};

describe('SpreadsheetWriter', function () {

	describe('ctor', function () {
		it('should only support xls and xlsx formats', function () {
			new SpreadsheetWriter({ format: 'xls' }).destroy();
			new SpreadsheetWriter({ format: 'xlsx' }).destroy();
			expect(function () {
				new SpreadsheetWriter({ format: 'something else' });
			}).to.throw(Error);
		});
	})

	describe('addSheet()', function () {
		it('should write a simple file with multiple sheets', function (done) {
			var file = 'test/output/multiple_sheets.xlsx';
			var writer = new SpreadsheetWriter();
			writer.addSheet('first');
			writer.addSheet('second');
			writer.addSheet('third');
			writer.saveAndRead(file, function (err, workbook) {
				if (err) throw err;
				expect(workbook.meta.sheets[0]).to.equal('first');
				expect(workbook.meta.sheets[1]).to.equal('second');
				expect(workbook.meta.sheets[2]).to.equal('third');
				done();
			});
		});
	});

	describe('addFormat()', function (done) {
		it('should write with a simple format', function (done) {
			var file = 'test/output/simple_format.xlsx';
			var writer = new SpreadsheetWriter();
			writer.addFormat('my format', { font: { bold: true, color: 'red' } });
			writer.write(0, 0, 'Hello World!', 'my format');
			writer.saveAndRead(file, function (err, workbook) {
				if (err) throw err;
				expect(workbook.sheets[0].rows[0][0].value).to.equal('Hello World!');
				done();
			});
		});
	});

	describe('write(cell)', function () {
		it('should write to the correct cell from its address', function (done) {
			var file = 'test/output/write1.xlsx';
			var writer = new SpreadsheetWriter();
			writer.write('A1', 1);
			writer.write('C1', 2);
			writer.write('B2', 3);
			writer.write('D4', 4);
			writer.saveAndRead(file, function (err, workbook) {
				if (err) throw err;
				expect(workbook.sheets[0].cell('A1').value).to.equal(1);
				expect(workbook.sheets[0].cell('C1').value).to.equal(2);
				expect(workbook.sheets[0].cell('B2').value).to.equal(3);
				expect(workbook.sheets[0].cell('D4').value).to.equal(4);
				done();
			});
		});
	});

	describe('write(row, col)', function () {
		it('should write to the correct cells from indexed positions', function (done) {
			var file = 'test/output/write2.xlsx';
			var writer = new SpreadsheetWriter();
			writer.write(0, 0, 1);
			writer.write(0, 2, 2);
			writer.write(1, 1, 3);
			writer.write(3, 3, 4);
			writer.saveAndRead(file, function (err, workbook) {
				if (err) throw err;
				expect(workbook.sheets[0].cell('A1').value).to.equal(1);
				expect(workbook.sheets[0].cell('C1').value).to.equal(2);
				expect(workbook.sheets[0].cell('B2').value).to.equal(3);
				expect(workbook.sheets[0].cell('D4').value).to.equal(4);
				done();
			});
		});
	});

	describe('write()', function () {
		it('should write supported data types', function (done) {
			var file = 'test/output/write3.xlsx';
			var writer = new SpreadsheetWriter();
			var number = 1934587.9812858,
				string = '<html>some html <body>tags</body></html>',
				date = new Date(),
				empty = null;

			writer.write(0, 0, number);
			writer.write(0, 1, string);
			writer.write(0, 2, date);
			writer.write(0, 3, true);

			writer.saveAndRead(file, function (err, workbook) {
				if (err) throw err;
				var row = workbook.sheets[0].rows[0];
				expect(row[0].value).to.equal(number);
				expect(row[1].value).to.equal(string);
				expect(row[2].value.getTime()).to.equal(new Date(
					date.getFullYear(), date.getMonth(), date.getDate(),
					date.getHours(), date.getMinutes(), date.getSeconds()).getTime()
				);
				//expect(row[3].value).to.equal(true);
				done();
			});
		});
	});

	describe('save()', function () {

		it('should save a hello world file', function (done) {
			var file = 'test/output/helloworld.xlsx';
			var writer = new SpreadsheetWriter();
			writer.write(0, 0, 'Hello World!');
			writer.saveAndRead(file, function (err, workbook) {
				if (err) throw err;
				expect(workbook.sheets[0].rows[0][0].value).to.equal('Hello World!');
				done();
			});
		});

	});

});
